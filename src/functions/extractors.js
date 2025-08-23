// ローカル開発環境用（Azure Functionsでは不要）
if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const axios = require('axios');
const FormData = require('form-data');

// Azure Document Intelligence 設定
const endpoint = process.env.EXTRACTOR_ENDPOINT;
const apiKey = process.env.AZURE_API_KEY;
const modelId = process.env.EXTRACTOR_MODEL_ID;
const apiVersion = "2024-11-30";

// Monday.com 設定
const MONDAY_API_TOKEN = process.env.MONDAY_API_KEY;
const BOARD_ID = 9857035666;
const MONDAY_API_VERSION = "2023-10";

// simple in-memory gate across invocations within this process
let lastMutationAt = 0;
const MUTATION_SPACING_MS = "2000";

function logMessage(message, context) {
    if (context && context.log) {
        context.log(message);
    } else {
        console.log(message);
    }
}

function handleError(error, phase, context) {
    if (context && context.log) {
        context.log.error(`[ERROR - ${phase}] ${error.message}`);
        if (error.response) {
            context.log.error(`[RESPONSE] ${JSON.stringify(error.response.data, null, 2)}`);
        }
        context.log.error(`[STACK] ${error.stack}`);
    } else {
        console.error(`[ERROR - ${phase}] ${error.message}`);
        if (error.response) {
            console.error(`[RESPONSE] ${JSON.stringify(error.response.data, null, 2)}`);
        }
        console.error(`[STACK] ${error.stack}`);
    }
}

async function extractImportantManagementData(context, base64BinFile, fileExtension) {
    try {
        logMessage("📤 Submitting to custom extraction model for 重要管理...", context);

        const response = await axios.post(
            `${endpoint}/documentintelligence/documentModels/${modelId}:analyze?api-version=${apiVersion}`,
            { base64Source: base64BinFile },
            {
                headers: {
                    'Ocp-Apim-Subscription-Key': apiKey,
                    'Content-Type': 'application/json'
                }
            }
        );

        const operationLocation = response.headers['operation-location'];
        logMessage(`📍 Extraction operation location: ${operationLocation}`, context);

        let result;
        let attempts = 0;
        const maxAttempts = 20;
        const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

        while (attempts < maxAttempts) {
            await delay(1000);
            const pollResponse = await axios.get(operationLocation, {
                headers: { 'Ocp-Apim-Subscription-Key': apiKey }
            });
            result = pollResponse.data;
            logMessage(`🔁 Extraction attempt ${attempts + 1}: ${result.status}`, context);
            if (result.status === "succeeded") break;
            attempts++;
        }

        const fields = result?.analyzeResult?.documents?.[0]?.fields;

        if (fields) {
            logMessage("🧾 Extracted fields:", context);
            console.log("📦 Extracted field keys:", Object.keys(fields));
            console.log("📦 Full fields object:", JSON.stringify(fields, null, 2));

            const rawYear = fields.year?.valueString || "0000";
            const rawMonth = fields.month?.valueString || "00";
            const location = fields.location?.valueString || "エラー";

            const year = rawYear && /^\d{4}$/.test(rawYear) ? rawYear : "0000";
            const month = rawMonth && /^\d{1,2}$/.test(rawMonth) ? rawMonth.padStart(2, '0') : "00";

            for (let day = 1; day <= 7; day++) {
                const dayKey = `day${day}`;
                const dayField = fields[dayKey];
                if (!dayField) {
                    logMessage(`⚠️ Missing field: ${dayKey}`, context);
                    continue;
                }

                const rawDay = dayField.valueString;
                const dayValue = rawDay && /^\d{1,2}$/.test(rawDay) ? rawDay.padStart(2, '0') : "00";

                const filename = `${location}-${year}-${month}-${dayValue}`;
                logMessage(`📄 Row name (unique ID): ${filename}`, context);

                for (let category = 1; category <= 5; category++) {
                    const gKey = `d${day}c${category}g`;
                    const ngKey = `d${day}c${category}ng`;
                    const g = fields[gKey]?.valueSelectionMark;
                    const ng = fields[ngKey]?.valueSelectionMark;

                    if (!g && !ng) {
                        logMessage(`⚠️ Missing category fields: ${gKey}, ${ngKey}`, context);
                    } else {
                        const status = getStatus(fields[gKey], fields[ngKey]);
                        logMessage(`  - ${gKey}: ${g || "not found"}, ${ngKey}: ${ng || "not found"} → ${status}`, context);
                    }
                }

                const dailyCheckStatus = getCheckStatus(fields[`d${day}dailyCheck`]);
                const approverStatus = getCheckStatus(fields[`d${day}approver`]);
                const comment = fields[`comment${day}`]?.valueString || "not found";

                logMessage(`  ✅ Daily Check: ${dailyCheckStatus}`, context);
                logMessage(`  💬 Comment: ${comment}`, context);
                logMessage(`  👤 Approver: ${approverStatus}`, context);

                const row = {
                    name: filename,
                    date4: `${year}-${month}-${dayValue}`,
                    text_mkv0z6d: location,
                    color_mkv02tqg: getStatus(fields[`d${day}c1g`], fields[`d${day}c1ng`]),
                    color_mkv0yb6g: getStatus(fields[`d${day}c2g`], fields[`d${day}c2ng`]),
                    color_mkv06e9z: getStatus(fields[`d${day}c3g`], fields[`d${day}c3ng`]),
                    color_mkv0x9mr: getStatus(fields[`d${day}c4g`], fields[`d${day}c4ng`]),
                    color_mkv0df43: getStatus(fields[`d${day}c5g`], fields[`d${day}c5ng`]),
                    color_mkv0ej57: dailyCheckStatus,
                    text_mkv0etfg: comment,
                    color_mkv0xnn4: approverStatus 
                };

                logMessage(`📤 Ready to upload row for Day ${day}:`, context);
                console.log("📤 Row data:", JSON.stringify(row, null, 2));

                await uploadToMonday(row, context, base64BinFile, `${filename}.${fileExtension}`);
            }
        } else {
            logMessage("⚠️ No fields extracted.", context);
            logMessage(`📎 Raw result: ${JSON.stringify(result, null, 2)}`, context);
        }

    } catch (error) {
        handleError(error, 'extract', context);
    }
}

function getStatus(gField, ngField) {
    const g = gField?.valueSelectionMark || "not found";
    const ng = ngField?.valueSelectionMark || "not found";

    if (g === "not found" && ng === "not found") return "エラー";
    if (g === "selected" && ng === "selected") return "エラー";
    if (g === "unselected" && ng === "unselected") return "未選択";
    //if (g == "not found" && ng == "selected") return "否"；
    //if (g == "not found" && ng == "unselected") return "良";
    if (g === "selected") return "良";
    if (ng === "selected") return "否";
    return "未選択";
}

function getCheckStatus(field) {
    const value = field?.valueSelectionMark || "not found";
    if (value === "selected") return "良";
    if (value === "not found") return "エラー";
    return "未選択";
}

async function gqlRequest(query, variables, context, opts = {}) {
  const {
    maxRetries = 5,
    baseDelayMs = 500,  // used when retry_in_seconds is missing
    apiVersion = '2023-10',
  } = opts;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const res = await axios.post(
        'https://api.monday.com/v2',
        { query, variables },
        {
          headers: {
            Authorization: `Bearer ${MONDAY_API_TOKEN}`,
            'Content-Type': 'application/json',
            'API-Version': apiVersion,
          },
          timeout: 20000,
          validateStatus: () => true, // handle 429 here
        }
      );

      // Handle HTTP 429 or a body that signals ComplexityException
      const status = res.status;
      const body = res.data;

      // Helper to sleep
      const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

      const isComplexity429 =
        status === 429 ||
        body?.error_code === 'ComplexityException' ||
        (Array.isArray(body?.errors) &&
         body.errors.some((m) => typeof m === 'string' && m.includes('Complexity budget exhausted')));

      if (isComplexity429) {
        // Respect retry hint if present
        let waitSec =
          body?.error_data?.retry_in_seconds ??
          // fallback: parse "... reset in 31 seconds"
          (() => {
            const m = (JSON.stringify(body) || '').match(/reset in (\d+) seconds/i);
            return m ? parseInt(m[1], 10) : null;
          })();

        if (waitSec == null) {
          // exponential backoff + jitter
          waitSec = Math.min(60, Math.pow(2, attempt)) + Math.random();
        }

        context.log(
          `⏳ 429/Complexity throttled. Waiting ${waitSec}s before retry (attempt ${attempt + 1}/${maxRetries}).`
        );
        await sleep(waitSec * 1000);
        continue; // retry
      }

      // Handle GraphQL errors (200 with errors array)
      if (body?.errors?.length) {
        context.log.error('🧩 GraphQL errors:', JSON.stringify(body.errors, null, 2));
        throw new Error(body.errors.map((e) => e.message || e).join('; '));
      }

      if (!body || !body.data) {
        context.log.error('❗Unexpected GraphQL response:', JSON.stringify(body, null, 2));
        throw new Error('Unexpected GraphQL response (no data).');
      }

      return body.data; // success
    } catch (err) {
      if (attempt === maxRetries) {
        throw err;
      }
      const delay = baseDelayMs * Math.pow(2, attempt) + Math.floor(Math.random() * 200);
      context.log(`♻️ Transient error. Backing off ${delay}ms. Attempt ${attempt + 1}/${maxRetries}`);
      await new Promise((r) => setTimeout(r, delay));
    }
  }
  throw new Error('gqlRequest: exhausted retries.');
}

// Helper: remove only undefined (keep null/"" if you really want to blank a field)
function stripUndefined(obj) {
  return Object.fromEntries(Object.entries(obj).filter(([, v]) => v !== undefined));
}

async function uploadToMonday(rowData, context, base64BinFile, originalFileName) {
  context.log("🚀 uploadToMonday() called");

  const itemName = (rowData.name || '').trim();
  const targetDate = formatDate(rowData.date4); // must be "YYYY-MM-DD"
  const targetLocation = (rowData.text_mkv0z6d || '').trim();

  // 1) Exact server-side lookup (date + location)
  const lookupQuery = `
    query ($boardId: ID!, $date: String!, $loc: String!) {
      items_page_by_column_values(
        board_id: $boardId,
        limit: 2,
        columns: [
          { column_id: "date4", column_values: [$date] },
          { column_id: "text_mkv0z6d", column_values: [$loc] }
        ]
      ) {
        items { id name }
      }
    }
  `;

  let existingItemId = null;

  await throttleMutation();
  const lookupData = await gqlRequest(
    lookupQuery,
    { boardId: String(BOARD_ID), date: targetDate, loc: targetLocation },
    context
  );
  const found = lookupData.items_page_by_column_values?.items || [];
  if (found.length > 0) {
    existingItemId = found[0].id;
    context.log(`✏️ Updating existing item with date "${targetDate}" and location "${targetLocation}" (ID: ${existingItemId})`);
  } else {
    context.log("🆕 No matching item found. Will create new.");
  }

  // 2) Prepare column values (date uses canonical JSON structure)
  const columnValues = stripUndefined({
    date4: { date: targetDate },          // <- important format for Date column
    text_mkv0z6d: targetLocation,
    color_mkv02tqg: rowData.color_mkv02tqg,
    color_mkv0yb6g: rowData.color_mkv0yb6g,
    color_mkv06e9z: rowData.color_mkv06e9z,
    color_mkv0x9mr: rowData.color_mkv0x9mr,
    color_mkv0df43: rowData.color_mkv0df43,
    color_mkv0ej57: rowData.color_mkv0ej57,
    color_mkv0xnn4: rowData.color_mkv0xnn4,
    text_mkv0etfg: rowData.text_mkv0etfg
  });

  let itemId;

  if (existingItemId) {
    // 2a) Update existing

    const updateMutation = `
    mutation UpdateItem($boardId: ID!, $itemId: ID!, $cols: JSON!) {
        change_multiple_column_values(
        board_id: $boardId,
        item_id: $itemId,
        column_values: $cols
        ) { id }
    }
    `;

    await throttleMutation();
    const updateData = await gqlRequest(
      updateMutation,
      {
        boardId: String(BOARD_ID),
        itemId: String(existingItemId),            // pass as ID (string)
        cols: JSON.stringify(columnValues)         // JSON! expects string
      },
      context
    );
    itemId = updateData.change_multiple_column_values.id;  
    context.log(`✅ Updated item ID: ${itemId}`);
  } else {
    // (Optional) tiny randomized backoff to reduce races, then re-check
    await new Promise(r => setTimeout(r, Math.floor(Math.random() * 80)));


    await throttleMutation();
    const reLookup = await gqlRequest(
      lookupQuery,
      { boardId: String(BOARD_ID), date: targetDate, loc: targetLocation },
      context
    );
    const again = reLookup.items_page_by_column_values?.items || [];
    if (again.length > 0) {
      itemId = again[0].id;
      context.log(`⚠️ Race avoided. Using existing item ID: ${itemId}`);
    } else {
      // 2b) Create new
      const createMutation = `
        mutation CreateItem($boardId: ID!, $name: String!, $cols: JSON!) {
          create_item(board_id: $boardId, item_name: $name, column_values: $cols) {
            id
          }
        }
      `;

      await throttleMutation();
      const createData = await gqlRequest(
        createMutation,
        {
          boardId: String(BOARD_ID),
          name: itemName,
          cols: JSON.stringify(columnValues)
        },
        context
      );
      itemId = createData.create_item.id;
      context.log(`✅ Created new item ID: ${itemId}`);
    }
  }

  // 3) Upload file to the file column
  const fileBuffer = Buffer.from(base64BinFile, 'base64');
  const form = new FormData();
  form.append('query', `
    mutation ($file: File!) {
      add_file_to_column (file: $file, item_id: ${itemId}, column_id: "file_mkv1kpsc") {
        id
      }
    }
  `);
  form.append('variables[file]', fileBuffer, { filename: originalFileName });

  const fileUploadResponse = await axios.post(
    'https://api.monday.com/v2/file',
    form,
    {
      headers: {
        Authorization: `Bearer ${MONDAY_API_TOKEN}`,
        ...form.getHeaders(),
        'API-Version': '2023-10',
      },
      timeout: 30000
    }
  );
  context.log("📎 File upload success:");
  context.log(JSON.stringify(fileUploadResponse.data, null, 2));
}

function formatDate(dateStr) {
    const [year, month, day] = dateStr.split('-');
    return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
}

async function throttleMutation() {
  const now = Date.now();
  const wait = Math.max(0, lastMutationAt + MUTATION_SPACING_MS - now);
  if (wait > 0) await new Promise((r) => setTimeout(r, wait));
  lastMutationAt = Date.now();
}

module.exports = {
    extractImportantManagementData
};