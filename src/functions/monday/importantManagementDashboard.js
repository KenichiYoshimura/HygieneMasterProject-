
// ローカル開発環境用（Azure Functionsでは不要）
if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}
const { logMessage, handleError, convertHeicToJpegIfNeeded } = require('../utils');
const axios = require('axios');
const FormData = require('form-data');

const MONDAY_API_TOKEN = process.env.MONDAY_API_KEY;
const BOARD_ID = 9857035666;

// simple in-memory gate across invocations within this process
let lastMutationAt = 0;
const MUTATION_SPACING_MS = 2000;


async function uploadToMonday(rowData, context, base64BinFile, originalFileName) {
  logMessage("🚀 uploadToMonday() called", context);

  const itemName = (rowData.name || '').trim();
  const targetDate = formatDate(rowData.date4); // must be "YYYY-MM-DD"
  const targetLocation = (rowData.text_mkv0z6d || '').trim();

  logMessage("itemName, targetdate and TargetLocation retrieved", context);

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

  logMessage("just about to lookup", context);

  await throttleMutation();
  const lookupData = await gqlRequest(
    lookupQuery,
    { boardId: String(BOARD_ID), date: targetDate, loc: targetLocation },
    context
  );

  logMessage("finished lookup", context);

  const found = lookupData.items_page_by_column_values?.items || [];
  if (found.length > 0) {
    existingItemId = found[0].id;
    logMessage(`✏️ Updating existing item with date "${targetDate}" and location "${targetLocation}" (ID: ${existingItemId})`, context);
  } else {
    logMessage("🆕 No matching item found. Will create new.", context);
  }

  logMessage("Just about to upload the data", context);

  // 2) Prepare column values (date uses canonical JSON structure)
  const columnValues = stripUndefined({
    date4: { date: targetDate },
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
    logMessage("Updating the existing record", context);
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
        itemId: String(existingItemId),
        cols: JSON.stringify(columnValues)
      },
      context
    );
    itemId = updateData.change_multiple_column_values.id;
    context.log(`✅ Updated item ID: ${itemId}`);
  } else {
    logMessage("Creating new record", context);
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

  logMessage("Just about to upload the file for record", context);

  // 3) Upload file to the file column
  let fileBuffer = Buffer.from(base64BinFile, 'base64');
  let fileNameToUpload = originalFileName;

  // Convert HEIC to JPEG if needed
  const converted = await convertHeicToJpegIfNeeded(fileBuffer, originalFileName, context);
  fileBuffer = converted.buffer;
  fileNameToUpload = converted.filename;

  const form = new FormData();
  form.append('query', `
    mutation ($file: File!) {
      add_file_to_column (file: $file, item_id: ${itemId}, column_id: "file_mkv1kpsc") {
        id
      }
    }
  `);
  form.append('variables[file]', fileBuffer, { filename: fileNameToUpload });

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

module.exports = {
  uploadToMonday
};
