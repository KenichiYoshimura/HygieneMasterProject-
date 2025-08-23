const axios = require('axios');
const FormData = require('form-data');
const { stripUndefined, formatDate, throttleMutation, gqlRequest } = require('../utils');

const MONDAY_API_TOKEN = process.env.MONDAY_API_KEY;
const BOARD_ID = 9857035666;

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

module.exports = {
  uploadToMonday
};
