if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}

const axios = require('axios');
const { logMessage, handleError } = require('../utils');

const endpoint = process.env.GENERAL_MANAGEMENT_EXTRACTOR_ENDPOINT;
const apiKey = process.env.GENERAL_MANAGEMENT_EXTRACTOR_ENDPOINT_AZURE_API_KEY;
const modelId = process.env.GENERAL_MANAGEMENT_EXTRACTOR_MODEL_ID;
const apiVersion = "2024-11-30";

async function extractGeneralManagementData(context, base64BinFile, fileExtension) {
  try {
    logMessage("📤 Submitting to custom extraction model for 一般管理...", context);

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

    logMessage("Checking the extracted data next!!!", context);
    const fields = result?.analyzeResult?.documents?.[0]?.fields;
    logMessage("Got fields", context);

    if (fields) {
      logMessage("🧾 Extracted fields:", context);
      console.log("📦 Extracted field keys:", Object.keys(fields));
      console.log("📦 Full fields object:", JSON.stringify(fields, null, 2));

      const rawYear = fields.year?.valueString || "0000";
      const rawMonth = fields.month?.valueString || "00";
      const location = fields.location?.valueString || "エラー";

      const year = rawYear && /^\d{4}$/.test(rawYear) ? rawYear : "0000";
      const month = rawMonth && /^\d{1,2}$/.test(rawMonth) ? rawMonth.padStart(2, '0') : "00";

      const categories = [];
      categories[0] = fields.Cat1 || "not found";
      categories[1] = fields.Cat2 || "not found";
      categories[2] = fields.Cat3 || "not found";
      categories[3] = fields.Cat4 || "not found";
      categories[4] = fields.Cat5 || "not found";
      categories[5] = fields.Cat6 || "not found";
      categories[6] = fields.Cat7 || "not found";

      logMessage('Categories:', context);
      categories.forEach((cat, index) => {
        logMessage(`  - Cat${index + 1}: ${cat.valueString || "not found"}`, context);
      });

      const extractedRows = [];

      for (let day = 1; day <= 7; day++) {
        const dayKey = `Day${day}`;
        const dayField = fields[dayKey];
        if (!dayField) {
          logMessage(`⚠️ Missing field: ${dayKey}`, context);
          continue;
        }

        const rawDay = dayField.valueString;
        const dayValue = rawDay && /^\d{1,2}$/.test(rawDay) ? rawDay.padStart(2, '0') : "00";

        const filename = `${location}-${year}-${month}-${dayValue}`;
        logMessage(`📄 Row name (unique ID): ${filename}`, context);

        for (let category = 1; category <= 7; category++) {
          const gKey = `C${category}D${day}G`;
          const ngKey = `C${category}D${day}NG`;
          const g = fields[gKey]?.valueSelectionMark;
          const ng = fields[ngKey]?.valueSelectionMark;

          if (!g && !ng) {
            logMessage(`⚠️ Missing category fields: ${gKey}, ${ngKey}`, context);
          } else {
            const status = getStatus(fields[gKey], fields[ngKey]);
            logMessage(`  - ${gKey}: ${g || "not found"}, ${ngKey}: ${ng || "not found"} → ${status}`, context);
          }
        }

        const approverStatus = getCheckStatus(fields[`D${day}Approver`]);
        const comment = fields[`D${day}comment`]?.valueString || "not found";

        logMessage(`  💬 Comment: ${comment}`, context);
        logMessage(`  👤 Approver: ${approverStatus}`, context);

        /*
          Monday General Management Form Columns:
          ID: name, Title: Name, Type: name
          ID: date4, Title: 日付, Type: date
          ID: text_mkv0z6d, Title: 店舗, Type: text
          ID: color_mkv02tqg, Title: Category1, Type: status
          ID: color_mkv0yb6g, Title: Category2, Type: status
          ID: color_mkv06e9z, Title: Category3, Type: status
          ID: color_mkv0x9mr, Title: Category4, Type: status
          ID: color_mkv0df43, Title: Category5, Type: status
          ID: color_mkv5fa8m, Title: Category6, Type: status
          ID: color_mkv59ent, Title: Category7, Type: status
          ID: text_mkv0etfg, Title: 特記事項, Type: text
          ID: color_mkv0xnn4, Title: 確認者, Type: status
          ID: file_mkv1kpsc, Title: 紙の帳票, Type: file
        */
        const row = {
          name: filename,
          date4: `${year}-${month}-${dayValue}`,
          text_mkv0z6d: location,
          color_mkv02tqg: getStatus(fields[`C1D${day}G`], fields[`C1D${day}NG`]),
          color_mkv0yb6g: getStatus(fields[`C2D${day}G`], fields[`C2D${day}NG`]),
          color_mkv06e9z: getStatus(fields[`C3D${day}G`], fields[`C3D${day}NG`]),
          color_mkv0x9mr: getStatus(fields[`C4D${day}G`], fields[`C4D${day}NG`]),
          color_mkv0df43: getStatus(fields[`C5D${day}G`], fields[`C5D${day}NG`]),
          color_mkv5fa8m: getStatus(fields[`C6D${day}G`], fields[`C6D${day}NG`]),
          color_mkv59ent: getStatus(fields[`C7D${day}G`], fields[`C7D${day}NG`]),
          text_mkv0etfg: comment,
          color_mkv0xnn4: approverStatus 
        };

        logMessage(`📤 Ready to upload row for Day ${day}:`, context);
        console.log("📤 Row data:", JSON.stringify(row, null, 2));

        extractedRows.push({
          row,
          fileName: `${filename}.${fileExtension}`
        });
      }

      return extractedRows;
    } else {
      logMessage("⚠️ No fields extracted.", context);
      logMessage(`📎 Raw result: ${JSON.stringify(result, null, 2)}`, context);
      return [];
    }

  } catch (error) {
    handleError(error, 'extract', context);
    return [];
  }
}

function getStatus(gField, ngField) {
  const g = gField?.valueSelectionMark || "not found";
  const ng = ngField?.valueSelectionMark || "not found";

  if (g === "not found" && ng === "not found") return "エラー";
  if (g === "selected" && ng === "selected") return "エラー";
  if (g === "unselected" && ng === "unselected") return "未選択";
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

module.exports = {
  extractGeneralManagementData
};
