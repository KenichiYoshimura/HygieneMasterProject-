if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}

const axios = require('axios');
const { logMessage, handleError } = require('./utils');

const endpoint = process.env.EXTRACTOR_ENDPOINT;
const apiKey = process.env.EXTRACTOR_ENDPOINT_AZURE_API_KEY;
const modelId = process.env.EXTRACTOR_MODEL_ID;
const apiVersion = "2024-11-30";

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

      const extractedRows = [];

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
  extractImportantManagementData
};
