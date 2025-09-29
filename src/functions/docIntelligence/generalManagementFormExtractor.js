if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}

const axios = require('axios');
const { logMessage, handleError } = require('../utils');

const endpoint = process.env.GENERAL_MANAGEMENT_EXTRACTOR_ENDPOINT;
const apiKey = process.env.GENERAL_MANAGEMENT_EXTRACTOR_ENDPOINT_AZURE_API_KEY;
const modelId = process.env.GENERAL_MANAGEMENT_EXTRACTOR_MODEL_ID;
const apiVersion = "2024-11-30";

/**
 * Extracts structured data from 一般管理 (General Management) forms using Azure Document Intelligence
 * 
 * This function processes PDF forms containing 7-day hygiene management records with the following structure:
 * - Document metadata (year, month, location)
 * - 7 management categories with descriptions
 * - Daily records for up to 7 days, each containing:
 *   - Date information
 * - Status for each category (良/否/未選択/エラー)
 *   - Daily comments
 *   - Approver status
 * 
 * @param {Object} context - Azure Functions execution context for logging
 * @param {string} base64BinFile - Base64 encoded PDF file content
 * @param {string} fileExtension - Original file extension (e.g., "pdf")
 * 
 * @returns {Object} Structured data object with the following schema:
 * {
 *   metadata: {
 *     year: "2025",                    // 4-digit year from form
 *     month: "05",                     // 2-digit month from form
 *     location: "中目黒",              // Store/location name from form
 *     yearMonth: "2025-05",            // Combined year-month for easy reference
 *     fileExtension: "pdf"             // Original file extension
 *   },
 *   categories: [                      // Array of 7 management categories
 *     {
 *       categoryNumber: 1,             // Category sequence number (1-7)
 *       categoryName: "1 原材料の受入の 確認"  // Full category description from form
 *     },
 *     // ... 6 more categories
 *   ],
 *   dailyRecords: [                    // Array of daily records (up to 7 days)
 *     {
 *       day: 10,                       // Day of month as integer
 *       date: "2025-05-10",           // Full ISO date string
 *       Cat1Status: "良",             // Status for category 1: "良"|"否"|"未選択"|"エラー"
 *       Cat2Status: "良",             // Status for category 2
 *       Cat3Status: "良",             // Status for category 3
 *       Cat4Status: "良",             // Status for category 4
 *       Cat5Status: "良",             // Status for category 5
 *       Cat6Status: "良",             // Status for category 6
 *       Cat7Status: "良",             // Status for category 7
 *       comment: "問題なく運営",       // Daily comment text or "not found"
 *       approverStatus: "未選択"       // Approver checkbox: "良"|"未選択"|"エラー"
 *     },
 *     // ... more daily records
 *   ],
 *   summary: {                         // Calculated summary statistics
 *     totalDays: 7,                    // Total number of daily records found
 *     recordedDays: 7,                 // Days with valid day numbers (> 0)
 *     daysWithComments: 5,             // Days that have comments (not "not found")
 *     approvedDays: 0                  // Days where approver status is "良"
 *   }
 * }
 * 
 * Status Values Explanation:
 * - "良" (Good): The "Good" checkbox was selected for this category/day
 * - "否" (No/Bad): The "NG" checkbox was selected for this category/day  
 * - "未選択" (Unselected): Neither checkbox was selected
 * - "エラー" (Error): Field not found or both checkboxes selected (invalid state)
 * 
 * Processing Logic:
 * 1. Submits form to Azure Document Intelligence for analysis
 * 2. Polls for completion (up to 20 attempts)
 * 3. Extracts metadata (year, month, location) from form header
 * 4. Extracts category descriptions (Cat1-Cat7)
 * 5. For each day (1-7):
 *    - Extracts day number and builds full date
 *    - For each category (1-7): determines status using getStatus()
 *    - Extracts daily comment and approver status
 * 6. Builds summary statistics
 * 7. Returns structured data optimized for report generation
 */
async function extractGeneralManagementData(context, base64BinFile, fileExtension) {
  try {
    logMessage("📤 Submitting to custom extraction model for 一般管理...", context);

    // Submit document to Azure Document Intelligence for analysis
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

    // Poll for completion of the analysis operation
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

      // Extract and validate metadata from form header
      const rawYear = fields.year?.valueString || "0000";
      const rawMonth = fields.month?.valueString || "00";
      const location = fields.location?.valueString || "エラー";

      const year = rawYear && /^\d{4}$/.test(rawYear) ? rawYear : "0000";
      const month = rawMonth && /^\d{1,2}$/.test(rawMonth) ? rawMonth.padStart(2, '0') : "00";

      // Extract category descriptions (Cat1-Cat7)
      const categories = [];
      categories[0] = fields.Cat1?.valueString || "not found";
      categories[1] = fields.Cat2?.valueString || "not found";
      categories[2] = fields.Cat3?.valueString || "not found";
      categories[3] = fields.Cat4?.valueString || "not found";
      categories[4] = fields.Cat5?.valueString || "not found";
      categories[5] = fields.Cat6?.valueString || "not found";
      categories[6] = fields.Cat7?.valueString || "not found";

      logMessage('📊 Extracted Categories:', context);
      categories.forEach((cat, index) => {
        logMessage(`  - Cat${index + 1}: ${cat}`, context);
      });

      // Process daily records (up to 7 days)
      const dailyRecords = [];

      for (let day = 1; day <= 7; day++) {
        const dayKey = `Day${day}`;
        const dayField = fields[dayKey];
        
        if (!dayField) {
          logMessage(`⚠️ Missing field: ${dayKey}`, context);
          continue;
        }

        // Extract and validate day number, build full date
        const rawDay = dayField.valueString;
        const dayValue = rawDay && /^\d{1,2}$/.test(rawDay) ? rawDay.padStart(2, '0') : "00";
        const fullDate = `${year}-${month}-${dayValue}`;

        logMessage(`📅 Processing Day ${day}: ${fullDate}`, context);

        // Extract category statuses for this day using getStatus logic
        const categoryStatuses = [];
        for (let category = 1; category <= 7; category++) {
          const gKey = `C${category}D${day}G`;    // Good checkbox field key
          const ngKey = `C${category}D${day}NG`;  // NG checkbox field key
          const status = getStatus(fields[gKey], fields[ngKey]);
          categoryStatuses.push(status);
          
          logMessage(`  - Cat${category}: ${status}`, context);
        }

        // Extract comment and approver status
        const comment = fields[`D${day}comment`]?.valueString || "not found";
        const approverStatus = getCheckStatus(fields[`D${day}Approver`]);

        logMessage(`  💬 Comment: ${comment}`, context);
        logMessage(`  👤 Approver: ${approverStatus}`, context);

        // Build structured daily record object
        const dailyRecord = {
          day: parseInt(dayValue),
          date: fullDate,
          Cat1Status: categoryStatuses[0],
          Cat2Status: categoryStatuses[1],
          Cat3Status: categoryStatuses[2],
          Cat4Status: categoryStatuses[3],
          Cat5Status: categoryStatuses[4],
          Cat6Status: categoryStatuses[5],
          Cat7Status: categoryStatuses[6],
          comment: comment,
          approverStatus: approverStatus
        };

        logMessage(`✅ Daily record for Day ${day} complete`, context);
        dailyRecords.push(dailyRecord);
      }

      // Build the complete structured data object
      const structuredData = {
        metadata: {
          year: year,
          month: month,
          location: location,
          yearMonth: `${year}-${month}`,
          fileExtension: fileExtension
        },
        categories: categories.map((cat, index) => ({
          categoryNumber: index + 1,
          categoryName: cat
        })),
        dailyRecords: dailyRecords,
        summary: {
          totalDays: dailyRecords.length,
          recordedDays: dailyRecords.filter(record => record.day > 0).length,
          daysWithComments: dailyRecords.filter(record => record.comment !== "not found").length,
          approvedDays: dailyRecords.filter(record => record.approverStatus === "良").length
        }
      };

      logMessage(`📊 Extraction complete: ${dailyRecords.length} daily records processed`, context);
      logMessage(`📊 Summary: ${structuredData.summary.recordedDays} recorded days, ${structuredData.summary.daysWithComments} with comments`, context);
      
      return structuredData;

    } else {
      logMessage("⚠️ No fields extracted.", context);
      // Return empty structure if no data was extracted
      return {
        metadata: { year: "0000", month: "00", location: "エラー", yearMonth: "0000-00", fileExtension },
        categories: [],
        dailyRecords: [],
        summary: { totalDays: 0, recordedDays: 0, daysWithComments: 0, approvedDays: 0 }
      };
    }

  } catch (error) {
    handleError(error, 'extract', context);
    // Return empty structure on error
    return {
      metadata: { year: "0000", month: "00", location: "エラー", yearMonth: "0000-00", fileExtension },
      categories: [],
      dailyRecords: [],
      summary: { totalDays: 0, recordedDays: 0, daysWithComments: 0, approvedDays: 0 }
    };
  }
}

/**
 * Determines the status of a category based on Good/NG checkbox states
 * 
 * @param {Object} gField - Good checkbox field from Azure Document Intelligence
 * @param {Object} ngField - NG checkbox field from Azure Document Intelligence
 * @returns {string} Status: "良" (good), "否" (no/bad), "未選択" (unselected), or "エラー" (error)
 */
function getStatus(gField, ngField) {
  const g = gField?.valueSelectionMark || "not found";
  const ng = ngField?.valueSelectionMark || "not found";

  if (g === "not found" && ng === "not found") return "エラー";    // Both fields missing
  if (g === "selected" && ng === "selected") return "エラー";      // Both checkboxes selected (invalid)
  if (g === "unselected" && ng === "unselected") return "未選択";  // Neither checkbox selected
  if (g === "selected") return "良";   // Good checkbox selected
  if (ng === "selected") return "否";  // NG checkbox selected
  return "未選択";                     // Default case
}

/**
 * Determines the status of a single checkbox (used for approver field)
 * 
 * @param {Object} field - Checkbox field from Azure Document Intelligence
 * @returns {string} Status: "良" (checked), "未選択" (unchecked), or "エラー" (error)
 */
function getCheckStatus(field) {
  const value = field?.valueSelectionMark || "not found";
  if (value === "selected") return "良";     // Checkbox is checked
  if (value === "not found") return "エラー";  // Field not found
  return "未選択";                           // Checkbox is unchecked
}

module.exports = {
  extractGeneralManagementData
};
