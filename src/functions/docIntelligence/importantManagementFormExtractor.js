if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}

const axios = require('axios');
const { logMessage, handleError } = require('../utils');

const endpoint = process.env.EXTRACTOR_ENDPOINT;
const apiKey = process.env.EXTRACTOR_ENDPOINT_AZURE_API_KEY;
const modelId = process.env.EXTRACTOR_MODEL_ID;
const apiVersion = "2024-11-30";

/**
 * Extracts structured data from 重要管理 (Important Management) forms using Azure Document Intelligence
 * 
 * This function processes PDF forms containing 7-day important hygiene management records with the following structure:
 * - Document metadata (year, month, location)
 * - 5 menu items with descriptions
 * - Daily records for up to 7 days, each containing:
 *   - Date information
 *   - Status for each menu item (良/否/未選択/エラー)
 *   - Daily check status
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
 *     month: "03",                     // 2-digit month from form
 *     location: "府中",                // Store/location name from form
 *     yearMonth: "2025-03",            // Combined year-month for easy reference
 *     fileExtension: "pdf"             // Original file extension
 *   },
 *   menuItems: [                       // Array of 5 menu items
 *     {
 *       menuNumber: 1,                 // Menu sequence number (1-5)
 *       menuName: "重要管理項目1"       // Full menu description from form
 *     },
 *     // ... 4 more menu items
 *   ],
 *   dailyRecords: [                    // Array of daily records (up to 7 days)
 *     {
 *       day: 5,                        // Day of month as integer
 *       date: "2025-03-05",           // Full ISO date string
 *       Menu1Status: "否",            // Status for menu 1: "良"|"否"|"未選択"|"エラー"
 *       Menu2Status: "否",            // Status for menu 2
 *       Menu3Status: "否",            // Status for menu 3
 *       Menu4Status: "良",            // Status for menu 4
 *       Menu5Status: "良",            // Status for menu 5
 *       dailyCheckStatus: "良",       // Daily check checkbox: "良"|"未選択"|"エラー"
 *       comment: "クレームあり",       // Daily comment text or "not found"
 *       approverStatus: "良"          // Approver checkbox: "良"|"未選択"|"エラー"
 *     },
 *     // ... more daily records
 *   ],
 *   summary: {                         // Calculated summary statistics
 *     totalDays: 7,                    // Total number of daily records found
 *     recordedDays: 7,                 // Days with valid day numbers (> 0)
 *     daysWithComments: 5,             // Days that have comments (not "not found")
 *     approvedDays: 7,                 // Days where approver status is "良"
 *     dailyCheckCompletedDays: 7       // Days where daily check status is "良"
 *   }
 * }
 * 
 * Status Values Explanation:
 * - "良" (Good): The "Good" checkbox was selected for this menu/day
 * - "否" (No/Bad): The "NG" checkbox was selected for this menu/day  
 * - "未選択" (Unselected): Neither checkbox was selected
 * - "エラー" (Error): Field not found or both checkboxes selected (invalid state)
 * 
 * Processing Logic:
 * 1. Submits form to Azure Document Intelligence for analysis
 * 2. Polls for completion (up to 20 attempts)
 * 3. Extracts metadata (year, month, location) from form header
 * 4. Extracts menu item descriptions (menu1-menu5)
 * 5. For each day (1-7):
 *    - Extracts day number and builds full date
 *    - For each menu (1-5): determines status using getStatus()
 *    - Extracts daily check status, comment, and approver status
 * 6. Builds summary statistics
 * 7. Returns structured data optimized for report generation
 */
async function extractImportantManagementData(context, base64BinFile, fileExtension) {
  try {
    logMessage("📤 Submitting to custom extraction model for 重要管理...", context);

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

      // Extract menu item descriptions (menu1-menu5)
      const menuItems = [];
      menuItems[0] = fields.menu1?.valueString || "not found";
      menuItems[1] = fields.menu2?.valueString || "not found";
      menuItems[2] = fields.menu3?.valueString || "not found";
      menuItems[3] = fields.menu4?.valueString || "not found";
      menuItems[4] = fields.menu5?.valueString || "not found";

      logMessage('📊 Extracted Menu Items:', context);
      menuItems.forEach((item, index) => {
        logMessage(`  - Menu${index + 1}: ${item}`, context);
      });

      // Process daily records (up to 7 days)
      const dailyRecords = [];

      for (let day = 1; day <= 7; day++) {
        const dayKey = `day${day}`;
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

        // Extract menu statuses for this day using getStatus logic
        const menuStatuses = [];
        for (let menu = 1; menu <= 5; menu++) {
          const gKey = `d${day}c${menu}g`;    // Good checkbox field key
          const ngKey = `d${day}c${menu}ng`;  // NG checkbox field key
          const status = getStatus(fields[gKey], fields[ngKey]);
          menuStatuses.push(status);
          
          logMessage(`  - Menu${menu}: ${status}`, context);
        }

        // Extract daily check, comment, and approver status
        const dailyCheckStatus = getCheckStatus(fields[`d${day}dailyCheck`]);
        const comment = fields[`comment${day}`]?.valueString || "not found";
        const approverStatus = getCheckStatus(fields[`d${day}approver`]);

        logMessage(`  ✅ Daily Check: ${dailyCheckStatus}`, context);
        logMessage(`  💬 Comment: ${comment}`, context);
        logMessage(`  👤 Approver: ${approverStatus}`, context);

        // Build structured daily record object
        const dailyRecord = {
          day: parseInt(dayValue),
          date: fullDate,
          Menu1Status: menuStatuses[0],
          Menu2Status: menuStatuses[1],
          Menu3Status: menuStatuses[2],
          Menu4Status: menuStatuses[3],
          Menu5Status: menuStatuses[4],
          dailyCheckStatus: dailyCheckStatus,
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
        menuItems: menuItems.map((item, index) => ({
          menuNumber: index + 1,
          menuName: item
        })),
        dailyRecords: dailyRecords,
        summary: {
          totalDays: dailyRecords.length,
          recordedDays: dailyRecords.filter(record => record.day > 0).length,
          daysWithComments: dailyRecords.filter(record => record.comment !== "not found").length,
          approvedDays: dailyRecords.filter(record => record.approverStatus === "良").length,
          dailyCheckCompletedDays: dailyRecords.filter(record => record.dailyCheckStatus === "良").length
        }
      };

      logMessage(`📊 Extraction complete: ${dailyRecords.length} daily records processed`, context);
      logMessage(`📊 Summary: ${structuredData.summary.recordedDays} recorded days, ${structuredData.summary.daysWithComments} with comments`, context);
      logMessage(`📊 Summary: ${structuredData.summary.approvedDays} approved days, ${structuredData.summary.dailyCheckCompletedDays} daily checks completed`, context);
      
      return structuredData;

    } else {
      logMessage("⚠️ No fields extracted.", context);
      // Return empty structure if no data was extracted
      return {
        metadata: { year: "0000", month: "00", location: "エラー", yearMonth: "0000-00", fileExtension },
        menuItems: [],
        dailyRecords: [],
        summary: { totalDays: 0, recordedDays: 0, daysWithComments: 0, approvedDays: 0, dailyCheckCompletedDays: 0 }
      };
    }

  } catch (error) {
    handleError(error, 'extract', context);
    // Return empty structure on error
    return {
      metadata: { year: "0000", month: "00", location: "エラー", yearMonth: "0000-00", fileExtension },
      menuItems: [],
      dailyRecords: [],
      summary: { totalDays: 0, recordedDays: 0, daysWithComments: 0, approvedDays: 0, dailyCheckCompletedDays: 0 }
    };
  }
}

/**
 * Determines the status of a menu item based on Good/NG checkbox states
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
 * Determines the status of a single checkbox (used for daily check and approver fields)
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
  extractImportantManagementData
};
