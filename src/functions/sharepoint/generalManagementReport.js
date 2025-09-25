const { logMessage, handleError, convertHeicToJpegIfNeeded} = require('../utils');
const { 
    uploadJsonToSharePoint, 
    uploadTextToSharePoint,
    uploadOriginalDocumentToSharePoint, 
    ensureSharePointFolder 
} = require('./sendToSharePoint');
const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');
const path = require('path');

async function prepareGeneralManagementReport(extractedRows, categories, context, base64BinFile, originalFileName) {
    logMessage("🚀 prepareGeneralManagementReport() called", context);
    
    try {
        // DEBUG: Print the exact structure we're receiving
        logMessage("🔍 DEBUG: Raw input analysis...", context);
        logMessage(`📊 extractedRows type: ${typeof extractedRows}`, context);
        logMessage(`📊 extractedRows length: ${Array.isArray(extractedRows) ? extractedRows.length : 'not array'}`, context);
        logMessage(`📊 extractedRows content:`, context);
        logMessage(`${JSON.stringify(extractedRows, null, 2)}`, context);
        
        logMessage(`📊 categories:`, context);
        logMessage(`${JSON.stringify(categories, null, 2)}`, context);
        
        logMessage(`📊 originalFileName: ${originalFileName}`, context);
        
        // Based on generalManagementDashboard.js, it expects a single rowData object
        // But extractedRows might be an array. Let's handle both cases:
        let rowDataArray = [];
        
        if (Array.isArray(extractedRows)) {
            // If it's an array, extract the .row property from each item
            rowDataArray = extractedRows.map(item => {
                if (item && item.row) {
                    return item.row;
                } else {
                    return item; // fallback if no .row property
                }
            });
        } else if (extractedRows && typeof extractedRows === 'object') {
            // If it's a single object, wrap it in an array
            rowDataArray = [extractedRows.row || extractedRows];
        } else {
            logMessage("❌ ERROR: extractedRows is neither array nor object", context);
            throw new Error("Invalid extractedRows format");
        }
        
        logMessage(`📊 Processed rowDataArray length: ${rowDataArray.length}`, context);
        if (rowDataArray.length > 0) {
            logMessage(`📊 First processed row:`, context);
            logMessage(`${JSON.stringify(rowDataArray[0], null, 2)}`, context);
            logMessage(`📊 Available keys: ${Object.keys(rowDataArray[0]).join(', ')}`, context);
        }
        
        // Generate structured JSON data
        const jsonReport = generateJsonReport(rowDataArray, categories, originalFileName, context);
        logMessage("✅ JSON report generated", context);
        
        // Generate text report - NOW PASSING CONTEXT
        const textReport = generateTextReport(rowDataArray, categories, originalFileName, context);
        logMessage("✅ Text report generated", context);
        
        // Upload to SharePoint
        logMessage("📤 Starting SharePoint upload...", context);
        await uploadReportsToSharePoint(jsonReport, textReport, base64BinFile, originalFileName, rowDataArray, context);
        logMessage("✅ SharePoint upload completed", context);
        
        return {
            json: jsonReport,
            text: textReport
        };
        
    } catch (error) {
        handleError(error, 'General Management Report Generation', context);
        throw error;
    }
}

async function uploadReportsToSharePoint(jsonReport, textReport, base64BinFile, originalFileName, extractedRows, context) {
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const baseFileName = originalFileName.replace(/\.[^/.]+$/, "");
        const location = extractedRows[0]?.text_mkv0z6d || extractedRows[0]?.store || 'unknown';
        const year = extractedRows[0]?.year || new Date().getFullYear();
        const month = extractedRows[0]?.month || new Date().getMonth() + 1;
        
        // Use environment variables for folder structure
        const basePath = process.env.SHAREPOINT_FOLDER_PATH?.replace(/^\/+|\/+$/g, '') || 'Form_Data';
        const folderPath = `${basePath}/GeneralManagement/${year}/${String(month).padStart(2, '0')}/${location}`;
        
        logMessage(`📁 Using configured base path: ${basePath}`, context);
        logMessage(`📁 Target SharePoint folder: ${folderPath}`, context);
        
        // IMPORTANT: Ensure folder exists BEFORE trying to upload files
        logMessage("📁 Creating folder structure before upload...", context);
        await ensureSharePointFolder(folderPath, context);
        logMessage("✅ Folder structure ready", context);
        
        // Generate file names
        const jsonFileName = `general-report-${baseFileName}-${timestamp}.json`;
        const textFileName = `general-report-${baseFileName}-${timestamp}.txt`;
        const originalDocFileName = `original-${originalFileName}`;
        
        logMessage(`📤 Uploading JSON report: ${jsonFileName}`, context);
        await uploadJsonToSharePoint(jsonReport, jsonFileName, folderPath, context);
        
        logMessage(`📤 Uploading text report: ${textFileName}`, context);
        await uploadTextToSharePoint(textReport, textFileName, folderPath, context);
        
        logMessage(`📤 Uploading original document: ${originalDocFileName}`, context);
        await uploadOriginalDocumentToSharePoint(base64BinFile, originalDocFileName, folderPath, context);
        
        logMessage("✅ All general management reports uploaded to SharePoint successfully", context);
        
    } catch (error) {
        logMessage(`❌ SharePoint upload process failed: ${error.message}`, context);
        handleError(error, 'SharePoint Upload', context);
        throw error;
    }
}

function generateJsonReport(rowDataArray, categories, originalFileName, context) {
    // Parse original filename for submission info
    const fileNameParts = parseFileName(originalFileName, context);
    
    // Get store and date info from first row
    const storeName = rowDataArray[0]?.text_mkv0z6d || "Unknown Store";
    const fullDate = rowDataArray[0]?.date4 || new Date().toISOString().split('T')[0];
    const yearMonth = fullDate.substring(0, 7); // YYYY-MM format
    
    const reportData = {
        // Report header (matching TXT exactly)
        title: "一般管理の実施記録",
        submissionDate: fileNameParts.submissionDate,
        submitter: fileNameParts.senderEmail,
        originalFileName: fileNameParts.originalFileName,
        storeName: storeName,
        yearMonth: yearMonth,
        
        // Category definitions (matching TXT exactly)
        categories: categories.map((category, index) => ({
            id: `Cat ${index + 1}`,
            name: category
        })),
        
        // Table headers (matching TXT exactly)
        tableHeaders: [
            "日付",
            "Cat 1",
            "Cat 2", 
            "Cat 3",
            "Cat 4",
            "Cat 5",
            "Cat 6",
            "Cat 7",
            "特記事項",
            "確認者"
        ],
        
        // Daily data (matching TXT table exactly)
        dailyData: rowDataArray.map(row => {
            const dayOnly = row.date4 ? row.date4.split('-')[2] : '--';
            
            return {
                日付: dayOnly,
                "Cat 1": row.color_mkv02tqg || '--',
                "Cat 2": row.color_mkv0yb6g || '--', 
                "Cat 3": row.color_mkv06e9z || '--',
                "Cat 4": row.color_mkv0x9mr || '--',
                "Cat 5": row.color_mkv0df43 || '--',
                "Cat 6": row.color_mkv5fa8m || '--',
                "Cat 7": row.color_mkv59ent || '--',
                特記事項: row.text_mkv0etfg || '--',
                確認者: row.color_mkv0xnn4 || '--'
            };
        }),
        
        // Footer (matching TXT exactly)
        footer: {
            generatedBy: "HygienMaster システム",
            generatedAt: new Date().toISOString(),
            note: "このレポートは HygienMaster システムにより自動生成されました"
        }
    };
    
    return reportData;
}

function generateTextReport(rowDataArray, categories, originalFileName, context) {
    // Parse original filename for submission info
    const fileNameParts = parseFileName(originalFileName, context);
    
    // Get store and date info from first row (if available)
    let storeName = 'Unknown Store';
    let yearMonth = new Date().toISOString().substring(0, 7);
    
    if (rowDataArray.length > 0 && rowDataArray[0]) {
        const firstRow = rowDataArray[0];
        storeName = firstRow.text_mkv0z6d || firstRow.store || 'Unknown Store';
        
        if (firstRow.date4) {
            yearMonth = firstRow.date4.substring(0, 7);
        } else if (firstRow.year && firstRow.month) {
            yearMonth = `${firstRow.year}-${String(firstRow.month).padStart(2, '0')}`;
        }
        
        // Remove the context parameter since it's not available in this function
        logMessage(`📊 Store: ${storeName}, Year-Month: ${yearMonth}`, context);
    }
    
    let textReport = `
一般管理の実施記録
提出日：${fileNameParts.submissionDate}
提出者：${fileNameParts.senderEmail}  
ファイル名：${fileNameParts.originalFileName}

店舗名：${storeName}
年月：${yearMonth}

管理カテゴリ：
`;

    // Add category descriptions
    if (categories && categories.length > 0) {
        categories.forEach((category, index) => {
            if (category && category !== 'not found') {
                textReport += `Cat ${index + 1}: ${category}\n`;
            }
        });
    } else {
        // Fallback category descriptions
        const defaultCategories = [
            '原材料の受入の確認',
            '庫内温度の確認 冷蔵庫・冷凍庫(°C)',
            '交差汚染・二次汚染の防止',
            '器具等の洗浄・消毒・殺菌',
            'トイレの洗浄・消毒',
            '従業員の健康管理等',
            '手洗いの実施'
        ];
        defaultCategories.forEach((category, index) => {
            textReport += `Cat ${index + 1}: ${category}\n`;
        });
    }

    textReport += '\n';

    // Create shorter table header
    const headerRow = `日付 | Cat 1 | Cat 2 | Cat 3 | Cat 4 | Cat 5 | Cat 6 | Cat 7 | 特記事項 | 確認者`;
    textReport += headerRow + '\n';
    textReport += ''.padEnd(headerRow.length, '-') + '\n';

    // Add data rows
    if (rowDataArray.length > 0) {
        rowDataArray.forEach(row => {
            if (row) {
                // Extract day from date4 (remove year-month part)
                let dayOnly = '--';
                if (row.date4) {
                    dayOnly = row.date4.split('-')[2] || '--';
                } else if (row.day) {
                    dayOnly = String(row.day).padStart(2, '0');
                }
                
                const dataRow = [
                    dayOnly.padEnd(4),
                    (row.color_mkv02tqg || '--').padEnd(6),
                    (row.color_mkv0yb6g || '--').padEnd(6), 
                    (row.color_mkv06e9z || '--').padEnd(6),
                    (row.color_mkv0x9mr || '--').padEnd(6),
                    (row.color_mkv0df43 || '--').padEnd(6),
                    (row.color_mkv5fa8m || '--').padEnd(6),
                    (row.color_mkv59ent || '--').padEnd(6),
                    (row.text_mkv0etfg && row.text_mkv0etfg !== 'not found' ? row.text_mkv0etfg.substring(0, 8) : '--').padEnd(8),
                    (row.color_mkv0xnn4 || '--')
                ].join('| ');
                
                textReport += dataRow + '\n';
            }
        });
    } else {
        // No data available
        textReport += 'データが見つかりませんでした。\n';
    }

    textReport += `
========================================
このレポートは HygienMaster システムにより自動生成されました
生成日時: ${new Date().toISOString()}
========================================
`;

    return textReport;
}

function parseFileName(fileName, context) {
    logMessage(`🔍 Parsing filename: ${fileName}`, context);
    
    try {
        let submissionTime = '';
        let senderEmail = '';
        let originalFileName = fileName;
        
        // Extract email (between parentheses)
        const emailMatch = fileName.match(/\(([^)]*@[^)]*)\)/);
        if (emailMatch) {
            senderEmail = emailMatch[1];
            logMessage(`📧 Found email: ${senderEmail}`, context);
            
            // Extract original filename - everything AFTER the (email) closing parenthesis
            const emailEndIndex = fileName.indexOf(emailMatch[0]) + emailMatch[0].length;
            originalFileName = fileName.substring(emailEndIndex);
            
            // Clean up any leading/trailing whitespace and remove leading special characters
            originalFileName = originalFileName.replace(/^\W+/, '').trim();
            
            logMessage(`📄 Found original filename: ${originalFileName}`, context);
        }
        
        // Extract timestamp (before first parenthesis)
        const timeMatch = fileName.match(/^([^(]+)/);
        if (timeMatch) {
            submissionTime = timeMatch[1];
            logMessage(`⏰ Found timestamp: ${submissionTime}`, context);
            
            // Try to parse the timestamp
            if (submissionTime.includes('T')) {
                try {
                    // Handle format like "20260826T050735"
                    const cleanTime = submissionTime.replace(/[^\d]/g, '');
                    if (cleanTime.length >= 8) {
                        const year = cleanTime.substring(0, 4);
                        const month = cleanTime.substring(4, 6);
                        const day = cleanTime.substring(6, 8);
                        const hour = cleanTime.substring(8, 10) || '00';
                        const minute = cleanTime.substring(10, 12) || '00';
                        
                        const isoString = `${year}-${month}-${day}T${hour}:${minute}:00`;
                        const date = new Date(isoString);
                        
                        if (!isNaN(date.getTime())) {
                            submissionTime = date.toLocaleDateString('ja-JP', {
                                year: 'numeric',
                                month: '2-digit', 
                                day: '2-digit',
                                hour: '2-digit',
                                minute: '2-digit'
                            });
                            logMessage(`📅 Parsed date: ${submissionTime}`, context);
                        }
                    }
                } catch (e) {
                    logMessage(`⚠️ Date parsing failed: ${e.message}`, context);
                }
            }
        }
        
        return {
            submissionDate: submissionTime || 'Unknown',
            senderEmail: senderEmail || 'Unknown',
            originalFileName: originalFileName || fileName
        };
        
    } catch (error) {
        logMessage(`❌ Filename parsing error: ${error.message}`, context);
        return {
            submissionDate: 'Unknown',
            senderEmail: 'Unknown', 
            originalFileName: fileName
        };
    }
}

function generateSummaryData(extractedRows, categories) {
    const totalDays = extractedRows.length;
    const approvedDays = extractedRows.filter(row => 
        (row.color_mkv0xnn4 || row.approverStatus) === '選択済み'
    ).length;
    const daysWithComments = extractedRows.filter(row => 
        (row.text_mkv0etfg || row.comment) && 
        (row.text_mkv0etfg || row.comment) !== 'not found'
    ).length;
    
    return {
        totalDays,
        approvedDays,
        approvalRate: totalDays > 0 ? (approvedDays / totalDays * 100).toFixed(1) : 0,
        daysWithComments,
        commentRate: totalDays > 0 ? (daysWithComments / totalDays * 100).toFixed(1) : 0
    };
}

function generateAnalyticsData(extractedRows, categories) {
    const analytics = {
        categoryPerformance: [],
        criticalDays: []
    };
    
    // Category performance analysis using Monday column mapping
    const categoryColumnMapping = {
        0: 'color_mkv02tqg', // Category1
        1: 'color_mkv0yb6g', // Category2
        2: 'color_mkv06e9z', // Category3
        3: 'color_mkv0x9mr', // Category4
        4: 'color_mkv0df43', // Category5
        5: 'color_mkv5fa8m', // Category6
        6: 'color_mkv59ent'  // Category7
    };
    
    categories.forEach((category, index) => {
        const mondayColumnId = categoryColumnMapping[index];
        const legacyKey = `category${index + 1}Status`;
        
        const okCount = extractedRows.filter(row => 
            (row[mondayColumnId] || row[legacyKey]) === '良'
        ).length;
        const ngCount = extractedRows.filter(row => 
            (row[mondayColumnId] || row[legacyKey]) === '否'
        ).length;
        
        analytics.categoryPerformance.push({
            categoryId: index + 1,
            categoryName: category,
            mondayColumnId: mondayColumnId,
            okCount,
            ngCount,
            successRate: extractedRows.length > 0 ? (okCount / extractedRows.length * 100).toFixed(1) : 0,
            riskLevel: ngCount > extractedRows.length * 0.2 ? "critical" : ngCount > 0 ? "high" : "normal"
        });
    });
    
    return analytics;
}

function getStatusCode(status) {
    switch(status) {
        case '良': return 1;
        case '否': return 0;
        default: return -1;
    }
}

module.exports = {
    prepareGeneralManagementReport
};