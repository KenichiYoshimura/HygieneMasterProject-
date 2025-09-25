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
        const jsonReport = generateJsonReport(rowDataArray, categories, originalFileName);
        logMessage("✅ JSON report generated", context);
        
        // Generate text report
        const textReport = generateTextReport(rowDataArray, categories, originalFileName);
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

function generateJsonReport(extractedRows, categories, originalFileName) {
    // Parse original filename for submission info
    const fileNameParts = parseFileName(originalFileName);
    
    // Get store and date info from first row
    const storeName = extractedRows[0]?.text_mkv0z6d || extractedRows[0]?.store || "unknown";
    const fullDate = extractedRows[0]?.date4 || extractedRows[0]?.year + '-' + String(extractedRows[0]?.month || new Date().getMonth() + 1).padStart(2, '0') + '-01';
    const yearMonth = fullDate.substring(0, 7); // YYYY-MM format
    
    const categoryColumnMapping = {
        0: 'color_mkv02tqg', // Category1
        1: 'color_mkv0yb6g', // Category2
        2: 'color_mkv06e9z', // Category3
        3: 'color_mkv0x9mr', // Category4
        4: 'color_mkv0df43', // Category5
        5: 'color_mkv5fa8m', // Category6
        6: 'color_mkv59ent'  // Category7
    };

    const reportData = {
        metadata: {
            reportType: "general_management_form",
            generatedAt: new Date().toISOString(),
            originalFileName: originalFileName,
            version: "1.0",
            mondayColumnMapping: {
                name: "name",
                date: "date4", 
                location: "text_mkv0z6d",
                comments: "text_mkv0etfg",
                approver: "color_mkv0xnn4",
                originalFile: "file_mkv1kpsc",
                categories: categoryColumnMapping
            }
        },
        
        // Report header information (same as text report)
        reportHeader: {
            title: "一般管理の実施記録",
            submissionDate: fileNameParts.submissionDate,
            submitter: fileNameParts.senderEmail,
            originalFileName: fileNameParts.originalFileName,
            storeName: storeName,
            yearMonth: yearMonth
        },
        
        // Categories with their Monday column mappings
        categories: categories.map((category, index) => ({
            id: index + 1,
            name: category,
            mondayColumnId: categoryColumnMapping[index] || `category${index + 1}`,
            key: `category${index + 1}`
        })),
        
        // Table headers (matching text report structure)
        tableHeaders: [
            "日付",
            categories[0] || "Category1",
            categories[1] || "Category2", 
            categories[2] || "Category3",
            categories[3] || "Category4",
            categories[4] || "Category5",
            categories[5] || "Category6",
            categories[6] || "Category7",
            "特記事項",
            "確認者"
        ],
        
        // Daily data rows (matching text report structure)
        dailyData: extractedRows.map(row => {
            // Extract day from date4 (remove year-month part)
            const dayOnly = row.date4 ? row.date4.split('-')[2] : (row.day ? String(row.day).padStart(2, '0') : '--');
            
            return {
                // Table row data (same order as headers)
                tableRow: [
                    dayOnly,
                    row.color_mkv02tqg || '--',
                    row.color_mkv0yb6g || '--', 
                    row.color_mkv06e9z || '--',
                    row.color_mkv0x9mr || '--',
                    row.color_mkv0df43 || '--',
                    row.color_mkv5fa8m || '--',
                    row.color_mkv59ent || '--',
                    row.text_mkv0etfg || '--',
                    row.color_mkv0xnn4 || '--'
                ],
                
                // Individual field access
                day: dayOnly,
                categoryStatuses: {
                    category1: row.color_mkv02tqg || '--',
                    category2: row.color_mkv0yb6g || '--',
                    category3: row.color_mkv06e9z || '--',
                    category4: row.color_mkv0x9mr || '--',
                    category5: row.color_mkv0df43 || '--',
                    category6: row.color_mkv5fa8m || '--',
                    category7: row.color_mkv59ent || '--'
                },
                comments: row.text_mkv0etfg || '--',
                approver: row.color_mkv0xnn4 || '--',
                
                // Raw Monday column data for reference
                mondayColumnData: {
                    name: row.name,
                    date4: row.date4,
                    text_mkv0z6d: row.text_mkv0z6d,
                    color_mkv02tqg: row.color_mkv02tqg,
                    color_mkv0yb6g: row.color_mkv0yb6g,
                    color_mkv06e9z: row.color_mkv06e9z,
                    color_mkv0x9mr: row.color_mkv0x9mr,
                    color_mkv0df43: row.color_mkv0df43,
                    color_mkv5fa8m: row.color_mkv5fa8m,
                    color_mkv59ent: row.color_mkv59ent,
                    text_mkv0etfg: row.text_mkv0etfg,
                    color_mkv0xnn4: row.color_mkv0xnn4
                },
                
                // Analysis fields
                statusCodes: {
                    category1: getStatusCode(row.color_mkv02tqg),
                    category2: getStatusCode(row.color_mkv0yb6g),
                    category3: getStatusCode(row.color_mkv06e9z),
                    category4: getStatusCode(row.color_mkv0x9mr),
                    category5: getStatusCode(row.color_mkv0df43),
                    category6: getStatusCode(row.color_mkv5fa8m),
                    category7: getStatusCode(row.color_mkv59ent)
                }
            };
        }),
        
        // Summary and analytics
        summary: generateSummaryData(extractedRows, categories),
        analytics: generateAnalyticsData(extractedRows, categories),
        
        // Footer information
        footer: {
            generatedBy: "HygienMaster システム",
            generatedAt: new Date().toISOString(),
            note: "このレポートは HygienMaster システムにより自動生成されました"
        }
    };
    
    return reportData;
}

function generateTextReport(rowData, categories, originalFileName) {
    // Parse original filename for submission info
    const fileNameParts = parseFileName(originalFileName);
    
    // Get store and date info from first row (if available)
    let storeName = 'Unknown Store';
    let yearMonth = new Date().toISOString().substring(0, 7);
    
    if (rowData.length > 0 && rowData[0]) {
        const firstRow = rowData[0];
        storeName = firstRow.text_mkv0z6d || firstRow.store || 'Unknown Store';
        
        if (firstRow.date4) {
            yearMonth = firstRow.date4.substring(0, 7);
        } else if (firstRow.year && firstRow.month) {
            yearMonth = `${firstRow.year}-${String(firstRow.month).padStart(2, '0')}`;
        }
        
        logMessage(`📊 Store: ${storeName}, Year-Month: ${yearMonth}`);
    }
    
    let textReport = `
一般管理の実施記録
提出日：${fileNameParts.submissionDate}
提出者：${fileNameParts.senderEmail}  
ファイル名：${fileNameParts.originalFileName}

店舗名：${storeName}
年月：${yearMonth}

`;

    // Create table header - handle missing categories gracefully
    const categoryNames = categories.length >= 7 ? categories : [
        '原材料の受入の確認',
        '庫内温度の確認',
        '交差汚染・二次汚染の防止',
        '器具等の洗浄・消毒・殺菌',
        'トイレの洗浄・消毒',
        '従業員の健康管理等',
        '手洗いの実施'
    ];
    
    const headerRow = `日付 | ${categoryNames.slice(0, 7).join(' | ')} | 特記事項 | 確認者`;
    textReport += headerRow + '\n';
    textReport += ''.padEnd(headerRow.length, '-') + '\n';

    // Add data rows
    if (rowData.length > 0) {
        rowData.forEach(row => {
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
                    (row.color_mkv02tqg || '--').padEnd(12),
                    (row.color_mkv0yb6g || '--').padEnd(25), 
                    (row.color_mkv06e9z || '--').padEnd(19),
                    (row.color_mkv0x9mr || '--').padEnd(20),
                    (row.color_mkv0df43 || '--').padEnd(16),
                    (row.color_mkv5fa8m || '--').padEnd(16),
                    (row.color_mkv59ent || '--').padEnd(12),
                    (row.text_mkv0etfg || '--').padEnd(8),
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

function parseFileName(fileName) {
    // Enhanced parsing for different filename formats
    logMessage(`🔍 Parsing filename: ${fileName}`);
    
    try {
        // Handle format: "timestamp(email)filename"
        // Example: "20260826T050735(kyosh1107@gmail.com)一般衛生管理シート V2.0 (記入後) 2-0.pdf"
        
        let submissionTime = '';
        let senderEmail = '';
        let originalFileName = fileName;
        
        // Extract email (between parentheses)
        const emailMatch = fileName.match(/\(([^)]*@[^)]*)\)/);
        if (emailMatch) {
            senderEmail = emailMatch[1];
            logMessage(`📧 Found email: ${senderEmail}`);
        }
        
        // Extract timestamp (before first parenthesis)
        const timeMatch = fileName.match(/^([^(]+)/);
        if (timeMatch) {
            submissionTime = timeMatch[1];
            logMessage(`⏰ Found timestamp: ${submissionTime}`);
            
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
                        
                        submissionTime = date.toLocaleDateString('ja-JP', {
                            year: 'numeric',
                            month: '2-digit', 
                            day: '2-digit',
                            hour: '2-digit',
                            minute: '2-digit'
                        });
                        logMessage(`📅 Parsed date: ${submissionTime}`);
                    }
                } catch (e) {
                    logMessage(`⚠️ Date parsing failed: ${e.message}`);
                }
            }
        }
        
        // Extract original filename (after last closing parenthesis)
        const fileMatch = fileName.match(/\)[^)]*(.+)$/);
        if (fileMatch) {
            originalFileName = fileMatch[1];
            logMessage(`📄 Found original filename: ${originalFileName}`);
        }
        
        return {
            submissionDate: submissionTime || 'Unknown',
            senderEmail: senderEmail || 'Unknown',
            originalFileName: originalFileName
        };
        
    } catch (error) {
        logMessage(`❌ Filename parsing error: ${error.message}`);
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