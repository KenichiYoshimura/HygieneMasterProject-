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

async function prepareImportantManagementReport(extractedRows, menuItems, context, base64BinFile, originalFileName) {
    logMessage("🚀 prepareImportantManagementReport() called", context);
    
    try {
        // DEBUG: Print the exact structure we're receiving
        logMessage("🔍 DEBUG: Raw input analysis...", context);
        logMessage(`📊 extractedRows type: ${typeof extractedRows}`, context);
        logMessage(`📊 extractedRows length: ${Array.isArray(extractedRows) ? extractedRows.length : 'not array'}`, context);
        logMessage(`📊 extractedRows content:`, context);
        logMessage(`${JSON.stringify(extractedRows, null, 2)}`, context);
        
        logMessage(`📊 menuItems:`, context);
        logMessage(`${JSON.stringify(menuItems, null, 2)}`, context);
        
        logMessage(`📊 originalFileName: ${originalFileName}`, context);
        
        // Handle both array and single object formats
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
        const jsonReport = generateJsonReport(rowDataArray, menuItems, originalFileName, context);
        logMessage("✅ JSON report generated", context);
        
        // Generate text report - NOW PASSING CONTEXT
        const textReport = generateTextReport(rowDataArray, menuItems, originalFileName, context);
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
        handleError(error, 'Important Management Report Generation', context);
        throw error;
    }
}

async function uploadReportsToSharePoint(jsonReport, textReport, base64BinFile, originalFileName, extractedRows, context) {
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const baseFileName = originalFileName.replace(/\.[^/.]+$/, "");
        const location = extractedRows[0]?.text_mkv0z6d || extractedRows[0]?.store || 'unknown';
        
        // Get date info
        const dateStr = extractedRows[0]?.date4 || new Date().toISOString().split('T')[0];
        const [year, month] = dateStr.split('-');
        
        // Use environment variables for folder structure
        const basePath = process.env.SHAREPOINT_FOLDER_PATH?.replace(/^\/+|\/+$/g, '') || 'Form_Data';
        const folderPath = `${basePath}/ImportantManagement/${year}/${String(month).padStart(2, '0')}/${location}`;
        
        logMessage(`📁 Target SharePoint folder: ${folderPath}`, context);
        
        // Create folder structure
        await ensureSharePointFolder(folderPath, context);
        
        // Generate file names
        const jsonFileName = `important-report-${baseFileName}-${timestamp}.json`;
        const textFileName = `important-report-${baseFileName}-${timestamp}.txt`;
        const originalDocFileName = `original-${originalFileName}`;
        
        // Upload files
        await uploadJsonToSharePoint(jsonReport, jsonFileName, folderPath, context);
        await uploadTextToSharePoint(textReport, textFileName, folderPath, context);
        await uploadOriginalDocumentToSharePoint(base64BinFile, originalDocFileName, folderPath, context);
        
        logMessage("✅ All important management reports uploaded to SharePoint successfully", context);
        
    } catch (error) {
        logMessage(`❌ SharePoint upload process failed: ${error.message}`, context);
        handleError(error, 'SharePoint Upload', context);
        throw error;
    }
}

function generateJsonReport(rowDataArray, menuItems, originalFileName, context) {
    // Parse original filename for submission info
    const fileNameParts = parseFileName(originalFileName, context);
    
    // Get store and date info from first row
    const storeName = rowDataArray[0]?.text_mkv0z6d || "unknown";
    const fullDate = rowDataArray[0]?.date4 || new Date().toISOString().split('T')[0];
    const yearMonth = fullDate.substring(0, 7); // YYYY-MM format
    
    const menuColumnMapping = {
        0: 'color_mkv02tqg', // Menu1
        1: 'color_mkv0yb6g', // Menu2
        2: 'color_mkv06e9z', // Menu3
        3: 'color_mkv0x9mr', // Menu4
        4: 'color_mkv0df43'  // Menu5
    };

    const reportData = {
        metadata: {
            reportType: "important_management_form",
            generatedAt: new Date().toISOString(),
            originalFileName: originalFileName,
            version: "1.0",
            mondayColumnMapping: {
                name: "name",
                date: "date4", 
                location: "text_mkv0z6d",
                comments: "text_mkv0etfg",
                approver: "color_mkv0xnn4",
                dailyCheck: "color_mkv0ej57",
                menuItems: menuColumnMapping
            }
        },
        
        // Report header information
        reportHeader: {
            title: "重要管理の実施記録",
            submissionDate: fileNameParts.submissionDate,
            submitter: fileNameParts.senderEmail,
            originalFileName: fileNameParts.originalFileName,
            storeName: storeName,
            yearMonth: yearMonth
        },
        
        // Menu items with their Monday column mappings
        menuItems: menuItems.map((item, index) => ({
            id: index + 1,
            name: item,
            mondayColumnId: menuColumnMapping[index] || `menu${index + 1}`,
            key: `menu${index + 1}`
        })),
        
        // Table headers (matching text report structure)
        tableHeaders: [
            "日付",
            "Menu 1",
            "Menu 2", 
            "Menu 3",
            "Menu 4",
            "Menu 5",
            "日常点検",
            "特記事項",
            "確認者"
        ],
        
        // Daily data rows
        dailyData: rowDataArray.map(row => {
            const dayOnly = row.date4 ? row.date4.split('-')[2] : '--';
            
            return {
                // Table row data (same order as headers)
                tableRow: [
                    dayOnly,
                    row.color_mkv02tqg || '--',
                    row.color_mkv0yb6g || '--', 
                    row.color_mkv06e9z || '--',
                    row.color_mkv0x9mr || '--',
                    row.color_mkv0df43 || '--',
                    row.color_mkv0ej57 || '--',
                    row.text_mkv0etfg || '--',
                    row.color_mkv0xnn4 || '--'
                ],
                
                // Individual field access
                day: dayOnly,
                menuStatuses: {
                    menu1: row.color_mkv02tqg || '--',
                    menu2: row.color_mkv0yb6g || '--',
                    menu3: row.color_mkv06e9z || '--',
                    menu4: row.color_mkv0x9mr || '--',
                    menu5: row.color_mkv0df43 || '--'
                },
                dailyCheck: row.color_mkv0ej57 || '--',
                comments: row.text_mkv0etfg || '--',
                approver: row.color_mkv0xnn4 || '--',
                
                // Raw Monday column data
                mondayColumnData: {
                    name: row.name,
                    date4: row.date4,
                    text_mkv0z6d: row.text_mkv0z6d,
                    color_mkv02tqg: row.color_mkv02tqg,
                    color_mkv0yb6g: row.color_mkv0yb6g,
                    color_mkv06e9z: row.color_mkv06e9z,
                    color_mkv0x9mr: row.color_mkv0x9mr,
                    color_mkv0df43: row.color_mkv0df43,
                    color_mkv0ej57: row.color_mkv0ej57,
                    text_mkv0etfg: row.text_mkv0etfg,
                    color_mkv0xnn4: row.color_mkv0xnn4
                },
                
                // Status codes for analysis
                statusCodes: {
                    menu1: getStatusCode(row.color_mkv02tqg),
                    menu2: getStatusCode(row.color_mkv0yb6g),
                    menu3: getStatusCode(row.color_mkv06e9z),
                    menu4: getStatusCode(row.color_mkv0x9mr),
                    menu5: getStatusCode(row.color_mkv0df43),
                    dailyCheck: getStatusCode(row.color_mkv0ej57),
                    approver: getStatusCode(row.color_mkv0xnn4)
                }
            };
        }),
        
        // Summary and analytics
        summary: generateSummaryData(rowDataArray, menuItems),
        analytics: generateAnalyticsData(rowDataArray, menuItems),
        
        // Footer information
        footer: {
            generatedBy: "HygienMaster システム",
            generatedAt: new Date().toISOString(),
            note: "このレポートは HygienMaster システムにより自動生成されました"
        }
    };
    
    return reportData;
}

function generateTextReport(rowDataArray, menuItems, originalFileName, context) {
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
        
        // Now we can use logMessage with context
        logMessage(`📊 Store: ${storeName}, Year-Month: ${yearMonth}`, context);
    }
    
    let textReport = `
重要管理の実施記録
提出日：${fileNameParts.submissionDate}
提出者：${fileNameParts.senderEmail}  
ファイル名：${fileNameParts.originalFileName}

店舗名：${storeName}
年月：${yearMonth}

重要管理項目：
`;

    // Add menu item descriptions
    if (menuItems && menuItems.length > 0) {
        menuItems.forEach((menuItem, index) => {
            if (menuItem && menuItem !== 'not found') {
                textReport += `Menu ${index + 1}: ${menuItem}\n`;
            }
        });
    } else {
        // Fallback menu item descriptions
        const defaultMenuItems = [
            '重要管理項目1',
            '重要管理項目2',
            '重要管理項目3',
            '重要管理項目4',
            '重要管理項目5'
        ];
        defaultMenuItems.forEach((menuItem, index) => {
            textReport += `Menu ${index + 1}: ${menuItem}\n`;
        });
    }

    textReport += '\n';

    // Create shorter table header
    const headerRow = `日付 | Menu 1 | Menu 2 | Menu 3 | Menu 4 | Menu 5 | 日常点検 | 特記事項 | 確認者`;
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
                    (row.color_mkv02tqg || '--').padEnd(7),
                    (row.color_mkv0yb6g || '--').padEnd(7), 
                    (row.color_mkv06e9z || '--').padEnd(7),
                    (row.color_mkv0x9mr || '--').padEnd(7),
                    (row.color_mkv0df43 || '--').padEnd(7),
                    (row.color_mkv0ej57 || '--').padEnd(8),
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
    // Now we can use logMessage with context
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
        
        // Extract original filename - improved regex to handle edge cases
        // Look for content after the closing parenthesis
        if (emailMatch) {
            const afterEmail = fileName.substring(fileName.indexOf(emailMatch[0]) + emailMatch[0].length);
            // Remove any leading non-alphanumeric characters except dots and spaces
            originalFileName = afterEmail.replace(/^[^\w\s.]+/, '').trim();
            if (originalFileName) {
                logMessage(`📄 Found original filename: ${originalFileName}`, context);
            } else {
                // Fallback: try to extract from the end
                const fallbackMatch = fileName.match(/[^)]*([^)]+\.[a-zA-Z]{2,4})$/);
                if (fallbackMatch) {
                    originalFileName = fallbackMatch[1].trim();
                    logMessage(`📄 Fallback original filename: ${originalFileName}`, context);
                } else {
                    originalFileName = fileName; // Use full filename as fallback
                }
            }
        } else {
            // No email found, try different approach
            const fileExtMatch = fileName.match(/([^/\\:*?"<>|]+\.[a-zA-Z]{2,4})$/);
            if (fileExtMatch) {
                originalFileName = fileExtMatch[1];
                logMessage(`📄 Extracted by extension: ${originalFileName}`, context);
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

function generateSummaryData(rowDataArray, menuItems) {
    const totalDays = rowDataArray.length;
    const approvedDays = rowDataArray.filter(row => 
        row.color_mkv0xnn4 === '良'
    ).length;
    const daysWithComments = rowDataArray.filter(row => 
        row.text_mkv0etfg && row.text_mkv0etfg !== 'not found'
    ).length;
    
    return {
        totalDays,
        approvedDays,
        approvalRate: totalDays > 0 ? (approvedDays / totalDays * 100).toFixed(1) : 0,
        daysWithComments,
        commentRate: totalDays > 0 ? (daysWithComments / totalDays * 100).toFixed(1) : 0
    };
}

function generateAnalyticsData(rowDataArray, menuItems) {
    const analytics = {
        menuPerformance: [],
        criticalDays: []
    };
    
    const menuColumnMapping = {
        0: 'color_mkv02tqg', // Menu1
        1: 'color_mkv0yb6g', // Menu2
        2: 'color_mkv06e9z', // Menu3
        3: 'color_mkv0x9mr', // Menu4
        4: 'color_mkv0df43'  // Menu5
    };
    
    menuItems.forEach((menuItem, index) => {
        const mondayColumnId = menuColumnMapping[index];
        
        const okCount = rowDataArray.filter(row => 
            row[mondayColumnId] === '良'
        ).length;
        const ngCount = rowDataArray.filter(row => 
            row[mondayColumnId] === '否'
        ).length;
        
        analytics.menuPerformance.push({
            menuId: index + 1,
            menuName: menuItem,
            mondayColumnId: mondayColumnId,
            okCount,
            ngCount,
            successRate: rowDataArray.length > 0 ? (okCount / rowDataArray.length * 100).toFixed(1) : 0,
            riskLevel: ngCount > rowDataArray.length * 0.2 ? "critical" : ngCount > 0 ? "high" : "normal"
        });
    });
    
    return analytics;
}

function getStatusCode(status) {
    switch(status) {
        case '良': return 1;
        case '否': return 0;
        case '未選択': return -1;
        case 'エラー': return -2;
        default: return -1;
    }
}

module.exports = {
    prepareImportantManagementReport
};
