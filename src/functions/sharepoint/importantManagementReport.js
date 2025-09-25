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
        // DEBUG: Print extractedRows and menuItems to understand the data structure
        logMessage("🔍 DEBUG: Analyzing extractedRows structure...", context);
        logMessage(`📊 extractedRows length: ${extractedRows.length}`, context);
        
        if (extractedRows.length > 0) {
            logMessage(`📋 First extractedRow sample:`, context);
            logMessage(`${JSON.stringify(extractedRows[0], null, 2)}`, context);
            
            // Access the nested .row property
            logMessage(`📋 First row data:`, context);
            logMessage(`${JSON.stringify(extractedRows[0].row, null, 2)}`, context);
            
            logMessage(`📋 All row keys from first row:`, context);
            logMessage(`${Object.keys(extractedRows[0].row).join(', ')}`, context);
        }
        
        logMessage("🔍 DEBUG: Analyzing menuItems structure...", context);
        logMessage(`📊 menuItems length: ${menuItems.length}`, context);
        logMessage(`📋 menuItems content:`, context);
        logMessage(`${JSON.stringify(menuItems, null, 2)}`, context);
        
        // Extract just the row data from extractedRows
        const rowData = extractedRows.map(item => item.row);
        
        // Generate structured JSON data
        const jsonReport = generateJsonReport(rowData, menuItems, originalFileName);
        logMessage("✅ JSON report generated", context);
        
        // Generate text report
        const textReport = generateTextReport(rowData, menuItems, originalFileName);
        logMessage("✅ Text report generated", context);
        
        // Upload to SharePoint
        logMessage("📤 Starting SharePoint upload...", context);
        await uploadReportsToSharePoint(jsonReport, textReport, base64BinFile, originalFileName, rowData, context);
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

async function uploadReportsToSharePoint(jsonReport, textReport, base64BinFile, originalFileName, rowData, context) {
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const baseFileName = originalFileName.replace(/\.[^/.]+$/, "");
        
        // Use the Monday.com column data directly
        const location = rowData[0]?.text_mkv0z6d || 'unknown';
        const dateStr = rowData[0]?.date4 || new Date().toISOString().split('T')[0];
        const [year, month] = dateStr.split('-');
        
        logMessage(`📋 Resolved location: ${location}`, context);
        logMessage(`📋 Resolved year: ${year}`, context);
        logMessage(`📋 Resolved month: ${month}`, context);
        
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

function generateJsonReport(rowData, menuItems, originalFileName) {
    // Parse original filename for submission info
    const fileNameParts = parseFileName(originalFileName);
    
    // Get store and date info from first row
    const storeName = rowData[0]?.text_mkv0z6d || "unknown";
    const fullDate = rowData[0]?.date4 || new Date().toISOString().split('T')[0];
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
            menuItems[0] || "Menu1",
            menuItems[1] || "Menu2", 
            menuItems[2] || "Menu3",
            menuItems[3] || "Menu4",
            menuItems[4] || "Menu5",
            "日常点検",
            "特記事項",
            "確認者"
        ],
        
        // Daily data rows
        dailyData: rowData.map(row => {
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
        summary: generateSummaryData(rowData, menuItems),
        analytics: generateAnalyticsData(rowData, menuItems),
        
        // Footer information
        footer: {
            generatedBy: "HygienMaster システム",
            generatedAt: new Date().toISOString(),
            note: "このレポートは HygienMaster システムにより自動生成されました"
        }
    };
    
    return reportData;
}

function generateTextReport(rowData, menuItems, originalFileName) {
    // Parse original filename for submission info
    const fileNameParts = parseFileName(originalFileName);
    
    // Get store and date info from first row
    const storeName = rowData[0]?.text_mkv0z6d || 'Unknown Store';
    const fullDate = rowData[0]?.date4 || new Date().toISOString().split('T')[0];
    const yearMonth = fullDate.substring(0, 7); // YYYY-MM format
    
    let textReport = `
重要管理の実施記録
提出日：${fileNameParts.submissionDate}
提出者：${fileNameParts.senderEmail}  
ファイル名：${fileNameParts.originalFileName}

店舗名：${storeName}
年月：${yearMonth}

`;

    // Create table header
    const headerRow = `日付 | ${menuItems[0]} | ${menuItems[1]} | ${menuItems[2]} | ${menuItems[3]} | ${menuItems[4]} | 日常点検 | 特記事項 | 確認者`;
    textReport += headerRow + '\n';
    textReport += ''.padEnd(headerRow.length, '-') + '\n';

    // Add data rows
    rowData.forEach(row => {
        const dayOnly = row.date4 ? row.date4.split('-')[2] : '--';
        
        const dataRow = [
            dayOnly.padEnd(4),
            (row.color_mkv02tqg || '--').padEnd(menuItems[0].length + 1),
            (row.color_mkv0yb6g || '--').padEnd(menuItems[1].length + 1), 
            (row.color_mkv06e9z || '--').padEnd(menuItems[2].length + 1),
            (row.color_mkv0x9mr || '--').padEnd(menuItems[3].length + 1),
            (row.color_mkv0df43 || '--').padEnd(menuItems[4].length + 1),
            (row.color_mkv0ej57 || '--').padEnd(8),
            (row.text_mkv0etfg || '--').padEnd(8),
            (row.color_mkv0xnn4 || '--')
        ].join('| ');
        
        textReport += dataRow + '\n';
    });

    textReport += `
========================================
このレポートは HygienMaster システムにより自動生成されました
生成日時: ${new Date().toISOString()}
========================================
`;

    return textReport;
}

function parseFileName(fileName) {
    // Same parsing logic as general management
    try {
        const timeMatch = fileName.match(/^([^(]+)/);
        let submissionTime = timeMatch ? timeMatch[1] : '';
        
        const emailMatch = fileName.match(/\(([^)]+)\)/);
        const senderEmail = emailMatch ? emailMatch[1] : '';
        
        const fileNameMatch = fileName.match(/\)[^)]*(.+)$/);
        let originalFileName = fileNameMatch ? fileNameMatch[1] : fileName;
        
        if (submissionTime.includes('T')) {
            try {
                const date = new Date(submissionTime);
                submissionTime = date.toLocaleDateString('ja-JP', {
                    year: 'numeric',
                    month: '2-digit', 
                    day: '2-digit',
                    hour: '2-digit',
                    minute: '2-digit'
                });
            } catch (e) {
                // Keep original if parsing fails
            }
        }
        
        return {
            submissionDate: submissionTime,
            senderEmail: senderEmail,
            originalFileName: originalFileName
        };
        
    } catch (error) {
        return {
            submissionDate: 'Unknown',
            senderEmail: 'Unknown', 
            originalFileName: fileName
        };
    }
}

function generateSummaryData(rowData, menuItems) {
    const totalDays = rowData.length;
    const approvedDays = rowData.filter(row => 
        row.color_mkv0xnn4 === '良'
    ).length;
    const daysWithComments = rowData.filter(row => 
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

function generateAnalyticsData(rowData, menuItems) {
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
        
        const okCount = rowData.filter(row => 
            row[mondayColumnId] === '良'
        ).length;
        const ngCount = rowData.filter(row => 
            row[mondayColumnId] === '否'
        ).length;
        
        analytics.menuPerformance.push({
            menuId: index + 1,
            menuName: menuItem,
            mondayColumnId: mondayColumnId,
            okCount,
            ngCount,
            successRate: rowData.length > 0 ? (okCount / rowData.length * 100).toFixed(1) : 0,
            riskLevel: ngCount > rowData.length * 0.2 ? "critical" : ngCount > 0 ? "high" : "normal"
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
