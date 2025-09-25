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
        // Generate structured JSON data
        const jsonReport = generateJsonReport(extractedRows, categories, originalFileName);
        logMessage("✅ JSON report generated", context);
        
        // Generate text report
        const textReport = generateTextReport(extractedRows, categories, originalFileName);
        logMessage("✅ Text report generated", context);
        
        // Upload to SharePoint
        logMessage("📤 Starting SharePoint upload...", context);
        await uploadReportsToSharePoint(jsonReport, textReport, base64BinFile, originalFileName, extractedRows, context);
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
    // 1. Parse original filename for submission info
    const fileNameParts = parseFileName(originalFileName);
    
    // 2. Get store and date info from first row
    const storeName = rowData[0]?.text_mkv0z6d || 'Unknown Store';
    const fullDate = rowData[0]?.date4 || new Date().toISOString().split('T')[0];
    const yearMonth = fullDate.substring(0, 7); // YYYY-MM format
    
    let textReport = `
一般管理の実施記録
提出日：${fileNameParts.submissionDate}
提出者：${fileNameParts.senderEmail}  
ファイル名：${fileNameParts.originalFileName}

店舗名：${storeName}
年月：${yearMonth}

`;

    // 3. Create table header
    const headerRow = `日付 | ${categories[0]} | ${categories[1]} | ${categories[2]} | ${categories[3]} | ${categories[4]} | ${categories[5]} | ${categories[6]} | 特記事項 | 確認者`;
    textReport += headerRow + '\n';
    textReport += ''.padEnd(headerRow.length, '-') + '\n';

    // 4. Add data rows
    rowData.forEach(row => {
        // Extract day from date4 (remove year-month part)
        const dayOnly = row.date4 ? row.date4.split('-')[2] : '--';
        
        const dataRow = [
            dayOnly.padEnd(4),
            (row.color_mkv02tqg || '--').padEnd(categories[0].length + 1),
            (row.color_mkv0yb6g || '--').padEnd(categories[1].length + 1), 
            (row.color_mkv06e9z || '--').padEnd(categories[2].length + 1),
            (row.color_mkv0x9mr || '--').padEnd(categories[3].length + 1),
            (row.color_mkv0df43 || '--').padEnd(categories[4].length + 1),
            (row.color_mkv5fa8m || '--').padEnd(categories[5].length + 1),
            (row.color_mkv59ent || '--').padEnd(categories[6].length + 1),
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
    // Parse format: "submission time(sender email)original-file-name"
    // Example: "20250925T103045(user@example.com)hygiene-form.pdf"
    
    try {
        // Extract submission time (before first parenthesis)
        const timeMatch = fileName.match(/^([^(]+)/);
        let submissionTime = timeMatch ? timeMatch[1] : '';
        
        // Extract sender email (between parentheses) 
        const emailMatch = fileName.match(/\(([^)]+)\)/);
        const senderEmail = emailMatch ? emailMatch[1] : '';
        
        // Extract original file name (after last parenthesis)
        const fileNameMatch = fileName.match(/\)[^)]*(.+)$/);
        let originalFileName = fileNameMatch ? fileNameMatch[1] : fileName;
        
        // Format submission time if it looks like ISO format
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
        // Fallback if parsing fails
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