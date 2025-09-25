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
    /*
    Monday General Management Form Column Mapping:
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
        formInfo: {
            location: extractedRows[0]?.text_mkv0z6d || extractedRows[0]?.store || "unknown",
            year: parseInt(extractedRows[0]?.year) || new Date().getFullYear(),
            month: parseInt(extractedRows[0]?.month) || new Date().getMonth() + 1,
            totalDays: extractedRows.length
        },
        categories: categories.map((category, index) => ({
            id: index + 1,
            name: category,
            mondayColumnId: categoryColumnMapping[index] || `category${index + 1}`,
            key: `category${index + 1}`
        })),
        dailyData: extractedRows.map(row => {
            const dailyEntry = {
                // Basic info using Monday column structure
                name: row.name || `${row.year}-${String(row.month).padStart(2, '0')}-${String(row.day).padStart(2, '0')}`,
                day: parseInt(row.day),
                date: `${row.year}-${String(row.month).padStart(2, '0')}-${String(row.day).padStart(2, '0')}`,
                date4: `${row.year}-${String(row.month).padStart(2, '0')}-${String(row.day).padStart(2, '0')}`, // Monday date format
                text_mkv0z6d: row.text_mkv0z6d || row.store || "unknown", // 店舗
                text_mkv0etfg: row.text_mkv0etfg || row.comment || null, // 特記事項
                color_mkv0xnn4: row.color_mkv0xnn4 || row.approverStatus || null, // 確認者
                
                // Category statuses using Monday column IDs
                categoryStatuses: categories.map((_, index) => {
                    const mondayColumnId = categoryColumnMapping[index];
                    const status = row[mondayColumnId] || row[`category${index + 1}Status`] || "unknown";
                    
                    return {
                        categoryId: index + 1,
                        categoryName: categories[index],
                        mondayColumnId: mondayColumnId,
                        status: status,
                        statusCode: getStatusCode(status),
                        rawValue: row[mondayColumnId] || row[`category${index + 1}Status`]
                    };
                }),
                
                // Include raw Monday column data for reference
                mondayColumnData: {
                    color_mkv02tqg: row.color_mkv02tqg,
                    color_mkv0yb6g: row.color_mkv0yb6g,
                    color_mkv06e9z: row.color_mkv06e9z,
                    color_mkv0x9mr: row.color_mkv0x9mr,
                    color_mkv0df43: row.color_mkv0df43,
                    color_mkv5fa8m: row.color_mkv5fa8m,
                    color_mkv59ent: row.color_mkv59ent
                },
                
                // Legacy format for compatibility
                comment: row.text_mkv0etfg || row.comment || null,
                approverStatus: row.color_mkv0xnn4 || row.approverStatus || null,
                isApproved: (row.color_mkv0xnn4 || row.approverStatus) === "選択済み"
            };
            
            return dailyEntry;
        }),
        summary: generateSummaryData(extractedRows, categories),
        analytics: generateAnalyticsData(extractedRows, categories)
    };
    
    return reportData;
}

function generateTextReport(extractedRows, categories, originalFileName) {
    const reportDate = new Date().toLocaleDateString('ja-JP', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
    const location = extractedRows[0]?.text_mkv0z6d || extractedRows[0]?.store || 'Unknown Location';
    const year = extractedRows[0]?.year || new Date().getFullYear();
    const month = extractedRows[0]?.month || new Date().getMonth() + 1;
    
    let textReport = `
========================================
📋 一般衛生管理フォーム 週間レポート
========================================

📋 基本情報:
  店舗: ${location}
  対象期間: ${year}年${month}月
  作成日: ${reportDate}
  元ファイル: ${originalFileName}

📊 管理カテゴリ:
${categories.map((category, index) => `  Category${index + 1}: ${category}`).join('\n')}

========================================
📅 日別管理状況
========================================

`;

    // Header row
    textReport += '日付    ';
    categories.forEach((_, index) => {
        textReport += `Cat${index + 1}  `;
    });
    textReport += '承認  コメント\n';
    textReport += ''.padEnd(80, '-') + '\n';

    // Data rows using Monday column structure
    extractedRows.forEach(row => {
        textReport += `${String(row.day).padEnd(6)}`;
        
        // Category statuses using Monday column mapping
        const categoryColumnMapping = {
            0: 'color_mkv02tqg', // Category1
            1: 'color_mkv0yb6g', // Category2
            2: 'color_mkv06e9z', // Category3
            3: 'color_mkv0x9mr', // Category4
            4: 'color_mkv0df43', // Category5
            5: 'color_mkv5fa8m', // Category6
            6: 'color_mkv59ent'  // Category7
        };
        
        categories.forEach((_, index) => {
            const mondayColumnId = categoryColumnMapping[index];
            const status = row[mondayColumnId] || row[`category${index + 1}Status`] || '—';
            const displayStatus = status === '良' ? '✓' : status === '否' ? '✗' : '?';
            textReport += `${displayStatus.padEnd(6)}`;
        });
        
        const approver = (row.color_mkv0xnn4 || row.approverStatus) === '選択済み' ? '✓' : '—';
        textReport += `${approver.padEnd(4)}`;
        
        const comment = row.text_mkv0etfg || row.comment || '—';
        if (comment && comment !== 'not found') {
            textReport += `${comment.substring(0, 30)}\n`;
        } else {
            textReport += '—\n';
        }
    });

    // Summary section
    const summary = generateSummaryData(extractedRows, categories);
    const analytics = generateAnalyticsData(extractedRows, categories);
    
    textReport += `
========================================
📈 週間サマリー
========================================

📊 全体統計:
  • 総日数: ${summary.totalDays}日
  • 承認済み: ${summary.approvedDays}日 (${summary.approvalRate}%)
  • コメント有り: ${summary.daysWithComments}日 (${summary.commentRate}%)

🚨 重要度レベル:
`;

    const criticalCategories = analytics.categoryPerformance.filter(cat => cat.riskLevel === 'critical');
    const highCategories = analytics.categoryPerformance.filter(cat => cat.riskLevel === 'high');

    textReport += `  • 緊急対応必要: ${criticalCategories.length}カテゴリ\n`;
    textReport += `  • 要注意: ${highCategories.length}カテゴリ\n\n`;

    if (criticalCategories.length > 0) {
        textReport += `⚠️ 問題発生カテゴリ:\n`;
        criticalCategories.forEach(cat => {
            textReport += `  • ${cat.categoryName}: ${cat.ngCount}件の問題 (成功率: ${cat.successRate}%)\n`;
        });
    }

    textReport += `
========================================
このレポートは HygienMaster システムにより自動生成されました
生成日時: ${new Date().toISOString()}
========================================
`;

    return textReport;
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