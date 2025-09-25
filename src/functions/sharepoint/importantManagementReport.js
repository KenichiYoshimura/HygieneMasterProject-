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
        // Generate structured JSON data
        const jsonReport = generateJsonReport(extractedRows, menuItems, originalFileName);
        logMessage("✅ JSON report generated", context);
        
        // Generate text report
        const textReport = generateTextReport(extractedRows, menuItems, originalFileName);
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
        handleError(error, 'Important Management Report Generation', context);
        throw error;
    }
}

async function uploadReportsToSharePoint(jsonReport, textReport, base64BinFile, originalFileName, extractedRows, context) {
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const baseFileName = originalFileName.replace(/\.[^/.]+$/, "");
        const location = extractedRows[0]?.store || 'unknown';
        const year = extractedRows[0]?.year || new Date().getFullYear();
        const month = extractedRows[0]?.month || new Date().getMonth() + 1;
        
        // Use environment variables for folder structure
        const basePath = process.env.SHAREPOINT_FOLDER_PATH?.replace(/^\/+|\/+$/g, '') || 'Form_Data';
        const folderPath = `${basePath}/ImportantManagement/${year}/${String(month).padStart(2, '0')}/${location}`;
        
        logMessage(`📁 Using configured base path: ${basePath}`, context);
        logMessage(`📁 Target SharePoint folder: ${folderPath}`, context);
        
        // Ensure folder exists
        logMessage("📁 Ensuring SharePoint folder exists...", context);
        await ensureSharePointFolder(folderPath, context);
        
        // Generate file names
        const jsonFileName = `important-report-${baseFileName}-${timestamp}.json`;
        const textFileName = `important-report-${baseFileName}-${timestamp}.txt`;
        const originalDocFileName = `original-${originalFileName}`;
        
        logMessage(`📤 Uploading JSON report: ${jsonFileName}`, context);
        await uploadJsonToSharePoint(jsonReport, jsonFileName, folderPath, context);
        
        logMessage(`📤 Uploading text report: ${textFileName}`, context);
        await uploadTextToSharePoint(textReport, textFileName, folderPath, context);
        
        logMessage(`📤 Uploading original document: ${originalDocFileName}`, context);
        await uploadOriginalDocumentToSharePoint(base64BinFile, originalDocFileName, folderPath, context);
        
        logMessage("✅ All important management reports uploaded to SharePoint successfully", context);
        
    } catch (error) {
        handleError(error, 'SharePoint Upload', context);
        throw error;
    }
}

function generateJsonReport(extractedRows, menuItems, originalFileName) {
    const reportData = {
        metadata: {
            reportType: "important_management_form",
            generatedAt: new Date().toISOString(),
            originalFileName: originalFileName,
            version: "1.0"
        },
        formInfo: {
            location: extractedRows[0]?.store || "unknown",
            year: parseInt(extractedRows[0]?.year) || new Date().getFullYear(),
            month: parseInt(extractedRows[0]?.month) || new Date().getMonth() + 1,
            totalDays: extractedRows.length
        },
        menuItems: menuItems.map((item, index) => ({
            id: index + 1,
            name: item,
            key: `menu${index + 1}`
        })),
        dailyData: extractedRows.map(row => ({
            day: parseInt(row.day),
            date: `${row.year}-${String(row.month).padStart(2, '0')}-${String(row.day).padStart(2, '0')}`,
            menuStatuses: menuItems.map((_, index) => ({
                menuId: index + 1,
                menuName: menuItems[index],
                status: row[`menu${index + 1}Status`] || "unknown",
                statusCode: getStatusCode(row[`menu${index + 1}Status`])
            })),
            comment: row.comment && row.comment !== "not found" ? row.comment : null,
            approverStatus: row.approverStatus,
            isApproved: row.approverStatus === "選択済み"
        })),
        summary: generateSummaryData(extractedRows, menuItems),
        analytics: generateAnalyticsData(extractedRows, menuItems)
    };
    
    return reportData;
}

// Changed from generatePdfReport to generateTextReport
function generateTextReport(extractedRows, menuItems, originalFileName) {
    const reportDate = new Date().toLocaleDateString('ja-JP', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
    const location = extractedRows[0]?.store || 'Unknown Location';
    const year = extractedRows[0]?.year || new Date().getFullYear();
    const month = extractedRows[0]?.month || new Date().getMonth() + 1;
    
    let textReport = `
========================================
🚨 重要管理フォーム 週間レポート
========================================

📋 基本情報:
  店舗: ${location}
  対象期間: ${year}年${month}月
  作成日: ${reportDate}
  元ファイル: ${originalFileName}

📊 重要管理項目:
${menuItems.map((item, index) => `  項目${index + 1}: ${item}`).join('\n')}

========================================
📅 日別管理状況
========================================

`;

    // Header row
    textReport += '日付    ';
    menuItems.forEach((_, index) => {
        textReport += `項目${index + 1}  `;
    });
    textReport += '承認  コメント\n';
    textReport += ''.padEnd(80, '-') + '\n';

    // Data rows
    extractedRows.forEach(row => {
        textReport += `${String(row.day).padEnd(6)}`;
        
        menuItems.forEach((_, index) => {
            const status = row[`menu${index + 1}Status`] || '—';
            const displayStatus = status === '良' ? '✓' : status === '否' ? '✗' : '?';
            textReport += `${displayStatus.padEnd(6)}`;
        });
        
        const approver = row.approverStatus === '選択済み' ? '✓' : '—';
        textReport += `${approver.padEnd(4)}`;
        
        const comment = row.comment && row.comment !== 'not found' ? row.comment : '—';
        textReport += `${comment.substring(0, 30)}\n`;
    });

    // Summary section
    const summary = generateSummaryData(extractedRows, menuItems);
    const analytics = generateAnalyticsData(extractedRows, menuItems);
    
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

    const criticalItems = analytics.menuPerformance.filter(menu => menu.riskLevel === 'critical');
    const highItems = analytics.menuPerformance.filter(menu => menu.riskLevel === 'high');

    textReport += `  • 緊急対応必要: ${criticalItems.length}項目\n`;
    textReport += `  • 要注意: ${highItems.length}項目\n\n`;

    if (criticalItems.length > 0) {
        textReport += `⚠️ 問題発生項目:\n`;
        criticalItems.forEach(item => {
            textReport += `  • ${item.menuName}: ${item.ngCount}件の問題 (成功率: ${item.successRate}%)\n`;
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

function generateSummaryData(extractedRows, menuItems) {
    const totalDays = extractedRows.length;
    const approvedDays = extractedRows.filter(row => row.approverStatus === '選択済み').length;
    const daysWithComments = extractedRows.filter(row => row.comment && row.comment !== 'not found').length;
    
    return {
        totalDays,
        approvedDays,
        approvalRate: totalDays > 0 ? (approvedDays / totalDays * 100).toFixed(1) : 0,
        daysWithComments,
        commentRate: totalDays > 0 ? (daysWithComments / totalDays * 100).toFixed(1) : 0
    };
}

function generateAnalyticsData(extractedRows, menuItems) {
    const analytics = {
        menuPerformance: [],
        criticalDays: []
    };
    
    // Menu item performance analysis
    menuItems.forEach((menuItem, index) => {
        const statusKey = `menu${index + 1}Status`;
        const okCount = extractedRows.filter(row => row[statusKey] === '良').length;
        const ngCount = extractedRows.filter(row => row[statusKey] === '否').length;
        
        analytics.menuPerformance.push({
            menuId: index + 1,
            menuName: menuItem,
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
    prepareImportantManagementReport
};
