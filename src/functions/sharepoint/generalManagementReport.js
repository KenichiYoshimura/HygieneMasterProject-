const { logMessage, handleError, convertHeicToJpegIfNeeded } = require('../utils');
const { 
    uploadJsonToSharePoint, 
    uploadTextToSharePoint,  // Changed from uploadPdfToSharePoint
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
        
        // Generate text report (changed from PDF)
        const textReport = generateTextReport(extractedRows, categories, originalFileName);
        logMessage("✅ Text report generated", context);
        
        // Upload to SharePoint
        logMessage("📤 Starting SharePoint upload...", context);
        await uploadReportsToSharePoint(jsonReport, textReport, base64BinFile, originalFileName, extractedRows, context);
        logMessage("✅ SharePoint upload completed", context);
        
        return {
            json: jsonReport,
            text: textReport  // Changed from pdf
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
        const location = extractedRows[0]?.store || 'unknown';
        const year = extractedRows[0]?.year || new Date().getFullYear();
        const month = extractedRows[0]?.month || new Date().getMonth() + 1;
        
        // Create folder structure: Reports/GeneralManagement/Year/Month/Store
        const folderPath = `Reports/GeneralManagement/${year}/${String(month).padStart(2, '0')}/${location}`;
        logMessage(`📁 Target SharePoint folder: ${folderPath}`, context);
        
        // Ensure folder exists
        logMessage("📁 Ensuring SharePoint folder exists...", context);
        await ensureSharePointFolder(folderPath, context);
        
        // Generate file names
        const jsonFileName = `general-report-${baseFileName}-${timestamp}.json`;
        const textFileName = `general-report-${baseFileName}-${timestamp}.txt`;  // Changed from .pdf
        const originalDocFileName = `original-${originalFileName}`;
        
        logMessage(`📤 Uploading JSON report: ${jsonFileName}`, context);
        await uploadJsonToSharePoint(jsonReport, jsonFileName, folderPath, context);
        
        logMessage(`📤 Uploading text report: ${textFileName}`, context);
        await uploadTextToSharePoint(textReport, textFileName, folderPath, context);  // Changed function call
        
        logMessage(`📤 Uploading original document: ${originalDocFileName}`, context);
        await uploadOriginalDocumentToSharePoint(base64BinFile, originalDocFileName, folderPath, context);
        
        logMessage("✅ All general management reports uploaded to SharePoint successfully", context);
        
    } catch (error) {
        handleError(error, 'SharePoint Upload', context);
        throw error;
    }
}

function generateJsonReport(extractedRows, categories, originalFileName) {
    const reportData = {
        metadata: {
            reportType: "general_management_form",
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
        categories: categories.map((cat, index) => ({
            id: index + 1,
            name: cat,
            key: `cat${index + 1}`
        })),
        dailyData: extractedRows.map(row => ({
            day: parseInt(row.day),
            date: `${row.year}-${String(row.month).padStart(2, '0')}-${String(row.day).padStart(2, '0')}`,
            categories: categories.map((_, index) => ({
                categoryId: index + 1,
                categoryName: categories[index],
                status: row[`cat${index + 1}Status`] || "unknown",
                statusCode: getStatusCode(row[`cat${index + 1}Status`])
            })),
            comment: row.comment && row.comment !== "not found" ? row.comment : null,
            approverStatus: row.approverStatus,
            isApproved: row.approverStatus === "選択済み"
        })),
        summary: generateSummaryData(extractedRows, categories),
        analytics: generateAnalyticsData(extractedRows, categories)
    };
    
    return reportData;
}

// Changed from generatePdfReport to generateTextReport
function generateTextReport(extractedRows, categories, originalFileName) {
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
📋 一般管理フォーム 週間レポート
========================================

📋 基本情報:
  店舗: ${location}
  対象期間: ${year}年${month}月
  作成日: ${reportDate}
  元ファイル: ${originalFileName}

📊 管理項目:
${categories.map((cat, index) => `  項目${index + 1}: ${cat}`).join('\n')}

========================================
📅 日別管理状況
========================================

`;

    // Header row
    textReport += '日付    ';
    categories.forEach((_, index) => {
        textReport += `項目${index + 1}  `;
    });
    textReport += '承認  コメント\n';
    textReport += ''.padEnd(80, '-') + '\n';

    // Data rows
    extractedRows.forEach(row => {
        textReport += `${String(row.day).padEnd(6)}`;
        
        categories.forEach((_, index) => {
            const status = row[`cat${index + 1}Status`] || '—';
            const displayStatus = status === '良' ? '✓' : status === '否' ? '✗' : '?';
            textReport += `${displayStatus.padEnd(6)}`;
        });
        
        const approver = row.approverStatus === '選択済み' ? '✓' : '—';
        textReport += `${approver.padEnd(4)}`;
        
        const comment = row.comment && row.comment !== 'not found' ? row.comment : '—';
        textReport += `${comment.substring(0, 30)}\n`;
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

📋 項目別統計:
`;

    analytics.categoryPerformance.forEach(cat => {
        textReport += `  • ${cat.categoryName}: 良${cat.okCount}件 / 否${cat.ngCount}件 (成功率: ${cat.successRate}%)\n`;
    });

    const riskCategories = analytics.categoryPerformance.filter(cat => cat.riskLevel === 'high');
    
    if (riskCategories.length > 0) {
        textReport += `
⚠️ 注意が必要な項目:
`;
        riskCategories.forEach(cat => {
            textReport += `  • ${cat.categoryName}: ${cat.ngCount}件の問題 (成功率: ${cat.successRate}%)\n`;
        });
    } else {
        textReport += `
✅ すべての項目が良好な状態です
`;
    }

    if (analytics.issuesDays.length > 0) {
        textReport += `
📅 問題発生日:
`;
        analytics.issuesDays.forEach(day => {
            textReport += `  • ${day.day}日: ${day.issueCount}件の問題 (${day.issues.join(', ')})\n`;
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

function generateAnalyticsData(extractedRows, categories) {
    const analytics = {
        categoryPerformance: [],
        trendData: [],
        issuesDays: []
    };
    
    // Category performance analysis
    categories.forEach((category, index) => {
        const statusKey = `cat${index + 1}Status`;
        const okCount = extractedRows.filter(row => row[statusKey] === '良').length;
        const ngCount = extractedRows.filter(row => row[statusKey] === '否').length;
        const unknownCount = extractedRows.filter(row => !row[statusKey] || row[statusKey] === 'not found').length;
        
        analytics.categoryPerformance.push({
            categoryId: index + 1,
            categoryName: category,
            okCount,
            ngCount,
            unknownCount,
            totalCount: extractedRows.length,
            successRate: extractedRows.length > 0 ? (okCount / extractedRows.length * 100).toFixed(1) : 0,
            riskLevel: ngCount > extractedRows.length * 0.3 ? "high" : ngCount > 0 ? "medium" : "low"
        });
    });
    
    // Days with issues
    extractedRows.forEach(row => {
        let issueCount = 0;
        let issues = [];
        
        categories.forEach((_, index) => {
            if (row[`cat${index + 1}Status`] === '否') {
                issueCount++;
                issues.push(`項目${index + 1}`);
            }
        });
        
        if (issueCount > 0) {
            analytics.issuesDays.push({
                day: row.day,
                issueCount,
                issues,
                hasComment: !!(row.comment && row.comment !== 'not found'),
                isApproved: row.approverStatus === '選択済み'
            });
        }
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