const { logMessage, handleError, convertHeicToJpegIfNeeded} = require('../utils');
const { 
    uploadJsonToSharePoint, 
    uploadPdfToSharePoint, 
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
        
        // Generate PDF report
        const pdfReport = await generatePdfReport(extractedRows, menuItems, originalFileName);
        logMessage("✅ PDF report generated", context);
        
        // Upload to SharePoint
        logMessage("📤 Starting SharePoint upload...", context);
        await uploadReportsToSharePoint(jsonReport, pdfReport, base64BinFile, originalFileName, extractedRows, context);
        logMessage("✅ SharePoint upload completed", context);
        
        return {
            json: jsonReport,
            pdf: pdfReport
        };
        
    } catch (error) {
        handleError(error, 'Important Management Report Generation', context);
        throw error;
    }
}

async function uploadReportsToSharePoint(jsonReport, pdfReport, base64BinFile, originalFileName, extractedRows, context) {
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const baseFileName = originalFileName.replace(/\.[^/.]+$/, "");
        const location = extractedRows[0]?.store || 'unknown';
        const year = extractedRows[0]?.year || new Date().getFullYear();
        const month = extractedRows[0]?.month || new Date().getMonth() + 1;
        
        // Create folder structure: Reports/ImportantManagement/Year/Month/Store
        const folderPath = `Reports/ImportantManagement/${year}/${String(month).padStart(2, '0')}/${location}`;
        logMessage(`📁 Target SharePoint folder: ${folderPath}`, context);
        
        // Ensure folder exists
        logMessage("📁 Ensuring SharePoint folder exists...", context);
        await ensureSharePointFolder(folderPath, context);
        
        // Generate file names
        const jsonFileName = `important-report-${baseFileName}-${timestamp}.json`;
        const pdfFileName = `important-report-${baseFileName}-${timestamp}.pdf`;
        const originalDocFileName = `original-${originalFileName}`;
        
        logMessage(`📤 Uploading JSON report: ${jsonFileName}`, context);
        await uploadJsonToSharePoint(jsonReport, jsonFileName, folderPath, context);
        
        logMessage(`📤 Uploading PDF report: ${pdfFileName}`, context);
        await uploadPdfToSharePoint(pdfReport, pdfFileName, folderPath, context);
        
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

async function generatePdfReport(extractedRows, menuItems, originalFileName) {
    const htmlContent = generateHtmlForPdf(extractedRows, menuItems, originalFileName);
    return htmlContent; // Will be converted to PDF in SharePoint upload function
}

function generateHtmlForPdf(extractedRows, menuItems, originalFileName) {
    const reportDate = new Date().toLocaleDateString('ja-JP');
    const location = extractedRows[0]?.store || 'Unknown Location';
    const year = extractedRows[0]?.year || new Date().getFullYear();
    const month = extractedRows[0]?.month || new Date().getMonth() + 1;
    
    return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: 'Yu Gothic', sans-serif; font-size: 12px; }
        .header { text-align: center; margin-bottom: 20px; border-bottom: 2px solid #d32f2f; padding-bottom: 10px; }
        .title { font-size: 18px; font-weight: bold; color: #d32f2f; }
        .important-badge { background-color: #d32f2f; color: white; padding: 2px 8px; border-radius: 12px; font-size: 10px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { border: 1px solid #333; padding: 8px; text-align: center; }
        th { background-color: #ffebee; font-weight: bold; color: #d32f2f; }
        .status-ok { background-color: #d4edda; }
        .status-ng { background-color: #f8d7da; }
        .status-unknown { background-color: #fff3cd; }
    </style>
</head>
<body>
    <div class="header">
        <div class="title">
            <span class="important-badge">重要</span>
            重要管理フォーム レポート
        </div>
        <p>店舗: ${location} | ${year}年${month}月 | 作成日: ${reportDate}</p>
        <div style="font-size: 10px; color: #888;">元ファイル: ${originalFileName}</div>
    </div>

    <table>
        <thead>
            <tr>
                <th>日付</th>
                ${menuItems.map((item, index) => `<th>項目${index + 1}<br>${item.length > 10 ? item.substring(0, 10) + '...' : item}</th>`).join('')}
                <th>承認</th>
                <th>コメント</th>
            </tr>
        </thead>
        <tbody>
            ${extractedRows.map(row => `
            <tr>
                <td><strong>${row.day}日</strong></td>
                ${menuItems.map((_, index) => {
                    const status = row[`menu${index + 1}Status`] || '—';
                    const cssClass = status === '良' ? 'status-ok' : status === '否' ? 'status-ng' : 'status-unknown';
                    return `<td class="${cssClass}">${status}</td>`;
                }).join('')}
                <td>${row.approverStatus === '選択済み' ? '✓' : '—'}</td>
                <td style="text-align: left; max-width: 150px;">${row.comment && row.comment !== 'not found' ? row.comment : '—'}</td>
            </tr>
            `).join('')}
        </tbody>
    </table>

    <div style="margin-top: 20px; text-align: center; color: #666; font-size: 10px;">
        <p>このレポートは HygienMaster システムにより自動生成されました (${new Date().toISOString()})</p>
    </div>
</body>
</html>`;
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
