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
        await uploadReportsToSharePoint(jsonReport, pdfReport, base64BinFile, originalFileName, extractedRows, context);
        
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
        
        // Ensure folder exists
        await ensureSharePointFolder(folderPath, context);
        
        // Generate file names
        const jsonFileName = `important-report-${baseFileName}-${timestamp}.json`;
        const pdfFileName = `important-report-${baseFileName}-${timestamp}.pdf`;
        const originalDocFileName = `original-${originalFileName}`;
        
        // Upload JSON report
        await uploadJsonToSharePoint(jsonReport, jsonFileName, folderPath, context);
        
        // Upload PDF report
        await uploadPdfToSharePoint(pdfReport, pdfFileName, folderPath, context);
        
        // Upload original document
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
    // Using HTML-to-PDF approach (requires puppeteer or similar)
    // For now, generating HTML that can be converted to PDF
    const htmlContent = generateHtmlForPdf(extractedRows, menuItems, originalFileName);
    
    // If you want to use puppeteer for actual PDF generation:
    /*
    const puppeteer = require('puppeteer');
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
    const pdfBuffer = await page.pdf({
        format: 'A4',
        printBackground: true,
        margin: { top: '20mm', bottom: '20mm', left: '10mm', right: '10mm' }
    });
    await browser.close();
    return pdfBuffer;
    */
    
    return htmlContent; // Return HTML for now, can be converted to PDF later
}

function generateHtmlForPdf(extractedRows, menuItems, originalFileName) {
    const reportDate = new Date().toLocaleDateString('ja-JP', {
        year: 'numeric',
        month: 'long',
        day: 'numeric'
    });
    const location = extractedRows[0]?.store || 'Unknown Location';
    const year = extractedRows[0]?.year || new Date().getFullYear();
    const month = extractedRows[0]?.month || new Date().getMonth() + 1;
    
    return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <style>
        @page { size: A4; margin: 20mm 15mm; }
        body { font-family: 'Yu Gothic', 'Hiragino Sans', sans-serif; font-size: 12px; line-height: 1.4; }
        .header { text-align: center; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 20px; }
        .title { font-size: 18px; font-weight: bold; margin-bottom: 5px; color: #d32f2f; }
        .subtitle { font-size: 14px; color: #666; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 10px; }
        th, td { border: 1px solid #333; padding: 6px 4px; text-align: center; }
        th { background-color: #ffebee; font-weight: bold; font-size: 9px; color: #d32f2f; }
        .status-ok { background-color: #d4edda; }
        .status-ng { background-color: #f8d7da; }
        .status-unknown { background-color: #fff3cd; }
        .comments { font-size: 9px; text-align: left; max-width: 150px; word-wrap: break-word; }
        .summary { background-color: #ffeaa7; padding: 10px; margin-top: 20px; border-left: 4px solid #fdcb6e; }
        .footer { text-align: center; font-size: 10px; color: #666; margin-top: 30px; }
        .important-badge { background-color: #d32f2f; color: white; padding: 2px 8px; border-radius: 12px; font-size: 10px; }
    </style>
</head>
<body>
    <div class="header">
        <div class="title">
            <span class="important-badge">重要</span>
            重要管理フォーム 週間レポート
        </div>
        <div class="subtitle">店舗: ${location} | ${year}年${month}月 | 作成日: ${reportDate}</div>
        <div style="font-size: 10px; color: #888; margin-top: 5px;">元ファイル: ${originalFileName}</div>
    </div>

    <table>
        <thead>
            <tr>
                <th rowspan="2">日付</th>
                <th colspan="${menuItems.length}">重要管理項目</th>
                <th rowspan="2">承認</th>
                <th rowspan="2">備考・コメント</th>
            </tr>
            <tr>
                ${menuItems.map((item, index) => `<th>項目${index + 1}<br>${item.length > 10 ? item.substring(0, 10) + '...' : item}</th>`).join('')}
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
                <td class="comments">${row.comment && row.comment !== 'not found' ? row.comment : '—'}</td>
            </tr>
            `).join('')}
        </tbody>
    </table>

    <div class="summary">
        <h3>🚨 重要管理項目 週間サマリー</h3>
        ${generateSummaryHtml(extractedRows, menuItems)}
    </div>

    <div class="footer">
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
        trendData: [],
        criticalDays: []
    };
    
    // Menu item performance analysis
    menuItems.forEach((menuItem, index) => {
        const statusKey = `menu${index + 1}Status`;
        const okCount = extractedRows.filter(row => row[statusKey] === '良').length;
        const ngCount = extractedRows.filter(row => row[statusKey] === '否').length;
        const unknownCount = extractedRows.filter(row => !row[statusKey] || row[statusKey] === 'not found').length;
        
        analytics.menuPerformance.push({
            menuId: index + 1,
            menuName: menuItem,
            okCount,
            ngCount,
            unknownCount,
            totalCount: extractedRows.length,
            successRate: extractedRows.length > 0 ? (okCount / extractedRows.length * 100).toFixed(1) : 0,
            riskLevel: ngCount > extractedRows.length * 0.2 ? "critical" : ngCount > 0 ? "high" : "normal",
            priority: ngCount > extractedRows.length * 0.3 ? "urgent" : "normal"
        });
    });
    
    // Critical days with issues
    extractedRows.forEach(row => {
        let criticalIssueCount = 0;
        let criticalIssues = [];
        
        menuItems.forEach((_, index) => {
            if (row[`menu${index + 1}Status`] === '否') {
                criticalIssueCount++;
                criticalIssues.push(`項目${index + 1}`);
            }
        });
        
        if (criticalIssueCount > 0) {
            analytics.criticalDays.push({
                day: row.day,
                criticalIssueCount,
                criticalIssues,
                hasComment: !!(row.comment && row.comment !== 'not found'),
                isApproved: row.approverStatus === '選択済み',
                severity: criticalIssueCount > menuItems.length * 0.5 ? "severe" : "moderate"
            });
        }
    });
    
    return analytics;
}

function generateSummaryHtml(extractedRows, menuItems) {
    const summary = generateSummaryData(extractedRows, menuItems);
    const analytics = generateAnalyticsData(extractedRows, menuItems);
    
    // Count critical and urgent items
    const criticalItems = analytics.menuPerformance.filter(menu => menu.riskLevel === 'critical');
    const urgentItems = analytics.menuPerformance.filter(menu => menu.priority === 'urgent');
    
    return `
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px; font-size: 11px;">
            <div>
                <strong>📊 全体統計</strong><br>
                • 総日数: ${summary.totalDays}日<br>
                • 承認済み: ${summary.approvedDays}日 (${summary.approvalRate}%)<br>
                • コメント有り: ${summary.daysWithComments}日<br>
                <br>
                <strong>🚨 重要度レベル</strong><br>
                • 緊急対応必要: ${urgentItems.length}項目<br>
                • 要注意: ${criticalItems.length}項目
            </div>
            <div>
                <strong>⚠️ 問題発生項目</strong><br>
                ${analytics.menuPerformance
                    .filter(menu => menu.riskLevel === 'critical')
                    .slice(0, 3)
                    .map(menu => `• ${menu.menuName}: ${menu.ngCount}件`).join('<br>') || '• 重大な問題なし'}<br>
                <br>
                <strong>📅 問題発生日</strong><br>
                ${analytics.criticalDays.length > 0 
                    ? analytics.criticalDays.map(day => `• ${day.day}日: ${day.criticalIssueCount}件`).join('<br>')
                    : '• 問題発生日なし'}
            </div>
        </div>
    `;
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
