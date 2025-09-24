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

async function prepareGeneralManagementReport(extractedRows, categories, context, base64BinFile, originalFileName) {
    logMessage("🚀 prepareGeneralManagementReport() called", context);
    
    try {
        // Generate structured JSON data
        const jsonReport = generateJsonReport(extractedRows, categories, originalFileName);
        logMessage("✅ JSON report generated", context);
        
        // Generate PDF report
        const pdfReport = await generatePdfReport(extractedRows, categories, originalFileName);
        logMessage("✅ PDF report generated", context);
        
        // Upload to SharePoint
        await uploadReportsToSharePoint(jsonReport, pdfReport, base64BinFile, originalFileName, extractedRows, context);
        
        return {
            json: jsonReport,
            pdf: pdfReport
        };
        
    } catch (error) {
        handleError(error, 'Report Generation', context);
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
        
        // Create folder structure: Reports/GeneralManagement/Year/Month/Store
        const folderPath = `Reports/GeneralManagement/${year}/${String(month).padStart(2, '0')}/${location}`;
        
        // Ensure folder exists
        await ensureSharePointFolder(folderPath, context);
        
        // Generate file names
        const jsonFileName = `general-report-${baseFileName}-${timestamp}.json`;
        const pdfFileName = `general-report-${baseFileName}-${timestamp}.pdf`;
        const originalDocFileName = `original-${originalFileName}`;
        
        // Upload JSON report
        await uploadJsonToSharePoint(jsonReport, jsonFileName, folderPath, context);
        
        // Upload PDF report
        await uploadPdfToSharePoint(pdfReport, pdfFileName, folderPath, context);
        
        // Upload original document
        await uploadOriginalDocumentToSharePoint(base64BinFile, originalDocFileName, folderPath, context);
        
        logMessage("✅ All reports uploaded to SharePoint successfully", context);
        
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

async function generatePdfReport(extractedRows, categories, originalFileName) {
    // Using HTML-to-PDF approach (requires puppeteer or similar)
    // For now, generating HTML that can be converted to PDF
    const htmlContent = generateHtmlForPdf(extractedRows, categories, originalFileName);
    
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

function generateHtmlForPdf(extractedRows, categories, originalFileName) {
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
        .title { font-size: 18px; font-weight: bold; margin-bottom: 5px; }
        .subtitle { font-size: 14px; color: #666; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 10px; }
        th, td { border: 1px solid #333; padding: 6px 4px; text-align: center; }
        th { background-color: #f0f0f0; font-weight: bold; font-size: 9px; }
        .status-ok { background-color: #d4edda; }
        .status-ng { background-color: #f8d7da; }
        .status-unknown { background-color: #fff3cd; }
        .comments { font-size: 9px; text-align: left; max-width: 150px; word-wrap: break-word; }
        .summary { background-color: #f8f9fa; padding: 10px; margin-top: 20px; }
        .footer { text-align: center; font-size: 10px; color: #666; margin-top: 30px; }
    </style>
</head>
<body>
    <div class="header">
        <div class="title">一般管理フォーム 週間レポート</div>
        <div class="subtitle">店舗: ${location} | ${year}年${month}月 | 作成日: ${reportDate}</div>
        <div style="font-size: 10px; color: #888; margin-top: 5px;">元ファイル: ${originalFileName}</div>
    </div>

    <table>
        <thead>
            <tr>
                <th rowspan="2">日付</th>
                <th colspan="${categories.length}">管理項目</th>
                <th rowspan="2">承認</th>
                <th rowspan="2">備考・コメント</th>
            </tr>
            <tr>
                ${categories.map((cat, index) => `<th>項目${index + 1}<br>${cat.length > 8 ? cat.substring(0, 8) + '...' : cat}</th>`).join('')}
            </tr>
        </thead>
        <tbody>
            ${extractedRows.map(row => `
            <tr>
                <td><strong>${row.day}日</strong></td>
                ${categories.map((_, index) => {
                    const status = row[`cat${index + 1}Status`] || '—';
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
        <h3>週間サマリー</h3>
        ${generateSummaryHtml(extractedRows, categories)}
    </div>

    <div class="footer">
        <p>このレポートは HygienMaster システムにより自動生成されました (${new Date().toISOString()})</p>
    </div>
</body>
</html>`;
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

function generateSummaryHtml(extractedRows, categories) {
    const summary = generateSummaryData(extractedRows, categories);
    const analytics = generateAnalyticsData(extractedRows, categories);
    
    return `
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px; font-size: 11px;">
            <div>
                <strong>全体統計</strong><br>
                • 総日数: ${summary.totalDays}日<br>
                • 承認済み: ${summary.approvedDays}日 (${summary.approvalRate}%)<br>
                • コメント有り: ${summary.daysWithComments}日<br>
            </div>
            <div>
                <strong>リスク項目</strong><br>
                ${analytics.categoryPerformance
                    .filter(cat => cat.riskLevel === 'high')
                    .map(cat => `• ${cat.categoryName}: ${cat.ngCount}件の問題`).join('<br>') || '• リスク項目なし'}
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
    prepareGeneralManagementReport
};
