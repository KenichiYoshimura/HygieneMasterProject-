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
        
        // Upload to SharePoint - ADD THIS CALL
        logMessage("📤 Starting SharePoint upload...", context);
        await uploadReportsToSharePoint(jsonReport, pdfReport, base64BinFile, originalFileName, extractedRows, context);
        logMessage("✅ SharePoint upload completed", context);
        
        return {
            json: jsonReport,
            pdf: pdfReport
        };
        
    } catch (error) {
        handleError(error, 'Report Generation', context);
        throw error;
    }
}

// ADD THIS FUNCTION if missing
async function uploadReportsToSharePoint(jsonReport, pdfReport, base64BinFile, originalFileName, extractedRows, context) {
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
        const pdfFileName = `general-report-${baseFileName}-${timestamp}.pdf`;
        const originalDocFileName = `original-${originalFileName}`;
        
        logMessage(`📤 Uploading JSON report: ${jsonFileName}`, context);
        await uploadJsonToSharePoint(jsonReport, jsonFileName, folderPath, context);
        
        logMessage(`📤 Uploading PDF report: ${pdfFileName}`, context);
        await uploadPdfToSharePoint(pdfReport, pdfFileName, folderPath, context);
        
        logMessage(`📤 Uploading original document: ${originalDocFileName}`, context);
        await uploadOriginalDocumentToSharePoint(base64BinFile, originalDocFileName, folderPath, context);
        
        logMessage("✅ All reports uploaded to SharePoint successfully", context);
        
    } catch (error) {
        handleError(error, 'SharePoint Upload', context);
        throw error;
    }
}

// ADD THESE FUNCTIONS if missing
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
        summary: generateSummaryData(extractedRows, categories)
    };
    
    return reportData;
}

async function generatePdfReport(extractedRows, categories, originalFileName) {
    const htmlContent = generateHtmlForPdf(extractedRows, categories, originalFileName);
    return htmlContent; // Will be converted to PDF in SharePoint upload function
}

function generateHtmlForPdf(extractedRows, categories, originalFileName) {
    // Your existing HTML generation code here
    const reportDate = new Date().toLocaleDateString('ja-JP');
    const location = extractedRows[0]?.store || 'Unknown Location';
    
    return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: 'Yu Gothic', sans-serif; }
        .header { text-align: center; margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #333; padding: 8px; text-align: center; }
    </style>
</head>
<body>
    <div class="header">
        <h1>一般管理フォーム レポート</h1>
        <p>店舗: ${location} | 作成日: ${reportDate}</p>
    </div>
    <table>
        <tr>
            <th>日付</th>
            ${categories.map((cat, i) => `<th>項目${i+1}</th>`).join('')}
            <th>コメント</th>
        </tr>
        ${extractedRows.map(row => `
        <tr>
            <td>${row.day}日</td>
            ${categories.map((_, i) => `<td>${row[`cat${i+1}Status`] || '—'}</td>`).join('')}
            <td>${row.comment || '—'}</td>
        </tr>
        `).join('')}
    </table>
</body>
</html>`;
}

function generateSummaryData(extractedRows, categories) {
    return {
        totalDays: extractedRows.length,
        approvedDays: extractedRows.filter(row => row.approverStatus === '選択済み').length
    };
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
