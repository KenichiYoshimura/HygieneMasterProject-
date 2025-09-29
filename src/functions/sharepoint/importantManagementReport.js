const { logMessage, handleError, convertHeicToJpegIfNeeded } = require('../utils');
const {
    uploadJsonToSharePoint,
    uploadTextToSharePoint,
    uploadOriginalDocumentToSharePoint,
    ensureSharePointFolder,
    uploadHtmlToSharePoint
} = require('./sendToSharePoint');
const { analyzeComment, supportedLanguages } = require('../analytics/sentimentAnalysis');
const axios = require('axios');
const { getReportStyles, getReportScripts } = require('./styles/sharedStyles');

/**
 * Prepares important management reports from structured data and uploads to SharePoint
 * 
 * @param {Object} structuredData - New structured data format from importantManagementFormExtractor
 * @param {Object} context - Azure Functions execution context
 * @param {string} base64BinFile - Base64 encoded original file
 * @param {string} originalFileName - Original filename for submission info
 */
async function prepareImportantManagementReport(structuredData, context, base64BinFile, originalFileName) {
    logMessage("🚀 prepareImportantManagementReport() called with structured data", context);
    
    try {
        logMessage("📊 Processing structured data:", context);
        logMessage(`  - Store: ${structuredData.metadata.location}`, context);
        logMessage(`  - Year-Month: ${structuredData.metadata.yearMonth}`, context);
        logMessage(`  - Daily Records: ${structuredData.dailyRecords.length}`, context);
        logMessage(`  - Menu Items: ${structuredData.menuItems.length}`, context);

        logMessage("🧠 Starting sentiment analysis for comments...", context);
        await addSentimentAnalysisToStructuredData(structuredData, context);
        logMessage("✅ Sentiment analysis completed", context);

        const jsonReport = generateJsonReport(structuredData, originalFileName, context);
        logMessage("✅ JSON report generated", context);

        const textReport = generateTextReport(structuredData, originalFileName, context);
        logMessage("✅ Text report generated", context);

        const htmlReport = generateHtmlReport(structuredData, originalFileName, context);
        logMessage("✅ HTML report generated", context);

        logMessage("📤 Starting SharePoint upload...", context);
        await uploadReportsToSharePoint(jsonReport, textReport, htmlReport, base64BinFile, originalFileName, structuredData, context);
        logMessage("✅ SharePoint upload completed", context);

        return {
            json: jsonReport,
            text: textReport,
            html: htmlReport
        };
        
    } catch (error) {
        handleError(error, 'Important Management Report Generation', context);
        throw error;
    }
}

/**
 * Adds sentiment analysis to comments in structured data
 */
async function addSentimentAnalysisToStructuredData(structuredData, context) {
    for (const record of structuredData.dailyRecords) {
        if (record.comment && record.comment !== "not found" && record.comment.trim()) {
            try {
                logMessage(`😊 Analyzing sentiment for comment: "${record.comment.substring(0, 30)}..."`, context);
                const sentimentResult = await analyzeComment(record.comment);
                
                // Add sentiment data to the record - the sentimentResult already contains the correct logic
                record.sentimentAnalysis = {
                    originalComment: sentimentResult.originalComment,
                    detectedLanguage: sentimentResult.detectedLanguage,
                    japaneseTranslation: sentimentResult.japaneseTranslation,
                    analysisLanguage: sentimentResult.sentimentAnalysisLanguage,
                    sentiment: sentimentResult.sentiment,
                    confidenceScores: sentimentResult.scores,
                    wasTranslated: sentimentResult.wasTranslated
                };
                
                const analysisInfo = sentimentResult.wasTranslated 
                    ? `translated from ${sentimentResult.detectedLanguage} to ja`
                    : `analyzed in original language ${sentimentResult.detectedLanguage}`;
                logMessage(`✅ Sentiment: ${sentimentResult.sentiment} (${Math.round(sentimentResult.scores[sentimentResult.sentiment] * 100)}% confidence) - ${analysisInfo}`, context);
                
            } catch (error) {
                logMessage(`❌ Sentiment analysis failed for comment: ${error.message}`, context);
                record.sentimentAnalysis = {
                    originalComment: record.comment,
                    detectedLanguage: 'unknown',
                    japaneseTranslation: null,
                    analysisLanguage: 'unknown',
                    error: error.message,
                    sentiment: "unknown",
                    wasTranslated: false
                };
            }
        }
    }
}

async function uploadReportsToSharePoint(jsonReport, textReport, htmlReport, base64BinFile, originalFileName, structuredData, context) {
    try {
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const baseFileName = originalFileName.replace(/\.[^/.]+$/, "");
        
        const location = structuredData.metadata.location;
        const [year, month] = structuredData.metadata.yearMonth.split('-');
        
        logMessage(`📋 Using structured data for folder: ${location}, ${year}-${month}`, context);
        
        const basePath = process.env.SHAREPOINT_FOLDER_PATH?.replace(/^\/+|\/+$/g, '') || '衛生管理日誌';
        const folderPath = `${basePath}/重要衛生管理の実施記録/${year}/${month}/${location}`;
        
        logMessage(`📁 Target SharePoint folder: ${folderPath}`, context);
        await ensureSharePointFolder(folderPath, context);

        // Use Japanese naming convention like legacy format
        const jsonFileName = `重要衛生管理レポート-${baseFileName}-${timestamp}.json`;
        const textFileName = `重要衛生管理レポート-${baseFileName}-${timestamp}.txt`;
        const htmlFileName = `重要衛生管理レポート-${baseFileName}-${timestamp}.html`;
        const originalDocFileName = `original-${originalFileName}`;

        await uploadJsonToSharePoint(jsonReport, jsonFileName, folderPath, context);
        await uploadTextToSharePoint(textReport, textFileName, folderPath, context);
        await uploadOriginalDocumentToSharePoint(base64BinFile, originalDocFileName, folderPath, context);
        await uploadHtmlToSharePoint(htmlReport, htmlFileName, folderPath, context);

        logMessage("✅ All important management reports uploaded to SharePoint successfully", context);
        
    } catch (error) {
        logMessage(`❌ SharePoint upload process failed: ${error.message}`, context);
        handleError(error, 'SharePoint Upload', context);
        throw error;
    }
}

function generateJsonReport(structuredData, originalFileName, context) {
    const fileNameParts = parseFileName(originalFileName, context);
    
    return {
        title: "重要衛生管理の実施記録",
        submissionDate: fileNameParts.submissionDate,
        submitter: fileNameParts.senderEmail,
        originalFileName: fileNameParts.originalFileName,
        storeName: structuredData.metadata.location,
        yearMonth: structuredData.metadata.yearMonth,
        
        menuItems: structuredData.menuItems.map((item, index) => ({
            id: `Menu ${index + 1}`,
            name: item.menuName
        })),
        
        tableHeaders: [
            "日付", "Menu 1", "Menu 2", "Menu 3", "Menu 4", "Menu 5", "日常点検", "特記事項", "確認者"
        ],
        
        dailyData: structuredData.dailyRecords.map(record => ({
            日付: String(record.day).padStart(2, '0'),
            "Menu 1": record.Menu1Status,
            "Menu 2": record.Menu2Status,
            "Menu 3": record.Menu3Status,
            "Menu 4": record.Menu4Status,
            "Menu 5": record.Menu5Status,
            日常点検: record.dailyCheckStatus,
            特記事項: record.comment !== "not found" ? record.comment : "--",
            確認者: record.approverStatus
        })),
        
        summary: {
            totalDays: structuredData.summary.totalDays,
            recordedDays: structuredData.summary.recordedDays,
            daysWithComments: structuredData.summary.daysWithComments,
            approvedDays: structuredData.summary.approvedDays,
            dailyCheckCompletedDays: structuredData.summary.dailyCheckCompletedDays,
            sentimentSummary: generateSentimentSummary(structuredData.dailyRecords)
        },
        
        footer: {
            generatedBy: "HygienMaster システム",
            generatedAt: new Date().toISOString(),
            note: "このレポートは HygienMaster システムにより自動生成されました"
        }
    };
}

function generateTextReport(structuredData, originalFileName, context) {
    const fileNameParts = parseFileName(originalFileName, context);
    
    let textReport = `
重要衛生管理の実施記録
提出日：${fileNameParts.submissionDate}
提出者：${fileNameParts.senderEmail}  
ファイル名：${fileNameParts.originalFileName}

店舗名：${structuredData.metadata.location}
年月：${structuredData.metadata.yearMonth}

重要管理項目：
`;

    structuredData.menuItems.forEach((menuItem, index) => {
        textReport += `Menu ${index + 1}: ${menuItem.menuName}\n`;
    });

    textReport += '\n';

    const headerRow = `日付 | Menu 1 | Menu 2 | Menu 3 | Menu 4 | Menu 5 | 日常点検 | 特記事項 | 確認者`;
    textReport += headerRow + '\n';
    textReport += ''.padEnd(headerRow.length, '-') + '\n';

    structuredData.dailyRecords.forEach(record => {
        const dataRow = [
            String(record.day).padStart(2, '0').padEnd(4),
            record.Menu1Status.padEnd(7),
            record.Menu2Status.padEnd(7),
            record.Menu3Status.padEnd(7),
            record.Menu4Status.padEnd(7),
            record.Menu5Status.padEnd(7),
            record.dailyCheckStatus.padEnd(8),
            (record.comment !== "not found" ? record.comment.substring(0, 8) : '--').padEnd(8),
            record.approverStatus
        ].join('| ');
        
        textReport += dataRow + '\n';
    });

    const sentimentSummary = generateSentimentSummary(structuredData.dailyRecords);
    textReport += `
========================================
感情分析サマリー：
ポジティブ: ${sentimentSummary.positive}件
ネガティブ: ${sentimentSummary.negative}件
ニュートラル: ${sentimentSummary.neutral}件
分析エラー: ${sentimentSummary.errors}件
========================================
このレポートは HygienMaster システムにより自動生成されました
生成日時: ${new Date().toISOString()}
========================================
`;

    return textReport;
}

function generateHtmlReport(structuredData, originalFileName, context) {
    const fileNameParts = parseFileName(originalFileName, context);

    const tableRows = structuredData.dailyRecords.map(record => {
        const statusClass = (status) => {
            switch(status) {
                case '良': return 'status-good';
                case '否': return 'status-bad';
                case '無': return 'status-none';
                default: return 'status-neutral';
            }
        };

        return `
        <tr class="data-row">
            <td class="date-cell">${String(record.day).padStart(2, '0')}</td>
            <td><span class="status-badge ${statusClass(record.Menu1Status)}">${record.Menu1Status}</span></td>
            <td><span class="status-badge ${statusClass(record.Menu2Status)}">${record.Menu2Status}</span></td>
            <td><span class="status-badge ${statusClass(record.Menu3Status)}">${record.Menu3Status}</span></td>
            <td><span class="status-badge ${statusClass(record.Menu4Status)}">${record.Menu4Status}</span></td>
            <td><span class="status-badge ${statusClass(record.Menu5Status)}">${record.Menu5Status}</span></td>
            <td><span class="status-badge ${statusClass(record.dailyCheckStatus)}">${record.dailyCheckStatus}</span></td>
            <td class="comment-cell">${record.comment !== "not found" ? record.comment : '--'}</td>
            <td><span class="status-badge ${statusClass(record.approverStatus)}">${record.approverStatus}</span></td>
        </tr>
        `;
    }).join('\n');

    // Updated sentiment rows to show all days with comments
    const sentimentRows = structuredData.dailyRecords
        .map(record => {
            const day = String(record.day).padStart(2, '0');
            
            // Check if sentiment analysis exists and was successful
            if (record.sentimentAnalysis && !record.sentimentAnalysis.error) {
                const sentiment = record.sentimentAnalysis;
                const sentimentClass = `sentiment-${sentiment.sentiment}`;
                const confidence = Math.round((sentiment.confidenceScores[sentiment.sentiment] || 0) * 100);
                
                return `
    <tr class="sentiment-row">
        <td class="date-cell">${day}</td>
        <td class="comment-text">${sentiment.originalComment}</td>
        <td class="language-tag">${sentiment.detectedLanguage}</td>
        <td class="translation-text">${sentiment.wasTranslated ? sentiment.japaneseTranslation : '<span class="no-translation">翻訳不要</span>'}</td>
        <td class="language-tag">${sentiment.analysisLanguage}</td>
        <td><span class="sentiment-badge ${sentimentClass}">${getSentimentIcon(sentiment.sentiment)} ${sentiment.sentiment}</span></td>
        <td class="confidence-bar">
            <div class="confidence-container">
                <div class="confidence-fill ${sentimentClass}" style="width: ${confidence}%"></div>
                <span class="confidence-text">${confidence}%</span>
            </div>
        </td>
    </tr>`;
            } else if (record.comment && record.comment !== "not found" && record.comment.trim()) {
                // Show why sentiment analysis wasn't performed
                let reason = 'コメントなし';
                if (record.sentimentAnalysis && record.sentimentAnalysis.error) {
                    reason = '分析エラー';
                } else {
                    reason = '未分析';
                }
                
                return `
    <tr class="sentiment-row no-analysis">
        <td class="date-cell">${day}</td>
        <td class="comment-text">${record.comment}</td>
        <td colspan="5" class="no-analysis-reason">${reason}</td>
    </tr>`;
            }
            return ''; // Skip records with no comments
        })
        .filter(row => row) // Remove empty rows
        .join('\n');

    const menuSummary = calculateMenuSummary(structuredData);
    const sentimentSummary = generateSentimentSummary(structuredData.dailyRecords);
    const complianceRate = Math.round((menuSummary.allGoodDays / structuredData.summary.recordedDays) * 100);
    const dailyCheckRate = Math.round((structuredData.summary.dailyCheckCompletedDays / structuredData.summary.recordedDays) * 100);

    return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>重要衛生管理の実施記録</title>
    ${getReportStyles()}
</head>
<body>
    <div class="report-container">
        <h1>重要衛生管理の実施記録</h1>
        <p>店舗名: ${structuredData.metadata.location}</p>
        <p>年月: ${structuredData.metadata.yearMonth}</p>
        <p>提出日: ${fileNameParts.submissionDate}</p>
        <p>提出者: ${fileNameParts.senderEmail}</p>
        <p>ファイル名: ${fileNameParts.originalFileName}</p>
        
        <h2>重要管理項目</h2>
        <ul>
            ${structuredData.menuItems.map(item => `<li>${item.menuName}</li>`).join('')}
        </ul>
        
        <h2>日次記録</h2>
        <table class="data-table">
            <thead>
                <tr>
                    ${structuredData.tableHeaders.map(header => `<th>${header}</th>`).join('')}
                </tr>
            </thead>
            <tbody>
                ${tableRows}
            </tbody>
        </table>
        
        <h2>感情分析結果</h2>
        <table class="sentiment-table">
            <thead>
                <tr>
                    <th>日付</th>
                    <th>コメント</th>
                    <th>検出言語</th>
                    <th>翻訳</th>
                    <th>分析言語</th>
                    <th>感情</th>
                    <th>信頼度</th>
                </tr>
            </thead>
            <tbody>
                ${sentimentRows}
            </tbody>
        </table>
        
        <h2>サマリー</h2>
        <p>記録された日数: ${structuredData.summary.recordedDays}日</p>
        <p>コメントのある日数: ${structuredData.summary.daysWithComments}日</p>
        <p>承認された日数: ${structuredData.summary.approvedDays}日</p>
        <p>日常点検完了日数: ${structuredData.summary.dailyCheckCompletedDays}日</p>
        <p>コンプライアンス率: ${complianceRate}%</p>
        <p>日常点検実施率: ${dailyCheckRate}%</p>
        
        <h2>感情分析サマリー</h2>
        <p>ポジティブ: ${sentimentSummary.positive}件</p>
        <p>ネガティブ: ${sentimentSummary.negative}件</p>
        <p>ニュートラル: ${sentimentSummary.neutral}件</p>
        <p>分析エラー: ${sentimentSummary.errors}件</p>
        
        <div class="footer">
            <p>このレポートは HygienMaster システムにより自動生成されました</p>
            <p>生成日時: ${new Date().toISOString()}</p>
        </div>
    </div>
    ${getReportScripts()}
</body>
</html>
`;
}

function parseFileName(originalFileName, context) {
    // Improved parsing logic to handle more cases and avoid crashes
    try {
        const nameParts = originalFileName.split('_');
        
        // Extract date in YYYYMMDD format
        const datePart = nameParts.find(part => /^\d{8}$/.test(part)) || '';
        const formattedDate = datePart.replace(/(\d{4})(\d{2})(\d{2})/, '$1-$2-$3');
        
        // Extract sender email (might be in the format name@example.com or just name)
        const senderEmail = nameParts.find(part => part.includes('@')) || '';
        
        // Original file name without extension
        const baseFileName = originalFileName.replace(/\.[^/.]+$/, "");
        
        return {
            submissionDate: formattedDate,
            senderEmail: senderEmail,
            originalFileName: baseFileName
        };
    } catch (error) {
        handleError(error, 'File Name Parsing', context);
        return {
            submissionDate: '',
            senderEmail: '',
            originalFileName: originalFileName // Fallback to original if parsing fails
        };
    }
}

function generateSentimentSummary(dailyRecords) {
    return dailyRecords.reduce((summary, record) => {
        if (record.sentimentAnalysis && !record.sentimentAnalysis.error) {
            summary[record.sentimentAnalysis.sentiment] = (summary[record.sentimentAnalysis.sentiment] || 0) + 1;
        } else {
            summary.errors = (summary.errors || 0) + 1;
        }
        return summary;
    }, { positive: 0, negative: 0, neutral: 0, errors: 0 });
}

function calculateMenuSummary(structuredData) {
    const summary = {
        allGoodDays: 0,
        someIssuesDays: 0,
        noRecordsDays: 0,
        totalDays: structuredData.summary.totalDays
    };
    
    structuredData.dailyRecords.forEach(record => {
        const menuStatuses = [record.Menu1Status, record.Menu2Status, record.Menu3Status, record.Menu4Status, record.Menu5Status];
        const allGood = menuStatuses.every(status => status === '良');
        const someIssues = menuStatuses.some(status => status === '否' || status === '無');
        
        if (allGood) {
            summary.allGoodDays++;
        } else if (someIssues) {
            summary.someIssuesDays++;
        } else {
            summary.noRecordsDays++;
        }
    });
    
    return summary;
}