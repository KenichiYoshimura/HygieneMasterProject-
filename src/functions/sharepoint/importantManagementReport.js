const { logMessage, handleError, convertHeicToJpegIfNeeded } = require('../utils');
const {
    uploadJsonToSharePoint,
    uploadTextToSharePoint,
    uploadOriginalDocumentToSharePoint,
    ensureSharePointFolder,
    uploadHtmlToSharePoint
} = require('./sendToSharePoint');
const { analyzeComment } = require('../analytics/sentimentAnalysis');
const axios = require('axios');

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
 * Modifies the structuredData object in place
 */
async function addSentimentAnalysisToStructuredData(structuredData, context) {
    for (const record of structuredData.dailyRecords) {
        if (record.comment && record.comment !== "not found" && record.comment.trim()) {
            try {
                logMessage(`😊 Analyzing sentiment for comment: "${record.comment.substring(0, 30)}..."`, context);
                const sentimentResult = await analyzeComment(record.comment);
                
                record.sentimentAnalysis = {
                    originalComment: sentimentResult.originalComment,
                    detectedLanguage: sentimentResult.detectedLanguage,
                    japaneseTranslation: sentimentResult.japaneseTranslation,
                    analysisLanguage: sentimentResult.analysisLanguage,
                    sentiment: sentimentResult.sentiment,
                    confidenceScores: sentimentResult.scores
                };
                
                logMessage(`✅ Sentiment: ${sentimentResult.sentiment} (${Math.round(sentimentResult.scores[sentimentResult.sentiment] * 100)}% confidence)`, context);
                
            } catch (error) {
                logMessage(`❌ Sentiment analysis failed for comment: ${error.message}`, context);
                record.sentimentAnalysis = {
                    originalComment: record.comment,
                    error: error.message,
                    sentiment: "unknown"
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

    const tableRows = structuredData.dailyRecords.map(record => `
        <tr>
            <td>${String(record.day).padStart(2, '0')}</td>
            <td>${record.Menu1Status}</td>
            <td>${record.Menu2Status}</td>
            <td>${record.Menu3Status}</td>
            <td>${record.Menu4Status}</td>
            <td>${record.Menu5Status}</td>
            <td>${record.dailyCheckStatus}</td>
            <td>${record.comment !== "not found" ? record.comment : '--'}</td>
            <td>${record.approverStatus}</td>
        </tr>
    `).join('\n');

    const sentimentRows = structuredData.dailyRecords
        .filter(record => record.sentimentAnalysis && !record.sentimentAnalysis.error)
        .map(record => {
            const sentiment = record.sentimentAnalysis;
            return `
        <tr>
            <td>${String(record.day).padStart(2, '0')}</td>
            <td>${sentiment.originalComment}</td>
            <td>${sentiment.detectedLanguage}</td>
            <td>${sentiment.japaneseTranslation}</td>
            <td>${sentiment.analysisLanguage}</td>
            <td>${sentiment.sentiment}</td>
            <td>
                👍 ${sentiment.confidenceScores.positive || 0} /
                😐 ${sentiment.confidenceScores.neutral || 0} /
                👎 ${sentiment.confidenceScores.negative || 0}
            </td>
        </tr>
            `;
        }).join('\n');

    const menuSummary = calculateMenuSummary(structuredData);

    return `
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>重要衛生管理の実施記録</title>
  <style>
    body { font-family: 'Meiryo', 'Yu Gothic', sans-serif; margin: 2em; }
    h1, h2, h3 { color: #333; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 2em; }
    th, td { border: 1px solid #aaa; padding: 0.5em; text-align: center; }
    th { background: #d0f5d8; }
    tr:nth-child(even) { background: #f9f9f9; }
    .summary { margin-bottom: 2em; }
  </style>
</head>
<body>
  <h1>重要衛生管理の実施記録</h1>
  <div class="summary">
    <strong>提出日：</strong>${fileNameParts.submissionDate}<br>
    <strong>提出者：</strong>${fileNameParts.senderEmail}<br>
    <strong>ファイル名：</strong>${fileNameParts.originalFileName}<br>
    <strong>店舗名：</strong>${structuredData.metadata.location}<br>
    <strong>年月：</strong>${structuredData.metadata.yearMonth}
  </div>
  <h3>管理記録表</h3>
  <table>
    <tr>
      <th>日付</th>
      <th>Menu 1</th>
      <th>Menu 2</th>
      <th>Menu 3</th>
      <th>Menu 4</th>
      <th>Menu 5</th>
      <th>日常点検</th>
      <th>特記事項</th>
      <th>確認者</th>
    </tr>
    ${tableRows}
  </table>
  <h3>サマリー</h3>
  <ul>
    <li>記録日数：${structuredData.summary.recordedDays}日</li>
    <li>全項目「良」達成日数：${menuSummary.allGoodDays}日</li>
    <li>「否」あり日数：${menuSummary.anyNgDays}日</li>
    <li>コメント記入日数：${structuredData.summary.daysWithComments}日</li>
    <li>日常点検完了日数：${structuredData.summary.dailyCheckCompletedDays}日</li>
    <li>メニューアイテムごとの「否」回数：
      <ul>
        ${menuSummary.ngCounts.map((count, index) => 
          `<li>Menu ${index + 1}: ${count}回</li>`
        ).join('\n')}
      </ul>
    </li>
  </ul>
  
    <h3>センチメント分析レポート</h3>
    <table>
      <tr>
        <th>日付</th>
        <th>コメント（原文）</th>
        <th>検出言語</th>
        <th>日本語訳</th>
        <th>分析言語</th>
        <th>センチメント</th>
        <th>スコア</th>
      </tr>
      ${sentimentRows}
    </table>
  
  <h3>重要管理項目</h3>
  <table>
    <tr>
      <th>メニュー</th>
      <th>説明</th>
    </tr>
    ${structuredData.menuItems.map((item, index) => `
      <tr>
        <td>Menu ${index + 1}</td>
        <td>${item.menuName}</td>
      </tr>
    `).join('\n')}
  </table>
  <div style="margin-top:2em;">
    このレポートは HygienMaster システムにより自動生成されました<br>
    生成日時: ${new Date().toISOString()}
  </div>
</body>
</html>`;
}

function calculateMenuSummary(structuredData) {
    const ngCounts = [0, 0, 0, 0, 0]; // Menu1-Menu5
    let allGoodDays = 0;
    let anyNgDays = 0;

    structuredData.dailyRecords.forEach(record => {
        const statuses = [
            record.Menu1Status,
            record.Menu2Status,
            record.Menu3Status,
            record.Menu4Status,
            record.Menu5Status
        ];

        let allGood = true;
        let hasNg = false;

        statuses.forEach((status, index) => {
            if (status === "否") {
                ngCounts[index]++;
                hasNg = true;
                allGood = false;
            } else if (status !== "良") {
                allGood = false;
            }
        });

        // Also check daily check status
        if (record.dailyCheckStatus !== "良") {
            allGood = false;
        }

        if (allGood) allGoodDays++;
        if (hasNg) anyNgDays++;
    });

    return {
        ngCounts,
        allGoodDays,
        anyNgDays
    };
}

function generateSentimentSummary(dailyRecords) {
    const summary = {
        positive: 0,
        negative: 0,
        neutral: 0,
        errors: 0
    };

    dailyRecords.forEach(record => {
        if (record.sentimentAnalysis) {
            if (record.sentimentAnalysis.error) {
                summary.errors++;
            } else {
                switch (record.sentimentAnalysis.sentiment) {
                    case 'positive':
                        summary.positive++;
                        break;
                    case 'negative':
                        summary.negative++;
                        break;
                    case 'neutral':
                        summary.neutral++;
                        break;
                    default:
                        summary.errors++;
                }
            }
        }
    });

    return summary;
}

function parseFileName(fileName, context) {
    logMessage(`🔍 Parsing filename: ${fileName}`, context);
    
    try {
        let submissionTime = '';
        let senderEmail = '';
        let originalFileName = fileName;
        
        const emailMatch = fileName.match(/\(([^)]*@[^)]*)\)/);
        if (emailMatch) {
            senderEmail = emailMatch[1];
            const emailEndIndex = fileName.indexOf(emailMatch[0]) + emailMatch[0].length;
            originalFileName = fileName.substring(emailEndIndex).replace(/^\W+/, '').trim();
        }
        
        const timeMatch = fileName.match(/^([^(]+)/);
        if (timeMatch) {
            submissionTime = timeMatch[1];
            if (submissionTime.includes('T')) {
                try {
                    const cleanTime = submissionTime.replace(/[^\d]/g, '');
                    if (cleanTime.length >= 8) {
                        const year = cleanTime.substring(0, 4);
                        const month = cleanTime.substring(4, 6);
                        const day = cleanTime.substring(6, 8);
                        const hour = cleanTime.substring(8, 10) || '00';
                        const minute = cleanTime.substring(10, 12) || '00';
                        
                        const isoString = `${year}-${month}-${day}T${hour}:${minute}:00`;
                        const date = new Date(isoString);
                        
                        if (!isNaN(date.getTime())) {
                            submissionTime = date.toLocaleDateString('ja-JP', {
                                year: 'numeric',
                                month: '2-digit',
                                day: '2-digit',
                                hour: '2-digit',
                                minute: '2-digit'
                            });
                        }
                    }
                } catch (e) {
                    logMessage(`⚠️ Date parsing failed: ${e.message}`, context);
                }
            }
        }
        
        return {
            submissionDate: submissionTime || 'Unknown',
            senderEmail: senderEmail || 'Unknown',
            originalFileName: originalFileName || fileName
        };
        
    } catch (error) {
        logMessage(`❌ Filename parsing error: ${error.message}`, context);
        return {
            submissionDate: 'Unknown',
            senderEmail: 'Unknown',
            originalFileName: fileName
        };
    }
}

module.exports = {
    prepareImportantManagementReport
};