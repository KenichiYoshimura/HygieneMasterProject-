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
 * Prepares general management reports from structured data and uploads to SharePoint
 */
async function prepareGeneralManagementReport(structuredData, context, base64BinFile, originalFileName) {
    logMessage("🚀 prepareGeneralManagementReport() called with structured data", context);
    
    try {
        logMessage("📊 Processing structured data:", context);
        logMessage(`  - Store: ${structuredData.metadata.location}`, context);
        logMessage(`  - Year-Month: ${structuredData.metadata.yearMonth}`, context);
        logMessage(`  - Daily Records: ${structuredData.dailyRecords.length}`, context);
        logMessage(`  - Categories: ${structuredData.categories.length}`, context);
        
        // Add sentiment analysis to structured data
        logMessage("🧠 Starting sentiment analysis for comments...", context);
        await addSentimentAnalysisToStructuredData(structuredData, context);
        logMessage("✅ Sentiment analysis completed", context);
        
        // Generate reports using structured data (now with sentiment analysis)
        const jsonReport = generateJsonReport(structuredData, originalFileName, context);
        logMessage("✅ JSON report generated", context);

        const textReport = generateTextReport(structuredData, originalFileName, context);
        logMessage("✅ Text report generated", context);

        const htmlReport = generateHtmlReport(structuredData, originalFileName, context);
        logMessage("✅ HTML report generated", context);

        // Upload to SharePoint
        logMessage("📤 Starting SharePoint upload...", context);
        await uploadReportsToSharePoint(jsonReport, textReport, htmlReport, base64BinFile, originalFileName, structuredData, context);
        logMessage("✅ SharePoint upload completed", context);

        return {
            json: jsonReport,
            text: textReport,
            html: htmlReport
        };
        
    } catch (error) {
        handleError(error, 'General Management Report Generation', context);
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
                
                // Add sentiment data to the record
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
        
        // Use form data for folder structure
        const location = structuredData.metadata.location;
        const [year, month] = structuredData.metadata.yearMonth.split('-');
        
        logMessage(`📋 Using structured data for folder: ${location}, ${year}-${month}`, context);
        
        const basePath = process.env.SHAREPOINT_FOLDER_PATH?.replace(/^\/+|\/+$/g, '') || '衛生管理日誌';
        const folderPath = `${basePath}/一般衛生管理の実施記録/${year}/${month}/${location}`;
        
        logMessage(`📁 Target SharePoint folder: ${folderPath}`, context);
        await ensureSharePointFolder(folderPath, context);

        // Use Japanese naming convention like legacy format
        const jsonFileName = `一般衛生管理レポート-${baseFileName}-${timestamp}.json`;
        const textFileName = `一般衛生管理レポート-${baseFileName}-${timestamp}.txt`;
        const htmlFileName = `一般衛生管理レポート-${baseFileName}-${timestamp}.html`;
        const originalDocFileName = `original-${originalFileName}`;

        await uploadJsonToSharePoint(jsonReport, jsonFileName, folderPath, context);
        await uploadTextToSharePoint(textReport, textFileName, folderPath, context);
        await uploadOriginalDocumentToSharePoint(base64BinFile, originalDocFileName, folderPath, context);
        await uploadHtmlToSharePoint(htmlReport, htmlFileName, folderPath, context);

        logMessage("✅ All general management reports uploaded to SharePoint successfully", context);
        
    } catch (error) {
        logMessage(`❌ SharePoint upload process failed: ${error.message}`, context);
        handleError(error, 'SharePoint Upload', context);
        throw error;
    }
}

function generateJsonReport(structuredData, originalFileName, context) {
    const fileNameParts = parseFileName(originalFileName, context);
    
    const reportData = {
        title: "一般衛生管理の実施記録",
        submissionDate: fileNameParts.submissionDate,
        submitter: fileNameParts.senderEmail,
        originalFileName: fileNameParts.originalFileName,
        storeName: structuredData.metadata.location,
        yearMonth: structuredData.metadata.yearMonth,
        
        categories: structuredData.categories.map((cat, index) => ({
            id: `Cat ${index + 1}`,
            name: cat.categoryName
        })),
        
        tableHeaders: [
            "日付", "Cat 1", "Cat 2", "Cat 3", "Cat 4", "Cat 5", "Cat 6", "Cat 7", "特記事項", "確認者"
        ],
        
        dailyData: structuredData.dailyRecords.map(record => ({
            日付: String(record.day).padStart(2, '0'),
            "Cat 1": record.Cat1Status,
            "Cat 2": record.Cat2Status,
            "Cat 3": record.Cat3Status,
            "Cat 4": record.Cat4Status,
            "Cat 5": record.Cat5Status,
            "Cat 6": record.Cat6Status,
            "Cat 7": record.Cat7Status,
            特記事項: record.comment !== "not found" ? record.comment : "--",
            確認者: record.approverStatus
        })),
        
        summary: {
            totalDays: structuredData.summary.totalDays,
            recordedDays: structuredData.summary.recordedDays,
            daysWithComments: structuredData.summary.daysWithComments,
            approvedDays: structuredData.summary.approvedDays,
            sentimentSummary: generateSentimentSummary(structuredData.dailyRecords)
        },
        
        footer: {
            generatedBy: "HygienMaster システム",
            generatedAt: new Date().toISOString(),
            note: "このレポートは HygienMaster システムにより自動生成されました"
        }
    };
    
    return reportData;
}

function generateTextReport(structuredData, originalFileName, context) {
    const fileNameParts = parseFileName(originalFileName, context);
    
    let textReport = `
一般衛生管理の実施記録
提出日：${fileNameParts.submissionDate}
提出者：${fileNameParts.senderEmail}  
ファイル名：${fileNameParts.originalFileName}

店舗名：${structuredData.metadata.location}
年月：${structuredData.metadata.yearMonth}

管理カテゴリ：
`;

    structuredData.categories.forEach((category, index) => {
        textReport += `Cat ${index + 1}: ${category.categoryName}\n`;
    });

    textReport += '\n';

    const headerRow = `日付 | Cat 1 | Cat 2 | Cat 3 | Cat 4 | Cat 5 | Cat 6 | Cat 7 | 特記事項 | 確認者`;
    textReport += headerRow + '\n';
    textReport += ''.padEnd(headerRow.length, '-') + '\n';

    structuredData.dailyRecords.forEach(record => {
        const dataRow = [
            String(record.day).padStart(2, '0').padEnd(4),
            record.Cat1Status.padEnd(6),
            record.Cat2Status.padEnd(6),
            record.Cat3Status.padEnd(6),
            record.Cat4Status.padEnd(6),
            record.Cat5Status.padEnd(6),
            record.Cat6Status.padEnd(6),
            record.Cat7Status.padEnd(6),
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
            <td>${record.Cat1Status}</td>
            <td>${record.Cat2Status}</td>
            <td>${record.Cat3Status}</td>
            <td>${record.Cat4Status}</td>
            <td>${record.Cat5Status}</td>
            <td>${record.Cat6Status}</td>
            <td>${record.Cat7Status}</td>
            <td>${record.comment !== "not found" ? record.comment : '--'}</td>
            <td>${record.approverStatus}</td>
        </tr>
    `).join('\n');

    // Generate detailed sentiment analysis table
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

    // Calculate category summary
    const categorySummary = calculateCategorySummary(structuredData);

    return `
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <title>一般衛生管理の実施記録</title>
  <style>
    body { font-family: 'Meiryo', 'Yu Gothic', sans-serif; margin: 2em; }
    h1, h2, h3 { color: #333; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 2em; }
    th, td { border: 1px solid #aaa; padding: 0.5em; text-align: center; }
    th { background: #e0f7fa; }
    tr:nth-child(even) { background: #f9f9f9; }
    .summary { margin-bottom: 2em; }
  </style>
</head>
<body>
  <h1>一般衛生管理の実施記録</h1>
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
      <th>Cat 1</th>
      <th>Cat 2</th>
      <th>Cat 3</th>
      <th>Cat 4</th>
      <th>Cat 5</th>
      <th>Cat 6</th>
      <th>Cat 7</th>
      <th>特記事項</th>
      <th>確認者</th>
    </tr>
    ${tableRows}
  </table>
  <h3>サマリー</h3>
  <ul>
    <li>記録日数：${structuredData.summary.recordedDays}日</li>
    <li>全項目「良」達成日数：${categorySummary.allGoodDays}日</li>
    <li>「否」あり日数：${categorySummary.anyNgDays}日</li>
    <li>コメント記入日数：${structuredData.summary.daysWithComments}日</li>
    <li>カテゴリごとの「否」回数：
      <ul>
        ${categorySummary.ngCounts.map((count, index) => 
          `<li>Cat ${index + 1}: ${count}回</li>`
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
  
  <h3>管理カテゴリ</h3>
  <table>
    <tr>
      <th>カテゴリ</th>
      <th>説明</th>
    </tr>
    ${structuredData.categories.map((cat, index) => `
      <tr>
        <td>Cat ${index + 1}</td>
        <td>${cat.categoryName}</td>
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

function calculateCategorySummary(structuredData) {
    const ngCounts = [0, 0, 0, 0, 0, 0, 0]; // Cat1-Cat7
    let allGoodDays = 0;
    let anyNgDays = 0;

    structuredData.dailyRecords.forEach(record => {
        const statuses = [
            record.Cat1Status,
            record.Cat2Status,
            record.Cat3Status,
            record.Cat4Status,
            record.Cat5Status,
            record.Cat6Status,
            record.Cat7Status
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
    prepareGeneralManagementReport
};