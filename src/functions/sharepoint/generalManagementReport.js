const { logMessage, handleError, convertHeicToJpegIfNeeded } = require('../utils');
const {
  uploadJsonToSharePoint,
  uploadTextToSharePoint,
  uploadOriginalDocumentToSharePoint,
  ensureSharePointFolder,
  uploadHtmlToSharePoint
} = require('./sendToSharePoint');
const { analyzeComment } = require('../analytics/sentimentAnalysis'); // Added sentiment analysis
const axios = require('axios');

/**
 * Prepares general management reports from structured data and uploads to SharePoint
 * 
 * @param {Object} structuredData - New structured data format from generalManagementFormExtractor
 * @param {Object} context - Azure Functions execution context
 * @param {string} base64BinFile - Base64 encoded original file
 * @param {string} originalFileName - Original filename for submission info
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
 * Modifies the structuredData object in place
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
        
        const basePath = process.env.SHAREPOINT_FOLDER_PATH?.replace(/^\/+|\/+$/g, '') || 'Form_Data';
        const folderPath = `${basePath}/衛生管理日誌/${year}/${month}/${location}`;
        
        logMessage(`📁 Target SharePoint folder: ${folderPath}`, context);
        await ensureSharePointFolder(folderPath, context);

        const jsonFileName = `general-report-${baseFileName}-${timestamp}.json`;
        const textFileName = `general-report-${baseFileName}-${timestamp}.txt`;
        const htmlFileName = `general-report-${baseFileName}-${timestamp}.html`;
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
        title: "一般管理の実施記録",
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
            "日付", "Cat 1", "Cat 2", "Cat 3", "Cat 4", "Cat 5", "Cat 6", "Cat 7", "特記事項", "感情分析", "確認者"
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
            感情分析: record.sentimentAnalysis || null,
            確認者: record.approverStatus
        })),
        
        summary: {
            totalDays: structuredData.summary.totalDays,
            recordedDays: structuredData.summary.recordedDays,
            daysWithComments: structuredData.summary.daysWithComments,
            approvedDays: structuredData.summary.approvedDays,
            // Add sentiment analysis summary
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
一般管理の実施記録
提出日：${fileNameParts.submissionDate}
提出者：${fileNameParts.senderEmail}  
ファイル名：${fileNameParts.originalFileName}

店舗名：${structuredData.metadata.location}
年月：${structuredData.metadata.yearMonth}

管理カテゴリ：
`;

    // Add category descriptions
    structuredData.categories.forEach((category, index) => {
        textReport += `Cat ${index + 1}: ${category.categoryName}\n`;
    });

    textReport += '\n';

    // Create table header
    const headerRow = `日付 | Cat 1 | Cat 2 | Cat 3 | Cat 4 | Cat 5 | Cat 6 | Cat 7 | 特記事項 | 感情 | 確認者`;
    textReport += headerRow + '\n';
    textReport += ''.padEnd(headerRow.length, '-') + '\n';

    // Add data rows
    structuredData.dailyRecords.forEach(record => {
        const sentiment = record.sentimentAnalysis?.sentiment || '--';
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
            sentiment.padEnd(4),
            record.approverStatus
        ].join('| ');
        
        textReport += dataRow + '\n';
    });

    // Add sentiment analysis section
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
        const sentimentEmoji = getSentimentEmoji(record.sentimentAnalysis?.sentiment);
        const sentimentTitle = record.sentimentAnalysis ? 
            `${record.sentimentAnalysis.sentiment} (${Math.round(record.sentimentAnalysis.confidenceScores[record.sentimentAnalysis.sentiment] * 100)}%)` : 
            '';
        
        return `
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
            <td title="${sentimentTitle}">${sentimentEmoji}</td>
            <td>${record.approverStatus}</td>
        </tr>
        `;
    }).join('\n');

    const sentimentSummary = generateSentimentSummary(structuredData.dailyRecords);

    return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <title>一般管理の実施記録</title>
    <style>
        body { 
            font-family: 'Meiryo', 'Yu Gothic', sans-serif; 
            margin: 2em; 
            background-color: #fff; 
        }
        h1, h2, h3 { 
            color: #333; 
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin-bottom: 2em; 
        }
        th, td { 
            border: 1px solid #aaa; 
            padding: 0.5em; 
            text-align: center; 
        }
        th { 
            background: #e1f5fe; 
        }
        tr:nth-child(even) { 
            background: #f9f9f9; 
        }
        .section-box {
            border-left: 6px solid #1976d2;
            background-color: #f5f5f5;
            padding: 1em;
            margin-bottom: 2em;
        }
        .sentiment-positive { color: #4caf50; font-weight: bold; }
        .sentiment-negative { color: #f44336; font-weight: bold; }
        .sentiment-neutral { color: #9e9e9e; }
    </style>
</head>
<body>
    <h1>一般管理の実施記録</h1>
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
            <th>感情</th>
            <th>確認者</th>
        </tr>
        ${tableRows}
    </table>

    <div class="section-box">
        <h3>管理カテゴリの各項目</h3>
        <ul>
            ${structuredData.categories.map((cat, idx) => `<li>Cat ${idx + 1}: ${cat.categoryName}</li>`).join('\n')}
        </ul>
    </div>

    <div class="section-box">
        <h3>サマリー</h3>
        <ul>
            <li>記録日数：${structuredData.summary.recordedDays}日</li>
            <li>コメント記入日数：${structuredData.summary.daysWithComments}日</li>
            <li>承認済み日数：${structuredData.summary.approvedDays}日</li>
        </ul>
        <h4>感情分析サマリー</h4>
        <ul>
            <li class="sentiment-positive">😊 ポジティブ: ${sentimentSummary.positive}件</li>
            <li class="sentiment-negative">😞 ネガティブ: ${sentimentSummary.negative}件</li>
            <li class="sentiment-neutral">😐 ニュートラル: ${sentimentSummary.neutral}件</li>
            <li>❓ 分析エラー: ${sentimentSummary.errors}件</li>
        </ul>
    </div>

    <div style="margin-top:2em;">
        このレポートは HygienMaster システムにより自動生成されました<br>
        生成日時: ${new Date().toISOString()}
    </div>
</body>
</html>`;
}

/**
 * Generates sentiment analysis summary from daily records
 */
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

/**
 * Returns emoji representation of sentiment
 */
function getSentimentEmoji(sentiment) {
    switch (sentiment) {
        case 'positive':
            return '😊';
        case 'negative':
            return '😞';
        case 'neutral':
            return '😐';
        default:
            return '--';
    }
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