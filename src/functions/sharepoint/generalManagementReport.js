const { logMessage, handleError, convertHeicToJpegIfNeeded } = require('../utils');
const {
  uploadJsonToSharePoint,
  uploadTextToSharePoint,
  uploadOriginalDocumentToSharePoint,
  ensureSharePointFolder,
  uploadHtmlToSharePoint
} = require('./sendToSharePoint');
const { analyzeComment, getLanguageNameInJapanese, formatInlineConfidenceDetails, supportedLanguages } = require('../analytics/sentimentAnalysis');
const axios = require('axios');
const { getReportStyles, getReportScripts } = require('./styles/sharedStyles');

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
                
                // Check if the analysis was successful
                if (sentimentResult.error) {
                    logMessage(`❌ Sentiment analysis returned error: ${sentimentResult.error}`, context);
                    record.sentimentAnalysis = {
                        originalComment: sentimentResult.originalComment,
                        detectedLanguage: sentimentResult.detectedLanguage,
                        japaneseTranslation: sentimentResult.japaneseTranslation,
                        analysisLanguage: sentimentResult.analysisLanguage,
                        sentiment: sentimentResult.sentiment,
                        confidenceScores: sentimentResult.confidenceScores,
                        wasTranslated: sentimentResult.wasTranslated,
                        error: sentimentResult.error
                    };
                    continue;
                }
                
                // Add sentiment data to the record - use the exact structure returned by analyzeComment
                record.sentimentAnalysis = {
                    originalComment: sentimentResult.originalComment,
                    detectedLanguage: sentimentResult.detectedLanguage,
                    japaneseTranslation: sentimentResult.japaneseTranslation,
                    analysisLanguage: sentimentResult.analysisLanguage,
                    sentiment: sentimentResult.sentiment,
                    confidenceScores: sentimentResult.confidenceScores,  // ✅ FIXED: was sentimentResult.scores
                    wasTranslated: sentimentResult.wasTranslated  // ✅ FIXED: use the value from analyzeComment
                };
                
                const confidence = Math.round((sentimentResult.confidenceScores[sentimentResult.sentiment] || 0) * 100);
                logMessage(`✅ Sentiment: ${sentimentResult.sentiment} (${confidence}% confidence) - Language: ${sentimentResult.detectedLanguage} -> Analysis: ${sentimentResult.analysisLanguage}`, context);
                
            } catch (error) {
                logMessage(`❌ Sentiment analysis failed for comment: ${error.message}`, context);
                record.sentimentAnalysis = {
                    originalComment: record.comment,
                    detectedLanguage: 'unknown',
                    japaneseTranslation: null,
                    analysisLanguage: 'unknown',
                    sentiment: 'unknown',
                    confidenceScores: { positive: 0, neutral: 0, negative: 0 },
                    wasTranslated: false,
                    error: error.message
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
            <td><span class="status-badge ${statusClass(record.Cat1Status)}">${record.Cat1Status}</span></td>
            <td><span class="status-badge ${statusClass(record.Cat2Status)}">${record.Cat2Status}</span></td>
            <td><span class="status-badge ${statusClass(record.Cat3Status)}">${record.Cat3Status}</span></td>
            <td><span class="status-badge ${statusClass(record.Cat4Status)}">${record.Cat4Status}</span></td>
            <td><span class="status-badge ${statusClass(record.Cat5Status)}">${record.Cat5Status}</span></td>
            <td><span class="status-badge ${statusClass(record.Cat6Status)}">${record.Cat6Status}</span></td>
            <td><span class="status-badge ${statusClass(record.Cat7Status)}">${record.Cat7Status}</span></td>
            <td class="comment-cell">${record.comment !== "not found" ? record.comment : '--'}</td>
            <td><span class="status-badge ${statusClass(record.approverStatus)}">${record.approverStatus}</span></td>
        </tr>
        `;
    }).join('\n');

    const sentimentRows = structuredData.dailyRecords
        .map(record => {
            const day = String(record.day).padStart(2, '0');
            
            // Check if sentiment analysis exists and was successful
            if (record.sentimentAnalysis && !record.sentimentAnalysis.error) {
                const sentiment = record.sentimentAnalysis;
                const sentimentClass = `sentiment-${sentiment.sentiment}`;
                const confidence = Math.round((sentiment.confidenceScores[sentiment.sentiment] || 0) * 100);
                const inlineDetails = formatInlineConfidenceDetails(sentiment.confidenceScores);
                
                // NEW LOGIC: Always show Japanese text in translation column
                let translationDisplay;
                if (sentiment.detectedLanguage === 'ja') {
                    // If original is Japanese, show the same text
                    translationDisplay = sentiment.originalComment;
                } else {
                    // If original is not Japanese, show translation (or original if translation failed)
                    translationDisplay = sentiment.japaneseTranslation || sentiment.originalComment;
                }
                
                return `
        <tr class="sentiment-row">
            <td class="date-cell">${day}</td>
            <td class="comment-text">${sentiment.originalComment}</td>
            <td class="language-tag">
                <span class="language-badge">${getLanguageNameInJapanese(sentiment.detectedLanguage)}</span>
            </td>
            <td class="translation-text">${translationDisplay}</td>
            <td class="language-tag">
                <span class="language-badge">${getLanguageNameInJapanese(sentiment.analysisLanguage)}</span>
            </td>
            <td><span class="sentiment-badge ${sentimentClass}">${getSentimentIcon(sentiment.sentiment)} ${sentiment.sentiment}</span></td>
            <td class="confidence-cell">
                <div class="confidence-container">
                    <div class="confidence-fill ${sentimentClass}" style="width: ${confidence}%"></div>
                    <span class="confidence-text">${confidence}%</span>
                </div>
                ${inlineDetails}
            </td>
        </tr>`;
            } else {
                // Show why sentiment analysis wasn't performed
                let reason = '';
                if (!record.comment || record.comment === "not found") {
                    reason = 'コメントなし';
                } else if (record.sentimentAnalysis && record.sentimentAnalysis.error) {
                    reason = '分析エラー';
                } else {
                    reason = '未分析';
                }
                
                return `
        <tr class="sentiment-row no-analysis">
            <td class="date-cell">${day}</td>
            <td class="comment-text">${record.comment !== "not found" ? record.comment : '--'}</td>
            <td colspan="5" class="no-analysis-reason">${reason}</td>
        </tr>`;
            }
        }).join('\n');

    const sentimentSummary = generateSentimentSummary(structuredData.dailyRecords);
    const complianceRate = Math.round((sentimentSummary.positive + sentimentSummary.neutral + sentimentSummary.negative) / structuredData.summary.recordedDays * 100);

    // Count analysis results for summary info
    const totalDaysWithComments = structuredData.dailyRecords.filter(r => r.comment && r.comment !== "not found" && r.comment.trim()).length;
    const successfulAnalyses = structuredData.dailyRecords.filter(r => r.sentimentAnalysis && !r.sentimentAnalysis.error).length;
    const failedAnalyses = totalDaysWithComments - successfulAnalyses;

    return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>一般衛生管理レポート - ${structuredData.metadata.location}</title>
    <style>
        ${getReportStyles('general')}
        
        /* Dynamic CSS variables for compliance rates */
        :root {
            --compliance-color: ${complianceRate >= 80 ? '#27ae60' : complianceRate >= 60 ? '#f39c12' : '#e74c3c'};
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Professional Header -->
        <header class="header">
            <h1>一般衛生管理レポート</h1>
            <div class="subtitle">${structuredData.metadata.location} | ${structuredData.metadata.yearMonth}</div>
        </header>

        <!-- Executive Summary Cards -->
        <div class="summary-cards">
            <div class="summary-card compliance">
                <div class="card-header">
                    <div class="card-icon">📊</div>
                    <div class="card-title">コンプライアンス率</div>
                </div>
                <div class="card-value">${complianceRate}%</div>
                <div class="card-description">全項目良好: ${sentimentSummary.positive + sentimentSummary.neutral + sentimentSummary.negative}/${structuredData.summary.recordedDays}日</div>
                <div class="progress-bar">
                    <div class="progress-fill ${complianceRate >= 80 ? '' : complianceRate >= 60 ? 'warning' : 'danger'}" 
                         style="width: ${complianceRate}%"></div>
                </div>
            </div>

            <div class="summary-card comments">
                <div class="card-header">
                    <div class="card-icon">💬</div>
                    <div class="card-title">コメント記入率</div>
                </div>
                <div class="card-value">${Math.round((structuredData.summary.daysWithComments / structuredData.summary.recordedDays) * 100)}%</div>
                <div class="card-description">${structuredData.summary.daysWithComments}/${structuredData.summary.recordedDays}日でコメント記入</div>
            </div>

            <div class="summary-card sentiment">
                <div class="card-header">
                    <div class="card-icon">😊</div>
                    <div class="card-title">感情分析</div>
                </div>
                <div class="card-value">${sentimentSummary.positive + sentimentSummary.neutral + sentimentSummary.negative}</div>
                <div class="card-description">
                    👍${sentimentSummary.positive} 😐${sentimentSummary.neutral} 👎${sentimentSummary.negative}
                </div>
            </div>
        </div>

        <!-- Submission Information -->
        <div class="section">
            <div class="section-header">
                <h3>📋 提出情報</h3>
            </div>
            <div class="section-content">
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px;">
                    <div><strong>提出日時:</strong> ${fileNameParts.submissionDate}</div>
                    <div><strong>提出者:</strong> ${fileNameParts.senderEmail}</div>
                    <div><strong>ファイル名:</strong> ${fileNameParts.originalFileName}</div>
                    <div><strong>店舗名:</strong> ${structuredData.metadata.location}</div>
                </div>
            </div>
        </div>

        <!-- Daily Records Table with Integrated Category Reference -->
        <div class="section">
            <div class="section-header">
                <h3>📅 日次管理記録</h3>
            </div>
            <div class="section-content">
                <!-- Category Reference (moved here) -->
                <div class="category-reference">
                    <h4 style="margin-bottom: 15px; color: #2c3e50;">📚 管理カテゴリ定義</h4>
                    <div class="category-grid">
                        ${structuredData.categories.map((cat, index) => `
                        <div class="category-item">
                            <strong>Cat ${index + 1}:</strong> ${cat.categoryName}
                        </div>
                        `).join('')}
                    </div>
                </div>

                <!-- Daily Records Table -->
                <table style="margin-top: 25px;">
                    <thead>
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
                    </thead>
                    <tbody>
                        ${tableRows}
                    </tbody>
                </table>
            </div>
        </div>

        <!-- Sentiment Analysis Section -->
        ${sentimentRows ? `
        <div class="section">
            <div class="section-header">
                <h3>🧠 感情分析詳細レポート</h3>
                <div class="section-description">
                    各コメントの感情分析結果と詳細スコアを表示します
                </div>
            </div>
            <div class="section-content">
                <div class="sentiment-summary">
                    <strong>📊 感情分析結果:</strong> ${totalDaysWithComments}件のコメント中 ${successfulAnalyses}件分析成功${failedAnalyses > 0 ? `、${failedAnalyses}件失敗` : ''}
                    <div class="hint-text">
                        💡 ヒント: 信頼度欄で全感情カテゴリのスコアを確認できます
                    </div>
                </div>
                <table>
                    <thead>
                        <tr>
                            <th>日付</th>
                            <th>コメント（原文）</th>
                            <th>検出言語</th>
                            <th>日本語訳</th>
                            <th>分析言語</th>
                            <th>感情判定</th>
                            <th>信頼度・詳細スコア</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${sentimentRows}
                    </tbody>
                </table>
            </div>
        </div>
        ` : ''}

        <!-- Footer -->
        <footer class="footer">
            <div>このレポートは <strong>HygienMaster システム</strong> により自動生成されました</div>
            <div class="timestamp">生成日時: ${new Date().toLocaleString('ja-JP')}</div>
        </footer>
    </div>

    <script>
        ${getReportScripts()}
    </script>
</body>
</html>`;
}

function getSentimentIcon(sentiment) {
    switch (sentiment) {
        case 'positive': return '😊';
        case 'negative': return '😞';
        case 'neutral': return '😐';
        default: return '❓';
    }
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