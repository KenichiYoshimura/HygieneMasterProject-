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
                let reason = '分析エラー';
                if (record.sentimentAnalysis && record.sentimentAnalysis.error) {
                    reason = `分析エラー: ${record.sentimentAnalysis.error}`;
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
    <title>重要衛生管理レポート - ${structuredData.metadata.location}</title>
    <style>
        ${getReportStyles('important')}
        
        /* Dynamic CSS variables for compliance rates */
        :root {
            --compliance-color: ${complianceRate >= 80 ? '#27ae60' : complianceRate >= 60 ? '#f39c12' : '#e74c3c'};
            --daily-check-color: ${dailyCheckRate >= 80 ? '#27ae60' : dailyCheckRate >= 60 ? '#f39c12' : '#e74c3c'};
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Professional Header -->
        <header class="header">
            <h1>重要衛生管理レポート</h1>
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
                <div class="card-description">全項目良好: ${menuSummary.allGoodDays}/${structuredData.summary.recordedDays}日</div>
                <div class="progress-bar">
                    <div class="progress-fill ${complianceRate >= 80 ? '' : complianceRate >= 60 ? 'warning' : 'danger'}" 
                         style="width: ${complianceRate}%"></div>
                </div>
            </div>

            <div class="summary-card daily-check">
                <div class="card-header">
                    <div class="card-icon">✅</div>
                    <div class="card-title">日常点検完了率</div>
                </div>
                <div class="card-value">${dailyCheckRate}%</div>
                <div class="card-description">${structuredData.summary.dailyCheckCompletedDays}/${structuredData.summary.recordedDays}日で完了</div>
                <div class="progress-bar">
                    <div class="progress-fill ${dailyCheckRate >= 80 ? '' : dailyCheckRate >= 60 ? 'warning' : 'danger'}" 
                         style="width: ${dailyCheckRate}%"></div>
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

        <!-- Daily Records Table -->
        <div class="section">
            <div class="section-header">
                <h3>📅 日次管理記録</h3>
            </div>
            <div class="section-content">
                <table>
                    <thead>
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
            </div>
            <div class="section-content">
                <div class="sentiment-summary">
                    <strong>📊 感情分析結果:</strong> ${totalDaysWithComments}件のコメント中 ${successfulAnalyses}件分析成功${failedAnalyses > 0 ? `、${failedAnalyses}件失敗` : ''}
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
                            <th>信頼度</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${sentimentRows}
                    </tbody>
                </table>
            </div>
        </div>
        ` : ''}

        <!-- Menu Item Reference -->
        <div class="section">
            <div class="section-header">
                <h3>🍽️ 重要管理項目定義</h3>
            </div>
            <div class="section-content">
                <table>
                    <thead>
                        <tr>
                            <th>メニュー</th>
                            <th>管理項目</th>
                            <th>NG回数</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${structuredData.menuItems.map((item, index) => `
                        <tr>
                            <td><strong>Menu ${index + 1}</strong></td>
                            <td style="text-align: left;">${item.menuName}</td>
                            <td>
                                <span class="status-badge ${menuSummary.ngCounts && menuSummary.ngCounts[index] > 0 ? 'status-bad' : 'status-good'}">
                                    ${menuSummary.ngCounts ? menuSummary.ngCounts[index] : 0}回
                                </span>
                            </td>
                        </tr>
                        `).join('\n')}
                    </tbody>
                </table>
            </div>
        </div>

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

// Add missing helper functions
function getSentimentIcon(sentiment) {
    switch (sentiment) {
        case 'positive': return '😊';
        case 'negative': return '😞';
        case 'neutral': return '😐';
        default: return '❓';
    }
}

function calculateMenuSummary(structuredData) {
    const summary = {
        allGoodDays: 0,
        someIssuesDays: 0,
        noRecordsDays: 0,
        totalDays: structuredData.summary.totalDays,
        ngCounts: [0, 0, 0, 0, 0] // Initialize for 5 menu items
    };
    
    structuredData.dailyRecords.forEach(record => {
        const menuStatuses = [
            record.Menu1Status, 
            record.Menu2Status, 
            record.Menu3Status, 
            record.Menu4Status, 
            record.Menu5Status
        ];
        
        // Count NG occurrences for each menu
        menuStatuses.forEach((status, index) => {
            if (status === '否') {
                summary.ngCounts[index]++;
            }
        });
        
        const allGood = menuStatuses.every(status => status === '良');
        const someIssues = menuStatuses.some(status => status === '否' || status === '無');
        
        if