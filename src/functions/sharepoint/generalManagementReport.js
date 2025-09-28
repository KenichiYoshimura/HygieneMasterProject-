const { logMessage, handleError, convertHeicToJpegIfNeeded } = require('../utils');
const {
  uploadJsonToSharePoint,
  uploadTextToSharePoint,
  uploadOriginalDocumentToSharePoint,
  ensureSharePointFolder,
  uploadHtmlToSharePoint // <-- Add this if you want to upload HTML
} = require('./sendToSharePoint');
const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');
const path = require('path');
const { analyzeComment } = require('./analytics/sentimentAnalysis');

// Main function
async function prepareGeneralManagementReport(extractedRows, categories, context, base64BinFile, originalFileName) {
  logMessage("🚀 prepareGeneralManagementReport() called", context);
  try {
    // DEBUG: Print the exact structure we're receiving
    logMessage("🔍 DEBUG: Raw input analysis...", context);
    logMessage(`📊 extractedRows type: ${typeof extractedRows}`, context);
    logMessage(`📊 extractedRows length: ${Array.isArray(extractedRows) ? extractedRows.length : 'not array'}`, context);
    logMessage(`📊 extractedRows content:`, context);
    logMessage(`${JSON.stringify(extractedRows, null, 2)}`, context);
    logMessage(`📊 categories:`, context);
    logMessage(`${JSON.stringify(categories, null, 2)}`, context);
    logMessage(`📊 originalFileName: ${originalFileName}`, context);

    // Handle both array and object input
    let rowDataArray = [];
    if (Array.isArray(extractedRows)) {
      // If it's an array, extract the .row property from each item
      rowDataArray = extractedRows.map(item => {
        if (item && item.row) {
          return item.row;
        } else {
          return item; // fallback if no .row property
        }
      });
    } else if (extractedRows && typeof extractedRows === 'object') {
      rowDataArray = [extractedRows.row || extractedRows];
    } else {
      logMessage("❌ ERROR: extractedRows is neither array nor object", context);
      throw new Error("Invalid extractedRows format");
    }
    logMessage(`📊 Processed rowDataArray length: ${rowDataArray.length}`, context);
    if (rowDataArray.length > 0) {
      logMessage(`📊 First processed row:`, context);
      logMessage(`${JSON.stringify(rowDataArray[0], null, 2)}`, context);
      logMessage(`📊 Available keys: ${Object.keys(rowDataArray[0]).join(', ')}`, context);
    }

    // Generate reports
    const jsonReport = generateJsonReport(rowDataArray, categories, originalFileName, context);
    logMessage("✅ JSON report generated", context);

    const textReport = generateTextReport(rowDataArray, categories, originalFileName, context);
    logMessage("✅ Text report generated", context);

    const htmlReport = await generateHtmlReport(rowDataArray, categories, originalFileName, context);
    logMessage("✅ HTML report generated", context);

    // Upload to SharePoint (add HTML upload if needed)
    logMessage("📤 Starting SharePoint upload...", context);
    await uploadReportsToSharePoint(jsonReport, textReport, htmlReport, base64BinFile, originalFileName, rowDataArray, context);
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

// Upload function (add HTML upload)
async function uploadReportsToSharePoint(jsonReport, textReport, htmlReport, base64BinFile, originalFileName, rowDataArray, context) {
  try {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const baseFileName = originalFileName.replace(/\.[^/.]+$/, "");
    const location = rowDataArray[0]?.text_mkv0z6d || rowDataArray[0]?.store || 'unknown';
    const dateStr = rowDataArray[0]?.date4 || new Date().toISOString().split('T')[0];
    const [year, month] = dateStr.split('-');
    logMessage(`📋 Resolved location from form data: ${location}`, context);
    logMessage(`📋 Resolved year from form data: ${year}`, context);
    logMessage(`📋 Resolved month from form data: ${month}`, context);
    logMessage(`📋 Form date used for folder structure: ${dateStr}`, context);
    const basePath = process.env.SHAREPOINT_FOLDER_PATH?.replace(/^\/+|\/+$/g, '') || '衛生管理日誌';
    const folderPath = `${basePath}/一般衛生管理の実施記録/${year}/${String(month).padStart(2, '0')}/${location}`;
    logMessage(`📁 Target SharePoint folder: ${folderPath}`, context);
    await ensureSharePointFolder(folderPath, context);

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

// --- HTML Generation Function ---
async function generateHtmlReport(rowDataArray, categories, originalFileName, context) {
  // Parse filename for metadata
  const fileNameParts = parseFileName(originalFileName, context);
  const storeName = rowDataArray[0]?.text_mkv0z6d || "Unknown Store";
  const fullDate = rowDataArray[0]?.date4 || new Date().toISOString().split('T')[0];
  const yearMonth = fullDate.substring(0, 7);

  // Category descriptions
  const categoryDescriptions = categories && categories.length > 0
    ? categories
    : [
      '原材料の受入の確認',
      '庫内温度の確認 冷蔵庫・冷凍庫(℃)',
      '交差汚染・二次汚染の防止',
      '器具等の洗浄・消毒・殺菌',
      'トイレの洗浄・消毒',
      '従業員の健康管理等',
      '手洗いの実施'
    ];

  // Table rows
  const tableRows = rowDataArray.map(row => `
    <tr>
      <td>${row.date4 ? row.date4.split('-')[2] : '--'}</td>
      <td>${row.color_mkv02tqg || '--'}</td>
      <td>${row.color_mkv0yb6g || '--'}</td>
      <td>${row.color_mkv06e9z || '--'}</td>
      <td>${row.color_mkv0x9mr || '--'}</td>
      <td>${row.color_mkv0df43 || '--'}</td>
      <td>${row.color_mkv5fa8m || '--'}</td>
      <td>${row.color_mkv59ent || '--'}</td>
      <td>${row.text_mkv0etfg || '--'}</td>
      <td>${row.color_mkv0xnn4 || '--'}</td>
    </tr>
  `).join('\n');

  // 1. Collects comments for sentiment analysis
  const commentRows = rowDataArray
  .filter(row => row.text_mkv0etfg && row.text_mkv0etfg !== 'not found')
  .map(row => ({
    date: row.date4 ? row.date4.split('-')[2] : '--',
    comment: row.text_mkv0etfg
  }));


  // 2. Perform sentiment analysis on comments
  const sentimentResults = await Promise.all(
  commentRows.map(async ({ date, comment }) => {
    const result = await analyzeComment(comment);
      return {
        date,
        ...result
      };
    })
  );

  // 3. Generate the sentiment analysis section
  const sentimentSection = generateSentimentReportTable(sentimentResults);

  // Category summary (NG count per category)
  const ngCounts = categoryDescriptions.map((cat, idx) => {
    const colId = [
      'color_mkv02tqg', 'color_mkv0yb6g', 'color_mkv06e9z',
      'color_mkv0x9mr', 'color_mkv0df43', 'color_mkv5fa8m', 'color_mkv59ent'
    ][idx];
    return rowDataArray.filter(row => row[colId] === '否').length;
  });

  // HTML template
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
    <strong>店舗名：</strong>${storeName}<br>
    <strong>年月：</strong>${yearMonth}
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
    <li>記録日数：${rowDataArray.length}日</li>
    <li>全項目「良」達成日数：${rowDataArray.filter(row =>
    ['color_mkv02tqg', 'color_mkv0yb6g', 'color_mkv06e9z', 'color_mkv0x9mr', 'color_mkv0df43', 'color_mkv5fa8m', 'color_mkv59ent']
      .every(col => row[col] === '良')).length}日</li>
    <li>「否」あり日数：${rowDataArray.filter(row =>
        ['color_mkv02tqg', 'color_mkv0yb6g', 'color_mkv06e9z', 'color_mkv0x9mr', 'color_mkv0df43', 'color_mkv5fa8m', 'color_mkv59ent']
          .some(col => row[col] === '否')).length}日</li>
    <li>コメント記入日数：${rowDataArray.filter(row => row.text_mkv0etfg && row.text_mkv0etfg !== 'not found').length}日</li>
    <li>カテゴリごとの「否」回数：
      <ul>
        ${ngCounts.map((count, idx) => `<li>Cat ${idx + 1}: ${count}回</li>`).join('\n')}
      </ul>
    </li>
  </ul>
  ${sentimentSection}
  <h3>管理カテゴリ</h3>
  <table>
    <tr>
      <th>カテゴリ</th>
      <th>説明</th>
    </tr>
    ${categoryDescriptions.map((desc, idx) => `
      <tr>
        <td>Cat ${idx + 1}</td>
        <td>${desc}</td>
      </tr>
    `).join('\n')}
  </table>
  <div style="margin-top:2em;">
    このレポートは HygienMaster システムにより自動生成されました<br>
    生成日時: ${new Date().toISOString()}
  </div>
</body>
</html>
  `;
}

function generateJsonReport(rowDataArray, categories, originalFileName, context) {
  // Parse original filename for submission info
  const fileNameParts = parseFileName(originalFileName, context);

  // Get store and date info from first row
  const storeName = rowDataArray[0]?.text_mkv0z6d || "Unknown Store";
  const fullDate = rowDataArray[0]?.date4 || new Date().toISOString().split('T')[0];
  const yearMonth = fullDate.substring(0, 7); // YYYY-MM format

  const reportData = {
    // Report header (matching TXT exactly)
    title: "一般衛生管理の実施記録",
    submissionDate: fileNameParts.submissionDate,
    submitter: fileNameParts.senderEmail,
    originalFileName: fileNameParts.originalFileName,
    storeName: storeName,
    yearMonth: yearMonth,

    // Category definitions (matching TXT exactly)
    categories: categories.map((category, index) => ({
      id: `Cat ${index + 1}`,
      name: category
    })),

    // Table headers (matching TXT exactly)
    tableHeaders: [
      "日付",
      "Cat 1",
      "Cat 2",
      "Cat 3",
      "Cat 4",
      "Cat 5",
      "Cat 6",
      "Cat 7",
      "特記事項",
      "確認者"
    ],

    // Daily data (matching TXT table exactly)
    dailyData: rowDataArray.map(row => {
      const dayOnly = row.date4 ? row.date4.split('-')[2] : '--';

      return {
        日付: dayOnly,
        "Cat 1": row.color_mkv02tqg || '--',
        "Cat 2": row.color_mkv0yb6g || '--',
        "Cat 3": row.color_mkv06e9z || '--',
        "Cat 4": row.color_mkv0x9mr || '--',
        "Cat 5": row.color_mkv0df43 || '--',
        "Cat 6": row.color_mkv5fa8m || '--',
        "Cat 7": row.color_mkv59ent || '--',
        特記事項: row.text_mkv0etfg || '--',
        確認者: row.color_mkv0xnn4 || '--'
      };
    }),

    // Footer (matching TXT exactly)
    footer: {
      generatedBy: "HygienMaster システム",
      generatedAt: new Date().toISOString(),
      note: "このレポートは HygienMaster システムにより自動生成されました"
    }
  };

  return reportData;
}

function generateTextReport(rowDataArray, categories, originalFileName, context) {
  // Parse original filename for submission info
  const fileNameParts = parseFileName(originalFileName, context);

  // Get store and date info from first row (if available)
  let storeName = 'Unknown Store';
  let yearMonth = new Date().toISOString().substring(0, 7);

  if (rowDataArray.length > 0 && rowDataArray[0]) {
    const firstRow = rowDataArray[0];
    storeName = firstRow.text_mkv0z6d || firstRow.store || 'Unknown Store';

    if (firstRow.date4) {
      yearMonth = firstRow.date4.substring(0, 7);
    } else if (firstRow.year && firstRow.month) {
      yearMonth = `${firstRow.year}-${String(firstRow.month).padStart(2, '0')}`;
    }

    // Remove the context parameter since it's not available in this function
    logMessage(`📊 Store: ${storeName}, Year-Month: ${yearMonth}`, context);
  }

  let textReport = `
一般衛生管理の実施記録
提出日：${fileNameParts.submissionDate}
提出者：${fileNameParts.senderEmail}  
ファイル名：${fileNameParts.originalFileName}

店舗名：${storeName}
年月：${yearMonth}

管理カテゴリ：
`;

  // Add category descriptions
  if (categories && categories.length > 0) {
    categories.forEach((category, index) => {
      if (category && category !== 'not found') {
        textReport += `Cat ${index + 1}: ${category}\n`;
      }
    });
  } else {
    // Fallback category descriptions
    const defaultCategories = [
      '原材料の受入の確認',
      '庫内温度の確認 冷蔵庫・冷凍庫(°C)',
      '交差汚染・二次汚染の防止',
      '器具等の洗浄・消毒・殺菌',
      'トイレの洗浄・消毒',
      '従業員の健康管理等',
      '手洗いの実施'
    ];
    defaultCategories.forEach((category, index) => {
      textReport += `Cat ${index + 1}: ${category}\n`;
    });
  }

  textReport += '\n';

  // Create shorter table header
  const headerRow = `日付 | Cat 1 | Cat 2 | Cat 3 | Cat 4 | Cat 5 | Cat 6 | Cat 7 | 特記事項 | 確認者`;
  textReport += headerRow + '\n';
  textReport += ''.padEnd(headerRow.length, '-') + '\n';

  // Add data rows
  if (rowDataArray.length > 0) {
    rowDataArray.forEach(row => {
      if (row) {
        // Extract day from date4 (remove year-month part)
        let dayOnly = '--';
        if (row.date4) {
          dayOnly = row.date4.split('-')[2] || '--';
        } else if (row.day) {
          dayOnly = String(row.day).padStart(2, '0');
        }

        const dataRow = [
          dayOnly.padEnd(4),
          (row.color_mkv02tqg || '--').padEnd(6),
          (row.color_mkv0yb6g || '--').padEnd(6),
          (row.color_mkv06e9z || '--').padEnd(6),
          (row.color_mkv0x9mr || '--').padEnd(6),
          (row.color_mkv0df43 || '--').padEnd(6),
          (row.color_mkv5fa8m || '--').padEnd(6),
          (row.color_mkv59ent || '--').padEnd(6),
          (row.text_mkv0etfg && row.text_mkv0etfg !== 'not found' ? row.text_mkv0etfg.substring(0, 8) : '--').padEnd(8),
          (row.color_mkv0xnn4 || '--')
        ].join('| ');

        textReport += dataRow + '\n';
      }
    });
  } else {
    // No data available
    textReport += 'データが見つかりませんでした。\n';
  }

  textReport += `
========================================
このレポートは HygienMaster システムにより自動生成されました
生成日時: ${new Date().toISOString()}
========================================
`;

  return textReport;
}

function parseFileName(fileName, context) {
  logMessage(`🔍 Parsing filename: ${fileName}`, context);

  try {
    let submissionTime = '';
    let senderEmail = '';
    let originalFileName = fileName;

    // Extract email (between parentheses)
    const emailMatch = fileName.match(/\(([^)]*@[^)]*)\)/);
    if (emailMatch) {
      senderEmail = emailMatch[1];
      logMessage(`📧 Found email: ${senderEmail}`, context);

      // Extract original filename - everything AFTER the (email) closing parenthesis
      const emailEndIndex = fileName.indexOf(emailMatch[0]) + emailMatch[0].length;
      originalFileName = fileName.substring(emailEndIndex);

      // Clean up any leading/trailing whitespace and remove leading special characters
      originalFileName = originalFileName.replace(/^\W+/, '').trim();

      logMessage(`📄 Found original filename: ${originalFileName}`, context);
    }

    // Extract timestamp (before first parenthesis)
    const timeMatch = fileName.match(/^([^(]+)/);
    if (timeMatch) {
      submissionTime = timeMatch[1];
      logMessage(`⏰ Found timestamp: ${submissionTime}`, context);

      // Try to parse the timestamp
      if (submissionTime.includes('T')) {
        try {
          // Handle format like "20260826T050735"
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
              logMessage(`📅 Parsed date: ${submissionTime}`, context);
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

function generateSummaryData(extractedRows, categories) {
  const totalDays = extractedRows.length;
  const approvedDays = extractedRows.filter(row =>
    (row.color_mkv0xnn4 || row.approverStatus) === '選択済み'
  ).length;
  const daysWithComments = extractedRows.filter(row =>
    (row.text_mkv0etfg || row.comment) &&
    (row.text_mkv0etfg || row.comment) !== 'not found'
  ).length;

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
    criticalDays: []
  };

  // Category performance analysis using Monday column mapping
  const categoryColumnMapping = {
    0: 'color_mkv02tqg', // Category1
    1: 'color_mkv0yb6g', // Category2
    2: 'color_mkv06e9z', // Category3
    3: 'color_mkv0x9mr', // Category4
    4: 'color_mkv0df43', // Category5
    5: 'color_mkv5fa8m', // Category6
    6: 'color_mkv59ent'  // Category7
  };

  categories.forEach((category, index) => {
    const mondayColumnId = categoryColumnMapping[index];
    const legacyKey = `category${index + 1}Status`;

    const okCount = extractedRows.filter(row =>
      (row[mondayColumnId] || row[legacyKey]) === '良'
    ).length;
    const ngCount = extractedRows.filter(row =>
      (row[mondayColumnId] || row[legacyKey]) === '否'
    ).length;

    analytics.categoryPerformance.push({
      categoryId: index + 1,
      categoryName: category,
      mondayColumnId: mondayColumnId,
      okCount,
      ngCount,
      successRate: extractedRows.length > 0 ? (okCount / extractedRows.length * 100).toFixed(1) : 0,
      riskLevel: ngCount > extractedRows.length * 0.2 ? "critical" : ngCount > 0 ? "high" : "normal"
    });
  });

  return analytics;
}

function generateSentimentReportTable(sentimentResults) {
  return `
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
      ${sentimentResults.map(res => `
        <tr>
          <td>${res.date}</td>
          <td>${res.originalComment}</td>
          <td>${res.detectedLanguage}</td>
          <td>${res.japaneseTranslation}</td>
          <td>${res.sentimentAnalysisLanguage}</td>
          <td>${res.sentiment}</td>
          <td>
            👍 ${res.scores.positive} /
            😐 ${res.scores.neutral} /
            👎 ${res.scores.negative}
          </td>
        </tr>
      `).join('')}
    </table>
  `;
}

module.exports = {
  prepareGeneralManagementReport
};