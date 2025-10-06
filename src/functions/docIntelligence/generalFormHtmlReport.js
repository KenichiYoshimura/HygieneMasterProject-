'use strict';

const path = require('path');
const { logMessage, handleError } = require('../utils');
const { analyzeComment, getLanguageNameInJapanese } = require('../analytics/sentimentAnalysis');

/* -----------------------------------------------------------------------------
  HTML Report Generation with SharePoint Upload
----------------------------------------------------------------------------- */

/**
 * Generate human-readable HTML report with spatial layout preservation
 * @param {Array} analyseOutput - Text regions with bbox, text, etc.
 * @param {string} originalFileName - Original filename
 * @param {Object} context - Azure Functions context
 * @param {Object} options - Configuration options
 * @returns {string} HTML content
 */
function produceHtml(analyseOutput, originalFileName, context, options = {}) {
  if (!Array.isArray(analyseOutput) || analyseOutput.length === 0) {
    return generateEmptyHtml(originalFileName);
  }

  try {
    logMessage(`📄 Generating HTML report for ${analyseOutput.length} text regions...`, context);

    // Configuration
    const config = {
      scaleFactor: options.scaleFactor || 0.3,  // Scale down coordinates for display
      cellPadding: options.cellPadding || 8,
      showBoundingBoxes: options.showBoundingBoxes !== false,
      showOrientation: options.showOrientation !== false,
      showHandwriting: options.showHandwriting !== false,
      groupByRows: options.groupByRows !== false,
      ...options
    };

    // Find document boundaries
    const allBboxes = analyseOutput.map(entry => entry.bbox).filter(Boolean);
    const docBounds = {
      minX: Math.min(...allBboxes.map(b => b[0])),
      minY: Math.min(...allBboxes.map(b => b[1])),
      maxX: Math.max(...allBboxes.map(b => b[2])),
      maxY: Math.max(...allBboxes.map(b => b[3]))
    };

    const docWidth = docBounds.maxX - docBounds.minX;
    const docHeight = docBounds.maxY - docBounds.minY;

    logMessage(`📐 Document bounds: ${docWidth}x${docHeight} pixels`, context);

    // Group text regions by approximate rows (if enabled)
    const textRegions = config.groupByRows 
      ? groupTextIntoRows(analyseOutput, docBounds)
      : analyseOutput.map(entry => ({ ...entry, rowIndex: 0 }));

    // Generate HTML
    const html = `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>文書解析レポート - ${escapeHtml(originalFileName)}</title>
    <style>
        ${generateCSS(config)}
    </style>
</head>
<body>
    <div class="container">
        ${generateHeader(originalFileName, analyseOutput, docBounds)}
        ${generateSpatialLayout(textRegions, docBounds, config)}
        ${generateDataTable(analyseOutput, config)}
        ${generateSummary(analyseOutput)}
    </div>
</body>
</html>`;

    logMessage(`✅ Generated HTML report: ${html.length} characters`, context);
    return html;

  } catch (error) {
    handleError(error, 'produceHtml', context);
    return generateErrorHtml(originalFileName, error);
  }
}

/**
 * Generate HTML report and upload to SharePoint
 * @param {Array} analyseOutput - Text regions with bbox, text, etc.
 * @param {string} originalFileName - Original filename
 * @param {Object} context - Azure Functions context
 * @param {string} companyName - Company name for folder organization
 * @param {string} folderPath - SharePoint folder path
 * @param {Object} options - Configuration options
 * @returns {Promise<Object|null>} SharePoint upload result or null
 */
async function generateHtmlReportToSharePoint(analyseOutput, originalFileName, context, companyName, folderPath, options = {}) {
  if (!Array.isArray(analyseOutput)) {
    handleError(new Error('Invalid analyseOutput input'), 'generateHtmlReportToSharePoint', context);
    return null;
  }

  try {
    logMessage(`📄 Generating HTML report for SharePoint upload...`, context);
    logMessage(`📁 Target folder: ${folderPath}`, context);

    // Generate HTML content
    const htmlContent = produceHtml(analyseOutput, originalFileName, context, options);

    if (!htmlContent) {
      logMessage(`❌ Failed to generate HTML content`, context);
      return null;
    }

    // Prepare SharePoint upload
    const baseName = path.basename(originalFileName, path.extname(originalFileName));
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const htmlFileName = `文書解析レポート-${baseName}-${timestamp}.html`;

    logMessage(`📤 Uploading HTML report to SharePoint: ${htmlFileName}`, context);

    // Import SharePoint helpers
    const { ensureSharePointFolder, uploadOriginalDocumentToSharePoint } = require('../sharepoint/sendToSharePoint');

    // Ensure folder exists
    await ensureSharePointFolder(folderPath, context);

    // Convert HTML to base64 for upload
    const base64HtmlContent = Buffer.from(htmlContent, 'utf8').toString('base64');

    // Upload to SharePoint
    const sharePointResult = await uploadOriginalDocumentToSharePoint(
      base64HtmlContent,
      htmlFileName,
      folderPath,
      context
    );

    if (sharePointResult) {
      logMessage(`✅ Successfully uploaded HTML report to SharePoint: ${htmlFileName}`, context);
      return {
        success: true,
        fileName: htmlFileName,
        fileSize: htmlContent.length,
        uploadResult: sharePointResult
      };
    } else {
      logMessage(`❌ Failed to upload HTML report to SharePoint`, context);
      return null;
    }

  } catch (error) {
    handleError(error, 'generateHtmlReportToSharePoint', context);
    return null;
  }
}

/**
 * Group text regions into approximate rows based on Y coordinates
 */
function groupTextIntoRows(analyseOutput, docBounds, rowThreshold = 40) {
  const regions = analyseOutput.map(entry => ({
    ...entry,
    centerY: (entry.bbox[1] + entry.bbox[3]) / 2
  }));

  // Sort by Y position
  regions.sort((a, b) => a.centerY - b.centerY);

  const rows = [];
  let currentRow = [];
  let currentRowY = null;

  regions.forEach(region => {
    if (currentRowY === null || Math.abs(region.centerY - currentRowY) <= rowThreshold) {
      currentRow.push(region);
      currentRowY = currentRowY === null ? region.centerY : (currentRowY + region.centerY) / 2;
    } else {
      if (currentRow.length > 0) {
        // Sort current row by X position
        currentRow.sort((a, b) => a.bbox[0] - b.bbox[0]);
        rows.push(currentRow);
      }
      currentRow = [region];
      currentRowY = region.centerY;
    }
  });

  if (currentRow.length > 0) {
    currentRow.sort((a, b) => a.bbox[0] - b.bbox[0]);
    rows.push(currentRow);
  }

  // Add row indices
  const result = [];
  rows.forEach((row, rowIndex) => {
    row.forEach(region => {
      result.push({ ...region, rowIndex });
    });
  });

  return result;
}

/**
 * Generate spatial layout using CSS positioning
 */
function generateSpatialLayout(textRegions, docBounds, config) {
  const scaledWidth = (docBounds.maxX - docBounds.minX) * config.scaleFactor;
  const scaledHeight = (docBounds.maxY - docBounds.minY) * config.scaleFactor;

  let html = `
    <div class="spatial-section">
      <h2>📍 空間レイアウト</h2>
      <div class="spatial-container" style="width: ${scaledWidth}px; height: ${scaledHeight}px;">`;

  textRegions.forEach((entry, idx) => {
    const bbox = entry.bbox;
    const scaledX = (bbox[0] - docBounds.minX) * config.scaleFactor;
    const scaledY = (bbox[1] - docBounds.minY) * config.scaleFactor;
    const scaledW = (bbox[2] - bbox[0]) * config.scaleFactor;
    const scaledH = (bbox[3] - bbox[1]) * config.scaleFactor;

    const isHandwritten = entry.isHandwritten;
    const orientationClass = Math.abs(entry.orientationDeg || 0) > 45 ? 'rotated' : 'horizontal';
    const handwritingClass = isHandwritten ? 'handwritten' : 'printed';

    html += `
      <div class="text-region ${handwritingClass} ${orientationClass}" 
           style="left: ${scaledX}px; top: ${scaledY}px; width: ${scaledW}px; height: ${scaledH}px;"
           title="Region ${idx}: ${escapeHtml(entry.displayText || '')}"
           data-index="${idx}">
        <div class="text-content">${escapeHtml(truncateText(entry.displayText || '', 50))}</div>
        ${config.showOrientation ? `<div class="orientation">${entry.orientationDeg || 0}°</div>` : ''}
      </div>`;
  });

  html += `
      </div>
    </div>`;

  return html;
}

/**
 * Generate data table with all extracted information
 */
function generateDataTable(analyseOutput, config) {
  let html = `
    <div class="table-section">
      <h2>📋 抽出データ一覧</h2>
      <table class="data-table">
        <thead>
          <tr>
            <th>#</th>
            <th>抽出テキスト</th>
            <th>座標 (X,Y,W,H)</th>
            ${config.showOrientation ? '<th>回転角度</th>' : ''}
            ${config.showHandwriting ? '<th>手書き</th>' : ''}
            ${config.showLanguage ? '<th>検出言語</th>' : ''}
            ${config.showTranslation ? '<th>日本語翻訳</th>' : ''}
            <th>OCRマッチ</th>
            <th>Layoutテキスト</th>
          </tr>
        </thead>
        <tbody>`;

  analyseOutput.forEach((entry, idx) => {
    const bbox = entry.bbox || [0, 0, 0, 0];
    const width = bbox[2] - bbox[0];
    const height = bbox[3] - bbox[1];
    const coordinates = `(${bbox[0]}, ${bbox[1]}, ${width}, ${height})`;
    
    html += `
      <tr data-index="${idx}" class="${entry.isHandwritten ? 'handwritten-row' : 'printed-row'}">
        <td class="index-cell">${idx + 1}</td>
        <td class="text-cell">
          <div class="display-text">${escapeHtml(entry.displayText || '')}</div>
        </td>
        <td class="coords-cell">${coordinates}</td>
        ${config.showOrientation ? `<td class="orientation-cell">${entry.orientationDeg || 0}°</td>` : ''}
        ${config.showHandwriting ? `<td class="handwriting-cell">${entry.isHandwritten ? '✍️ 手書き' : '🖨️ 印刷'}</td>` : ''}
        ${config.showLanguage ? `
          <td class="language-cell">
            <div class="language-info">
              <span class="language-name">${getLanguageNameInJapanese(entry.detectedLanguage)}</span>
              ${entry.languageConfidence > 0 ? `<span class="confidence">(${Math.round(entry.languageConfidence * 100)}%)</span>` : ''}
            </div>
          </td>
        ` : ''}
        ${config.showTranslation ? `
          <td class="translation-cell">
            ${entry.needsTranslation && entry.japaneseTranslation 
              ? `<div class="translation-text">${escapeHtml(entry.japaneseTranslation)}</div>`
              : entry.detectedLanguage === 'ja'
                ? '<em class="no-translation">翻訳不要</em>'
                : '<em class="no-translation">翻訳なし</em>'
            }
          </td>
        ` : ''}
        <td class="ocr-cell">
          ${(entry.matchedOCR || []).length > 0 
            ? (entry.matchedOCR || []).map(text => `<div class="ocr-match">${escapeHtml(text)}</div>`).join('')
            : '<em>なし</em>'
          }
        </td>
        <td class="layout-cell">
          ${entry.layoutText ? `<div class="layout-text">${escapeHtml(entry.layoutText)}</div>` : '<em>なし</em>'}
        </td>
      </tr>`;
  });

  html += `
        </tbody>
      </table>
    </div>`;

  return html;
}

/**
 * Generate summary statistics
 */
function generateSummary(analyseOutput) {
  const totalRegions = analyseOutput.length;
  const handwrittenCount = analyseOutput.filter(entry => entry.isHandwritten).length;
  const printedCount = totalRegions - handwrittenCount;
  const ocrMatches = analyseOutput.filter(entry => entry.matchedOCR && entry.matchedOCR.length > 0).length;
  const layoutMatches = analyseOutput.filter(entry => entry.layoutText).length;

  return `
    <div class="summary-section">
      <h2>📊 解析サマリー</h2>
      <div class="summary-grid">
        <div class="summary-item">
          <div class="summary-value">${totalRegions}</div>
          <div class="summary-label">総テキスト領域数</div>
        </div>
        <div class="summary-item">
          <div class="summary-value">${printedCount}</div>
          <div class="summary-label">印刷テキスト</div>
        </div>
        <div class="summary-item">
          <div class="summary-value">${handwrittenCount}</div>
          <div class="summary-label">手書きテキスト</div>
        </div>
        <div class="summary-item">
          <div class="summary-value">${ocrMatches}</div>
          <div class="summary-label">OCR検出</div>
        </div>
        <div class="summary-item">
          <div class="summary-value">${layoutMatches}</div>
          <div class="summary-label">Layout検出</div>
        </div>
      </div>
    </div>`;
}

/**
 * Generate header section
 */
function generateHeader(originalFileName, analyseOutput, docBounds) {
  const processedTime = new Date().toLocaleString('ja-JP');
  const docWidth = docBounds.maxX - docBounds.minX;
  const docHeight = docBounds.maxY - docBounds.minY;

  return `
    <div class="header-section">
      <h1>📄 文書解析レポート</h1>
      <div class="header-info">
        <div class="info-item"><strong>ファイル名:</strong> ${escapeHtml(originalFileName)}</div>
        <div class="info-item"><strong>処理時刻:</strong> ${processedTime}</div>
        <div class="info-item"><strong>文書サイズ:</strong> ${docWidth} × ${docHeight} ピクセル</div>
        <div class="info-item"><strong>検出領域数:</strong> ${analyseOutput.length}</div>
      </div>
    </div>`;
}

/**
 * Generate CSS styles
 */
function generateCSS(config) {
  return `
    * { box-sizing: border-box; }
    
    body {
      font-family: 'Noto Sans JP', 'Hiragino Sans', sans-serif;
      line-height: 1.6;
      margin: 0;
      padding: 20px;
      background-color: #f5f5f5;
    }
    
    .container {
      max-width: 1400px;
      margin: 0 auto;
      background: white;
      border-radius: 8px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
      overflow: hidden;
    }
    
    .header-section {
      background: linear-gradient(135deg, #1976D2, #2196F3);
      color: white;
      padding: 20px;
    }
    
    .header-section h1 {
      margin: 0 0 15px 0;
      font-size: 28px;
    }
    
    .header-info {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
      gap: 10px;
      font-size: 14px;
    }
    
    .spatial-section, .table-section, .summary-section {
      padding: 20px;
      border-bottom: 1px solid #eee;
    }
    
    .spatial-section h2, .table-section h2, .summary-section h2 {
      margin: 0 0 20px 0;
      color: #333;
      border-bottom: 2px solid #2196F3;
      padding-bottom: 10px;
    }
    
    .spatial-container {
      position: relative;
      border: 2px solid #ddd;
      margin: 20px 0;
      background: #fafafa;
      overflow: auto;
      min-height: 300px;
    }
    
    .text-region {
      position: absolute;
      border: 2px solid;
      background: rgba(25, 118, 210, 0.1);
      font-size: 10px;
      overflow: hidden;
      cursor: pointer;
      transition: all 0.2s ease;
    }
    
    .text-region:hover {
      transform: scale(1.05);
      z-index: 10;
      box-shadow: 0 4px 12px rgba(0,0,0,0.3);
    }
    
    .text-region.handwritten {
      border-color: #2E7D32;
      background: rgba(46, 125, 50, 0.1);
    }
    
    .text-region.printed {
      border-color: #1976D2;
      background: rgba(25, 118, 210, 0.1);
    }
    
    .text-content {
      padding: 2px;
      font-weight: bold;
      color: #333;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    
    .orientation {
      position: absolute;
      top: -20px;
      right: 0;
      background: rgba(0,0,0,0.7);
      color: white;
      padding: 2px 4px;
      font-size: 8px;
      border-radius: 2px;
    }
    
    .data-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 14px;
    }
    
    .data-table th, .data-table td {
      border: 1px solid #ddd;
      padding: ${config.cellPadding}px;
      text-align: left;
      vertical-align: top;
    }
    
    .data-table th {
      background: #f8f9fa;
      font-weight: bold;
      position: sticky;
      top: 0;
      z-index: 5;
    }
    
    .data-table tr:hover {
      background: #f0f7ff;
    }
    
    .handwritten-row {
      background: rgba(46, 125, 50, 0.05);
    }
    
    .printed-row {
      background: rgba(25, 118, 210, 0.05);
    }
    
    .index-cell {
      width: 50px;
      text-align: center;
      font-weight: bold;
    }
    
    .text-cell {
      min-width: 200px;
      max-width: 300px;
    }
    
    .display-text {
      font-weight: bold;
      color: #333;
      word-break: break-all;
    }
    
    .coords-cell {
      font-family: monospace;
      font-size: 12px;
      color: #666;
    }
    
    .orientation-cell {
      text-align: center;
      font-family: monospace;
    }
    
    .handwriting-cell {
      text-align: center;
    }
    
    .ocr-match, .layout-text {
      margin: 2px 0;
      padding: 2px 4px;
      background: #f8f9fa;
      border-radius: 3px;
      font-size: 12px;
    }
    
    .summary-section {
      border-bottom: none;
    }
    
    .summary-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
      gap: 15px;
    }
    
    .summary-item {
      text-align: center;
      padding: 15px;
      background: #f8f9fa;
      border-radius: 6px;
      border-left: 4px solid #2196F3;
    }
    
    .summary-value {
      font-size: 24px;
      font-weight: bold;
      color: #1976D2;
    }
    
    .summary-label {
      font-size: 12px;
      color: #666;
      margin-top: 5px;
    }
    
    @media (max-width: 768px) {
      .container { margin: 10px; }
      .header-info { grid-template-columns: 1fr; }
      .data-table { font-size: 12px; }
      .data-table th, .data-table td { padding: 4px; }
    }`;
}

/**
 * Utility functions
 */
function escapeHtml(text) {
  const textNode = text || '';
  return textNode
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function truncateText(text, maxLength) {
  if (!text || text.length <= maxLength) return text;
  return text.slice(0, maxLength) + '...';
}

function generateEmptyHtml(originalFileName) {
  return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>文書解析レポート - ${escapeHtml(originalFileName)}</title>
</head>
<body>
    <div style="text-align: center; padding: 50px;">
        <h1>📄 文書解析レポート</h1>
        <p>ファイル: ${escapeHtml(originalFileName)}</p>
        <p>⚠️ テキスト領域が検出されませんでした。</p>
    </div>
</body>
</html>`;
}

function generateErrorHtml(originalFileName, error) {
  return `
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>エラー - ${escapeHtml(originalFileName)}</title>
</head>
<body>
    <div style="text-align: center; padding: 50px;">
        <h1>❌ エラーが発生しました</h1>
        <p>ファイル: ${escapeHtml(originalFileName)}</p>
        <p>エラー: ${escapeHtml(error.message)}</p>
    </div>
</body>
</html>`;
}

/**
 * Enhanced function to detect language and translate text regions using existing sentimentAnalysis logic
 */
async function enhanceTextRegionsWithLanguage(analyseOutput, context) {
  logMessage(`🌐 Detecting languages and translating text regions...`, context);
  
  const enhancedRegions = [];
  
  for (let i = 0; i < analyseOutput.length; i++) {
    const entry = analyseOutput[i];
    const text = entry.displayText || '';
    
    if (!text.trim()) {
      // Empty text - skip language detection
      enhancedRegions.push({
        ...entry,
        detectedLanguage: 'N/A',
        languageConfidence: 0,
        japaneseTranslation: '',
        needsTranslation: false
      });
      continue;
    }

    try {
      // ✅ Use existing analyzeComment function which already does language detection + translation
      const analysisResult = await analyzeComment(text);
      
      enhancedRegions.push({
        ...entry,
        detectedLanguage: analysisResult.detectedLanguage || 'unknown',
        languageConfidence: analysisResult.confidenceScores ? 
          Math.max(
            analysisResult.confidenceScores.positive || 0,
            analysisResult.confidenceScores.neutral || 0,
            analysisResult.confidenceScores.negative || 0
          ) : 0,
        japaneseTranslation: analysisResult.japaneseTranslation || '',
        needsTranslation: analysisResult.detectedLanguage !== 'ja' && analysisResult.japaneseTranslation,
        // ✅ Also store sentiment data for potential future use
        sentiment: analysisResult.sentiment,
        sentimentScores: analysisResult.confidenceScores
      });
      
    } catch (error) {
      logMessage(`❌ Language detection failed for region ${i}: ${error.message}`, context);
      enhancedRegions.push({
        ...entry,
        detectedLanguage: 'Error',
        languageConfidence: 0,
        japaneseTranslation: '',
        needsTranslation: false
      });
    }
  }
  
  logMessage(`✅ Language detection complete for ${enhancedRegions.length} regions`, context);
  return enhancedRegions;
}

/**
 * ✅ Generate language detection summary (updated to use existing function)
 */
function generateLanguageSummary(analyseOutput) {
  const languageStats = {};
  const translationNeeded = analyseOutput.filter(entry => entry.needsTranslation).length;
  
  analyseOutput.forEach(entry => {
    const lang = entry.detectedLanguage || 'Unknown';
    languageStats[lang] = (languageStats[lang] || 0) + 1;
  });

  const sortedLanguages = Object.entries(languageStats)
    .sort(([,a], [,b]) => b - a)
    .slice(0, 5); // Top 5 languages

  return `
    <div class="language-summary-section">
      <h2>🌐 言語検出サマリー</h2>
      <div class="language-grid">
        ${sortedLanguages.map(([lang, count]) => `
          <div class="language-item">
            <div class="language-name">${getLanguageNameInJapanese(lang)}</div>
            <div class="language-count">${count}件</div>
          </div>
        `).join('')}
        <div class="translation-item">
          <div class="translation-label">翻訳対象</div>
          <div class="translation-count">${translationNeeded}件</div>
        </div>
      </div>
    </div>`;
}

/* -----------------------------------------------------------------------------
  Exports
----------------------------------------------------------------------------- */
module.exports = {
  produceHtml,
  generateHtmlReportToSharePoint
};