'use strict';

const path = require('path');
const { logMessage, handleError } = require('../utils');
const { detectLanguageAndTranslate, getLanguageNameInJapanese } = require('../analytics/sentimentAnalysis');

/* -----------------------------------------------------------------------------
  HTML Report Generation with SharePoint Upload
----------------------------------------------------------------------------- */

/**
 * Enhanced function using the new lightweight language detection
 */
async function enhanceTextRegionsWithLanguage(analyseOutput, context) {
  logMessage(`🌐 Starting language detection for ${analyseOutput.length} text regions...`, context);
  
  const enhancedRegions = [];
  let successCount = 0;
  let errorCount = 0;
  
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
      // Use the lightweight language detection function
      const languageResult = await detectLanguageAndTranslate(text);
      
      if (languageResult.error) {
        logMessage(`⚠️ Language detection error for region ${i + 1}: ${languageResult.error}`, context);
        errorCount++;
        enhancedRegions.push({
          ...entry,
          detectedLanguage: 'Error',
          languageConfidence: 0,
          japaneseTranslation: '',
          needsTranslation: false
        });
      } else {
        successCount++;
        enhancedRegions.push({
          ...entry,
          detectedLanguage: languageResult.detectedLanguage,
          languageConfidence: languageResult.languageConfidence,
          japaneseTranslation: languageResult.japaneseTranslation || '',
          needsTranslation: languageResult.needsTranslation
        });
      }
      
    } catch (error) {
      logMessage(`❌ Language detection exception for region ${i + 1}: ${error.message}`, context);
      errorCount++;
      enhancedRegions.push({
        ...entry,
        detectedLanguage: 'Error',
        languageConfidence: 0,
        japaneseTranslation: '',
        needsTranslation: false
      });
    }
  }
  
  logMessage(`✅ Language detection complete: ${successCount} success, ${errorCount} errors`, context);
  return enhancedRegions;
}

/**
 * Fallback HTML generation if language detection fails
 */
async function produceHtml(analyseOutput, originalFileName, context, options = {}) {
  if (!Array.isArray(analyseOutput) || analyseOutput.length === 0) {
    return generateEmptyHtml(originalFileName);
  }

  try {
    logMessage(`📄 Generating HTML report for ${analyseOutput.length} text regions...`, context);

    // Try language detection, but fall back if it fails
    let enhancedRegions;
    try {
      logMessage(`🌐 Attempting language detection...`, context);
      enhancedRegions = await enhanceTextRegionsWithLanguage(analyseOutput, context);
      logMessage(`✅ Language detection completed`, context);
    } catch (languageError) {
      logMessage(`⚠️ Language detection failed, using original data: ${languageError.message}`, context);
      // Fall back to original data without language info
      enhancedRegions = analyseOutput.map(entry => ({
        ...entry,
        detectedLanguage: 'N/A',
        languageConfidence: 0,
        japaneseTranslation: '',
        needsTranslation: false
      }));
    }

    // Configuration
    const config = {
      scaleFactor: options.scaleFactor || 0.3,
      cellPadding: options.cellPadding || 8,
      showBoundingBoxes: options.showBoundingBoxes !== false,
      showOrientation: options.showOrientation !== false,
      showHandwriting: options.showHandwriting !== false,
      showLanguage: options.showLanguage !== false,
      showTranslation: options.showTranslation !== false,
      groupByRows: options.groupByRows !== false,
      ...options
    };

    // Find document boundaries
    const allBboxes = enhancedRegions.map(entry => entry.bbox).filter(Boolean);
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
      ? groupTextIntoRows(enhancedRegions, docBounds)
      : enhancedRegions.map(entry => ({ ...entry, rowIndex: 0 }));

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
        ${generateHeader(originalFileName, enhancedRegions, docBounds)}
        ${generateLanguageSummary(enhancedRegions)}
        ${generateSpatialLayout(textRegions, docBounds, config)}
        ${generateDataTable(enhancedRegions, config)}
        ${generateSummary(enhancedRegions)}
    </div>
    
    <script>
        // JavaScript for click-to-scroll functionality
        function scrollToTableRow(index) {
            const previousHighlighted = document.querySelector('.data-table tr.highlighted');
            if (previousHighlighted) {
                previousHighlighted.classList.remove('highlighted');
            }
            
            const targetRow = document.querySelector('.data-table tbody tr[data-index="' + index + '"]');
            if (targetRow) {
                targetRow.classList.add('highlighted');
                targetRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
                setTimeout(() => targetRow.classList.remove('highlighted'), 3000);
            }
        }
        
        document.addEventListener('DOMContentLoaded', function() {
            const regions = document.querySelectorAll('.text-region');
            regions.forEach(region => {
                region.addEventListener('mouseenter', function() {
                    this.style.zIndex = '100';
                });
                region.addEventListener('mouseleave', function() {
                    this.style.zIndex = '1';
                });
            });
        });
    </script>
</body>
</html>`;

    logMessage(`✅ Generated HTML report: ${html.length} characters`, context);
    return html;

  } catch (error) {
    logMessage(`❌ HTML generation failed: ${error.message}`, context);
    logMessage(`❌ Error stack: ${error.stack}`, context);
    handleError(error, 'produceHtml', context);
    return generateErrorHtml(originalFileName, error);
  }
}

/**
 * Generate HTML report and upload to SharePoint (with enhanced debugging)
 */
async function generateHtmlReportToSharePoint(analyseOutput, originalFileName, context, companyName, folderPath, options = {}) {
  if (!Array.isArray(analyseOutput)) {
    handleError(new Error('Invalid analyseOutput input'), 'generateHtmlReportToSharePoint', context);
    return null;
  }

  try {
    logMessage(`📄 Starting HTML report generation for SharePoint upload...`, context);
    logMessage(`📁 Target folder: ${folderPath}`, context);
    logMessage(`📊 Input data: ${analyseOutput.length} text regions`, context);

    // Generate HTML content (now async for language detection)
    logMessage(`🌐 Generating HTML content with language detection...`, context);
    const htmlContent = await produceHtml(analyseOutput, originalFileName, context, options);

    if (!htmlContent) {
      logMessage(`❌ Failed to generate HTML content - content is null/undefined`, context);
      return null;
    }

    logMessage(`✅ HTML content generated successfully: ${htmlContent.length} characters`, context);

    // Prepare SharePoint upload
    const baseName = path.basename(originalFileName, path.extname(originalFileName));
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const htmlFileName = `文書解析レポート-${baseName}-${timestamp}.html`;

    logMessage(`📤 Preparing SharePoint upload: ${htmlFileName}`, context);

    // Import SharePoint helpers
    try {
      const { ensureSharePointFolder, uploadOriginalDocumentToSharePoint } = require('../sharepoint/sendToSharePoint');
      logMessage(`✅ SharePoint helpers imported successfully`, context);

      // Ensure folder exists
      logMessage(`📁 Ensuring SharePoint folder exists: ${folderPath}`, context);
      await ensureSharePointFolder(folderPath, context);
      logMessage(`✅ SharePoint folder ensured`, context);

      // Convert HTML to base64 for upload
      logMessage(`🔄 Converting HTML to base64...`, context);
      const base64HtmlContent = Buffer.from(htmlContent, 'utf8').toString('base64');
      logMessage(`✅ HTML converted to base64: ${base64HtmlContent.length} characters`, context);

      // Upload to SharePoint
      logMessage(`📤 Starting SharePoint upload...`, context);
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
        logMessage(`❌ SharePoint upload returned null/false`, context);
        return null;
      }

    } catch (importError) {
      logMessage(`❌ Failed to import SharePoint helpers: ${importError.message}`, context);
      logMessage(`❌ Import error stack: ${importError.stack}`, context);
      throw importError;
    }

  } catch (error) {
    logMessage(`❌ HTML report generation failed: ${error.message}`, context);
    logMessage(`❌ Error stack: ${error.stack}`, context);
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
 * Generate spatial layout using CSS positioning with enhanced interactivity
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

    // ✅ Enhanced with number display and click functionality
    html += `
      <div class="text-region ${handwritingClass} ${orientationClass} clickable-region" 
           style="left: ${scaledX}px; top: ${scaledY}px; width: ${scaledW}px; height: ${scaledH}px;"
           title="Region ${idx + 1}: ${escapeHtml(entry.displayText || '')}"
           data-index="${idx}"
           onclick="scrollToTableRow(${idx})">
        <div class="region-number">${idx + 1}</div>
        ${config.showOrientation ? `<div class="orientation">${entry.orientationDeg || 0}°</div>` : ''}
      </div>`;
  });

  html += `
      </div>
      <div class="spatial-instructions">
        💡 ヒント: ボックスをクリックすると、対応するデータテーブル行にジャンプします
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
 * ✅ Generate language detection summary
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

/**
 * Generate header section with JST time
 */
function generateHeader(originalFileName, analyseOutput, docBounds) {
  // ✅ Fix: Proper JST conversion - use simpler approach
  const now = new Date();
  
  // Create JST time directly using toLocaleString with Asia/Tokyo timezone
  const processedTime = now.toLocaleString('ja-JP', {
    timeZone: 'Asia/Tokyo',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit'
  });

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
      max-width: 1600px;
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
    
    .language-summary-section {
      padding: 20px;
      border-bottom: 1px solid #eee;
      background: #f8f9fa;
    }
    
    .language-summary-section h2 {
      margin: 0 0 15px 0;
      color: #333;
      border-bottom: 2px solid #FF9800;
      padding-bottom: 10px;
    }
    
    .language-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
      gap: 10px;
    }
    
    .language-item, .translation-item {
      text-align: center;
      padding: 10px;
      background: white;
      border-radius: 6px;
      border-left: 4px solid #FF9800;
    }
    
    .language-name, .translation-label {
      font-size: 12px;
      color: #666;
      margin-bottom: 5px;
    }
    
    .language-count, .translation-count {
      font-size: 18px;
      font-weight: bold;
      color: #FF9800;
    }
    
    .translation-item {
      border-left-color: #4CAF50;
    }
    
    .translation-count {
      color: #4CAF50;
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
    
    .spatial-instructions {
      margin-top: 10px;
      padding: 10px;
      background: #e3f2fd;
      border-left: 4px solid #2196F3;
      border-radius: 4px;
      font-size: 14px;
      color: #1976D2;
    }
    
    .text-region {
      position: absolute;
      border: 2px solid;
      background: rgba(25, 118, 210, 0.1);
      font-size: 12px;
      overflow: hidden;
      transition: all 0.2s ease;
      display: flex;
      align-items: center;
      justify-content: center;
    }
    
    .text-region.clickable-region {
      cursor: pointer;
    }
    
    .text-region:hover {
      transform: scale(1.1);
      z-index: 10;
      box-shadow: 0 6px 16px rgba(0,0,0,0.4);
      border-width: 3px;
    }
    
    .text-region.handwritten {
      border-color: #2E7D32;
      background: rgba(46, 125, 50, 0.1);
    }
    
    .text-region.printed {
      border-color: #1976D2;
      background: rgba(25, 118, 210, 0.1);
    }
    
    .region-number {
      font-weight: bold;
      font-size: 14px;
      color: #fff;
      background: rgba(0, 0, 0, 0.8);
      border-radius: 50%;
      width: 24px;
      height: 24px;
      display: flex;
      align-items: center;
      justify-content: center;
      position: absolute;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      border: 2px solid #fff;
      box-shadow: 0 2px 4px rgba(0,0,0,0.3);
    }
    
    .text-region.handwritten .region-number {
      background: rgba(46, 125, 50, 0.9);
    }
    
    .text-region.printed .region-number {
      background: rgba(25, 118, 210, 0.9);
    }
    
    .text-region:hover .region-number {
      transform: translate(-50%, -50%) scale(1.2);
      box-shadow: 0 4px 8px rgba(0,0,0,0.4);
    }
    
    .orientation {
      position: absolute;
      top: -20px;
      right: 2px;
      background: rgba(0,0,0,0.7);
      color: white;
      padding: 2px 4px;
      font-size: 8px;
      border-radius: 2px;
    }
    
    .data-table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
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
    
    .data-table tr.highlighted {
      background: #ffeb3b !important;
      animation: highlightPulse 2s ease-in-out;
    }
    
    @keyframes highlightPulse {
      0% { background: #ff9800; }
      50% { background: #ffeb3b; }
      100% { background: #fff3c4; }
    }
    
    .handwritten-row {
      background: rgba(46, 125, 50, 0.05);
    }
    
    .printed-row {
      background: rgba(25, 118, 210, 0.05);
    }
    
    .index-cell {
      width: 40px;
      text-align: center;
      font-weight: bold;
    }
    
    .text-cell {
      min-width: 150px;
      max-width: 200px;
    }
    
    .display-text {
      font-weight: bold;
      color: #333;
      word-break: break-all;
    }
    
    .coords-cell {
      font-family: monospace;
      font-size: 11px;
      color: #666;
      min-width: 120px;
    }
    
    .orientation-cell {
      text-align: center;
      font-family: monospace;
      width: 60px;
    }
    
    .handwriting-cell {
      text-align: center;
      width: 80px;
    }
    
    .language-cell {
      min-width: 100px;
      text-align: center;
    }
    
    .language-info {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 2px;
    }
    
    .language-name {
      font-weight: bold;
      color: #FF9800;
    }
    
    .confidence {
      font-size: 11px;
      color: #666;
    }
    
    .translation-cell {
      min-width: 150px;
      max-width: 200px;
    }
    
    .translation-text {
      background: #e8f5e8;
      padding: 4px 8px;
      border-radius: 4px;
      border-left: 3px solid #4CAF50;
      font-size: 12px;
      word-break: break-all;
    }
    
    .no-translation {
      color: #999;
      font-style: italic;
      font-size: 12px;
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
      .data-table { font-size: 11px; }
      .data-table th, .data-table td { padding: 3px; }
      .language-grid { grid-template-columns: 1fr 1fr; }
      .text-cell, .translation-cell { min-width: 120px; max-width: 150px; }
      .region-number { width: 20px; height: 20px; font-size: 12px; }
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

/* -----------------------------------------------------------------------------
  Exports
----------------------------------------------------------------------------- */
module.exports = {
  produceHtml,
  generateHtmlReportToSharePoint
};