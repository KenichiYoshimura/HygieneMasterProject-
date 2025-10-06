'use strict';

if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const { AzureKeyCredential, DocumentAnalysisClient } = require("@azure/ai-form-recognizer");
const { logMessage, handleError } = require('../utils');

// ✅ Import image rendering functions from the new module
const { generateAnnotatedImage, generateAnnotatedImageToSharePoint } = require('./generalFormImageAnnotator');

/* -----------------------------------------------------------------------------
  Azure setup
----------------------------------------------------------------------------- */
const endpoint = process.env['CLASSIFIER_ENDPOINT'];
const apiKey = process.env['CLASSIFIER_ENDPOINT_AZURE_API_KEY'];

logMessage('🔧 Endpoint: ' + endpoint, null);
logMessage('🔧 API Key: ' + (apiKey ? '[REDACTED]' : '❌ Missing API Key'), null);

if (!endpoint || !apiKey) {
  throw new Error('Missing CLASSIFIER_ENDPOINT or CLASSIFIER_ENDPOINT_AZURE_API_KEY');
}
const client = new DocumentAnalysisClient(endpoint, new AzureKeyCredential(apiKey));

/* -----------------------------------------------------------------------------
  Geometry & angle helpers
----------------------------------------------------------------------------- */
function getBoundingBoxFromPolygon(polygon = []) {
  if (!polygon || polygon.length === 0) return [];
  const xs = polygon.map(p => p.x);
  const ys = polygon.map(p => p.y);
  const x1 = Math.min(...xs);
  const y1 = Math.min(...ys);
  const x2 = Math.max(...xs);
  const y2 = Math.max(...ys);
  return [x1, y1, x2, y2];
}

function calculateIoU(box1, box2) {
  const x1 = Math.max(box1[0], box2[0]);
  const y1 = Math.max(box1[1], box2[1]);
  const x2 = Math.min(box1[2], box2[2]);
  const y2 = Math.min(box1[3], box2[3]);
  const intersection = Math.max(0, x2 - x1) * Math.max(0, y2 - y1);
  const area1 = (box1[2] - box1[0]) * (box1[3] - box1[1]);
  const area2 = (box2[2] - box2[0]) * (box2[3] - box2[1]);
  const union = area1 + area2 - intersection;
  return union === 0 ? 0 : intersection / union;
}

function bboxOverlap(b1, b2) {
  const x1 = Math.max(b1[0], b2[0]);
  const y1 = Math.max(b1[1], b2[1]);
  const x2 = Math.min(b1[2], b2[2]);
  const y2 = Math.min(b1[3], b2[3]);
  return x2 > x1 && y2 > y1;
}

// Compute angle (deg in [0, 180)) from a polygon's first edge
function angleFromPolygon(polygon) {
  if (!polygon || polygon.length < 2) return null;
  const dx = polygon[1].x - polygon[0].x;
  const dy = polygon[1].y - polygon[0].y;
  const rad = Math.atan2(dy, dx);
  let deg = rad * 180 / Math.PI;
  deg = ((deg % 180) + 180) % 180;
  return deg;
}

function median(arr) {
  if (!arr || arr.length === 0) return null;
  const a = arr.slice().sort((x, y) => x - y);
  const mid = Math.floor(a.length / 2);
  return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

function snapAngle0or90(deg) {
  if (deg == null || Number.isNaN(deg)) return 0;
  const d0 = Math.min(deg, 180 - deg);
  const d90 = Math.abs(90 - deg);
  return d0 <= d90 ? 0 : 90;
}

/* -----------------------------------------------------------------------------
  CJK helpers & text normalization
----------------------------------------------------------------------------- */
function isMostlyCJK(str) {
  const s = String(str || '');
  const cjkMatches = s.match(/[\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]/g) || [];
  return cjkMatches.length > 0 && (cjkMatches.length / s.length) > 0.3;
}

/**
 * Consolidate layoutText and matchedOCR into one single display string.
 * PRIORITY: Use OCR text first, fallback to Layout text only if OCR is empty.
 */
function toSingleMergedText(layoutText, matchedOCRArr) {
  const rawLayout = (layoutText || '').trim();
  const rawOCR = (matchedOCRArr || []).join(' ').trim();
  // ✅ OCR first, fallback to Layout
  let base = rawOCR || rawLayout || '';
  if (isMostlyCJK(base)) {
    base = base.replace(/\s*\n\s*/g, '');
    base = base.replace(
      /([\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF])\s+([\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF])/g,
      '$1$2'
    );
    base = base.replace(/\s*([。、・「」『』（）〔〕［］])\s*/g, '$1');
  } else {
    base = base.replace(/\s*\n\s*/g, ' ').replace(/\s+/g, ' ');
  }
  return base.trim();
}

/* -----------------------------------------------------------------------------
  Handwritten detection via Layout.styles spans
----------------------------------------------------------------------------- */
function spansOverlap(spanA, spanB) {
  const aStart = spanA.offset;
  const aEnd = spanA.offset + spanA.length;
  const bStart = spanB.offset;
  const bEnd = spanB.offset + spanB.length;
  return Math.min(aEnd, bEnd) > Math.max(aStart, bStart);
}

function isLineHandwritten(lineSpans = [], handwrittenStyleSpans = []) {
  if (!lineSpans.length || !handwrittenStyleSpans.length) return false;
  for (const ls of lineSpans) {
    for (const hs of handwrittenStyleSpans) {
      if (spansOverlap(ls, hs)) return true;
    }
  }
  return false;
}

/* -----------------------------------------------------------------------------
  Group by table cells
----------------------------------------------------------------------------- */
function groupByTableCells(layoutResult, structuredOutput, context) {
  if (!layoutResult.tables || !layoutResult.tables.length) {
    logMessage('No tables detected by layout model.', context);
    return null;
  }
  const handwrittenStyleSpans = (layoutResult.styles || [])
    .filter(st => st?.isHandwritten)
    .flatMap(st => st.spans || []);
  logMessage(`Handwritten style spans count: ${handwrittenStyleSpans.length}`, context);

  const used = new Set();
  const groups = [];

  logMessage(`Tables detected: ${layoutResult.tables.length}`, context);

  layoutResult.tables.forEach((table, tIdx) => {
    logMessage(` Table #${tIdx} cells: ${table.cells?.length ?? 0}`, context);

    (table.cells || []).forEach((cell, cIdx) => {
      const region = cell.boundingRegions?.[0];
      const cellPolygon = region?.polygon;
      if (!cellPolygon || cellPolygon.length === 0) {
        logMessage(` Cell #${cIdx}: missing polygon, skipped.`, context);
        return;
      }

      const cellBBox = getBoundingBoxFromPolygon(cellPolygon);

      const linesInCell = structuredOutput
        .map((line, idx) => ({ line, idx }))
        .filter(({ line, idx }) => !used.has(idx) && bboxOverlap(line.bbox, cellBBox));

      if (linesInCell.length) {
        linesInCell.forEach(({ idx }) => used.add(idx));
        const mergedText = linesInCell.map(({ line }) => line.layoutText).join('\n');
        const mergedOCR = linesInCell.flatMap(({ line }) => line.matchedOCR);

        const xs = linesInCell.flatMap(({ line }) => [line.bbox[0], line.bbox[2]]);
        const ys = linesInCell.flatMap(({ line }) => [line.bbox[1], line.bbox[3]]);
        const bbox = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];

        const angles = linesInCell
          .map(({ line }) => angleFromPolygon(line.polygon))
          .filter(a => a != null && Number.isFinite(a));
        const medianAngle = median(angles);
        const snappedDeg = snapAngle0or90(medianAngle) || 0;
        const groupIsHandwritten = linesInCell.some(({ line }) => line.isHandwritten);

        const displayText = toSingleMergedText(mergedText, mergedOCR);
        groups.push({
          layoutText: mergedText,
          matchedOCR: mergedOCR,
          displayText,
          bbox,
          polygon: cellPolygon,
          orientationDeg: snappedDeg,
          isHandwritten: groupIsHandwritten
        });

        logMessage(
          ` Cell #${cIdx}: grouped ${linesInCell.length} lines; `
          + `bbox=${bbox.map(n => n.toFixed(1)).join(', ')}, `
          + `ori=${snappedDeg}°, handwritten=${groupIsHandwritten}`,
          context
        );
      }
    });
  });

  logMessage(`Grouped by table cells: ${groups.length} groups`, context);
  return groups.length ? { groups, used } : null;
}

/* -----------------------------------------------------------------------------
  Fallback: group bounding boxes by adjacency/alignment
----------------------------------------------------------------------------- */
function groupBoundingBoxes(boxes, yThreshold = 40, xThreshold = 30, context) {
  logMessage(`Fallback grouping: yThreshold=${yThreshold}, xThreshold=${xThreshold}`, context);

  boxes.sort((a, b) => a.bbox[1] - b.bbox[1]);
  const groups = [];
  let currentGroup = [];

  for (let i = 0; i < boxes.length; i++) {
    if (currentGroup.length === 0) {
      currentGroup.push(boxes[i]);
      continue;
    }
    const prev = currentGroup[currentGroup.length - 1];
    const curr = boxes[i];

    const verticalGap = curr.bbox[1] - prev.bbox[3];

    const leftAligned = Math.abs(curr.bbox[0] - prev.bbox[0]) < xThreshold;
    const rightAligned = Math.abs(curr.bbox[2] - prev.bbox[2]) < xThreshold;
    const centerPrev = (prev.bbox[0] + prev.bbox[2]) / 2;
    const centerCurr = (curr.bbox[0] + curr.bbox[2]) / 2;
    const centerAligned = Math.abs(centerCurr - centerPrev) < xThreshold;

    const horizontallyAligned = leftAligned || rightAligned || centerAligned;

    if (verticalGap < yThreshold && horizontallyAligned) {
      currentGroup.push(curr);
    } else {
      groups.push(currentGroup);
      currentGroup = [curr];
    }
  }
  if (currentGroup.length > 0) groups.push(currentGroup);

  logMessage(`Fallback grouping produced ${groups.length} groups.`, context);
  groups.forEach((g, idx) => logMessage(` Group #${idx} size=${g.length}`, context));

  return groups;
}

function mergeGroups(groups, context) {
  const merged = groups.map(group => {
    const allText = group.map(g => g.layoutText).join('\n');
    const allMatchedOCR = group.flatMap(g => g.matchedOCR);

    const xs = group.flatMap(g => [g.bbox[0], g.bbox[2]]);
    const ys = group.flatMap(g => [g.bbox[1], g.bbox[3]]);
    const bbox = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];

    const angles = group
      .map(g => angleFromPolygon(g.polygon))
      .filter(a => a != null && Number.isFinite(a));
    const medianAngle = median(angles);
    const snappedDeg = snapAngle0or90(medianAngle) || 0;

    const displayText = toSingleMergedText(allText, allMatchedOCR);
    const groupIsHandwritten = group.some(g => g.isHandwritten);

    return {
      layoutText: allText,
      matchedOCR: allMatchedOCR,
      displayText,
      bbox,
      polygon: group[0].polygon,
      orientationDeg: snappedDeg,
      isHandwritten: groupIsHandwritten
    };
  });

  logMessage(`Merged ${merged.length} groups into consolidated entries.`, context);
  return merged;
}

/* -----------------------------------------------------------------------------
  Analysis pipeline
----------------------------------------------------------------------------- */
async function analyseAndExtract(buffer, mimeType, context) {
  try {
    logMessage('Starting analyseAndExtract...', context);
    console.time('analyseAndExtract');

    // 1) Layout (pages, tables, styles, spans, etc.)
    logMessage('Calling Azure prebuilt-layout...', context);
    const layoutPoller = await client.beginAnalyzeDocument("prebuilt-layout", buffer, { contentType: mimeType });
    const layoutResult = await layoutPoller.pollUntilDone();

    // 2) Read (OCR lines—PRIMARY SOURCE for text content)
    logMessage('Calling Azure prebuilt-read...', context);
    const readPoller = await client.beginAnalyzeDocument("prebuilt-read", buffer, { contentType: mimeType });
    const readResult = await readPoller.pollUntilDone();

    // 3) Layout lines (include spans for handwriting mapping and structure)
    const layoutSections = (layoutResult.pages || []).flatMap((page, pIdx) =>
      (page.lines || []).map((line, lIdx) => ({
        layoutText: line.content,
        bbox: getBoundingBoxFromPolygon(line.polygon),
        polygon: line.polygon,
        spans: line.spans || [],
        pageIndex: pIdx,
        lineIndex: lIdx
      }))
    );
    logMessage(`Layout pages: ${(layoutResult.pages || []).length}, layout lines: ${layoutSections.length}`, context);

    // 4) OCR lines (PRIMARY TEXT SOURCE)
    const ocrLines = (readResult.pages || []).flatMap((page, pIdx) =>
      (page.lines || []).map((line, lIdx) => ({
        text: line.content,
        bbox: getBoundingBoxFromPolygon(line.polygon),
        polygon: line.polygon,
        pageIndex: pIdx,
        lineIndex: lIdx
      }))
    );
    logMessage(`Read pages: ${(readResult.pages || []).length}, OCR lines: ${ocrLines.length}`, context);

    // 4.1) Handwritten style spans from Layout
    const handwrittenStyleSpans = (layoutResult.styles || [])
      .filter(st => st?.isHandwritten)
      .flatMap(st => st.spans || []);
    logMessage(`Handwritten style spans (layout.styles): ${handwrittenStyleSpans.length}`, context);

    // 5) Cross-reference Layout to OCR (IoU > 0.1) and prioritize OCR text
    const IOU_THRESHOLD = 0.1;
    logMessage(`Cross-referencing Layout->OCR with IoU threshold ${IOU_THRESHOLD}...`, context);

    const structuredOutput = layoutSections.map(section => {
      const matchedOCRLines = ocrLines.filter(ocr => calculateIoU(section.bbox, ocr.bbox) > IOU_THRESHOLD);
      const matchedTexts = matchedOCRLines.map(ocr => ocr.text);
      const isHW = isLineHandwritten(section.spans || [], handwrittenStyleSpans);

      return {
        layoutText: section.layoutText,
        matchedOCR: matchedTexts,
        bbox: section.bbox,
        polygon: section.polygon,
        spans: section.spans || [],
        isHandwritten: isHW,
        pageIndex: section.pageIndex
      };
    });

    logMessage(`Structured output lines: ${structuredOutput.length}`, context);

    // 6) Table-cell grouping + leftover headers outside the table
    let merged = [];
    const tableGrouping = groupByTableCells(layoutResult, structuredOutput, context);
    if (tableGrouping) {
      merged.push(...tableGrouping.groups);

      const leftovers = structuredOutput.filter((_, idx) => !tableGrouping.used.has(idx));
      logMessage(`Leftover lines outside tables: ${leftovers.length}`, context);

      if (leftovers.length) {
        const groupedLeftovers = groupBoundingBoxes(leftovers, 40, 30, context);
        const mergedLeftovers = mergeGroups(groupedLeftovers, context);
        merged.push(...mergedLeftovers);
        logMessage(`Appended ${mergedLeftovers.length} leftover groups (non-table regions).`, context);
      }
    } else {
      const grouped = groupBoundingBoxes(structuredOutput, 40, 30, context);
      merged = mergeGroups(grouped, context);
    }

    console.timeEnd('analyseAndExtract');
    logMessage(`Returning ${merged.length} merged entries.`, context);

    return merged;
  } catch (error) {
    handleError(error, 'analyseAndExtract', context);
    return null;
  }
}

/* -----------------------------------------------------------------------------
  Complete processing pipeline for unknown file types
----------------------------------------------------------------------------- */
async function processUnknownDocument(imageBuffer, mimeType, base64Raw, originalFileName, companyName, context) {
  try {
    logMessage(`🧠 Starting complete document processing pipeline`, context);
    logMessage(`📄 File: ${originalFileName}, Company: ${companyName}, MIME: ${mimeType}`, context);

    // Step 1: Extract and analyze the document
    logMessage(`📖 Starting document analysis...`, context);
    const analyseOutput = await analyseAndExtract(imageBuffer, mimeType, context);

    if (!analyseOutput || analyseOutput.length === 0) {
      logMessage(`❌ No text regions detected in the document`, context);
      return {
        success: false,
        reason: 'no_text_detected',
        textRegions: 0,
        uploads: { original: false, json: false, annotatedImage: false }
      };
    }

    logMessage(`✅ Analysis complete! Found ${analyseOutput.length} text regions`, context);
    
    // Log sample extracted text for debugging
    if (analyseOutput.length > 0) {
      logMessage(`📝 Sample extracted text regions:`, context);
      analyseOutput.slice(0, 3).forEach((region, idx) => {
        const text = region.displayText || '';
        const handwritten = region.isHandwritten ? '✍️' : '🖨️';
        logMessage(`  [${idx}] ${handwritten} "${text.slice(0, 50)}${text.length > 50 ? '...' : ''}"`, context);
      });
    }

    // Prepare SharePoint variables
    const baseName = originalFileName.replace(/\.[^/.]+$/, "");
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    
    // ✅ Better readable format: YYYY-MM-DD_HH-MM-SS-mmm
    const currentDateUTC = new Date();
    const currentDateJST = new Date(currentDateUTC.getTime() + (9 * 60 * 60 * 1000));
    const dateFolder = [
      currentDateJST.getUTCFullYear(),
      String(currentDateJST.getUTCMonth() + 1).padStart(2, '0'),
      String(currentDateJST.getUTCDate()).padStart(2, '0')
    ].join('-') + '_' + [
      String(currentDateJST.getUTCHours()).padStart(2, '0'),
      String(currentDateJST.getUTCMinutes()).padStart(2, '0'),
      String(currentDateJST.getUTCSeconds()).padStart(2, '0'),
      String(currentDateJST.getUTCMilliseconds()).padStart(3, '0')
    ].join('-');
    
    const basePath = process.env.SHAREPOINT_ETC_FOLDER_PATH?.replace(/^\/+|\/+$/g, '') || 'その他';
    const folderPath = `${basePath}/${companyName}/${dateFolder}`;
    logMessage(`📁 Target SharePoint folder (JST): ${folderPath}`, context);

    // Import SharePoint functions
    const { ensureSharePointFolder, uploadJsonToSharePoint, uploadOriginalDocumentToSharePoint } = require('../sharepoint/sendToSharePoint');
    
    // Ensure folder exists
    await ensureSharePointFolder(folderPath, context);

    const uploadResults = { original: false, json: false, annotatedImage: false };

    // Step 2: Upload original document
    logMessage(`📤 Uploading original document to SharePoint...`, context);
    const originalDocFileName = `original-${originalFileName}`;
    
    try {
      const originalDocUploadResult = await uploadOriginalDocumentToSharePoint(
        base64Raw, originalDocFileName, folderPath, context
      );
      
      if (originalDocUploadResult) {
        logMessage(`✅ Successfully uploaded original document: ${originalDocFileName}`, context);
        uploadResults.original = true;
      } else {
        logMessage(`⚠️ Failed to upload original document, but continuing...`, context);
      }
    } catch (error) {
      logMessage(`❌ Error uploading original document: ${error.message}`, context);
    }

    // Step 3: Upload JSON analysis
    logMessage(`📤 Uploading analysis JSON to SharePoint...`, context);
    
    const analysisJsonReport = {
      metadata: {
        originalFileName, processedDate: new Date().toISOString(),
        companyName, mimeType,
        totalTextRegions: analyseOutput.length,
        handwrittenRegions: analyseOutput.filter(region => region.isHandwritten).length,
        printedRegions: analyseOutput.filter(region => !region.isHandwritten).length,
        documentType: 'unknown_general_extraction'
      },
      extractedData: {
        textRegions: analyseOutput,
        extractionMethod: 'Azure Document Intelligence (Layout + Read)',
        ocrPriority: true,
        tableDetection: analyseOutput.some(region => region.bbox),
        handwritingDetection: analyseOutput.some(region => region.isHandwritten)
      },
      summary: {
        extractedTextSample: analyseOutput.slice(0, 5).map((region, idx) => ({
          regionIndex: idx,
          text: region.displayText?.slice(0, 100) + (region.displayText?.length > 100 ? '...' : ''),
          isHandwritten: region.isHandwritten,
          boundingBox: region.bbox,
          orientation: region.orientationDeg
        }))
      }
    };
    
    const jsonFileName = `一般抽出データ-${baseName}-${timestamp}.json`;
    
    try {
      const jsonUploadResult = await uploadJsonToSharePoint(
        analysisJsonReport, jsonFileName, folderPath, context
      );
      
      if (jsonUploadResult) {
        logMessage(`✅ Successfully uploaded analysis JSON: ${jsonFileName}`, context);
        uploadResults.json = true;
      } else {
        logMessage(`⚠️ Failed to upload analysis JSON, but continuing...`, context);
      }
    } catch (error) {
      logMessage(`❌ Error uploading analysis JSON: ${error.message}`, context);
    }

    // Step 4: Generate and upload annotated image
    logMessage(`🖼️ Generating and uploading annotated image to SharePoint...`, context);
    
    try {
      const sharePointResult = await generateAnnotatedImageToSharePoint(
        analyseOutput, imageBuffer, originalFileName, context, companyName, folderPath
      );

      if (sharePointResult) {
        logMessage(`✅ Successfully uploaded annotated image to SharePoint`, context);
        uploadResults.annotatedImage = true;
      } else {
        logMessage(`⚠️ Failed to upload annotated image`, context);
      }
    } catch (error) {
      logMessage(`❌ Error uploading annotated image: ${error.message}`, context);
    }

    // Return comprehensive results
    const successfulUploads = Object.values(uploadResults).filter(Boolean).length;
    const totalUploads = Object.keys(uploadResults).length;
    
    logMessage(`📊 Processing complete! Text regions: ${analyseOutput.length}, Uploads: ${successfulUploads}/${totalUploads}`, context);
    
    return {
      success: true,
      textRegions: analyseOutput.length,
      handwrittenRegions: analyseOutput.filter(region => region.isHandwritten).length,
      printedRegions: analyseOutput.filter(region => !region.isHandwritten).length,
      uploads: uploadResults,
      sharePointFolder: folderPath,
      analysisData: analyseOutput
    };

  } catch (error) {
    handleError(error, 'processUnknownDocument', context);
    return {
      success: false,
      reason: 'processing_error',
      error: error.message,
      textRegions: 0,
      uploads: { original: false, json: false, annotatedImage: false }
    };
  }
}

/* -----------------------------------------------------------------------------
  Exports (CommonJS)
----------------------------------------------------------------------------- */
module.exports = {
  analyseAndExtract,
  generateAnnotatedImage,        // ✅ Re-export from generalFormImageAnnotator
  generateAnnotatedImageToSharePoint, // ✅ Re-export from generalFormImageAnnotator
  processUnknownDocument
};