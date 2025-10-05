'use strict';

if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const { AzureKeyCredential, DocumentAnalysisClient } = require("@azure/ai-form-recognizer");
const { createCanvas, loadImage } = require('canvas');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { logMessage, handleError } = require('../utils');

/* -----------------------------------------------------------------------------
  Debug + rendering flags (env)
----------------------------------------------------------------------------- */
const DEBUG = process.env.DEBUG === 'true' || process.env.NODE_ENV !== 'production';

// Orientation & wrapping
const FORCE_HORIZONTAL = process.env.DISPLAY_FORCE_HORIZONTAL === 'true'; // force 0°
const ENABLE_WRAP = process.env.DISPLAY_WRAP !== 'false';                 // default: true

// Unlimited wrapping by default (bounded by height). If set >0, hard cap.
const WRAP_MAX_LINES_RAW = Number.parseInt(process.env.WRAP_MAX_LINES || '0', 10);
const WRAP_MAX_LINES = Number.isFinite(WRAP_MAX_LINES_RAW) && WRAP_MAX_LINES_RAW > 0
  ? WRAP_MAX_LINES_RAW
  : Infinity;

const WRAP_LINE_HEIGHT_MULT = Number.parseFloat(process.env.WRAP_LINE_HEIGHT_MULT || '1.15'); // tighter leading

// Font & fill sizing
const MAX_FONT_SIZE = Number.parseInt(process.env.MAX_FONT_SIZE || '72', 10);
const MIN_FONT_SIZE = Number.parseInt(process.env.MIN_FONT_SIZE || '12', 10);
const FILL_RATIO   = Number.parseFloat(process.env.FILL_RATIO   || '0.98'); // fill more of the box

// Colors (category base colors)
const PRINTED_FILL_COLOR     = process.env.FILL_PRINTED_COLOR     || '#1976D2'; // Blue 700
const HANDWRITTEN_FILL_COLOR = process.env.FILL_HANDWRITTEN_COLOR || '#2E7D32'; // Green 700

// Text colors (configurable via environment variables)
const TEXT_COLOR = process.env.TEXT_COLOR || '#FFFFFF';           // White text by default
const TEXT_OUTLINE_COLOR = process.env.TEXT_OUTLINE_COLOR || '#000000'; // Black outline by default

// Opacity (box fill transparency ~30% by default)
const BOX_ALPHA_RAW = Number.parseFloat(process.env.BOX_ALPHA || '0.3');
const BOX_ALPHA = Number.isFinite(BOX_ALPHA_RAW) ? Math.max(0, Math.min(1, BOX_ALPHA_RAW)) : 0.3;

// Orientation snapping (0° or 90°)
const ORIENTATION_SNAP = process.env.ORIENTATION_SNAP !== 'false'; // default true

/* -----------------------------------------------------------------------------
  Azure setup
----------------------------------------------------------------------------- */
const endpoint = process.env['CLASSIFIER_ENDPOINT'];
const apiKey   = process.env['CLASSIFIER_ENDPOINT_AZURE_API_KEY'];

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
  let deg = rad * 180 / Math.PI; // [-180, 180]
  deg = ((deg % 180) + 180) % 180;
  return deg;
}

// Median of numeric array
function median(arr) {
  if (!arr || arr.length === 0) return null;
  const a = arr.slice().sort((x, y) => x - y);
  const mid = Math.floor(a.length / 2);
  return a.length % 2 ? a[mid] : (a[mid - 1] + a[mid]) / 2;
}

// Snap to closest of 0° or 90°
function snapAngle0or90(deg) {
  if (deg == null || Number.isNaN(deg)) return 0;
  const d0  = Math.min(deg, 180 - deg);
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
 * - For CJK: remove line-breaks WITHOUT adding spaces; strip spaces between adjacent CJK chars.
 * - For Latin: collapse newlines to single spaces and normalize whitespace.
 */
function toSingleMergedText(layoutText, matchedOCRArr) {
  const rawLayout = (layoutText || '').trim();
  const rawOCR    = (matchedOCRArr || []).join(' ').trim();
  
  // ✅ FIXED: Prioritize OCR, fallback to Layout only if OCR is empty
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
  const aEnd   = spanA.offset + spanA.length;
  const bStart = spanB.offset;
  const bEnd   = spanB.offset + spanB.length;
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
  Group by table cells (returns { groups, used } to allow leftover handling)
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
    logMessage(`  Table #${tIdx} cells: ${table.cells?.length ?? 0}`, context);

    (table.cells || []).forEach((cell, cIdx) => {
      const region = cell.boundingRegions?.[0];
      const cellPolygon = region?.polygon;
      if (!cellPolygon || cellPolygon.length === 0) {
        logMessage(`    Cell #${cIdx}: missing polygon, skipped.`, context);
        return;
      }
      const cellBBox = getBoundingBoxFromPolygon(cellPolygon);

      const linesInCell = structuredOutput
        .map((line, idx) => ({ line, idx }))
        .filter(({ line, idx }) => !used.has(idx) && bboxOverlap(line.bbox, cellBBox));

      if (linesInCell.length) {
        linesInCell.forEach(({ idx }) => used.add(idx));

        const mergedText = linesInCell.map(({ line }) => line.layoutText).join('\n');
        const mergedOCR  = linesInCell.flatMap(({ line }) => line.matchedOCR);
        const xs         = linesInCell.flatMap(({ line }) => [line.bbox[0], line.bbox[2]]);
        const ys         = linesInCell.flatMap(({ line }) => [line.bbox[1], line.bbox[3]]);
        const bbox       = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];

        const angles = linesInCell
          .map(({ line }) => angleFromPolygon(line.polygon))
          .filter(a => a != null && Number.isFinite(a));
        const medianAngle = median(angles);
        const snappedDeg  = ORIENTATION_SNAP ? snapAngle0or90(medianAngle) : (medianAngle ?? 0);

        const groupIsHandwritten = linesInCell.some(({ line }) => line.isHandwritten);
        const displayText = toSingleMergedText(mergedText, mergedOCR);

        groups.push({
          layoutText: mergedText,
          matchedOCR: mergedOCR,
          displayText,
          bbox,
          polygon: cellPolygon,
          orientationDeg: snappedDeg,        // 0 or 90 by default
          isHandwritten: groupIsHandwritten  // boolean
        });

        logMessage(`    Cell #${cIdx}: grouped ${linesInCell.length} lines; bbox=${bbox.map(n => n.toFixed(1)).join(', ')}, ori=${snappedDeg}°, handwritten=${groupIsHandwritten}`, context);
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
  boxes.sort((a, b) => a.bbox[1] - b.bbox[1]); // sort by top y
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

    const leftAligned  = Math.abs(curr.bbox[0] - prev.bbox[0]) < xThreshold;
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
  groups.forEach((g, idx) => logMessage(`  Group #${idx} size=${g.length}`, context));
  return groups;
}

function mergeGroups(groups, context) {
  const merged = groups.map(group => {
    const allText       = group.map(g => g.layoutText).join('\n');
    const allMatchedOCR = group.flatMap(g => g.matchedOCR);
    const xs            = group.flatMap(g => [g.bbox[0], g.bbox[2]]);
    const ys            = group.flatMap(g => [g.bbox[1], g.bbox[3]]);
    const bbox          = [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];

    const angles      = group.map(g => angleFromPolygon(g.polygon)).filter(a => a != null && Number.isFinite(a));
    const medianAngle = median(angles);
    const snappedDeg  = ORIENTATION_SNAP ? snapAngle0or90(medianAngle) : (medianAngle ?? 0);

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
        spans: line.spans || [], // offset/length spans
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
      // Find OCR lines that overlap with this layout section
      const matchedOCRLines = ocrLines.filter(ocr => 
        calculateIoU(section.bbox, ocr.bbox) > IOU_THRESHOLD
      );
      
      // Extract text from matched OCR lines
      const matchedTexts = matchedOCRLines.map(ocr => ocr.text);

      const isHW = isLineHandwritten(section.spans || [], handwrittenStyleSpans);

      // Log OCR vs Layout comparison for debugging
      if (matchedTexts.length > 0) {
        const ocrText = matchedTexts.join(' ').trim();
        const layoutText = section.layoutText.trim();
        if (ocrText !== layoutText && DEBUG) {
          logMessage(`Text difference detected:`, context);
          logMessage(`  Layout: "${layoutText}"`, context);
          logMessage(`  OCR:    "${ocrText}"`, context);
          logMessage(`  Using:  "${ocrText}" (OCR priority)`, context);
        }
      }

      return {
        layoutText: section.layoutText,    // Keep for fallback
        matchedOCR: matchedTexts,         // Primary text source
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
      // (6a) add cell-based groups
      merged.push(...tableGrouping.groups);

      // (6b) group leftover lines (headers/titles/etc.) not inside any cell
      const leftovers = structuredOutput.filter((_, idx) => !tableGrouping.used.has(idx));
      logMessage(`Leftover lines outside tables: ${leftovers.length}`, context);

      if (leftovers.length) {
        const groupedLeftovers = groupBoundingBoxes(leftovers, 40, 30, context);
        const mergedLeftovers  = mergeGroups(groupedLeftovers, context);
        merged.push(...mergedLeftovers);
        logMessage(`Appended ${mergedLeftovers.length} leftover groups (non-table regions).`, context);
      }
    } else {
      // No table detected — use spatial grouping for all lines
      const grouped = groupBoundingBoxes(structuredOutput, 40, 30, context);
      merged = mergeGroups(grouped, context);
    }

    console.timeEnd('analyseAndExtract');
    logMessage(`Returning ${merged.length} merged entries.`, context);
    
    // ✅ Enhanced logging to show OCR priority in action
    if (DEBUG) {
      logMessage('Sample merged entries with OCR priority:', context);
      merged.slice(0, 3).forEach((entry, idx) => {
        const hasOCR = entry.matchedOCR && entry.matchedOCR.length > 0;
        const ocrText = hasOCR ? entry.matchedOCR.join(' ').trim() : '';
        const layoutText = entry.layoutText || '';
        const usedText = entry.displayText || '';
        
        logMessage(`  [${idx}] OCR: "${ocrText}" | Layout: "${layoutText}" | Used: "${usedText}" | Source: ${hasOCR ? 'OCR' : 'Layout'}`, context);
      });
    }
    
    return merged;
  } catch (error) {
    handleError(error, 'analyseAndExtract', context);
    return null;
  }
}

/* -----------------------------------------------------------------------------
  Text fitting & wrapping (height-driven, unlimited lines by default)
----------------------------------------------------------------------------- */
function fitTextToBox(ctx, text, effWidth, effHeight, maxFontSize = MAX_FONT_SIZE, minFontSize = MIN_FONT_SIZE) {
  let fontSize = maxFontSize;
  while (fontSize >= minFontSize) {
    ctx.font = `${fontSize}px sans-serif`;
    const w = ctx.measureText(text).width;
    const h = fontSize * WRAP_LINE_HEIGHT_MULT;
    if (w <= effWidth * FILL_RATIO && h <= effHeight * FILL_RATIO) break;
    fontSize -= 2;
  }
  if (fontSize < minFontSize) fontSize = minFontSize;
  if (fontSize > maxFontSize) fontSize = maxFontSize;
  return fontSize;
}

function wrapTextToBox(ctx, text, effWidth, effHeight, maxFontSize = MAX_FONT_SIZE, minFontSize = MIN_FONT_SIZE) {
  const isCJK = isMostlyCJK(text);

  function buildLines(fontSize) {
    ctx.font = `${fontSize}px sans-serif`;
    const maxWidth = effWidth * FILL_RATIO;
    const lines = [];

    if (isCJK) {
      let current = '';
      for (const ch of text) {
        const trial = current + ch;
        const w = ctx.measureText(trial).width;
        if (w <= maxWidth) {
          current = trial;
        } else {
          lines.push(current);
          current = ch;
        }
      }
      if (current) lines.push(current);
    } else {
      const words = text.split(/\s+/).filter(Boolean);
      let current = '';
      for (const word of words) {
        const trial = current ? `${current} ${word}` : word;
        const w = ctx.measureText(trial).width;
        if (w <= maxWidth) {
          current = trial;
        } else {
          lines.push(current);
          current = word;
        }
      }
      if (current) lines.push(current);
    }

    return lines;
  }

  let fontSize = maxFontSize;
  while (fontSize >= minFontSize) {
    const lines = buildLines(fontSize);
    const lineHeight  = fontSize * WRAP_LINE_HEIGHT_MULT;
    const totalHeight = lines.length * lineHeight;

    if (lines.length > 0 && totalHeight <= effHeight * FILL_RATIO && lines.length <= WRAP_MAX_LINES) {
      return { fontSize, lines };
    }

    fontSize -= 2;
  }

  const lines = buildLines(minFontSize);
  const clipped = lines.slice(0, WRAP_MAX_LINES);
  return { fontSize: minFontSize, lines: clipped };
}

/* -----------------------------------------------------------------------------
  Annotated image generation with SharePoint upload
----------------------------------------------------------------------------- */
/**
 * Generate annotated image and upload to SharePoint using existing SharePoint functions
 */
async function generateAnnotatedImageToSharePoint(analyseOutput, originalImageBuffer, originalFileName, context, companyName) {
  if (!Array.isArray(analyseOutput)) {
    handleError(new Error('Invalid analyseOutput input'), 'generateAnnotatedImageToSharePoint', context);
    return null;
  }

  try {
    logMessage(`🖼️ Generating annotated image for SharePoint upload...`, context);
    logMessage(`Render flags: FORCE_HORIZONTAL=${FORCE_HORIZONTAL}, ENABLE_WRAP=${ENABLE_WRAP}, WRAP_MAX_LINES=${WRAP_MAX_LINES === Infinity ? '∞' : WRAP_MAX_LINES}, WRAP_LINE_HEIGHT_MULT=${WRAP_LINE_HEIGHT_MULT}, MAX_FONT_SIZE=${MAX_FONT_SIZE}, FILL_RATIO=${FILL_RATIO}, BOX_ALPHA=${BOX_ALPHA}`, context);
    logMessage(`Colors: PRINTED=${PRINTED_FILL_COLOR}, HANDWRITTEN=${HANDWRITTEN_FILL_COLOR}, TEXT=${TEXT_COLOR}, OUTLINE=${TEXT_OUTLINE_COLOR}`, context);

    // Load image from buffer
    const image = await loadImage(originalImageBuffer);
    const canvas = createCanvas(image.width, image.height);
    const ctx = canvas.getContext('2d');
    ctx.drawImage(image, 0, 0, image.width, image.height);

    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';

    logMessage(`Drawing ${analyseOutput.length} merged entries...`, context);
    analyseOutput.forEach((entry, idx) => {
      const bbox    = entry.bbox;
      const polygon = entry.polygon;
      if (bbox && bbox.length === 4 && polygon && polygon.length >= 2) {
        const boxWidth  = bbox[2] - bbox[0];
        const boxHeight = bbox[3] - bbox[1];

        const isHW = !!entry.isHandwritten;
        const baseColor = isHW ? HANDWRITTEN_FILL_COLOR : PRINTED_FILL_COLOR;

        // Fill (with transparency)
        ctx.save();
        ctx.globalAlpha = BOX_ALPHA;
        ctx.fillStyle   = baseColor;
        ctx.fillRect(bbox[0], bbox[1], boxWidth, boxHeight);
        ctx.restore();

        // Border (opaque)
        ctx.save();
        ctx.strokeStyle = baseColor;
        ctx.lineWidth   = 3;
        ctx.strokeRect(bbox[0], bbox[1], boxWidth, boxHeight);
        ctx.restore();

        // Text rendering logic...
        const computedDeg = (entry.orientationDeg != null) ? entry.orientationDeg : angleFromPolygon(polygon) || 0;
        const angleDeg    = FORCE_HORIZONTAL ? 0 : (ORIENTATION_SNAP ? snapAngle0or90(computedDeg) : computedDeg);
        const angleRad    = angleDeg * Math.PI / 180;

        let effW = boxWidth;
        let effH = boxHeight;
        if (Math.abs(Math.round(angleDeg)) === 90) {
          effW = boxHeight;
          effH = boxWidth;
        }

        const text = (entry.displayText || '').trim();

        let fontSize, lines;
        if (ENABLE_WRAP) {
          const wrapped = wrapTextToBox(ctx, text, effW, effH, MAX_FONT_SIZE, MIN_FONT_SIZE);
          fontSize = wrapped.fontSize;
          lines    = wrapped.lines;
        } else {
          fontSize = fitTextToBox(ctx, text, effW, effH, MAX_FONT_SIZE, MIN_FONT_SIZE);
          lines    = [text];
        }

        ctx.save();
        ctx.font        = `bold ${fontSize}px sans-serif`;
        ctx.fillStyle   = TEXT_COLOR;
        ctx.strokeStyle = TEXT_OUTLINE_COLOR;
        ctx.lineWidth   = Math.max(1, fontSize / 20);
        ctx.textBaseline= 'middle';
        ctx.textAlign   = 'center';

        const centerX = bbox[0] + boxWidth  / 2;
        const centerY = bbox[1] + boxHeight / 2;

        ctx.translate(centerX, centerY);
        ctx.rotate(angleRad);

        const lineHeight  = fontSize * WRAP_LINE_HEIGHT_MULT;
        const totalHeight = lineHeight * (lines.length - 1);
        lines.forEach((line, i) => {
          const y = i * lineHeight - totalHeight / 2;
          if (TEXT_OUTLINE_COLOR !== TEXT_COLOR) {
            ctx.strokeText(line, 0, y);
          }
          ctx.fillText(line, 0, y);
        });

        ctx.restore();

        logMessage(`  [${idx}] bbox=(${bbox.map(n => n.toFixed(1)).join(', ')}) ori=${angleDeg}° effW=${effW.toFixed(1)} effH=${effH.toFixed(1)} font=BOLD ${fontSize}px lines=${lines.length} handwritten=${isHW} text="${text.slice(0,120)}${text.length>120?'…':''}"`, context);
      } else {
        logMessage(`  [${idx}] Skipped: invalid bbox or polygon.`, context);
      }
    });

    // Generate PNG buffer
    const pngBuffer = canvas.toBuffer('image/png');
    logMessage(`✅ Generated annotated image: ${pngBuffer.length} bytes`, context);

    // Prepare SharePoint upload
    const baseName = path.basename(originalFileName, path.extname(originalFileName));
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const annotatedFileName = `${baseName}_ANNOTATED_${timestamp}.png`;

    // Create SharePoint folder path for general form extractions
    const basePath = process.env.SHAREPOINT_ETC_FOLDER_PATH?.replace(/^\/+|\/+$/g, '') || 'その他';
    
    const folderPath = `${basePath}/${companyName}`;

    logMessage(`📁 Target SharePoint folder: ${folderPath}`, context);

    // ✅ CORRECT: Import the functions that actually exist in sendToSharePoint.js
    const { ensureSharePointFolder, uploadOriginalDocumentToSharePoint } = require('../sharepoint/sendToSharePoint');

    // Ensure folder exists
    await ensureSharePointFolder(folderPath, context);

    // Convert PNG buffer to base64 for upload
    const base64ImageContent = pngBuffer.toString('base64');

    // Upload annotated image using existing SharePoint function
    logMessage(`📤 Uploading annotated image to SharePoint: ${annotatedFileName}`, context);
    
    const sharePointResult = await uploadOriginalDocumentToSharePoint(
      base64ImageContent,
      annotatedFileName,
      folderPath,
      context
    );

    if (sharePointResult) {
      logMessage(`✅ Successfully uploaded annotated image to SharePoint: ${annotatedFileName}`, context);
      return sharePointResult;
    } else {
      logMessage(`❌ Failed to upload annotated image to SharePoint`, context);
      return null;
    }

  } catch (error) {
    handleError(error, 'generateAnnotatedImageToSharePoint', context);
    return null;
  }
}

/**
 * Legacy function for local file generation (kept for backward compatibility)
 * Updated to work with buffer input instead of file path
 */
async function generateAnnotatedImage(analyseOutput, imageBuffer, originalFileName, context) {
  if (!Array.isArray(analyseOutput)) {
    handleError(new Error('Invalid input'), 'generateAnnotatedImage', context);
    return null;
  }

  try {
    // Create temporary file for local processing
    const tempDir = os.tmpdir();
    const tempImagePath = path.join(tempDir, `temp_${Date.now()}_${originalFileName}`);
    const outFile = path.join(tempDir, `${path.basename(originalFileName, path.extname(originalFileName))}_ANNOTATED.png`);

    // Write buffer to temp file
    fs.writeFileSync(tempImagePath, imageBuffer);

    logMessage(`Loading image from buffer: ${originalFileName}`, context);
    logMessage(`Render flags: FORCE_HORIZONTAL=${FORCE_HORIZONTAL}, ENABLE_WRAP=${ENABLE_WRAP}, WRAP_MAX_LINES=${WRAP_MAX_LINES === Infinity ? '∞' : WRAP_MAX_LINES}, WRAP_LINE_HEIGHT_MULT=${WRAP_LINE_HEIGHT_MULT}, MAX_FONT_SIZE=${MAX_FONT_SIZE}, FILL_RATIO=${FILL_RATIO}, BOX_ALPHA=${BOX_ALPHA}`, context);
    logMessage(`Colors: PRINTED=${PRINTED_FILL_COLOR}, HANDWRITTEN=${HANDWRITTEN_FILL_COLOR}, TEXT=${TEXT_COLOR}, OUTLINE=${TEXT_OUTLINE_COLOR}`, context);

    const image = await loadImage(tempImagePath);
    const canvas = createCanvas(image.width, image.height);
    const ctx = canvas.getContext('2d');
    ctx.drawImage(image, 0, 0, image.width, image.height);

    ctx.textBaseline = 'middle';
    ctx.textAlign = 'center';

    logMessage(`Drawing ${analyseOutput.length} merged entries...`, context);
    analyseOutput.forEach((entry, idx) => {
      const bbox    = entry.bbox;
      const polygon = entry.polygon;
      if (bbox && bbox.length === 4 && polygon && polygon.length >= 2) {
        const boxWidth  = bbox[2] - bbox[0];
        const boxHeight = bbox[3] - bbox[1];

        const isHW = !!entry.isHandwritten;
        const baseColor = isHW ? HANDWRITTEN_FILL_COLOR : PRINTED_FILL_COLOR;

        // Fill (with transparency)
        ctx.save();
        ctx.globalAlpha = BOX_ALPHA;
        ctx.fillStyle   = baseColor;
        ctx.fillRect(bbox[0], bbox[1], boxWidth, boxHeight);
        ctx.restore();

        // Border (opaque)
        ctx.save();
        ctx.strokeStyle = baseColor;
        ctx.lineWidth   = 3;
        ctx.strokeRect(bbox[0], bbox[1], boxWidth, boxHeight);
        ctx.restore();

        // Text rendering logic...
        const computedDeg = (entry.orientationDeg != null) ? entry.orientationDeg : angleFromPolygon(polygon) || 0;
        const angleDeg    = FORCE_HORIZONTAL ? 0 : (ORIENTATION_SNAP ? snapAngle0or90(computedDeg) : computedDeg);
        const angleRad    = angleDeg * Math.PI / 180;

        let effW = boxWidth;
        let effH = boxHeight;
        if (Math.abs(Math.round(angleDeg)) === 90) {
          effW = boxHeight;
          effH = boxWidth;
        }

        const text = (entry.displayText || '').trim();

        let fontSize, lines;
        if (ENABLE_WRAP) {
          const wrapped = wrapTextToBox(ctx, text, effW, effH, MAX_FONT_SIZE, MIN_FONT_SIZE);
          fontSize = wrapped.fontSize;
          lines    = wrapped.lines;
        } else {
          fontSize = fitTextToBox(ctx, text, effW, effH, MAX_FONT_SIZE, MIN_FONT_SIZE);
          lines    = [text];
        }

        ctx.save();
        ctx.font        = `bold ${fontSize}px sans-serif`;
        ctx.fillStyle   = TEXT_COLOR;
        ctx.strokeStyle = TEXT_OUTLINE_COLOR;
        ctx.lineWidth   = Math.max(1, fontSize / 20);
        ctx.textBaseline= 'middle';
        ctx.textAlign   = 'center';

        const centerX = bbox[0] + boxWidth  / 2;
        const centerY = bbox[1] + boxHeight / 2;

        ctx.translate(centerX, centerY);
        ctx.rotate(angleRad);

        const lineHeight  = fontSize * WRAP_LINE_HEIGHT_MULT;
        const totalHeight = lineHeight * (lines.length - 1);
        lines.forEach((line, i) => {
          const y = i * lineHeight - totalHeight / 2;
          if (TEXT_OUTLINE_COLOR !== TEXT_COLOR) {
            ctx.strokeText(line, 0, y);
          }
          ctx.fillText(line, 0, y);
        });

        ctx.restore();

        logMessage(`  [${idx}] bbox=(${bbox.map(n => n.toFixed(1)).join(', ')}) ori=${angleDeg}° effW=${effW.toFixed(1)} effH=${effH.toFixed(1)} font=BOLD ${fontSize}px lines=${lines.length} handwritten=${isHW} text="${text.slice(0,120)}${text.length>120?'…':''}"`, context);
      } else {
        logMessage(`  [${idx}] Skipped: invalid bbox or polygon.`, context);
      }
    });

    const out = fs.createWriteStream(outFile);
    const stream = canvas.createPNGStream();
    stream.pipe(out);

    return new Promise((resolve) => {
      out.on('finish', () => {
        // Clean up temp files
        try {
          fs.unlinkSync(tempImagePath);
        } catch (e) {
          // Ignore cleanup errors
        }
        
        logMessage(`✅ Annotated image saved as ${outFile}`, context);
        resolve(outFile);
      });
      out.on('error', (err) => {
        // Clean up temp files on error
        try {
          fs.unlinkSync(tempImagePath);
        } catch (e) {
          // Ignore cleanup errors
        }
        
        handleError(err, 'generateAnnotatedImage', context);
        resolve(null);
      });
    });
  } catch (error) {
    handleError(error, 'generateAnnotatedImage', context);
    return null;
  }
}

/* -----------------------------------------------------------------------------
  Exports (CommonJS)
----------------------------------------------------------------------------- */
module.exports = {
  analyseAndExtract,
  generateAnnotatedImage,
  generateAnnotatedImageToSharePoint
};
