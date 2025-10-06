'use strict';

const { createCanvas, loadImage, GlobalFonts } = require('@napi-rs/canvas');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { logMessage, handleError } = require('../utils');

/* -----------------------------------------------------------------------------
  Font registration & constants (Noto Sans JP)
----------------------------------------------------------------------------- */
try {
  // Register Regular & Bold under the same family name
  GlobalFonts.registerFromPath(
    path.join(__dirname, './fonts/NotoSansJP-Regular.ttf'),
    'Noto Sans JP'
  );
  GlobalFonts.registerFromPath(
    path.join(__dirname, './fonts/NotoSansJP-Bold.ttf'),
    'Noto Sans JP'
  );
} catch (e) {
  // Log but don't fail indexing if fonts are missing
  console.warn(`⚠️ Font registration failed: ${e.message}`);
}

// Use this family consistently for measure & draw
const FONT_FAMILY = '"Noto Sans JP", "Noto Sans", sans-serif';
// Bold improves legibility in overlays; adjust if needed
const DEFAULT_WEIGHT = 'bold';

/* -----------------------------------------------------------------------------
  Rendering configuration from environment variables
----------------------------------------------------------------------------- */
const DEBUG = process.env.DEBUG === 'true' || process.env.NODE_ENV !== 'production';

// Orientation & wrapping
const FORCE_HORIZONTAL = process.env.DISPLAY_FORCE_HORIZONTAL === 'true';
const ENABLE_WRAP = process.env.DISPLAY_WRAP !== 'false';

// Wrapping limits
const WRAP_MAX_LINES_RAW = Number.parseInt(process.env.WRAP_MAX_LINES || '0', 10);
const WRAP_MAX_LINES = Number.isFinite(WRAP_MAX_LINES_RAW) && WRAP_MAX_LINES_RAW > 0
  ? WRAP_MAX_LINES_RAW
  : Infinity;

const WRAP_LINE_HEIGHT_MULT = Number.parseFloat(process.env.WRAP_LINE_HEIGHT_MULT || '1.15');

// Font & fill sizing
const MAX_FONT_SIZE = Number.parseInt(process.env.MAX_FONT_SIZE || '72', 10);
const MIN_FONT_SIZE = Number.parseInt(process.env.MIN_FONT_SIZE || '12', 10);
const FILL_RATIO = Number.parseFloat(process.env.FILL_RATIO || '0.98');

// Colors
const PRINTED_FILL_COLOR = process.env.FILL_PRINTED_COLOR || '#1976D2';
const HANDWRITTEN_FILL_COLOR = process.env.FILL_HANDWRITTEN_COLOR || '#2E7D32';
const TEXT_COLOR = process.env.TEXT_COLOR || '#FFFFFF';
const TEXT_OUTLINE_COLOR = process.env.TEXT_OUTLINE_COLOR || '#000000';

// Opacity
const BOX_ALPHA_RAW = Number.parseFloat(process.env.BOX_ALPHA || '0.3');
const BOX_ALPHA = Number.isFinite(BOX_ALPHA_RAW) ? Math.max(0, Math.min(1, BOX_ALPHA_RAW)) : 0.3;

// Orientation snapping
const ORIENTATION_SNAP = process.env.ORIENTATION_SNAP !== 'false';

/* -----------------------------------------------------------------------------
  Helper functions
----------------------------------------------------------------------------- */
function isMostlyCJK(str) {
  const s = String(str || '');
  const cjkMatches = s.match(/[\u3040-\u30FF\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]/g) || [];
  return cjkMatches.length > 0 && (cjkMatches.length / s.length) > 0.3;
}

function angleFromPolygon(polygon) {
  if (!polygon || polygon.length < 2) return null;
  const dx = polygon[1].x - polygon[0].x;
  const dy = polygon[1].y - polygon[0].y;
  const rad = Math.atan2(dy, dx);
  let deg = rad * 180 / Math.PI;
  deg = ((deg % 180) + 180) % 180;
  return deg;
}

function snapAngle0or90(deg) {
  if (deg == null || Number.isNaN(deg)) return 0;
  const d0 = Math.min(deg, 180 - deg);
  const d90 = Math.abs(90 - deg);
  return d0 <= d90 ? 0 : 90;
}

/* -----------------------------------------------------------------------------
  Text fitting & wrapping functions
----------------------------------------------------------------------------- */
function fitTextToBox(ctx, text, effWidth, effHeight, maxFontSize = MAX_FONT_SIZE, minFontSize = MIN_FONT_SIZE) {
  let fontSize = maxFontSize;
  const targetW = effWidth * FILL_RATIO;
  const targetH = effHeight * FILL_RATIO;

  while (fontSize >= minFontSize) {
    ctx.font = `${DEFAULT_WEIGHT} ${fontSize}px ${FONT_FAMILY}`;
    const m = ctx.measureText(text);
    const w = m.width;
    const h = fontSize * WRAP_LINE_HEIGHT_MULT;
    if (w <= targetW && h <= targetH) break;
    fontSize -= 2;
  }
  
  if (fontSize < minFontSize) fontSize = minFontSize;
  if (fontSize > maxFontSize) fontSize = maxFontSize;
  return fontSize;
}

function wrapTextToBox(ctx, text, effWidth, effHeight, maxFontSize = MAX_FONT_SIZE, minFontSize = MIN_FONT_SIZE) {
  const isCJK = isMostlyCJK(text);

  function buildLines(fontSize) {
    ctx.font = `${DEFAULT_WEIGHT} ${fontSize}px ${FONT_FAMILY}`;
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
    const lineHeight = fontSize * WRAP_LINE_HEIGHT_MULT;
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
  Core rendering functions
----------------------------------------------------------------------------- */
/**
 * Renders text regions onto a canvas with bounding boxes and text overlays
 * @param {Object} canvas - Canvas object
 * @param {Array} analyseOutput - Array of text regions with bbox, text, etc.
 * @param {Object} context - Azure Functions context for logging
 */
function renderTextRegionsOnCanvas(canvas, analyseOutput, context) {
  const ctx = canvas.getContext('2d');
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';

  logMessage(`Drawing ${analyseOutput.length} text regions...`, context);

  analyseOutput.forEach((entry, idx) => {
    const bbox = entry.bbox;
    const polygon = entry.polygon;

    if (bbox && bbox.length === 4 && polygon && polygon.length >= 2) {
      const boxWidth = bbox[2] - bbox[0];
      const boxHeight = bbox[3] - bbox[1];

      const isHW = !!entry.isHandwritten;
      const baseColor = isHW ? HANDWRITTEN_FILL_COLOR : PRINTED_FILL_COLOR;

      // Fill (with transparency)
      ctx.save();
      ctx.globalAlpha = BOX_ALPHA;
      ctx.fillStyle = baseColor;
      ctx.fillRect(bbox[0], bbox[1], boxWidth, boxHeight);
      ctx.restore();

      // Border (opaque)
      ctx.save();
      ctx.strokeStyle = baseColor;
      ctx.lineWidth = 3;
      ctx.strokeRect(bbox[0], bbox[1], boxWidth, boxHeight);
      ctx.restore();

      // Text rendering logic
      const computedDeg = (entry.orientationDeg != null) ? entry.orientationDeg : angleFromPolygon(polygon) || 0;
      const angleDeg = FORCE_HORIZONTAL ? 0 : (ORIENTATION_SNAP ? snapAngle0or90(computedDeg) : computedDeg);
      const angleRad = angleDeg * Math.PI / 180;

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
        lines = wrapped.lines;
      } else {
        fontSize = fitTextToBox(ctx, text, effW, effH, MAX_FONT_SIZE, MIN_FONT_SIZE);
        lines = [text];
      }

      ctx.save();
      ctx.font = `${DEFAULT_WEIGHT} ${fontSize}px ${FONT_FAMILY}`;
      ctx.fillStyle = TEXT_COLOR;
      ctx.strokeStyle = TEXT_OUTLINE_COLOR;
      ctx.lineWidth = Math.max(1, fontSize / 20);
      ctx.textBaseline = 'middle';
      ctx.textAlign = 'center';

      const centerX = bbox[0] + boxWidth / 2;
      const centerY = bbox[1] + boxHeight / 2;

      ctx.translate(centerX, centerY);
      ctx.rotate(angleRad);

      const lineHeight = fontSize * WRAP_LINE_HEIGHT_MULT;
      const totalHeight = lineHeight * (lines.length - 1);
      lines.forEach((line, i) => {
        const y = i * lineHeight - totalHeight / 2;
        if (TEXT_OUTLINE_COLOR !== TEXT_COLOR) {
          ctx.strokeText(line, 0, y);
        }
        ctx.fillText(line, 0, y);
      });

      ctx.restore();

      if (DEBUG) {
        logMessage(
          `  [${idx}] bbox=(${bbox.map(n => n.toFixed(1)).join(', ')}) `
          + `ori=${angleDeg}° effW=${effW.toFixed(1)} effH=${effH.toFixed(1)} `
          + `font=${DEFAULT_WEIGHT.toUpperCase()} ${fontSize}px lines=${lines.length} `
          + `handwritten=${isHW} text="${text.slice(0, 120)}${text.length > 120 ? '…' : ''}"`,
          context
        );
      }
    } else {
      logMessage(`  [${idx}] Skipped: invalid bbox or polygon.`, context);
    }
  });
}

/**
 * Generate annotated image and upload to SharePoint
 * @param {Array} analyseOutput - Text regions with bbox, text, etc.
 * @param {Buffer} originalImageBuffer - Original image buffer
 * @param {string} originalFileName - Original filename
 * @param {Object} context - Azure Functions context
 * @param {string} companyName - Company name for folder organization
 * @param {string} folderPath - SharePoint folder path
 * @returns {Promise<Object|null>} SharePoint upload result or null
 */
async function generateAnnotatedImageToSharePoint(analyseOutput, originalImageBuffer, originalFileName, context, companyName, folderPath) {
  if (!Array.isArray(analyseOutput)) {
    handleError(new Error('Invalid analyseOutput input'), 'generateAnnotatedImageToSharePoint', context);
    return null;
  }

  try {
    logMessage(`🖼️ Generating annotated image for SharePoint upload...`, context);
    logMessage(
      `Render flags: FORCE_HORIZONTAL=${FORCE_HORIZONTAL}, ENABLE_WRAP=${ENABLE_WRAP}, `
      + `WRAP_MAX_LINES=${WRAP_MAX_LINES === Infinity ? '∞' : WRAP_MAX_LINES}, WRAP_LINE_HEIGHT_MULT=${WRAP_LINE_HEIGHT_MULT}, `
      + `MAX_FONT_SIZE=${MAX_FONT_SIZE}, FILL_RATIO=${FILL_RATIO}, BOX_ALPHA=${BOX_ALPHA}`,
      context
    );
    logMessage(
      `Colors: PRINTED=${PRINTED_FILL_COLOR}, HANDWRITTEN=${HANDWRITTEN_FILL_COLOR}, `
      + `TEXT=${TEXT_COLOR}, OUTLINE=${TEXT_OUTLINE_COLOR}`,
      context
    );

    // Load image from buffer
    const image = await loadImage(originalImageBuffer);
    const canvas = createCanvas(image.width, image.height);
    const ctx = canvas.getContext('2d');

    // Draw original image
    ctx.drawImage(image, 0, 0, image.width, image.height);

    // Render text regions
    renderTextRegionsOnCanvas(canvas, analyseOutput, context);

    // Generate PNG buffer
    const pngBuffer = canvas.toBuffer('image/png');
    logMessage(`✅ Generated annotated image: ${pngBuffer.length} bytes`, context);

    // Prepare SharePoint upload
    const baseName = path.basename(originalFileName, path.extname(originalFileName));
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const annotatedFileName = `${baseName}_ANNOTATED_${timestamp}.png`;

    logMessage(`📁 Target SharePoint folder: ${folderPath}`, context);

    // Import SharePoint helpers
    const { ensureSharePointFolder, uploadOriginalDocumentToSharePoint } = require('../sharepoint/sendToSharePoint');

    await ensureSharePointFolder(folderPath, context);

    // Convert PNG to base64 for upload
    const base64ImageContent = pngBuffer.toString('base64');

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
 * Generate annotated image to local file (for testing/development)
 * @param {Array} analyseOutput - Text regions with bbox, text, etc.
 * @param {Buffer} imageBuffer - Original image buffer
 * @param {string} originalFileName - Original filename
 * @param {Object} context - Azure Functions context
 * @returns {Promise<string|null>} Path to generated file or null
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

    // Write buffer to temp file
    fs.writeFileSync(tempImagePath, imageBuffer);

    logMessage(`Loading image from buffer: ${originalFileName}`, context);

    const image = await loadImage(tempImagePath);
    const canvas = createCanvas(image.width, image.height);
    const ctx = canvas.getContext('2d');

    // Draw original image
    ctx.drawImage(image, 0, 0, image.width, image.height);

    // Render text regions
    renderTextRegionsOnCanvas(canvas, analyseOutput, context);

    // Generate PNG buffer and write to disk
    const pngBuffer = canvas.toBuffer('image/png');
    logMessage(`✅ Generated annotated image: ${pngBuffer.length} bytes`, context);

    const outFile = path.join(
      tempDir,
      `${path.basename(originalFileName, path.extname(originalFileName))}_ANNOTATED.png`
    );
    fs.writeFileSync(outFile, pngBuffer);

    // Clean up temp input image
    try { fs.unlinkSync(tempImagePath); } catch (_) {}

    return outFile;
  } catch (error) {
    handleError(error, 'generateAnnotatedImage', context);
    return null;
  }
}

/* -----------------------------------------------------------------------------
  Exports
----------------------------------------------------------------------------- */
module.exports = {
  renderTextRegionsOnCanvas,
  generateAnnotatedImage,
  generateAnnotatedImageToSharePoint,
  // Export configuration for external use
  config: {
    FONT_FAMILY,
    DEFAULT_WEIGHT,
    MAX_FONT_SIZE,
    MIN_FONT_SIZE,
    FILL_RATIO,
    PRINTED_FILL_COLOR,
    HANDWRITTEN_FILL_COLOR,
    TEXT_COLOR,
    TEXT_OUTLINE_COLOR,
    BOX_ALPHA,
    FORCE_HORIZONTAL,
    ENABLE_WRAP,
    WRAP_MAX_LINES,
    ORIENTATION_SNAP
  }
};