
const { app } = require('@azure/functions');
const path = require('path');
const { BlobServiceClient } = require('@azure/storage-blob');

const { extractImportantManagementData } = require('./docIntelligence/importantManagementFormExtractor');
const { uploadToMonday } = require('./monday/importantManagementDashboard');
const { classifyDocument } = require('./docIntelligence/documentClassifier');

const INCOMING_CONTAINER = 'incoming-emails';
const INVALID_CONTAINER = 'invalid-attachments';
const STORAGE_CONNECTION = process.env.hygienemasterstorage_STORAGE;

const DEFAULT_ALLOWED_EXTS = ['.pdf', '.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp'];
const DEFAULT_ALLOWED_MIMES = [
  'application/pdf', 'image/jpeg', 'image/png', 'image/tiff', 'image/bmp'
];

if (process.env.ALLOW_HEIC === 'true') {
  DEFAULT_ALLOWED_EXTS.push('.heic', '.heif');
  DEFAULT_ALLOWED_MIMES.push('image/heic', 'image/heif');
}

function detectFromMagic(buffer) {
  if (!Buffer.isBuffer(buffer)) return { ext: '', mime: '', confidence: 'low' };
  const b = buffer.subarray(0, 12);
  if (b[0] === 0x25 && b[1] === 0x50 && b[2] === 0x44 && b[3] === 0x46 && b[4] === 0x2D)
    return { ext: '.pdf', mime: 'application/pdf', confidence: 'high' };
  if (b[0] === 0xFF && b[1] === 0xD8)
    return { ext: '.jpg', mime: 'image/jpeg', confidence: 'high' };
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47)
    return { ext: '.png', mime: 'image/png', confidence: 'high' };
  const isTiffLittle = b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2A && b[3] === 0x00;
  const isTiffBig = b[0] === 0x4D && b[1] === 0x4D && b[2] === 0x00 && b[3] === 0x2A;
  if (isTiffLittle || isTiffBig)
    return { ext: '.tiff', mime: 'image/tiff', confidence: 'high' };
  if (b[0] === 0x42 && b[1] === 0x4D)
    return { ext: '.bmp', mime: 'image/bmp', confidence: 'high' };
  if (process.env.ALLOW_HEIC === 'true') {
    const s = b.toString('ascii');
    if (s.includes('ftypheic') || s.includes('ftypheif') || s.includes('ftypmif1') || s.includes('ftypheix'))
      return { ext: '.heic', mime: 'image/heic', confidence: 'low' };
  }
  return { ext: '', mime: '', confidence: 'low' };
}

function decideType({ magic, fileExt, blobContentType }) {
  const ALLOWED_EXTS = DEFAULT_ALLOWED_EXTS;
  const ALLOWED_MIMES = DEFAULT_ALLOWED_MIMES;
  if (magic.mime && ALLOWED_MIMES.includes(magic.mime)) {
    return { fileExtension: magic.ext, mimeType: magic.mime, source: 'magic' };
  }
  if (fileExt && ALLOWED_EXTS.includes(fileExt)) {
    const map = {
      '.pdf': 'application/pdf', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg',
      '.png': 'image/png', '.tif': 'image/tiff', '.tiff': 'image/tiff',
      '.bmp': 'image/bmp', '.heic': 'image/heic', '.heif': 'image/heif'
    };
    return { fileExtension: fileExt, mimeType: map[fileExt] || '', source: 'extension' };
  }
  const ct = (blobContentType || '').toLowerCase();
  if (ct && ALLOWED_MIMES.includes(ct)) {
    const map = {
      'application/pdf': '.pdf', 'image/jpeg': '.jpg', 'image/png': '.png',
      'image/tiff': '.tiff', 'image/bmp': '.bmp', 'image/heic': '.heic', 'image/heif': '.heif'
    };
    return { fileExtension: map[ct] || '', mimeType: ct, source: 'contentType' };
  }
  return { fileExtension: '', mimeType: '', source: 'unknown' };
}

function parseBlobName(blobName) {
  const baseName = path.basename(blobName);
  const parts = baseName.split('-');
  if (parts.length < 4) return {
    senderEmail: 'unknown',
    customerInboxName: 'unknown',
    receivedUtc: new Date().toISOString(),
    originalFileName: baseName
  };
  return {
    senderEmail: parts[0],
    customerInboxName: parts[1],
    receivedUtc: parts[2],
    originalFileName: parts.slice(3).join('-')
  };
}

async function moveToInvalidAttachments(context, buf, originalName, opts) {
  const service = BlobServiceClient.fromConnectionString(STORAGE_CONNECTION);
  const incomingCont = service.getContainerClient(INCOMING_CONTAINER);
  const invalidCont = service.getContainerClient(INVALID_CONTAINER);
  await invalidCont.createIfNotExists({ access: 'off' });

  const baseName = path.basename(originalName);
  const customerName = opts.customerInboxName || 'unknown';
  const yyyymmdd = opts.receivedUtc.replace(/[^0-9]/g, '').slice(0, 8);
  const destPath = `${customerName}/${yyyymmdd}/${baseName}`;
  const destBlob = invalidCont.getBlockBlobClient(destPath);

  context.log(`📦 Moving unsupported file → ${INVALID_CONTAINER}/${destPath}`);

  const metadata = {
    senderEmail: opts.senderEmail || '',
    customerInboxName: opts.customerInboxName || '',
    originalBlobUrl: opts.originalBlobUrl || '',
    reason: opts.reason || 'unsupported_file_type'
  };

  await destBlob.uploadData(buf, {
    blobHTTPHeaders: { blobContentType: opts.mimeType || 'application/octet-stream' },
    metadata
  });

  try {
    await incomingCont.getBlobClient(originalName).deleteIfExists();
  } catch (err) {
    context.log.warn(`⚠️ Could not delete original blob ${originalName}: ${err.message}`);
  }

  return `${INVALID_CONTAINER}/${destPath}`;
}

app.storageBlob('FormProcessor', {
  path: `${INCOMING_CONTAINER}/{name}`,
  connection: 'hygienemasterstorage_STORAGE',
  handler: async (blob, context) => {
    const blobName = context.triggerMetadata.name;
    const buf = Buffer.isBuffer(blob) ? blob : Buffer.from(blob);
    context.log(`📄 File uploaded: ${blobName}. Starting pre-checks...`);

    const { senderEmail, customerInboxName, receivedUtc, originalFileName } = parseBlobName(blobName);
    const fileExt = path.extname(originalFileName).toLowerCase();
    const blobContentType = context.bindingData.properties?.contentType || '';

    const magic = detectFromMagic(buf);
    const { fileExtension, mimeType, source } = decideType({ magic, fileExt, blobContentType });

    context.log(`🔎 Sender: ${senderEmail}`);
    context.log(`🔎 Inbox: ${customerInboxName}`);
    context.log(`🔎 Type decision: ext=${fileExtension || '(none)'} mime=${mimeType || '(none)'} via=${source}`);

    if (!mimeType || !fileExtension) {
      const dest = await moveToInvalidAttachments(context, buf, blobName, {
        customerInboxName,
        senderEmail,
        mimeType: blobContentType || 'application/octet-stream',
        receivedUtc,
        originalBlobUrl: context.bindingData.uri,
        reason: 'unsupported_or_unknown_type'
      });
      context.log(`⛔ Unsupported file moved to: ${dest}`);
      return;
    }

    context.log("🧠 Starting classification with Azure Document Intelligence...");
    const classification = await classifyDocument(context, buf, blobName);
    if (!classification || !classification.result || !classification.result.analyzeResult?.documents?.length) {
      const dest = await moveToInvalidAttachments(context, buf, blobName, {
        customerInboxName,
        senderEmail,
        mimeType,
        receivedUtc,
        originalBlobUrl: context.bindingData.uri,
        reason: 'classification_failed_or_empty'
      });
      context.log(`⛔ Classification failed or returned no documents. File moved to: ${dest}`);
      return;
    }

    const { result, base64Raw } = classification;
    const doc = result.analyzeResult.documents[0];
    context.log(`📄 Document Type predicted: ${doc.docType}`);

    const extractedRows = await extractImportantManagementData(context, base64Raw, fileExtension);
    for (const { row, fileName } of extractedRows) {
      await uploadToMonday(row, context, base64Raw, fileName);
    }
  }
});