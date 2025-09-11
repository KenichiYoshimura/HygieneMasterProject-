const { BlobServiceClient } = require('@azure/storage-blob');
const heicConvert = require('heic-convert'); 

function logMessage(message, context) {
    if (context && context.log) {
        context.log(message);
    } else {
        console.log(message);
    }
}

function handleError(error, phase, context) {
    const log = context?.log?.error || console.error;
    log(`[ERROR - ${phase}] ${error.message}`);
    if (error.response) {
        log(`[RESPONSE] ${JSON.stringify(error.response.data, null, 2)}`);
    }
    log(`[STACK] ${error.stack}`);
}

async function moveBlob(context, blobName, {
  connectionString,
  sourceContainerName,
  targetContainerName,
  targetSubfolder
}) {
  try {
    context.log(`🔧 Starting moveBlob for "${blobName}"`);
    context.log(`🔧 Source container: ${sourceContainerName}`);
    context.log(`🔧 Target container: ${targetContainerName}`);
    context.log(`🔧 Target subfolder: ${targetSubfolder}`);

    const blobServiceClient = BlobServiceClient.fromConnectionString(connectionString);
    context.log(`🔧 BlobServiceClient initialized.`);

    const sourceContainer = blobServiceClient.getContainerClient(sourceContainerName);
    const targetContainer = blobServiceClient.getContainerClient(targetContainerName);
    context.log(`🔧 Container clients retrieved.`);

    const sourceBlobClient = sourceContainer.getBlobClient(blobName);
    const targetBlobPath = `${targetSubfolder}/${blobName}`;
    const targetBlobClient = targetContainer.getBlobClient(targetBlobPath);
    context.log(`🔧 Source blob URL: ${sourceBlobClient.url}`);
    context.log(`🔧 Target blob path: ${targetBlobPath}`);

    context.log(`🔄 Initiating copy from source to target...`);
    const copyPoller = await targetBlobClient.beginCopyFromURL(sourceBlobClient.url);
    await copyPoller.pollUntilDone();
    context.log(`✅ Copy completed.`);

    context.log(`🗑️ Deleting source blob...`);
    await sourceBlobClient.delete();
    context.log(`✅ Source blob deleted.`);

    context.log(`📦 Moved blob "${blobName}" to ${targetContainerName}/${targetSubfolder}/`);
  } catch (error) {
    context.log(`❌ moveBlob failed for "${blobName}"`);
    context.log(`❌ Error message: ${error.message}`);
    if (error.response) {
      context.log(`❌ Error response: ${JSON.stringify(error.response.data, null, 2)}`);
    }
    context.log(`❌ Stack trace: ${error.stack}`);
    throw error;
  }
}

// Add HEIC to JPEG conversion utility
async function convertHeicToJpegIfNeeded(buffer, originalFileName, context) {
  if (originalFileName.toLowerCase().endsWith('.heic')) {
    context?.log?.("🔄 Converting HEIC to JPEG...");
    const jpegBuffer = await heicConvert({
      buffer,
      format: 'JPEG',
      quality: 1
    });
    const newFileName = originalFileName.replace(/\.heic$/i, '.jpg');
    context?.log?.("✅ HEIC converted to JPEG.");
    return { buffer: jpegBuffer, filename: newFileName };
  }
  return { buffer, filename: originalFileName };
}

module.exports = {
    logMessage,
    handleError,
    moveBlob,
    convertHeicToJpegIfNeeded
};