
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
  const blobServiceClient = BlobServiceClient.fromConnectionString(connectionString);

  const sourceContainer = blobServiceClient.getContainerClient(sourceContainerName);
  const targetContainer = blobServiceClient.getContainerClient(targetContainerName);

  const sourceBlobClient = sourceContainer.getBlobClient(blobName);
  const targetBlobClient = targetContainer.getBlobClient(`${targetSubfolder}/${blobName}`);

  const copyPoller = await targetBlobClient.beginCopyFromURL(sourceBlobClient.url);
  await copyPoller.pollUntilDone();

  await sourceBlobClient.delete();
  context.log(`📦 Moved blob "${blobName}" to ${targetContainerName}/${targetSubfolder}/ and deleted original.`);
}


module.exports = {
    logMessage,
    handleError,
    moveBlob
};