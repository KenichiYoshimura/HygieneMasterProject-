if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}
const { logMessage, handleError } = require('./utils');
const { app } = require('@azure/functions');
const { extractImportantManagementData } = require('./docIntelligence/importantManagementFormExtractor');
const { uploadToMonday } = require('./monday/importantManagementDashboard');
const { classifyDocument } = require('./docIntelligence/documentClassifier');

const supportedExtensions = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff'];

function parseBlobName(blobName, context) {
  logMessage(`blob name : ${blobName}`, context);
  const regex = /^(.+?)\((.+?)\)(.+)$/;
  const match = blobName.match(regex);

  logMessage(`performed match : ${blobName}`, context);
  logMessage(`value of match is ${match}`, context);
  if (!match) {
    logMessage(`❌ Invalid blob name format: ${blobName}`, context);
    return null;
  }
  logMessage(`blob name is ok : ${blobName}`);

  const timestamp = match[1];
  const senderEmail = match[2];
  const fileNameWithExt = match[3];
  const extension = fileNameWithExt.slice(fileNameWithExt.lastIndexOf('.')).toLowerCase();

  logMessage(`timestamp : ${timestamp}`, context);
  logMessage(`senderEmail : ${senderEmail}`, context);
  logMessage(`fileNameWithExt : ${fileNameWithExt}`, context);
  logMessage(`extension : ${extension}`, context);

  if (!supportedExtensions.includes(extension)) {
    logMessage(`❌ Unsupported file type: ${extension} in blob ${blobName}`, context);
    return null;
  }

  logMessage(`doen parseBlobName`, context);
  
  return {
    timestamp,
    senderEmail,
    fileName: fileNameWithExt,
    extension
  };
}

app.storageBlob('FormProcessor', {
  path: 'incoming-emails/{name}',
  connection: 'hygienemasterstorage_STORAGE',
  handler: async (blob, context) => {
    try {
      const blobName = context.triggerMetadata.name;
      logMessage(`📄 File uploaded: ${blobName}`, context);

      const parsed = parseBlobName(blobName, context);
      if (!parsed) {
        logMessage("⏭️ Skipping file due to format or unsupported type.", context);
        return;
      }

      logMessage("📄 Starting classification...", context);
      const classification = await classifyDocument(context, blob, parsed.fileName);
      if (!classification) return;

      const { result, mimeType, fileExtension, base64Raw } = classification;

      if (result?.analyzeResult?.documents?.length > 0) {
        const doc = result.analyzeResult.documents[0];
        logMessage(`📄 Got the document Type: ${doc.docType}`, context);

        const extractedRows = await extractImportantManagementData(context, base64Raw, fileExtension);
        for (const { row, fileName } of extractedRows) {
          await uploadToMonday(row, context, base64Raw, fileName);
        }
      } else {
        logMessage("⚠️ No classification result found.", context);
        logMessage(`📎 Raw result: ${JSON.stringify(result, null, 2)}`, context);
      }
    } catch (error) {
      logMessage("❌ Unexpected error occurred:", error.message, context);
      logMessage(error.stack, context);
    }
  }
});

