if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}

const { BlobServiceClient } = require('@azure/storage-blob');
const { logMessage, handleError, moveBlob } = require('./utils');
const { app } = require('@azure/functions');
const { extractImportantManagementData } = require('./docIntelligence/importantManagementFormExtractor');
const { uploadToMonday } = require('./monday/importantManagementDashboard');
const { classifyDocument } = require('./docIntelligence/documentClassifier');

const supportedExtensions = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff'];

function getCustomerID(senderEmail) {
  const domain = senderEmail.split('@')[1];
  return { name: domain };
}

function parseBlobName(blobName, context) {
  logMessage(`blob name : ${blobName}`, context);
  const regex = /^(.+?)\((.+?)\)(.+)$/;
  const match = blobName.match(regex);

  if (!match) {
    logMessage(`❌ Invalid blob name format: ${blobName}`, context);
    return { isValid: false, reason: 'invalid-filename' };
  }

  const timestamp = match[1];
  const senderEmail = match[2];
  const fileNameWithExt = match[3];
  const extension = fileNameWithExt.slice(fileNameWithExt.lastIndexOf('.')).toLowerCase();
  const companyName = getCustomerID(senderEmail).name;

  logMessage(`timestamp : ${timestamp}`, context);
  logMessage(`senderEmail : ${senderEmail}`, context);
  logMessage(`fileNameWithExt : ${fileNameWithExt}`, context);
  logMessage(`extension : ${extension}`, context);
  logMessage(`companyName : ${companyName}`, context);

  if (!supportedExtensions.includes(extension)) {
    logMessage(`❌ Unsupported file type: ${extension} in blob ${blobName}`, context);
    return { isValid: false, reason: companyName };
  }

  return {
    isValid: true,
    timestamp,
    senderEmail,
    fileName: fileNameWithExt,
    extension,
    companyName
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
      if (!parsed?.isValid) {
        await moveBlob(context, blobName, {
          connectionString: process.env['hygienemasterstorage_STORAGE'],
          sourceContainerName: 'incoming-emails',
          targetContainerName: 'invalid-attachments',
          targetSubfolder: parsed.reason
        });

        logMessage("⏭️ Skipped and moved file due to format or unsupported type.", context);
        return;
      }

      logMessage("📄 Starting classification...", context);
      const classification = await classifyDocument(context, blob, parsed.fileName);
      if (!classification) {
        await moveBlob(context, blobName, {
          connectionString: process.env['hygienemasterstorage_STORAGE'],
          sourceContainerName: 'incoming-emails',
          targetContainerName: 'processed-attachments',
          targetSubfolder: `${parsed.companyName}/classification-failures`
        });

        logMessage("⚠️ Classification failed. Moved to classification-failures folder.", context);
        return;
      }

      const { result, mimeType, fileExtension, base64Raw } = classification;

      if (result?.analyzeResult?.documents?.length > 0) {
        const doc = result.analyzeResult.documents[0];
        logMessage(`📄 Got the document Type: ${doc.docType}`, context);

        const extractedRows = await extractImportantManagementData(context, base64Raw, fileExtension);
        for (const { row, fileName } of extractedRows) {
          await uploadToMonday(row, context, base64Raw, fileName);
        }

        await moveBlob(context, blobName, {
          connectionString: process.env['hygienemasterstorage_STORAGE'],
          sourceContainerName: 'incoming-emails',
          targetContainerName: 'processed-attachments',
          targetSubfolder: parsed.companyName
        });

      } else {
        await moveBlob(context, blobName, {
          connectionString: process.env['hygienemasterstorage_STORAGE'],
          sourceContainerName: 'incoming-emails',
          targetContainerName: 'processed-attachments',
          targetSubfolder: `${parsed.companyName}/classification-failures`
        });

        logMessage("⚠️ No classification result found.", context);
        logMessage(`📎 Raw result: ${JSON.stringify(result, null, 2)}`, context);
      }
    } catch (error) {
      handleError("❌ Unexpected error occurred", error, context);
    }
  }
});
