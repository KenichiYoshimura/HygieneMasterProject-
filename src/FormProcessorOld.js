if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}

const { BlobServiceClient } = require('@azure/storage-blob');
const { logMessage, handleError, moveBlob } = require('./utils');
const { app } = require('@azure/functions');
const { extractGeneralManagementData } = require('./docIntelligence/generalManagementFormExtractor');
const { extractImportantManagementData } = require('./docIntelligence/importantManagementFormExtractor');
const { uploadToMondayGeneralManagementBoard } = require('./monday/generaltManagementDashboard');
const { uploadToMonday } = require('./monday/importantManagementDashboard');
const { classifyDocument } = require('./docIntelligence/documentClassifier');
const { detectTitleFromDocument, GENERAL_MANAGEMENT_FORM, IMPORTANT_MANAGEMENT_FORM} = require('./docIntelligence/ocrTitleDetector');

const supportedExtensions = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff'];

function getCustomerID(senderEmail) {
  const domain = senderEmail.split('@')[1];
  return { name: domain };
}

const INVALID_ATTACHED_FILE_NAME = 'invalid-filename';
const UNSUPPORTED_FILE_TYPE = 'invalid-file-type';

function parseBlobName(blobName, context) {
  logMessage(`blob name : ${blobName}`, context);
  const regex = /^(.+?)\((.+?)\)(.+)$/;
  const match = blobName.match(regex);

  if (!match) {
    logMessage(`❌ Invalid blob name format: ${blobName}`, context);
    return { isValid: false, reason: INVALID_ATTACHED_FILE_NAME };
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
    return { isValid: false, reason: UNSUPPORTED_FILE_TYPE, timestamp: timestamp, 
      senderEmail: senderEmail, fileName: fileNameWithExt, extension, extension, companyName: companyName };
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
        logMessage(`📄 Error reason is : ${parsed.reason}`, context);
  
        if (parsed.reason === INVALID_ATTACHED_FILE_NAME) {
          await moveBlob(context, blobName, {
            connectionString: process.env['hygienemasterstorage_STORAGE'],
            sourceContainerName: 'incoming-emails',
            targetContainerName: 'invalid-attachments',
            targetSubfolder: parsed.reason
          });
        } else {
          await moveBlob(context, blobName, {
            connectionString: process.env['hygienemasterstorage_STORAGE'],
            sourceContainerName: 'incoming-emails',
            targetContainerName: 'processed-attachments',
            targetSubfolder: `${parsed.companyName}/invalid-attachments`
          });
        }

        logMessage("⏭️ Skipped and moved file due to format or unsupported type.", context);
        return;
      }

        // Try OCR-based title detection first
      const detectedTitle = await detectTitleFromDocument(context, blob, parsed.extension === '.pdf' ? 'application/pdf' : `image/${parsed.extension.replace('.', '')}`);
      if (detectedTitle) {
        logMessage(`🔍 Title detected via OCR: ${detectedTitle}`, context);
        const base64Raw = blob.toString('base64');
        const fileExtension = parsed.extension.replace('.', '');
        const companyName = parsed.companyName;
        await processExtractedData(context, {
              title: detectedTitle,
              base64Raw,
              fileExtension,
              blobName,
              companyName
          });
          return;
      } else {
        logMessage(`failed to detect title via OCR`, context);
        return;
      }
    } catch (error) {
      handleError("❌ Unexpected error occurred", error, context);
    }
  }
});

async function processExtractedData(context, {
  title,
  base64Raw,
  fileExtension,
  blobName,
  companyName
}) {
  try {
    let extractedRows;

    if (title === GENERAL_MANAGEMENT_FORM) {
      extractedRows = await extractGeneralManagementData(context, base64Raw, fileExtension);
      for (const { row, fileName } of extractedRows) {
        await uploadToMondayGeneralManagementBoard(row, context, base64Raw, fileName);
      }
    } else if (title === IMPORTANT_MANAGEMENT_FORM) {
      extractedRows = await extractImportantManagementData(context, base64Raw, fileExtension);
      for (const { row, fileName } of extractedRows) {
        await uploadToMonday(row, context, base64Raw, fileName);
      }
    } else {
      logMessage(`⚠️ Unknown form title: ${title}`, context);
      return;
    }

    await moveBlob(context, blobName, {
      connectionString: process.env['hygienemasterstorage_STORAGE'],
      sourceContainerName: 'incoming-emails',
      targetContainerName: 'processed-attachments',
      targetSubfolder: companyName
    });

    logMessage(`✅ Successfully processed and moved blob: ${blobName}`, context);
  } catch (error) {
    handleError("❌ Error during data extraction/upload", error, context);
  }
}
