if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}

const { BlobServiceClient } = require('@azure/storage-blob');
const { logMessage, handleError, moveBlob } = require('./utils');
const { app } = require('@azure/functions');
const { extractGeneralManagementData } = require('./docIntelligence/generalManagementFormExtractor');
const { extractImportantManagementData } = require('./docIntelligence/importantManagementFormExtractor');
const { uploadToMondayGeneralManagementBoard } = require('./monday/generalManagementDashboard');
const { uploadToMonday } = require('./monday/importantManagementDashboard');
const { classifyDocument } = require('./docIntelligence/documentClassifier');
const { detectTitleFromDocument, GENERAL_MANAGEMENT_FORM, IMPORTANT_MANAGEMENT_FORM } = require('./docIntelligence/ocrTitleDetector');
const { prepareGeneralManagementReport} = require('./sharepoint/generalManagementReport');
const { prepareImportantManagementReport} = require('./sharepoint/importantManagementReport');

const supportedExtensions = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.heic'];

const INVALID_ATTACHED_FILE_NAME = 'invalid-filename';
const UNSUPPORTED_FILE_TYPE = 'invalid-file-type';

function getCustomerID(senderEmail) {
  const domain = senderEmail.split('@')[1];
  return { name: domain };
}

function parseBlobName(blobName, context) {
  logMessage(`🔍 Parsing blob name: ${blobName}.. Progressing...`, context);
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

  logMessage(`🧩 Parsed values → timestamp: ${timestamp}, senderEmail: ${senderEmail}, fileName: ${fileNameWithExt}, extension: ${extension}, companyName: ${companyName}`, context);

  if (!supportedExtensions.includes(extension)) {
    logMessage(`❌ Unsupported file type: ${extension}`, context);
    return {
      isValid: false,
      reason: UNSUPPORTED_FILE_TYPE,
      timestamp,
      senderEmail,
      fileName: fileNameWithExt,
      extension,
      companyName
    };
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
      logMessage(`📥 Blob triggered: ${blobName}`, context);

      const parsed = parseBlobName(blobName, context);
      if (!parsed?.isValid) {
        logMessage(`📄 Invalid file. Reason: ${parsed.reason}`, context);

        const targetContainer = parsed.reason === INVALID_ATTACHED_FILE_NAME ? 'invalid-attachments' : 'processed-attachments';
        const targetSubfolder = parsed.reason === INVALID_ATTACHED_FILE_NAME ? parsed.reason : `${parsed.companyName}/invalid-attachments`;

        await moveBlob(context, blobName, {
          connectionString: process.env['hygienemasterstorage_STORAGE'],
          sourceContainerName: 'incoming-emails',
          targetContainerName: targetContainer,
          targetSubfolder
        });

        logMessage(`📦 Moved invalid file to ${targetContainer}/${targetSubfolder}`, context);
        return;
      }

      logMessage(`🔍 Starting OCR title detection...`, context);
      const mimeType = parsed.extension === '.pdf' ? 'application/pdf' : parsed.extension === '.heic' ? 'image/heif' : `image/${parsed.extension.replace('.', '')}`;

      const detectedTitle = await detectTitleFromDocument(context, blob, mimeType);

      if (detectedTitle) {
        logMessage(`📘 OCR detected title: ${detectedTitle}`, context);
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
        logMessage(`⚠️ OCR failed to detect title. Skipping file.`, context);
        return;
      }
    } catch (error) {
      handleError("❌ Unexpected error occurred in Blob handler", error, context);
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
    logMessage(`🧠 Starting data extraction for title: ${title}`, context);
    let extractedRows;

    if (title === GENERAL_MANAGEMENT_FORM) {
      extractedRows = await extractGeneralManagementData(context, base64Raw, fileExtension);
      logMessage(`📊 Extracted ${extractedRows.length} rows from 一般管理フォーム`, context);

      logMessage('just about to start the report preparation', context);
      await prepareGeneralManagementReport(extractedRows, context, base64Raw, blobName);

      logMessage(`Finished generating the report`, context);
      /* Uncomment below to upload to Monday.com
      for (const { row, fileName } of extractedRows) {
        logMessage(`📤 Uploading row to Monday.com (一般管理): ${fileName}`, context);
        await uploadToMondayGeneralManagementBoard(row, context, base64Raw, fileName);
      }
      */
    } else if (title === IMPORTANT_MANAGEMENT_FORM) {
      extractedRows = await extractImportantManagementData(context, base64Raw, fileExtension);
      logMessage(`📊 Extracted ${extractedRows.length} rows from 重要管理フォーム`, context);

      logMessage('just about to start the report preparation', context);
      await prepareImportantManagementReport(extractedRows, context, base64Raw, blobName);

      logMessage(`Finished generating the report`, context);
      /* Uncomment below to upload to Monday.com 
      for (const { row, fileName } of extractedRows) {
        logMessage(`📤 Uploading row to Monday.com (重要管理): ${fileName}`, context);
        await uploadToMonday(row, context, base64Raw, fileName);
      }
      */
    } else {
      logMessage(`⚠️ Unknown form title: ${title}. Extraction skipped.`, context);
      return;
    }

    logMessage(`📦 Moving processed blob to processed-attachments/${companyName}`, context);

    await moveBlob(context, blobName, {
      connectionString: process.env['hygienemasterstorage_STORAGE'],
      sourceContainerName: 'incoming-emails',
      targetContainerName: 'processed-attachments',
      targetSubfolder: companyName
    });

    logMessage(`✅ Successfully processed and moved blob: ${blobName} to processed-attachments/${companyName}`, context);
  } catch (error) {
    handleError("❌ Error during data extraction/upload", error, context);
  }
}
