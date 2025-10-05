if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}

const { BlobServiceClient } = require('@azure/storage-blob');
const { logMessage, handleError, moveBlob } = require('./utils');
const { app } = require('@azure/functions');
// Updated to use new extractors (removed /legacy/ path)
const { extractGeneralManagementData } = require('./docIntelligence/generalManagementFormExtractor');
const { extractImportantManagementData } = require('./docIntelligence/importantManagementFormExtractor');
//const { uploadToMondayGeneralManagementBoard } = require('./monday/generalManagementDashboard');
//const { uploadToMonday } = require('./monday/importantManagementDashboard');
//const { classifyDocument } = require('./docIntelligence/documentClassifier');
const { detectTitleFromDocument, GENERAL_MANAGEMENT_FORM, IMPORTANT_MANAGEMENT_FORM } = require('./docIntelligence/ocrTitleDetector');
const { prepareGeneralManagementReport } = require('./sharepoint/generalManagementReport');
const { prepareImportantManagementReport } = require('./sharepoint/importantManagementReport');
const { analyseAndExtract, generateAnnotatedImageToSharePoint } = require('./docIntelligence/generalFormExtractor');

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
        logMessage(`⚠️ OCR failed to detect title. Trying to extract contents from it using general AI.`, context);
        await processUnknownFileType(context, {
          title: detectedTitle,
          base64Raw,
          fileExtension,
          blobName,
          companyName
        });
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
    
    if (title === GENERAL_MANAGEMENT_FORM) {
      // Use new structured extractor
      const structuredData = await extractGeneralManagementData(context, base64Raw, fileExtension);
      
      logMessage(`📊 Extracted structured data from 一般管理フォーム:`, context);
      logMessage(`  - Location: ${structuredData.metadata.location}`, context);
      logMessage(`  - Year-Month: ${structuredData.metadata.yearMonth}`, context);
      logMessage(`  - Daily Records: ${structuredData.dailyRecords.length}`, context);
      logMessage(`  - Categories: ${structuredData.categories.length}`, context);

      logMessage('🚀 Starting report preparation for 一般管理...', context);
      
      // Pass structured data directly to report generator
      await prepareGeneralManagementReport(structuredData, context, base64Raw, blobName);

      logMessage(`✅ Finished generating 一般管理 report`, context);
      
      /* Uncomment below to upload to Monday.com (will need legacy conversion)
      const legacyData = convertStructuredToLegacyFormat(structuredData, 'general');
      for (const { row, fileName } of legacyData.extractedRows) {
        logMessage(`📤 Uploading row to Monday.com (一般管理): ${fileName}`, context);
        await uploadToMondayGeneralManagementBoard(row, context, base64Raw, fileName);
      }
      */
      
    } else if (title === IMPORTANT_MANAGEMENT_FORM) {
      // Use new structured extractor
      const structuredData = await extractImportantManagementData(context, base64Raw, fileExtension);
      
      logMessage(`📊 Extracted structured data from 重要管理フォーム:`, context);
      logMessage(`  - Location: ${structuredData.metadata.location}`, context);
      logMessage(`  - Year-Month: ${structuredData.metadata.yearMonth}`, context);
      logMessage(`  - Daily Records: ${structuredData.dailyRecords.length}`, context);
      logMessage(`  - Menu Items: ${structuredData.menuItems.length}`, context);

      logMessage('🚀 Starting report preparation for 重要管理...', context);
      
      // Pass structured data directly to report generator
      await prepareImportantManagementReport(structuredData, context, base64Raw, blobName);

      logMessage(`✅ Finished generating 重要管理 report`, context);
      
      /* Uncomment below to upload to Monday.com (will need legacy conversion)
      const legacyData = convertStructuredToLegacyFormat(structuredData, 'important');
      for (const { row, fileName } of legacyData.extractedRows) {
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

/**  
 * Function to extract data from any document to demonstrate the capability of AI
*/
async function processUnknownFileType(context, {
  title,
  base64Raw,
  fileExtension,
  blobName,
  companyName
}) {
  try {
    logMessage(`🧠 Starting general form extraction for unknown file type`, context);
    logMessage(`📄 File: ${blobName}, Company: ${companyName}, Extension: ${fileExtension}`, context);

    // Convert base64 to buffer for processing
    const imageBuffer = Buffer.from(base64Raw, 'base64');
    
    // Determine MIME type
    const mimeType = fileExtension === 'pdf' ? 'application/pdf' : 
                    fileExtension === 'heic' ? 'image/heif' : 
                    `image/${fileExtension}`;

    logMessage(`🔍 Processing with MIME type: ${mimeType}`, context);

    // Step 1: Extract and analyze the document using general form extractor
    logMessage(`📖 Starting document analysis...`, context);
    const analyseOutput = await analyseAndExtract(imageBuffer, mimeType, context);

    if (!analyseOutput || analyseOutput.length === 0) {
      logMessage(`❌ No text regions detected in the document`, context);
      
      // Move to processed folder even if no text found
      await moveBlob(context, blobName, {
        sourceContainerName: 'incoming-emails',
        targetContainerName: 'processed-attachments',
        targetSubfolder: `${companyName}/no-text-detected`
      });
      
      return;
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

    // Step 2: Generate and upload annotated image to SharePoint
    logMessage(`🖼️ Generating annotated image for SharePoint upload...`, context);
    
    const originalFileName = blobName.split(')')[1] || `unknown_${Date.now()}.${fileExtension}`;
    
    const sharePointResult = await generateAnnotatedImageToSharePoint(
      analyseOutput,           // Analyzed text regions
      imageBuffer,             // Original image buffer
      originalFileName,        // Original filename for naming
      context,                 // Azure Functions context
      companyName             // Company name for folder organization
    );

    if (sharePointResult) {
      logMessage(`✅ Successfully uploaded annotated image to SharePoint`, context);
      logMessage(`🔗 SharePoint result: ${JSON.stringify(sharePointResult)}`, context);
    } else {
      logMessage(`⚠️ Failed to upload annotated image to SharePoint, but continuing...`, context);
    }

    // Step 3: Move original file to processed folder
    logMessage(`📦 Moving original file to processed-attachments/${companyName}/general-extraction`, context);

    await moveBlob(context, blobName, {
      sourceContainerName: 'incoming-emails',
      targetContainerName: 'processed-attachments',
      targetSubfolder: `${companyName}/general-extraction`
    });

    logMessage(`✅ Successfully processed unknown file type: ${blobName}`, context);
    logMessage(`📊 Summary: Found ${analyseOutput.length} text regions, uploaded annotated image to SharePoint`, context);

  } catch (error) {
    logMessage(`❌ Error during general form extraction: ${error.message}`, context);
    handleError("❌ Error during general form extraction", error, context);
    
    // Move to error folder on failure
    try {
      await moveBlob(context, blobName, {
        sourceContainerName: 'incoming-emails',
        targetContainerName: 'processed-attachments',
        targetSubfolder: `${companyName}/extraction-errors`
      });
      logMessage(`📦 Moved failed file to extraction-errors folder`, context);
    } catch (moveError) {
      logMessage(`❌ Failed to move error file: ${moveError.message}`, context);
    }
  }
}

/**
 * Converts the new structured data format back to legacy format for Monday.com compatibility
 * (Only needed if Monday.com upload is enabled)
 */
function convertStructuredToLegacyFormat(structuredData, formType) {
  if (formType === 'general') {
    const extractedRows = structuredData.dailyRecords.map(record => ({
      row: {
        text_mkv0z6d: structuredData.metadata.location,
        date4: record.date,
        color_mkv02tqg: record.Cat1Status,
        color_mkv0yb6g: record.Cat2Status,
        color_mkv06e9z: record.Cat3Status,
        color_mkv0x9mr: record.Cat4Status,
        color_mkv0df43: record.Cat5Status,
        color_mkv5fa8m: record.Cat6Status,
        color_mkv59ent: record.Cat7Status,
        text_mkv0etfg: record.comment,
        color_mkv0xnn4: record.approverStatus
      }
    }));
    const categories = structuredData.categories.map(cat => cat.categoryName);
    return { extractedRows, categories };
    
  } else if (formType === 'important') {
    const extractedRows = structuredData.dailyRecords.map(record => ({
      row: {
        text_mkv0z6d: structuredData.metadata.location,
        date4: record.date,
        color_mkv02tqg: record.Menu1Status,
        color_mkv0yb6g: record.Menu2Status,
        color_mkv06e9z: record.Menu3Status,
        color_mkv0x9mr: record.Menu4Status,
        color_mkv0df43: record.Menu5Status,
        color_mkv0ej57: record.dailyCheckStatus,
        text_mkv0etfg: record.comment,
        color_mkv0xnn4: record.approverStatus
      }
    }));
    const menuItems = structuredData.menuItems.map(item => item.menuName);
    return { extractedRows, menuItems };
  }

  throw new Error(`Unknown form type: ${formType}`);
}
