if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}

const { app } = require('@azure/functions');
const { extractImportantManagementData } = require('./docIntelligence/importantManagementFormExtractor');
const { uploadToMonday } = require('./monday/importantManagementDashboard');
const { classifyDocument } = require('./docIntelligence/documentClassifier');

const supportedExtensions = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff'];

function parseBlobName(blobName, context) {
  context.log(`blob name : ${blobName}`);
  const regex = /^(.+?)\((.+?)\)(.+)$/;
  const match = blobName.match(regex);

  context.log(`performed match : ${blobName}`);
  if (!match) {
    context.log.error(`❌ Invalid blob name format: ${blobName}`);
    return null;
  }
  context.log(`blob name is ok : ${blobName}`);

  const timestamp = match[1];
  const senderEmail = match[2];
  const fileNameWithExt = match[3];
  const extension = fileNameWithExt.slice(fileNameWithExt.lastIndexOf('.')).toLowerCase();

  context.log(`timestamp : ${timestamp}`);
  context.log(`senderEmail : ${senderEmail}`);
  context.log(`fileNameWithExt : ${fileNameWithExt}`);
  context.log(`extension : ${extension}`);

  if (!supportedExtensions.includes(extension)) {
    context.log.error(`❌ Unsupported file type: ${extension} in blob ${blobName}`);
    return null;
  }

  context.log(`doen parseBlobName`);
  
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
    const blobName = context.triggerMetadata.name;
    context.log(`📄 File uploaded: ${blobName}`);

    const parsed = parseBlobName(blobName, context);
    if (!parsed) {
      context.log("⏭️ Skipping file due to format or unsupported type.");
      return;
    }

    context.log("📄 Starting classification...");
    const classification = await classifyDocument(context, blob, parsed.fileName);
    if (!classification) return;

    const { result, mimeType, fileExtension, base64Raw } = classification;

    if (result?.analyzeResult?.documents?.length > 0) {
      const doc = result.analyzeResult.documents[0];
      context.log(`📄 Got the document Type: ${doc.docType}`);

      const extractedRows = await extractImportantManagementData(context, base64Raw, fileExtension);
      for (const { row, fileName } of extractedRows) {
        await uploadToMonday(row, context, base64Raw, fileName);
      }
    } else {
      context.log("⚠️ No classification result found.");
      context.log(`📎 Raw result: ${JSON.stringify(result, null, 2)}`);
    }
  }
});

