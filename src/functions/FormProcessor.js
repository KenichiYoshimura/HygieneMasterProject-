
// ローカル開発環境用（Azure Functionsでは不要）
if (!process.env.WEBSITE_SITE_NAME) {
  require('dotenv').config();
}

const { app } = require('@azure/functions');
const { extractImportantManagementData } = require('./extractors');
const { classifyDocument } = require('./documentClassifier');

app.storageBlob('FormProcessor', {
  path: 'incoming-emails/{name}',
  connection: 'hygienemasterstorage_STORAGE',
  handler: async (blob, context) => {
    context.log("📄 File uploaded. Starting classification...");
    const fileName = context.triggerMetadata.name;

    const classification = await classifyDocument(context, blob, fileName);
    if (!classification) return;

    const { result, mimeType, fileExtension, base64Raw } = classification;

    if (result?.analyzeResult?.documents?.length > 0) {
      const doc = result.analyzeResult.documents[0];
      context.log(`📄 Got the document Type: ${doc.docType}`);
      await extractImportantManagementData(context, base64Raw, fileExtension);
    } else {
      context.log("⚠️ No classification result found.");
      context.log(`📎 Raw result: ${JSON.stringify(result, null, 2)}`);
    }
  }
});