// ローカル開発環境用（Azure Functionsでは不要）
if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const { app } = require('@azure/functions');
const mime = require('mime-types');
const { extractImportantManagementData } = require('./extractors');
const axios = require('axios');

// 環境変数から設定を取得
const endpoint = process.env.CLASSIFIER_ENDPOINT;
const apiKey = process.env.AZURE_API_KEY;
const classifierId = process.env.CLASSIFIER_ID;
const apiVersion = "2024-11-30"; 

app.storageBlob('FormProcessor', {
  path: 'incoming-emails/{name}',
  connection: 'hygienemasterstorage_STORAGE',
  handler: async (blob, context) => {
    context.log("📄 File uploaded. Starting classification...");
    const fileName = context.triggerMetadata.name;
    const mimeType = mime.lookup(fileName) || 'application/pdf';
    const fileExtension = mime.extension(mimeType) || 'bin';
    const base64Raw = blob.toString('base64');

    try {
      context.log("🚀 Submitting document for classification...");
      const submitResponse = await axios.post(
        `${endpoint}/documentintelligence/documentClassifiers/${classifierId}:analyze?api-version=${apiVersion}`,
        { base64Source: base64Raw },
        {
          headers: {
            'Ocp-Apim-Subscription-Key': apiKey,
            'Content-Type': 'application/json'
          }
        }
      );

      const operationLocation = submitResponse.headers['operation-location'];
      context.log(`📍 Operation location: ${operationLocation}`);

      let result;
      let attempts = 0;
      const maxAttempts = 10;
      const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

      context.log("⏳ Polling for classification result...");
      while (attempts < maxAttempts) {
        await delay(1000);
        const pollResponse = await axios.get(operationLocation, {
          headers: { 'Ocp-Apim-Subscription-Key': apiKey }
        });
        result = pollResponse.data;
        context.log(`🔁 Attempt ${attempts + 1}: Status ${pollResponse.status}`);
        context.log(`📊 Poll result status: ${result.status}`);
        if (result.status === "succeeded") break;
        attempts++;
      }

      if (result?.analyzeResult?.documents?.length > 0) {
        const doc = result.analyzeResult.documents[0];
        context.log(`📄 Document Type: ${doc.docType}`);
        context.log(`🔢 Confidence: ${doc.confidence}`);
        context.log(`📦 Detected MIME type: ${mimeType}`);
        context.log(`📦 File extension: ${fileExtension}`);
        await extractImportantManagementData(context, base64Raw, fileExtension);
      } else {
        context.log("⚠️ No classification result found.");
        context.log(`📎 Raw result: ${JSON.stringify(result, null, 2)}`);
      }
    } catch (error) {
      context.log.error("❌ Error during classification:", error.message);
      if (error.response) {
        context.log.error("📥 Response data:", JSON.stringify(error.response.data, null, 2));
        context.log.error("📋 Response headers:", JSON.stringify(error.response.headers, null, 2));
      }
      context.log.error("🧠 Stack trace:", error.stack);
    }
  }
});
