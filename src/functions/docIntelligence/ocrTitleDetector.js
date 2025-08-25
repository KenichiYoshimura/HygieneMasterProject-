const { AzureKeyCredential, DocumentAnalysisClient } = require("@azure/ai-form-recognizer");

const endpoint = process.env['CLASSIFIER_ENDPOINT'];
const apiKey = process.env['CLASSIFIER_ENDPOINT_AZURE_API_KEY'];

const client = new DocumentAnalysisClient(endpoint, new AzureKeyCredential(apiKey));

async function detectTitleFromDocument(context, buffer, mimeType) {
  try {
    const poller = await client.beginAnalyzeDocument("prebuilt-layout", buffer, {
      contentType: mimeType,
    });

    const result = await poller.pollUntilDone();
    const fullText = result?.content || "";

    if (fullText.includes("一般管理の実施記録")) return "一般衛生管理シート";
    if (fullText.includes("重要管理の実施記録")) return "重要管理シート";

    return null;
  } catch (error) {
    context.log(`❌ OCR title detection failed: ${error.message}`);
    return null;
  }
}

module.exports = { detectTitleFromDocument };
