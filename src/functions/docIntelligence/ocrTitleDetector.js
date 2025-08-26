const { AzureKeyCredential, DocumentAnalysisClient } = require("@azure/ai-form-recognizer");

const GENERAL_MANAGEMENT_FORM = `一般衛生管理シート`;
const IMPORTANT_MANAGEMENT_FORM = `重要管理シート`;

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

    if (fullText.includes("一般管理の実施記録")) return GENERAL_MANAGEMENT_FORM;
    if (fullText.includes("重要管理の実施記録")) return IMPORTANT_MANAGEMENT_FORM;

    return null;
  } catch (error) {
    context.log(`❌ OCR title detection failed: ${error.message}`);
    return null;
  }
}

module.exports = { detectTitleFromDocument, GENERAL_MANAGEMENT_FORM, IMPORTANT_MANAGEMENT_FORM };
