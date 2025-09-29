const axios = require('axios');

// Azure Translator credentials from environment variables
const translatorEndpoint = process.env.AZURE_TRANSLATOR_ENDPOINT || 'https://api.cognitive.microsofttranslator.com/';
const translatorKey = process.env.AZURE_TRANSLATOR_KEY;
const translatorRegion = process.env.AZURE_TRANSLATOR_REGION || 'japaneast';

// Azure Language Service credentials from environment variables
const languageEndpoint = process.env.AZURE_LANGUAGE_ENDPOINT;
const languageKey = process.env.AZURE_LANGUAGE_KEY;

// Validation: Check if required environment variables are set
if (!translatorKey) {
    throw new Error('AZURE_TRANSLATOR_KEY environment variable is required');
}
if (!languageEndpoint) {
    throw new Error('AZURE_LANGUAGE_ENDPOINT environment variable is required');
}
if (!languageKey) {
    throw new Error('AZURE_LANGUAGE_KEY environment variable is required');
}

// Supported languages for sentiment analysis
const supportedLanguages = new Set([
    "af", "am", "ar", "az", "bg", "bn", "bs", "ca", "cs", "cy", "da", "de", "el", "en", "es", "et", "fa", "fi", "fil",
    "fr", "ga", "gu", "he", "hi", "hr", "hu", "hy", "id", "is", "it", "ja", "jv", "kk", "km", "kn", "ko", "lo", "lt",
    "lv", "mk", "ml", "mn", "mr", "ms", "mt", "my", "nb", "ne", "nl", "or", "pa", "pl", "ps", "pt", "ro", "ru", "si",
    "sk", "sl", "sq", "sr", "sv", "sw", "ta", "te", "th", "tl", "tr", "uk", "ur", "uz", "vi", "zh"
]);

// Unified function for sentiment analysis
async function analyzeComment(text) {
    try {
        // Step 1: Detect language
        console.log('🔍 Detecting language for text:', text.substring(0, 50));
        const detectRes = await axios.post(
            `${languageEndpoint}language/:analyze-text?api-version=2024-11-01`,
            {
                kind: "LanguageDetection",
                parameters: { modelVersion: "latest" },
                analysisInput: { documents: [{ id: "1", text }] }
            },
            {
                headers: {
                    'Ocp-Apim-Subscription-Key': languageKey,
                    'Content-Type': 'application/json',
                }
            }
        );
        const detectedLanguage = detectRes.data.results.documents[0].detectedLanguage.iso6391Name;
        console.log('🌐 Detected language:', detectedLanguage);

        // Step 2: Check if we need translation
        const isLanguageSupported = supportedLanguages.has(detectedLanguage);
        let japaneseTranslation = null;
        
        if (!isLanguageSupported) {
            console.log('🔄 Language not supported for sentiment analysis, translating to Japanese...');
            const translateRes = await axios.post(
                `${translatorEndpoint}translate?api-version=3.0&to=ja`,
                [{ text }],
                {
                    headers: {
                        'Ocp-Apim-Subscription-Key': translatorKey,
                        'Ocp-Apim-Subscription-Region': translatorRegion,
                        'Content-Type': 'application/json'
                    }
                }
            );
            japaneseTranslation = translateRes.data[0].translations[0].text;
            console.log('🇯🇵 Japanese translation:', japaneseTranslation);
        } else {
            console.log('✅ Language supported, using original text for sentiment analysis');
        }

        // Step 3: Sentiment analysis
        const sentimentLanguage = isLanguageSupported ? detectedLanguage : 'ja';
        const textToAnalyze = isLanguageSupported ? text : japaneseTranslation;
        console.log('😊 Analyzing sentiment using language:', sentimentLanguage);
        
        const sentimentRes = await axios.post(
            `${languageEndpoint}language/:analyze-text?api-version=2024-11-01`,
            {
                kind: "SentimentAnalysis",
                parameters: { modelVersion: "latest" },
                analysisInput: {
                    documents: [{ id: "1", language: sentimentLanguage, text: textToAnalyze }]
                }
            },
            {
                headers: {
                    'Ocp-Apim-Subscription-Key': languageKey,
                    'Content-Type': 'application/json',
                }
            }
        );
        const sentimentDoc = sentimentRes.data.results.documents[0];

        // Return result
        const result = {
            originalComment: text,
            detectedLanguage,
            japaneseTranslation,
            sentimentAnalysisLanguage: sentimentLanguage,
            sentiment: sentimentDoc.sentiment,
            scores: sentimentDoc.confidenceScores,
            wasTranslated: !isLanguageSupported
        };

        console.log('✅ Analysis complete:', result.sentiment, `(analyzed in ${sentimentLanguage})`);
        return result;

    } catch (error) {
        console.error('❌ Sentiment analysis failed:', error.message);
        if (error.response) {
            console.error('❌ Response status:', error.response.status);
            console.error('❌ Response data:', JSON.stringify(error.response.data, null, 2));
        }
        throw error;
    }
}

/*
// Sample usage
async function main() {
    const comments = [
        "良好",
        "サービスに満足しています。",
        "The food was cold and the service was slow.",
        "El servicio fue excelente y la comida deliciosa.",
        "บริการดีมากและอาหารอร่อยมาก"
    ];

    for (const comment of comments) {
        try {
            const result = await analyzeComment(comment);
            console.log(JSON.stringify(result, null, 2));
        } catch (error) {
            console.error(`Failed to analyze comment: "${comment}"`, error.message);
        }
    }
}

//Uncomment to test
main();
*/

module.exports = { analyzeComment, supportedLanguages };