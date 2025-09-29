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

/**
 * Converts language code to Japanese language name
 * @param {string} languageCode - ISO language code (e.g., 'en', 'ja')
 * @returns {string} Japanese language name
 */
function getLanguageNameInJapanese(languageCode) {
    const languageNames = {
        'ja': '日本語',
        'en': '英語',
        'zh': '中国語',
        'zh-cn': '中国語（簡体）',
        'zh-tw': '中国語（繁体）',
        'ko': '韓国語',
        'es': 'スペイン語',
        'fr': 'フランス語',
        'de': 'ドイツ語',
        'it': 'イタリア語',
        'pt': 'ポルトガル語',
        'ru': 'ロシア語',
        'ar': 'アラビア語',
        'hi': 'ヒンディー語',
        'th': 'タイ語',
        'vi': 'ベトナム語',
        'id': 'インドネシア語',
        'ms': 'マレー語',
        'tl': 'フィリピン語',
        'nl': 'オランダ語',
        'sv': 'スウェーデン語',
        'da': 'デンマーク語',
        'no': 'ノルウェー語',
        'fi': 'フィンランド語',
        'pl': 'ポーランド語',
        'tr': 'トルコ語',
        'he': 'ヘブライ語',
        'unknown': '不明'
    };
    
    return languageNames[languageCode?.toLowerCase()] || `${languageCode?.toUpperCase() || '不明'}`;
}

/**
 * Formats confidence scores into detailed breakdown
 * @param {Object} confidenceScores - Sentiment confidence scores
 * @returns {string} Formatted HTML string with confidence details
 */
function formatConfidenceDetails(confidenceScores) {
    if (!confidenceScores) return '';
    
    const scores = [
        { label: 'ポジティブ', value: confidenceScores.positive || 0, emoji: '😊', class: 'positive' },
        { label: 'ニュートラル', value: confidenceScores.neutral || 0, emoji: '😐', class: 'neutral' },
        { label: 'ネガティブ', value: confidenceScores.negative || 0, emoji: '😞', class: 'negative' }
    ].sort((a, b) => b.value - a.value); // Sort by highest confidence
    
    return scores.map(score => 
        `<div class="confidence-score-item ${score.class}">
            <span class="score-emoji">${score.emoji}</span>
            <span class="score-label">${score.label}</span>
            <span class="score-value">${Math.round(score.value * 100)}%</span>
        </div>`
    ).join('');
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

module.exports = { 
    analyzeComment, 
    supportedLanguages, 
    getLanguageNameInJapanese, 
    formatConfidenceDetails 
};