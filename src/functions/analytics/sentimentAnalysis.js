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
 * @param {string} languageCode - ISO language code or Azure-specific code
 * @returns {string} Japanese language name
 */
function getLanguageNameInJapanese(languageCode) {
    const languageNames = {
        // Standard ISO codes
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
        
        // Azure Cognitive Services specific language codes
        'zh_chs': '中国語（簡体）',
        'zh_cht': '中国語（繁体）',
        'zh-hans': '中国語（簡体）',
        'zh-hant': '中国語（繁体）',
        'en_us': '英語（米国）',
        'en_gb': '英語（英国）',
        'en_au': '英語（豪州）',
        'en_ca': '英語（カナダ）',
        'pt_br': 'ポルトガル語（ブラジル）',
        'pt_pt': 'ポルトガル語（ポルトガル）',
        'es_es': 'スペイン語（スペイン）',
        'es_mx': 'スペイン語（メキシコ）',
        'fr_fr': 'フランス語（フランス）',
        'fr_ca': 'フランス語（カナダ）',
        'de_de': 'ドイツ語（ドイツ）',
        'it_it': 'イタリア語（イタリア）',
        'ja_jp': '日本語（日本）',
        'ko_kr': '韓国語（韓国）',
        'ru_ru': 'ロシア語（ロシア）',
        'ar_sa': 'アラビア語（サウジアラビア）',
        'hi_in': 'ヒンディー語（インド）',
        'th_th': 'タイ語（タイ）',
        'vi_vn': 'ベトナム語（ベトナム）',
        'id_id': 'インドネシア語（インドネシア）',
        'ms_my': 'マレー語（マレーシア）',
        'tl_ph': 'フィリピン語（フィリピン）',
        'nl_nl': 'オランダ語（オランダ）',
        'sv_se': 'スウェーデン語（スウェーデン）',
        'da_dk': 'デンマーク語（デンマーク）',
        'no_no': 'ノルウェー語（ノルウェー）',
        'fi_fi': 'フィンランド語（フィンランド）',
        'pl_pl': 'ポーランド語（ポーランド）',
        'tr_tr': 'トルコ語（トルコ）',
        'he_il': 'ヘブライ語（イスラエル）',
        
        // Additional common variations
        'cmn': '中国語（標準）',
        'yue': '中国語（広東）',
        'wuu': '中国語（呉語）',
        
        // Fallback
        'unknown': '不明'
    };
    
    // Convert to lowercase for case-insensitive matching
    const normalizedCode = languageCode?.toLowerCase();
    
    // Try exact match first
    if (languageNames[normalizedCode]) {
        return languageNames[normalizedCode];
    }
    
    // Try without underscores (convert zh_chs to zh-chs)
    const dashFormat = normalizedCode?.replace('_', '-');
    if (languageNames[dashFormat]) {
        return languageNames[dashFormat];
    }
    
    // Try just the main language part (zh_chs -> zh)
    const mainLang = normalizedCode?.split(/[_-]/)[0];
    if (languageNames[mainLang]) {
        return `${languageNames[mainLang]}（詳細不明）`;
    }
    
    // Final fallback - return the original code in uppercase
    return languageCode?.toUpperCase() || '不明';
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

/**
 * Formats confidence scores into expandable details structure
 * @param {Object} confidenceScores - Sentiment confidence scores
 * @param {string} recordId - Unique identifier for this record (e.g., "day-5")
 * @returns {string} Formatted HTML string with expandable details
 */
function formatExpandableConfidenceDetails(confidenceScores, recordId) {
    if (!confidenceScores) return '';
    
    const scores = [
        { label: 'ポジティブ', value: confidenceScores.positive || 0, emoji: '😊', class: 'positive' },
        { label: 'ニュートラル', value: confidenceScores.neutral || 0, emoji: '😐', class: 'neutral' },
        { label: 'ネガティブ', value: confidenceScores.negative || 0, emoji: '😞', class: 'negative' }
    ].sort((a, b) => b.value - a.value); // Sort by highest confidence
    
    const detailsContent = scores.map(score => 
        `<div class="detail-score-item ${score.class}">
            <span class="detail-emoji">${score.emoji}</span>
            <span class="detail-label">${score.label}</span>
            <div class="detail-bar">
                <div class="detail-fill ${score.class}" style="width: ${Math.round(score.value * 100)}%"></div>
                <span class="detail-percentage">${Math.round(score.value * 100)}%</span>
            </div>
        </div>`
    ).join('');
    
    return `
        <div class="expandable-details">
            <button class="details-toggle" onclick="toggleDetails('${recordId}')" aria-expanded="false" type="button">
                <span class="toggle-text">詳細</span>
                <span class="toggle-icon">▼</span>
            </button>
            <div id="details-${recordId}" class="details-content" style="display: none;">
                <div class="details-header">
                    <strong>感情分析スコア詳細</strong>
                </div>
                ${detailsContent}
            </div>
        </div>
    `;
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
    formatExpandableConfidenceDetails
};