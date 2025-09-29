const axios = require('axios');

// Azure Translator credentials from environment variables (restored original names)
const translatorKey = process.env.AZURE_TRANSLATOR_KEY;
const translatorEndpoint = process.env.AZURE_TRANSLATOR_ENDPOINT;
const translatorRegion = process.env.AZURE_TRANSLATOR_REGION;

// Azure Language Service credentials (restored original names)
const languageKey = process.env.AZURE_LANGUAGE_KEY;
const languageEndpoint = process.env.AZURE_LANGUAGE_ENDPOINT;

// Languages supported by Azure Sentiment Analysis
const supportedLanguages = new Set([
    'ar', 'bg', 'ca', 'zh', 'zh-hans', 'zh-hant', 'hr', 'cs', 'da', 'nl', 'en', 'et',
    'fi', 'fr', 'de', 'el', 'he', 'hi', 'hu', 'is', 'id', 'it', 'ja', 'kk', 'ko',
    'lv', 'lt', 'ms', 'nb', 'fa', 'pl', 'pt', 'ro', 'ru', 'sr', 'sk', 'sl', 'es',
    'sv', 'ta', 'te', 'th', 'tr', 'uk', 'ur', 'vi'
]);

// Unified function for sentiment analysis
async function analyzeComment(text) {
    try {
        console.log('🔍 Starting sentiment analysis for text:', text.substring(0, 50));
        
        // Step 1: Detect language
        console.log('🔍 Detecting language...');
        const detectRes = await axios.post(
            `${languageEndpoint}language/:analyze-text?api-version=2023-04-01`,
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
        
        // Check if language detection was successful
        if (!detectRes.data.results || !detectRes.data.results.documents || detectRes.data.results.documents.length === 0) {
            throw new Error('Language detection failed - no results returned');
        }
        
        const languageDoc = detectRes.data.results.documents[0];
        if (languageDoc.error) {
            throw new Error(`Language detection error: ${languageDoc.error.message}`);
        }
        
        const detectedLanguage = languageDoc.detectedLanguage.iso6391Name;
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
            
            if (!translateRes.data || translateRes.data.length === 0 || !translateRes.data[0].translations) {
                throw new Error('Translation failed - no translation returned');
            }
            
            japaneseTranslation = translateRes.data[0].translations[0].text;
            console.log('🇯🇵 Japanese translation:', japaneseTranslation);
        } else {
            console.log('✅ Language supported, using original text for sentiment analysis');
        }

        // Step 3: Sentiment analysis
        const sentimentLanguage = isLanguageSupported ? detectedLanguage : 'ja';
        const textToAnalyze = isLanguageSupported ? text : japaneseTranslation;
        console.log('😊 Analyzing sentiment using language:', sentimentLanguage);
        console.log('😊 Text to analyze:', textToAnalyze);
        
        const sentimentRes = await axios.post(
            `${languageEndpoint}language/:analyze-text?api-version=2023-04-01`,
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
        
        console.log('🔍 Sentiment API response structure:', JSON.stringify(sentimentRes.data, null, 2));

        // Check if sentiment analysis was successful
        if (!sentimentRes.data.results || !sentimentRes.data.results.documents || sentimentRes.data.results.documents.length === 0) {
            throw new Error('Sentiment analysis failed - no results returned');
        }
        
        const sentimentDoc = sentimentRes.data.results.documents[0];
        if (sentimentDoc.error) {
            throw new Error(`Sentiment analysis error: ${sentimentDoc.error.message}`);
        }

        // Extract confidence scores safely
        let confidenceScores = {};
        if (sentimentDoc.confidenceScores) {
            confidenceScores = {
                positive: sentimentDoc.confidenceScores.positive || 0,
                neutral: sentimentDoc.confidenceScores.neutral || 0,
                negative: sentimentDoc.confidenceScores.negative || 0
            };
        } else {
            // Fallback: create default confidence scores
            console.warn('⚠️ No confidence scores found, using defaults');
            const sentiment = sentimentDoc.sentiment;
            confidenceScores = {
                positive: sentiment === 'positive' ? 0.8 : 0.1,
                neutral: sentiment === 'neutral' ? 0.8 : 0.1,
                negative: sentiment === 'negative' ? 0.8 : 0.1
            };
        }

        // Return result with proper structure
        const result = {
            originalComment: text,
            detectedLanguage,
            japaneseTranslation,
            analysisLanguage: sentimentLanguage,
            sentiment: sentimentDoc.sentiment,
            confidenceScores: confidenceScores,
            wasTranslated: !isLanguageSupported
        };

        console.log('✅ Analysis complete:', result.sentiment, `(confidence: ${Math.round((confidenceScores[result.sentiment] || 0) * 100)}%)`);
        return result;

    } catch (error) {
        console.error('❌ Sentiment analysis failed:', error.message);
        if (error.response) {
            console.error('❌ Response status:', error.response.status);
            console.error('❌ Response headers:', error.response.headers);
            console.error('❌ Response data:', JSON.stringify(error.response.data, null, 2));
        }
        
        // Return error object that the report can handle
        return {
            originalComment: text,
            error: error.message,
            detectedLanguage: 'unknown',
            japaneseTranslation: null,
            analysisLanguage: 'unknown',
            sentiment: 'unknown',
            confidenceScores: { positive: 0, neutral: 0, negative: 0 },
            wasTranslated: false
        };
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
 * Formats confidence scores into detailed breakdown (legacy function for backward compatibility)
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
 * Formats confidence scores into simple inline display
 * @param {Object} confidenceScores - Sentiment confidence scores
 * @returns {string} Formatted HTML string with inline confidence details
 */
function formatInlineConfidenceDetails(confidenceScores) {
    if (!confidenceScores) return '';
    
    const scores = [
        { label: 'ポジティブ', value: confidenceScores.positive || 0, emoji: '😊', class: 'positive' },
        { label: 'ニュートラル', value: confidenceScores.neutral || 0, emoji: '😐', class: 'neutral' },
        { label: 'ネガティブ', value: confidenceScores.negative || 0, emoji: '😞', class: 'negative' }
    ].sort((a, b) => b.value - a.value);
    
    return `
        <div class="inline-confidence-details">
            ${scores.map(score => 
                `<div class="inline-score-item ${score.class}">
                    <span class="score-emoji">${score.emoji}</span>
                    <span class="score-label">${score.label}</span>
                    <span class="score-value">${Math.round(score.value * 100)}%</span>
                </div>`
            ).join('')}
        </div>
    `;
}

module.exports = { 
    analyzeComment, 
    supportedLanguages, 
    getLanguageNameInJapanese, 
    formatConfidenceDetails,          // Keep for backward compatibility
    formatInlineConfidenceDetails     // New simple inline version
};