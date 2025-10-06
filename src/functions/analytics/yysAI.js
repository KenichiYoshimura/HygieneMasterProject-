'use strict';

/**
 * yysAI.js (Japanese-first, domain-agnostic summarizer)
 *
 * Input:
 *   - mergedGroups: Array returned by generalFormExtractor.js (each entry has:
 *       displayText: string
 *       bbox: [x1,y1,x2,y2]
 *       polygon: [{x,y}...]
 *       isHandwritten: boolean
 *       matchedOCR: string[]
 *       orientationDeg: number
 *     )
 *
 * Output:
 *   - summaryText: string (Japanese summary + English supplement)
 *   - summaryJson: object (generic JSON schema; no domain assumptions)
 *   - txtPath/jsonPath: (optional) file paths when baseOutPathWithoutExt is provided
 *
 * Behavior:
 *   - Japanese-first summarization (summary_ja is primary; summary_en optional).
 *   - If Azure OpenAI env is present: use chat completions with structured outputs.
 *   - Otherwise: use local heuristic summary (still generic).
 */

if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const fs = require('fs');
const path = require('path');
const { logMessage, handleError } = require('../utils');

// Attempt to load Azure OpenAI client from the official OpenAI SDK.
// If not installed or env not set, we gracefully fall back to local summarizer.
let AzureOpenAI = null;
try {
  ({ AzureOpenAI } = require('openai'));
} catch {
  /* no-op; we'll use heuristic summary */
}

/**
 * Public API:
 * aiAnalysis(mergedGroups, options, context)
 * @param {Array} mergedGroups - array returned by generalFormExtractor.js
 * @param {Object} options
 *   - baseOutPathWithoutExt?: string  -> write "<base>_summary.txt" and "<base>_summary.json"
 *   - maxChars?: number               -> cap the text corpus length (default 12000)
 * @param {Object} context - Azure Functions context for logging
 * @returns {Promise<{ summaryText:string, summaryJson:Object, txtPath?:string, jsonPath?:string }>}
 */
async function aiAnalysis(mergedGroups, options = {}, context) {
  if (!Array.isArray(mergedGroups) || mergedGroups.length === 0) {
    throw new Error('aiAnalysis: mergedGroups must be a non-empty array.');
  }

  const maxChars = Number.isFinite(options.maxChars) ? options.maxChars : 12000;

  logMessage(`Starting AI analysis with ${mergedGroups.length} groups, maxChars=${maxChars}`, context);

  // 1) Build a generic text corpus (top-to-bottom, then left-to-right)
  const sorted = mergedGroups
    .filter(g => (g && typeof g.displayText === 'string' && g.displayText.trim().length > 0))
    .slice()
    .sort((a, b) => {
      // Sort primarily by Y (top), secondarily by X (left)
      const ay = (a.bbox?.[1] ?? 0);
      const by = (b.bbox?.[1] ?? 0);
      if (ay !== by) return ay - by;
      const ax = (a.bbox?.[0] ?? 0);
      const bx = (b.bbox?.[0] ?? 0);
      return ax - bx;
    });

  // Deduplicate exact lines to reduce noise
  const lines = [];
  const seen = new Set();
  for (const g of sorted) {
    const t = (g.displayText || '').trim();
    if (!t) continue;
    if (!seen.has(t)) { seen.add(t); lines.push(t); }
  }

  let corpus = lines.join('\n');
  if (corpus.length > maxChars) corpus = corpus.slice(0, maxChars) + '\n…(truncated)…';

  logMessage(`Built corpus: ${lines.length} unique lines, ${corpus.length} characters`, context);

  // 2) Minimal metadata (Japanese-first)
  const meta = {
    total_groups: mergedGroups.length,
    lines_count: lines.length,
    characters: corpus.length,
    handwritten_ratio: (() => {
      const hw = mergedGroups.filter(g => g.isHandwritten).length;
      return mergedGroups.length ? (hw / mergedGroups.length) : 0;
    })(),
    mostlyCJK: true, // <- Your requirement: expect mostly Japanese; treat content as JA
  };

  logMessage(`Metadata: ${meta.total_groups} groups, ${meta.lines_count} lines, ${(meta.handwritten_ratio * 100).toFixed(1)}% handwritten`, context);

  // 3) Choose summarizer: Azure OpenAI if configured, else heuristic
  const hasAOAI =
    !!AzureOpenAI &&
    !!process.env.AZURE_OPENAI_ENDPOINT &&
    !!process.env.AZURE_OPENAI_API_KEY &&
    !!process.env.AZURE_OPENAI_DEPLOYMENT_NAME;

  logMessage(`Using summarizer: ${hasAOAI ? 'Azure OpenAI' : 'Local heuristic'}`, context);

  let summaryText, summaryJson;
  if (hasAOAI) {
    try {
      ({ summaryText, summaryJson } = await summarizeWithAzureOpenAI_ja(corpus, meta, context));
    } catch (e) {
      logMessage(`⚠️ Azure OpenAI summarization failed; using local heuristic: ${e.message}`, context);
      ({ summaryText, summaryJson } = localHeuristicSummary_ja(corpus, meta, context));
    }
  } else {
    ({ summaryText, summaryJson } = localHeuristicSummary_ja(corpus, meta, context));
  }

  // 4) Optionally write TXT & JSON to disk (only if baseOutPathWithoutExt is provided)
  let txtPath, jsonPath;
  if (options.baseOutPathWithoutExt) {
    txtPath  = `${options.baseOutPathWithoutExt}_summary.txt`;
    jsonPath = `${options.baseOutPathWithoutExt}_summary.json`;
    try {
      fs.writeFileSync(txtPath, summaryText, 'utf-8');
      fs.writeFileSync(jsonPath, JSON.stringify(summaryJson, null, 2), 'utf-8');
      logMessage(`🧠 Saved TXT:  ${txtPath}`, context);
      logMessage(`🧠 Saved JSON: ${jsonPath}`, context);
    } catch (err) {
      handleError(err, 'aiAnalysis file save', context);
    }
  } else {
    logMessage(`🧠 AI analysis completed (no file output requested)`, context);
  }

  logMessage('AI analysis completed successfully', context);
  return { summaryText, summaryJson, txtPath, jsonPath };
}

/* =============================================================================
   Azure OpenAI (Japanese-first) – Structured Outputs
   ============================================================================= */

const genericSummarySchema_ja = {
  type: "object",
  properties: {
    doc_type_guess: { type: "string" },     // arbitrary guess (e.g., "form", "letter", "report")
    languages_detected: { type: "array", items: { type: "string" } },
    topics: { type: "array", items: { type: "string" } },
    key_points: { type: "array", items: { type: "string" } },
    entities: {
      type: "object",
      properties: {
        persons:   { type: "array", items: { type: "string" } },
        orgs:      { type: "array", items: { type: "string" } },
        locations: { type: "array", items: { type: "string" } },
        dates:     { type: "array", items: { type: "string" } },
        amounts:   { type: "array", items: { type: "string" } },
        emails:    { type: "array", items: { type: "string" } },
        phones:    { type: "array", items: { type: "string" } },
      },
      additionalProperties: false
    },
    risks_or_issues: { type: "array", items: { type: "string" } },
    actions: { type: "array", items: { type: "string" } },
    quotes: { type: "array", items: { type: "string" } },         // short representative fragments
    summary_ja: { type: "string" },                               // primary
    summary_en: { type: "string" }                                // optional
  },
  required: ["summary_ja"]
};

async function summarizeWithAzureOpenAI_ja(corpus, meta, context) {
  logMessage('Starting Azure OpenAI summarization', context);

  const client = new AzureOpenAI({
    endpoint:   process.env.AZURE_OPENAI_ENDPOINT,
    apiKey:     process.env.AZURE_OPENAI_API_KEY,
    deployment: process.env.AZURE_OPENAI_DEPLOYMENT_NAME,
    apiVersion: process.env.AZURE_OPENAI_API_VERSION || '2024-10-21',
  });

  const systemPrompt =
    'あなたは中立的な文書アナリストです。与えられたテキストのみを使い、日本語で簡潔な要約とJSON構造化分析を作成してください。事実のみを記載し、推測は避けてください。';

  const userPrompt = `
以下のテキストドキュメントを分析してください。

1) 日本語で人間が読める要約を作成してください（必須）。
2) 可能であれば英語要約も付けてください（任意）。
3) JSON形式で構造化された分析をschema通りに返してください。

- テキストに現れた内容のみを使い、事実のみを記載してください。
- エンティティ（人名、組織、金額、メール、電話番号等）があれば抽出し、なければ空配列で返してください。
- 代表的な短い引用文も含めてください。
- 中立的なトーンでまとめてください。

[メタ情報]
${JSON.stringify(meta, null, 2)}

[本文]
---
${corpus}
---
`;

  logMessage(`Calling Azure OpenAI with ${corpus.length} characters`, context);

  // Request structured outputs. If not supported by your SDK/API, catch and fallback.
  const res = await client.chat.completions.create({
    model: process.env.AZURE_OPENAI_DEPLOYMENT_NAME,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user',   content: userPrompt },
    ],
    temperature: 0.2,
    max_tokens: 2000,
    response_format: {
      type: 'json_schema',
      json_schema: { name: 'generic_doc_summary', schema: genericSummarySchema_ja, strict: true }
    }
  });

  const jsonText = res.choices?.[0]?.message?.content || '{}';
  const summaryJson = JSON.parse(jsonText);

  logMessage('Azure OpenAI response received and parsed successfully', context);

  const ja = (summaryJson.summary_ja || '').trim();
  const en = (summaryJson.summary_en || '').trim();

  const summaryText =
    `【要約 (日本語)】\n${ja}\n\n— — — — — — — — — — —\n[Summary (English)]\n${en}`;

  return { summaryText, summaryJson };
}

/* =============================================================================
   Local heuristic (Japanese-first) – No external calls
   ============================================================================= */

function localHeuristicSummary_ja(corpus, meta, context) {
  logMessage('Using local heuristic summarization', context);

  const kws = topKeywords(corpus, 12);

  // JA: primary
  const ja = [];
  ja.push(`文書長さ: 文字数 ${meta.characters}, 行数 ${meta.lines_count}, グループ数 ${meta.total_groups}`);
  ja.push(`手書き比率(推定): ${(meta.handwritten_ratio * 100).toFixed(1)}%`);
  if (kws.length) ja.push(`主要キーワード: ${kws.join(', ')}`);
  ja.push('要点(参考):');

  const points = corpus.split('\n').map(s => s.trim()).filter(Boolean).slice(0, 3);
  points.forEach((p, i) => ja.push(`  ${i + 1}. ${trimTo(p, 160)}`));
  const summary_ja = ja.join('\n');

  // EN: supplementary
  const en = [];
  en.push(`Document length: ${meta.characters} chars, ${meta.lines_count} lines, ${meta.total_groups} groups`);
  en.push(`Handwritten ratio (approx): ${(meta.handwritten_ratio * 100).toFixed(1)}%`);
  if (kws.length) en.push(`Top keywords: ${kws.join(', ')}`);
  en.push('Key points (heuristic):');
  points.forEach((p, i) => en.push(`  ${i + 1}. ${trimTo(p, 160)}`));
  const summary_en = en.join('\n');

  const summaryText =
    `【要約 (日本語)】\n${summary_ja}\n\n— — — — — — — — — — —\n[Summary (English)]\n${summary_en}`;

  const summaryJson = {
    doc_type_guess: '',
    languages_detected: ['ja'],
    topics: kws,
    key_points: points,
    entities: { persons: [], orgs: [], locations: [], dates: [], amounts: [], emails: [], phones: [] },
    risks_or_issues: [],
    actions: [],
    quotes: points.map(p => trimTo(p, 120)),
    summary_ja,
    summary_en
  };

  logMessage(`Local heuristic summary completed: ${kws.length} keywords, ${points.length} key points`, context);

  return { summaryText, summaryJson };
}

/* =============================================================================
   Generic helpers
   ============================================================================= */

function topKeywords(text, k = 12) {
  // Simple keyword extractor (language-agnostic; surfaces Latin tokens
  // and numbers). For Japanese, this still helps (e.g., product codes, dates).
  const tokens = String(text || '')
    .toLowerCase()
    .split(/[^\p{L}\p{N}]+/u)
    .filter(w => w && w.length >= 3 && !stopwords.has(w));

  const freq = new Map();
  for (const t of tokens) freq.set(t, (freq.get(t) || 0) + 1);

  return [...freq.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, k)
    .map(([w]) => w);
}

const stopwords = new Set([
  // English common words (kept minimal)
  'the','and','for','that','with','this','from','have','been','are','was','were','not','you','your',
  'into','their','them','our','out','but','can','will','would','in','on','of','to','as','by','at',
  'it','is','a','an','be','or','if','we','i','they',
  // Light romanized Japanese/functional proxies (optional)
  'desu','masu','kara','made','koto','mono',
]);

function trimTo(s, n) {
  if (!s) return s;
  return s.length > n ? (s.slice(0, n - 1) + '…') : s;
}

module.exports = { aiAnalysis };