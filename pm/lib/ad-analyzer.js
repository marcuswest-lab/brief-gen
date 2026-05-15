// Claude vision API client for categorizing ad images.
//
// Uses BYO API key from localStorage. Never sends the key anywhere except
// directly to api.anthropic.com from the browser (no proxy).
//
// Concurrency: limited to N parallel requests to avoid rate limits.
// Retries: on 429, exponential backoff (max 3 attempts).

const API_KEY_STORAGE = 'pbg.anthropicApiKey.v1';
const MODEL_STORAGE = 'pbg.anthropicModel.v1';
const CONCURRENCY_STORAGE = 'pbg.anthropicConcurrency.v1';

export const DEFAULT_MODEL = 'claude-haiku-4-5-20251001';
export const SONNET_MODEL = 'claude-sonnet-4-6';
export const OPUS_MODEL = 'claude-opus-4-6';
export const DEFAULT_CONCURRENCY = 3;

// -------- Settings --------

export function getApiKey() {
  return localStorage.getItem(API_KEY_STORAGE) || '';
}

export function setApiKey(key) {
  if (key) localStorage.setItem(API_KEY_STORAGE, key);
  else localStorage.removeItem(API_KEY_STORAGE);
}

export function getModel() {
  return localStorage.getItem(MODEL_STORAGE) || DEFAULT_MODEL;
}

export function setModel(m) {
  localStorage.setItem(MODEL_STORAGE, m);
}

export function getConcurrency() {
  const v = parseInt(localStorage.getItem(CONCURRENCY_STORAGE) || '', 10);
  return (Number.isFinite(v) && v >= 1 && v <= 5) ? v : DEFAULT_CONCURRENCY;
}

export function setConcurrency(n) {
  localStorage.setItem(CONCURRENCY_STORAGE, String(n));
}

// -------- Test connection --------

/** Returns { ok: true } or { ok: false, error: string }. */
export async function testApiKey(apiKey) {
  try {
    // Make a tiny request to verify the key works
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify({
        model: DEFAULT_MODEL,
        max_tokens: 16,
        messages: [{ role: 'user', content: 'Say "ok".' }],
      }),
    });
    if (!res.ok) {
      const body = await res.text();
      return { ok: false, error: `HTTP ${res.status}: ${body.slice(0, 200)}` };
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// -------- Image → base64 --------

async function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result;
      // result is "data:image/png;base64,XXXX" — strip prefix
      const idx = result.indexOf(',');
      resolve(result.slice(idx + 1));
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function mediaTypeFor(filename) {
  const lower = filename.toLowerCase();
  if (lower.endsWith('.png')) return 'image/png';
  if (lower.endsWith('.jpg') || lower.endsWith('.jpeg')) return 'image/jpeg';
  if (lower.endsWith('.webp')) return 'image/webp';
  if (lower.endsWith('.gif')) return 'image/gif';
  return 'image/png';
}

// -------- Categorization prompt --------

const SYSTEM_PROMPT = `You categorize static ad creatives for the BAD Marketing Creative Tracker. Return STRICT JSON, nothing else, no markdown fences.`;

function buildUserPrompt(filename, filenameHints) {
  let hintsBlock = '';
  if (filenameHints && Object.keys(filenameHints).length > 0) {
    const lines = [];
    if (filenameHints.ideaName) lines.push(`- Likely Idea Name (from filename): "${filenameHints.ideaName}"`);
    if (filenameHints.leadType) lines.push(`- Lead Type (from filename, authoritative): "${filenameHints.leadType}"`);
    if (filenameHints.cid) lines.push(`- CID (from filename, authoritative): "${filenameHints.cid}"`);
    if (lines.length > 0) hintsBlock = `\n\nFilename hints (use as starting point, override fields below):\n${lines.join('\n')}\n`;
  }

  return `Analyze the attached ad image and return this exact JSON structure (and nothing else):

{
  "awarenessLevel": "<one of: Most Aware | Solution Aware | Problem Aware | Unaware>",
  "leadType": "<one of: Offer | Promise | Problem-Solution | Secret | Proclamation | Story>",
  "variationType": "<one of: Copy | Visual>",
  "ideaName": "<2-4 words describing the core ad concept>",
  "angleName": "<2-4 words describing the persuasion angle>",
  "imageStyle": "<2-4 words describing the visual format, e.g. Bold Text Card, UGC Style, Multi quote visual, Clipart>",
  "rationale": "<one sentence explaining the choices>"
}

Field guidance:
- "awarenessLevel": stage of customer awareness this ad targets
- "leadType": copywriting lead style — what's the headline/hook doing
- "variationType": "Copy" if the ad varies primarily by text/copy, "Visual" if by visual treatment. Default to "Copy" unless clearly visual-driven.
- "ideaName": short noun phrase capturing the central idea
- "angleName": short noun phrase capturing the persuasion angle
- "imageStyle": describe the visual treatment/template

Filename: ${filename}${hintsBlock}

Return ONLY the JSON object. No preamble, no explanation, no markdown.`;
}

// -------- Single-image analysis --------

/**
 * @param {File} file
 * @param {Object} filenameHints { ideaName, leadType, cid }
 * @param {Object} opts { apiKey, model, signal? }
 */
export async function analyzeAdImage(file, filenameHints, opts) {
  const apiKey = opts.apiKey || getApiKey();
  if (!apiKey) throw new Error('No API key configured. Open Settings to add one.');
  const model = opts.model || getModel();

  const base64 = await fileToBase64(file);
  const mediaType = mediaTypeFor(file.name);

  const body = {
    model,
    max_tokens: 400,
    system: SYSTEM_PROMPT,
    messages: [
      {
        role: 'user',
        content: [
          {
            type: 'image',
            source: { type: 'base64', media_type: mediaType, data: base64 },
          },
          {
            type: 'text',
            text: buildUserPrompt(file.name, filenameHints),
          },
        ],
      },
    ],
  };

  // Retry on 429 with exponential backoff
  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify(body),
      signal: opts.signal,
    });

    if (res.status === 429 && attempt < maxAttempts) {
      await new Promise(r => setTimeout(r, 1000 * Math.pow(2, attempt - 1)));
      continue;
    }
    if (!res.ok) {
      const errBody = await res.text();
      throw new Error(`API error ${res.status}: ${errBody.slice(0, 300)}`);
    }

    const json = await res.json();
    const text = json.content?.[0]?.text || '';
    return parseJsonResponse(text);
  }

  throw new Error('Rate-limited after retries');
}

function parseJsonResponse(text) {
  // Claude may occasionally wrap in markdown fences; strip them
  const cleaned = text.trim()
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```$/, '')
    .trim();
  try {
    return JSON.parse(cleaned);
  } catch (e) {
    // Fallback: extract first {...} block
    const m = cleaned.match(/\{[\s\S]*\}/);
    if (m) return JSON.parse(m[0]);
    throw new Error('Failed to parse JSON from API response: ' + cleaned.slice(0, 200));
  }
}

// -------- Concurrent batch analyzer --------

/**
 * Analyze a batch of {file, filenameHints} jobs concurrently.
 * Calls onProgress(index, status, result|error) for each job.
 *
 * @param {Array<{file: File, filenameHints: object}>} jobs
 * @param {Function} onProgress
 * @param {Object} opts { concurrency?, model?, apiKey? }
 * @returns {Promise<Array>} array of { ok, result?, error? } per job (in input order)
 */
export async function analyzeBatch(jobs, onProgress, opts = {}) {
  const concurrency = opts.concurrency || getConcurrency();
  const apiKey = opts.apiKey || getApiKey();
  const model = opts.model || getModel();
  const results = new Array(jobs.length);

  let nextIndex = 0;
  const workers = Array.from({ length: Math.min(concurrency, jobs.length) }, async () => {
    while (true) {
      const i = nextIndex++;
      if (i >= jobs.length) return;
      const job = jobs[i];
      onProgress?.(i, 'analyzing');
      try {
        const result = await analyzeAdImage(job.file, job.filenameHints, { apiKey, model });
        results[i] = { ok: true, result };
        onProgress?.(i, 'done', result);
      } catch (e) {
        results[i] = { ok: false, error: e.message };
        onProgress?.(i, 'error', e);
      }
    }
  });
  await Promise.all(workers);
  return results;
}

// -------- Cost estimate --------

/** Rough cost estimate in USD for a batch. */
export function estimateCost(numImages, model) {
  // Per image: ~1500 input tokens (image + prompt) + ~200 output tokens
  const m = model || getModel();
  let inputCostPerMTok, outputCostPerMTok;
  if (m.includes('haiku')) {
    inputCostPerMTok = 1.0;   outputCostPerMTok = 5.0;     // Haiku 4.5
  } else if (m.includes('opus')) {
    inputCostPerMTok = 15.0;  outputCostPerMTok = 75.0;    // Opus 4.6
  } else {
    inputCostPerMTok = 3.0;   outputCostPerMTok = 15.0;    // Sonnet 4.6
  }
  const totalInputTok = numImages * 1500;
  const totalOutputTok = numImages * 200;
  return (totalInputTok / 1_000_000) * inputCostPerMTok + (totalOutputTok / 1_000_000) * outputCostPerMTok;
}
