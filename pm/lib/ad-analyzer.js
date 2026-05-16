// Claude vision API client for categorizing ad images.
//
// Uses BYO API key from localStorage. Never sends the key anywhere except
// directly to api.anthropic.com from the browser (no proxy).
//
// Concurrency: limited to N parallel requests to avoid rate limits.
// Retries: on 429, exponential backoff (max 3 attempts).

const API_KEY_STORAGE = 'pbg.anthropicApiKey.v1';
const OPENAI_KEY_STORAGE = 'pbg.openaiApiKey.v1';
const MODEL_STORAGE = 'pbg.anthropicModel.v1';
const CONCURRENCY_STORAGE = 'pbg.anthropicConcurrency.v1';

export const DEFAULT_MODEL = 'claude-haiku-4-5-20251001';
export const SONNET_MODEL = 'claude-sonnet-4-6';
export const OPUS_MODEL = 'claude-opus-4-6';
// Lowered from 3 to 2: heavy video batches with Whisper + Anthropic in
// parallel × 3 was hanging the browser for some users.
export const DEFAULT_CONCURRENCY = 2;
// Hard per-video timeout (ms). If a single video doesn't finish by this,
// abort it and let the rest of the batch continue.
export const PER_ITEM_TIMEOUT_MS = 120_000;

// -------- Settings --------

export function getApiKey() {
  return localStorage.getItem(API_KEY_STORAGE) || '';
}

export function setApiKey(key) {
  if (key) localStorage.setItem(API_KEY_STORAGE, key);
  else localStorage.removeItem(API_KEY_STORAGE);
}

export function getOpenAIKey() {
  return localStorage.getItem(OPENAI_KEY_STORAGE) || '';
}

export function setOpenAIKey(key) {
  if (key) localStorage.setItem(OPENAI_KEY_STORAGE, key);
  else localStorage.removeItem(OPENAI_KEY_STORAGE);
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

// -------- Transcription (OpenAI Whisper) --------

const WHISPER_MAX_BYTES = 25 * 1024 * 1024; // OpenAI's hard limit

/**
 * Transcribe a video (or audio) file via OpenAI Whisper.
 *
 * Strategy:
 * 1. If the file is already small enough (<25 MB), upload as-is — Whisper
 *    will demux audio itself.
 * 2. Otherwise extract just the audio track in the browser (decode →
 *    downsample to mono 16kHz → WAV). A typical full-length video drops
 *    to ~1 MB per minute as a 16kHz mono WAV, so a 60-min video lands
 *    well under the 25 MB cap.
 *
 * @returns {Promise<{text: string, status: 'ok'|'no-key'|'too-big'|'error', error?: string}>}
 */
export async function transcribeVideoFile(file, opts = {}) {
  const apiKey = opts.openaiKey || getOpenAIKey();
  if (!apiKey) return { text: '', status: 'no-key' };

  let uploadFile = file;

  // For files over the limit, extract audio first
  if (file.size > WHISPER_MAX_BYTES) {
    try {
      const audioBlob = await extractAudioWav(file);
      if (audioBlob.size > WHISPER_MAX_BYTES) {
        return {
          text: '',
          status: 'too-big',
          error: `Even after extracting audio, file is ${(audioBlob.size / 1024 / 1024).toFixed(1)} MB \u2014 still over Whisper\u2019s 25 MB cap.`,
        };
      }
      uploadFile = new File([audioBlob], 'audio.wav', { type: 'audio/wav' });
    } catch (e) {
      return { text: '', status: 'error', error: 'Audio extraction failed: ' + e.message };
    }
  }

  const form = new FormData();
  form.append('file', uploadFile);
  form.append('model', 'whisper-1');
  form.append('response_format', 'text');
  form.append('language', 'en');

  try {
    const res = await fetch('https://api.openai.com/v1/audio/transcriptions', {
      method: 'POST',
      headers: { 'Authorization': `Bearer ${apiKey}` },
      body: form,
    });
    if (!res.ok) {
      const errText = await res.text();
      console.warn('Whisper transcription failed:', res.status, errText.slice(0, 200));
      return {
        text: '',
        status: 'error',
        error: `HTTP ${res.status}: ${errText.slice(0, 200)}`,
      };
    }
    const text = await res.text();
    return { text: text.trim(), status: 'ok' };
  } catch (e) {
    console.warn('Whisper transcription error:', e);
    return { text: '', status: 'error', error: e.message };
  }
}

/**
 * Extract the audio track from a video file as a 16kHz mono WAV blob.
 * Uses Web Audio API's decodeAudioData which handles MP4/MOV/WebM/etc.
 *
 * Note: decodeAudioData reads the ENTIRE file into memory. Very long videos
 * (>30 min HD) may run out of memory. For ad creatives (typically <2 min)
 * this is a non-issue.
 */
async function extractAudioWav(file) {
  const ab = await file.arrayBuffer();
  const ctx = new (window.AudioContext || window.webkitAudioContext)();
  let audioBuffer;
  try {
    audioBuffer = await ctx.decodeAudioData(ab);
  } finally {
    ctx.close();
  }
  // Downmix to mono and resample to 16kHz using OfflineAudioContext
  const targetRate = 16000;
  const targetLen = Math.ceil(audioBuffer.duration * targetRate);
  const offline = new OfflineAudioContext(1, targetLen, targetRate);
  const src = offline.createBufferSource();
  src.buffer = audioBuffer;
  src.connect(offline.destination);
  src.start(0);
  const rendered = await offline.startRendering();
  return audioBufferToWav(rendered);
}

/**
 * Encode an AudioBuffer as a 16-bit PCM WAV Blob.
 */
function audioBufferToWav(buffer) {
  const numChannels = buffer.numberOfChannels;
  const sampleRate = buffer.sampleRate;
  const numFrames = buffer.length;
  const bytesPerSample = 2; // 16-bit
  const blockAlign = numChannels * bytesPerSample;
  const byteRate = sampleRate * blockAlign;
  const dataSize = numFrames * blockAlign;
  const headerSize = 44;
  const totalSize = headerSize + dataSize;

  const ab = new ArrayBuffer(totalSize);
  const dv = new DataView(ab);

  // RIFF header
  writeString(dv, 0, 'RIFF');
  dv.setUint32(4, totalSize - 8, true);
  writeString(dv, 8, 'WAVE');
  // fmt chunk
  writeString(dv, 12, 'fmt ');
  dv.setUint32(16, 16, true);          // chunk size
  dv.setUint16(20, 1, true);           // PCM format
  dv.setUint16(22, numChannels, true);
  dv.setUint32(24, sampleRate, true);
  dv.setUint32(28, byteRate, true);
  dv.setUint16(32, blockAlign, true);
  dv.setUint16(34, 16, true);          // bits per sample
  // data chunk
  writeString(dv, 36, 'data');
  dv.setUint32(40, dataSize, true);

  // Interleave channels + write samples
  let offset = headerSize;
  const channels = [];
  for (let c = 0; c < numChannels; c++) channels.push(buffer.getChannelData(c));
  for (let i = 0; i < numFrames; i++) {
    for (let c = 0; c < numChannels; c++) {
      let s = Math.max(-1, Math.min(1, channels[c][i]));
      s = s < 0 ? s * 0x8000 : s * 0x7FFF;
      dv.setInt16(offset, s, true);
      offset += 2;
    }
  }
  return new Blob([ab], { type: 'audio/wav' });
}

function writeString(dv, offset, str) {
  for (let i = 0; i < str.length; i++) dv.setUint8(offset + i, str.charCodeAt(i));
}

// -------- Video keyframe extraction --------

const NUM_KEYFRAMES = 5;
// Reduced from 1024 to 640 to cut Claude payload roughly in half per frame
// (faster upload, less browser memory, fewer timeouts on big batches).
// Captions in ad creatives are still readable at 640px.
const VIDEO_FRAME_MAX_DIM = 640;

/**
 * Returns true if file appears to be a video by extension.
 */
export function isVideoFile(file) {
  if (!file?.name) return false;
  return /\.(mp4|mov|webm|m4v|avi)$/i.test(file.name);
}

/**
 * Extract N evenly-spaced keyframes from a video file.
 * Returns { duration: seconds, frames: Array<{ base64, mediaType, timestamp }> }.
 */
export async function extractKeyframes(file, numFrames = NUM_KEYFRAMES) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const video = document.createElement('video');
    video.preload = 'auto';
    video.muted = true;
    video.playsInline = true;
    video.src = url;

    let duration = 0;
    let canvas = null;
    let ctx = null;

    const cleanup = () => URL.revokeObjectURL(url);

    video.addEventListener('loadedmetadata', () => {
      duration = video.duration || 0;
      if (!isFinite(duration) || duration <= 0) {
        cleanup();
        return reject(new Error('Could not read video duration'));
      }
      // Compute scaled canvas size (preserve aspect ratio, cap longest side)
      const vw = video.videoWidth || 640;
      const vh = video.videoHeight || 360;
      const scale = Math.min(1, VIDEO_FRAME_MAX_DIM / Math.max(vw, vh));
      const cw = Math.round(vw * scale);
      const ch = Math.round(vh * scale);
      canvas = document.createElement('canvas');
      canvas.width = cw;
      canvas.height = ch;
      ctx = canvas.getContext('2d');
      // Compute timestamps: start, evenly spaced, end (slightly inside bounds)
      const timestamps = [];
      for (let i = 0; i < numFrames; i++) {
        const t = (duration * (i + 0.5)) / numFrames;
        timestamps.push(Math.min(Math.max(t, 0.05), duration - 0.05));
      }
      sampleNext(0, timestamps, []);
    });

    function sampleNext(idx, timestamps, frames) {
      if (idx >= timestamps.length) {
        cleanup();
        return resolve({ duration, frames });
      }
      const onSeek = () => {
        video.removeEventListener('seeked', onSeek);
        try {
          ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
          // JPEG keeps file size much smaller than PNG for photo-like frames
          const dataUrl = canvas.toDataURL('image/jpeg', 0.85);
          const base64 = dataUrl.slice(dataUrl.indexOf(',') + 1);
          frames.push({ base64, mediaType: 'image/jpeg', timestamp: timestamps[idx] });
          sampleNext(idx + 1, timestamps, frames);
        } catch (e) {
          cleanup();
          reject(e);
        }
      };
      video.addEventListener('seeked', onSeek, { once: true });
      video.currentTime = timestamps[idx];
    }

    video.addEventListener('error', () => {
      cleanup();
      reject(new Error('Video decode failed'));
    });
  });
}

// -------- Categorization prompt --------

const SYSTEM_PROMPT = `You categorize ad creatives for the BAD Marketing Creative Tracker. Return STRICT JSON, nothing else, no markdown fences.`;

const VIDEO_SYSTEM_PROMPT = `You categorize video ad creatives for the BAD Marketing Creative Tracker. You will be shown 5 keyframes sampled across the video's duration (start, 25%, 50%, 75%, end). Use them together to understand the ad's narrative arc. Return STRICT JSON, nothing else, no markdown fences.`;

function buildVideoUserPrompt(filename, durationSec, filenameHints, transcript, briefCandidates) {
  let hintsBlock = '';
  if (filenameHints && Object.keys(filenameHints).length > 0) {
    const lines = [];
    if (filenameHints.ideaName) lines.push(`- Likely Idea Name (from filename): "${filenameHints.ideaName}"`);
    if (filenameHints.leadType) lines.push(`- Lead Type (from filename, authoritative): "${filenameHints.leadType}"`);
    if (filenameHints.cid) lines.push(`- CID (from filename, authoritative): "${filenameHints.cid}"`);
    if (lines.length > 0) hintsBlock = `\n\nFilename hints (use as starting point, override fields below):\n${lines.join('\n')}\n`;
  }

  const transcriptBlock = transcript
    ? `\n\n=== TRANSCRIPT (full audio of the video, authoritative for copy/script questions) ===\n${transcript}\n=== END TRANSCRIPT ===\n`
    : `\n\n(No transcript available — rely on visible on-screen captions in the keyframes. Read every visible caption carefully.)\n`;

  let briefMatchBlock = '';
  let briefMatchInstructions = '';
  if (briefCandidates && briefCandidates.length > 0) {
    const candidatesText = briefCandidates.map((c, i) =>
      `[${i}] Brief: ${c.briefDisplayLabel} | Variation ${c.variationIndex + 1}\n  Lead Script: ${(c.leadScript || '(blank)').slice(0, 600)}\n  Body Script: ${(c.bodyScript || '(blank)').slice(0, 600)}`
    ).join('\n\n');
    briefMatchBlock = `\n\n=== CANDIDATE BRIEF VARIATIONS (scripts that may have been recorded as this video) ===\n${candidatesText}\n=== END CANDIDATES ===\n`;
    briefMatchInstructions = `

BRIEF MATCHING (very important):
- The transcript is a recording of a real human delivering one of the candidate scripts above.
- Talent improvises, paraphrases, drops words, restructures sentences, and adds filler. The transcript will RARELY be word-for-word.
- Match generously: if the transcript covers the same TOPIC + KEY PHRASES as a candidate's Lead Script (and roughly the same Body Script), that's a match.
- Pay attention to the OPENING: does the talent's first 1-2 sentences riff on the candidate's Lead Script?
- Don't require exact word match. Don't require all sentences to align. Don't require the same length.
- If 2+ candidates seem similar, pick the one whose Lead Script is the closest match to the opening 1-2 sentences of the transcript.
- Set "matchedCandidateIndex" to the integer index of the best match (e.g., 0, 1, 2...).
- Only set to null if NO candidate's topic / key phrases overlap meaningfully (e.g., the video is about a completely different subject).
- ALSO populate "matchRationale": one sentence explaining WHY you chose that index (or null), QUOTING the strongest matching phrase from the candidate's Lead Script.`;
  }

  const durStr = isFinite(durationSec) && durationSec > 0 ? ` (${Math.round(durationSec)}s long)` : '';
  return `You're shown 5 keyframes from a video ad${durStr}, sampled at 10%, 30%, 50%, 70%, and 90% of duration. Use the full sequence to understand the ad's hook, body, and CTA.

The MOST important signal for Lead Type is the OPENING WORDS \u2014 the first sentence the speaker says. Lead Type is determined by what the hook line is doing rhetorically, not by the visuals.${transcriptBlock}

Return this exact JSON structure (and nothing else):

{
  "awarenessLevel": "<one of: Most Aware | Solution Aware | Problem Aware | Unaware>",
  "leadType": "<one of: Offer | Promise | Problem-Solution | Secret | Proclamation | Story>",
  "variationType": "<one of: Lead | Pattern Interrupt | Body | CTA>",
  "ideaName": "<2-4 words describing the core ad concept>",
  "angleName": "<2-4 words describing the persuasion angle>",
  "styleName": "<2-4 words describing the video format, e.g. UGC Talking Head, Animated Explainer, Podcast Clip, Testimonial>",
  "rationale": "<one sentence explaining your Lead Type choice. QUOTE THE OPENING LINE OF THE TRANSCRIPT (or first visible caption) verbatim.>",
  "matchedCandidateIndex": <integer index of matching candidate brief variation, or null if no match. Always include this field.>,
  "matchRationale": "<one sentence explaining the match decision. If matched, QUOTE the strongest overlapping phrase. If null, explain why no candidate fit. Empty string if no candidates were provided.>"
}

Lead Type definitions \u2014 pick based on the opening line specifically:
- "Offer": opens with a deal/discount/free thing ("Get the framework free...", "$500 off if you book today")
- "Promise": opens with a desirable outcome ("I'll show you how to make $10K/mo", "You're about to learn the secret to...")
- "Problem-Solution": opens by naming a pain point + a fix ("Tired of X? Here's how to Y")
- "Secret": opens with a hidden truth or insider knowledge ("Almost nobody knows this about...", "Here's what they don't tell you")
- "Proclamation": opens with a bold contrarian claim ("Going viral is a trap", "Most webinars are broken")
- "Story": opens with personal narrative ("I delivered pizzas for 7 years before...", "Last year I lost $40K because...")

Other field guidance:
- "awarenessLevel": stage of customer awareness this ad targets
- "variationType": almost always "Lead" for new ads (the hook is the variation). "Pattern Interrupt" for unusual visual openers, "Body" if testing variations of body copy with same hook, "CTA" if testing endings
- "ideaName": short noun phrase capturing the central idea
- "angleName": short noun phrase capturing the persuasion angle
- "styleName": describe the video format / style

Filename: ${filename}${hintsBlock}${briefMatchBlock}${briefMatchInstructions}

Return ONLY the JSON object. No preamble, no explanation, no markdown.`;
}

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

/**
 * Extract keyframes from a video file and send them as a multi-image
 * request to Claude. Returns { ...categorization, durationSec, keyframes }.
 *
 * `keyframes` is the same array returned by extractKeyframes() so the UI
 * can show what Claude saw.
 */
export async function analyzeVideoFile(file, filenameHints, opts) {
  const apiKey = opts.apiKey || getApiKey();
  if (!apiKey) throw new Error('No API key configured. Open Settings to add one.');
  const model = opts.model || getModel();

  // Transcribe + extract keyframes in parallel — both are independent
  const [{ duration, frames }, transcribeResult] = await Promise.all([
    extractKeyframes(file),
    transcribeVideoFile(file, { openaiKey: opts.openaiKey }),
  ]);
  const transcript = transcribeResult.text;
  const transcriptStatus = transcribeResult.status;
  const transcriptError = transcribeResult.error;

  // Build the user content: each keyframe as an image block + a final text block
  const content = [];
  for (const frame of frames) {
    content.push({
      type: 'image',
      source: { type: 'base64', media_type: frame.mediaType, data: frame.base64 },
    });
  }
  content.push({
    type: 'text',
    text: buildVideoUserPrompt(file.name, duration, filenameHints, transcript, opts.briefCandidates),
  });

  const body = {
    model,
    max_tokens: 500,
    system: VIDEO_SYSTEM_PROMPT,
    messages: [{ role: 'user', content }],
  };

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
    const parsed = parseJsonResponse(text);
    return {
      ...parsed,
      durationSec: duration,
      keyframes: frames,
      transcript,
      transcriptStatus,
      transcriptError,
    };
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
  const openaiKey = opts.openaiKey || getOpenAIKey();
  const model = opts.model || getModel();
  const briefCandidates = opts.briefCandidates || null;
  const results = new Array(jobs.length);

  let nextIndex = 0;
  const workers = Array.from({ length: Math.min(concurrency, jobs.length) }, async () => {
    while (true) {
      const i = nextIndex++;
      if (i >= jobs.length) return;
      const job = jobs[i];
      onProgress?.(i, 'analyzing');
      try {
        const fn = isVideoFile(job.file) ? analyzeVideoFile : analyzeAdImage;
        const callOpts = isVideoFile(job.file)
          ? { apiKey, openaiKey, model, briefCandidates }
          : { apiKey, openaiKey, model };
        // Race the call against a hard per-item timeout so one slow/hung
        // video can't block the whole batch indefinitely.
        const result = await withTimeout(
          fn(job.file, job.filenameHints, callOpts),
          PER_ITEM_TIMEOUT_MS,
          `Timed out after ${PER_ITEM_TIMEOUT_MS / 1000}s`,
        );
        results[i] = { ok: true, result };
        onProgress?.(i, 'done', result);
      } catch (e) {
        results[i] = { ok: false, error: e.message };
        onProgress?.(i, 'error', e);
      }
      // Yield to the UI thread between jobs so the browser stays responsive.
      await new Promise(r => setTimeout(r, 0));
    }
  });
  await Promise.all(workers);
  return results;
}

function withTimeout(promise, ms, msg) {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error(msg)), ms);
    promise.then(
      v => { clearTimeout(t); resolve(v); },
      e => { clearTimeout(t); reject(e); },
    );
  });
}

// -------- Cost estimate --------

/** Rough cost estimate in USD for a batch. */
export function estimateCost(numItems, model, mediaKind = 'static') {
  // Per static: ~1200 input tokens (image + prompt) + ~200 output tokens
  // Per video: ~3500 input tokens (5 keyframes at 640px + prompt) + ~250 output
  const m = model || getModel();
  let inputCostPerMTok, outputCostPerMTok;
  if (m.includes('haiku')) {
    inputCostPerMTok = 1.0;   outputCostPerMTok = 5.0;
  } else if (m.includes('opus')) {
    inputCostPerMTok = 15.0;  outputCostPerMTok = 75.0;
  } else {
    inputCostPerMTok = 3.0;   outputCostPerMTok = 15.0;
  }
  const inputPerItem = mediaKind === 'video' ? 3500 : 1200;
  const outputPerItem = mediaKind === 'video' ? 250 : 200;
  const totalInputTok = numItems * inputPerItem;
  const totalOutputTok = numItems * outputPerItem;
  return (totalInputTok / 1_000_000) * inputCostPerMTok + (totalOutputTok / 1_000_000) * outputCostPerMTok;
}
