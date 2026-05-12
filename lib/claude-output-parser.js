// Parse Claude-generated brief text into a structured brief.
//
// Expected format (loose — handles markdown bold, colons, multi-line values):
//
//   Idea Name: Quiet Overwhelm
//   Angle Name: Cost Anchor
//   Task: Some multi-line task
//     description that wraps
//   General Notes: ...
//
//   Creative 1
//   File Name: Quiet Overwhelm | SV8_Gmail
//   Variation Type: Copy
//   Lead Type: Offer
//   Copy: The actual ad copy here.
//
//   Creative 2
//   ...
//
// Variation markers can be: "Creative 1", "Creative 1:", "## Creative 1",
// "**Creative 1**", "Variation 1", "Video 1", "Copy 1". Numbers 1-10.
//
// Returns: { briefType, overview: {field: value}, creatives: [{field: value}, ...] }
//   briefType is 'static' | 'video' | 'copy' | null (let caller decide)

import { TEMPLATES } from './templates-config.js';

// Build the canonical-name -> {kind, briefTypes} map from TEMPLATES so the
// parser can normalize variant labels (e.g. "AI Allowed" -> "AI Allowed?").
function buildFieldIndex() {
  const index = new Map(); // normalized lowercase label -> { canonical, scope: 'overview'|'creative', briefTypes: Set }
  for (const [briefType, cfg] of Object.entries(TEMPLATES)) {
    for (const def of cfg.overview) {
      const key = normalizeFieldName(def.field);
      if (!index.has(key)) index.set(key, { canonical: def.field, scope: 'overview', briefTypes: new Set() });
      index.get(key).briefTypes.add(briefType);
    }
    for (const def of cfg.creative) {
      const key = normalizeFieldName(def.field);
      if (!index.has(key)) index.set(key, { canonical: def.field, scope: 'creative', briefTypes: new Set() });
      index.get(key).briefTypes.add(briefType);
    }
  }
  // Common synonyms / shorthand the parser should accept
  const aliases = [
    ['ai allowed', 'AI Allowed?', 'overview'],
    ['ratio format', 'Ratio Format(s)', 'overview'],
    ['ratios', 'Ratio Format(s)', 'overview'],
    ['ratio formats', 'Ratio Format(s)', 'overview'],
    ['net new', 'Net New/Iteration', 'overview'],
    ['iteration', 'Net New/Iteration', 'overview'],
    ['filename', 'File Name', 'creative'],
    ['file', 'File Name', 'creative'], // ambiguous but more common than the brief's "File" url cell
    ['video url', 'Video File', 'creative'],
    ['headline 1', 'Headline', 'creative'],
    ['copywriter name', 'Copywriter', 'overview'],
    ['conversion', 'Conversion Objective', 'overview'],
    ['landing page', 'Landing Page URL', 'overview'],
    ['url', 'Landing Page URL', 'overview'],
    ['cta', 'CTA', 'creative'], // for body copy variation type maybe
  ];
  for (const [variant, canonical, scope] of aliases) {
    const k = normalizeFieldName(variant);
    if (!index.has(k)) {
      index.set(k, { canonical, scope, briefTypes: new Set(['static', 'video', 'copy']) });
    }
  }
  return index;
}

function normalizeFieldName(s) {
  return String(s)
    .toLowerCase()
    .replace(/[\*_`]/g, '')   // strip markdown emphasis chars
    .replace(/[\(\)]/g, '')   // strip parens (Brand Voice (variant) -> Brand Voice variant)
    .replace(/[?:]/g, '')     // strip trailing ? :
    .trim();
}

const FIELD_INDEX = buildFieldIndex();

function lookupField(label) {
  return FIELD_INDEX.get(normalizeFieldName(label)) || null;
}

// Match lines like "Field Name: value" — but not URLs or scripts that contain colons mid-text.
// We only treat as a field if the part before the first colon is short (<= 60 chars) AND
// matches a known field name.
function tryParseFieldLine(line) {
  const idx = line.indexOf(':');
  if (idx < 0 || idx > 60) return null;
  const labelRaw = line.slice(0, idx);
  const value = line.slice(idx + 1).trim();
  // Strip leading markdown markers (#, *, -, >, digits.)
  const cleanedLabel = labelRaw.replace(/^[\s>#*\-\d.]+/, '').replace(/[\*_`]/g, '').trim();
  if (!cleanedLabel) return null;
  const field = lookupField(cleanedLabel);
  if (!field) return null;
  return { field: field.canonical, scope: field.scope, value };
}

// Match variation header lines: "Creative 1", "Variation 1", "Video 1", "Copy 1"
// (with optional markdown #, **, : trailing). Also "Creative #1".
const VARIATION_HEADER_RE = /^[\s>#*]*\**\s*(?:Creative|Variation|Video|Copy|Static|Ad)\s*#?\s*(\d+)\b\s*[:\-\u2014]?\s*\**\s*$/i;

function isVariationHeader(line) {
  return VARIATION_HEADER_RE.test(line.trim());
}

/**
 * Detect brief type from the parsed fields. Look for type-distinctive fields.
 *   Video: Footage Folder, Editing Notes, Lead Script, Body Script, Video Type, Video File
 *   Copy:  Copy Type, Headline + Body Copy
 *   Static: Photo Folder, Design Notes, "Copy" cell on a creative
 */
function detectBriefType(overview, creatives) {
  const allFields = new Set([
    ...Object.keys(overview),
    ...creatives.flatMap(c => Object.keys(c)),
  ]);
  // Strong signals
  if (allFields.has('Video Type') || allFields.has('Footage Folder') || allFields.has('Lead Script') || allFields.has('Body Script') || allFields.has('Video File') || allFields.has('Editing Notes')) return 'video';
  if (allFields.has('Copy Type') || (allFields.has('Headline') && allFields.has('Body Copy'))) return 'copy';
  if (allFields.has('Photo Folder') || allFields.has('Design Notes') || (allFields.has('Copy') && !allFields.has('Body Copy'))) return 'static';
  // Header-line clue
  // (caller can also look at the input text directly)
  return null;
}

/**
 * Main parse entry point.
 * @param {string} text Pasted Claude output
 * @returns {{briefType: string|null, overview: object, creatives: object[], unrecognizedFields: string[]}}
 */
export function parseClaudeOutput(text) {
  if (!text || !text.trim()) {
    return { briefType: null, overview: {}, creatives: [], unrecognizedFields: [] };
  }

  const lines = text.replace(/\r\n/g, '\n').split('\n');

  const overview = {};
  const creatives = [];
  let currentBlock = overview;       // either overview or a creative {} object
  let currentField = null;           // the most-recent field name we wrote to
  const unrecognized = [];

  // Process header line ("Brief for Dan Henry | ...") if present
  // (not strictly needed; just skip lines until the first field/header)

  for (const rawLine of lines) {
    const line = rawLine; // preserve internal whitespace; trim only as needed

    // Variation header? -> start a new creative block
    if (isVariationHeader(line)) {
      const newCreative = {};
      creatives.push(newCreative);
      currentBlock = newCreative;
      currentField = null;
      continue;
    }

    // Field line?
    const fv = tryParseFieldLine(line);
    if (fv) {
      // Decide which block to write to: if scope is 'creative' but we haven't
      // yet entered a creative block, start an implicit one (e.g. user only
      // pasted creative-level fields).
      if (fv.scope === 'creative' && currentBlock === overview) {
        const newCreative = {};
        creatives.push(newCreative);
        currentBlock = newCreative;
      }
      currentBlock[fv.field] = fv.value;
      currentField = fv.field;
      continue;
    }

    // Unknown bracket-leading or all-caps section header — note as unrecognized
    // but keep going (don't write to current field if line looks like a header).
    const trimmed = line.trim();
    const looksLikeHeader = /^[A-Z][A-Z\s\-=]{2,}$/.test(trimmed) && !trimmed.includes(' the ');
    const isOnlyMarkdown = /^[#*\->\s]+$/.test(trimmed);
    if (looksLikeHeader || isOnlyMarkdown || trimmed === '') {
      // End the current field's run (a blank/header line breaks continuation)
      currentField = null;
      // Track unknown headers for diagnostic feedback
      if (looksLikeHeader && trimmed.length < 50) {
        // E.g. "OVERVIEW", "CREATIVES" — likely a section break
        // We won't surface these as unrecognized since they're harmless
      }
      continue;
    }

    // Continuation line for the current field (multi-line value)?
    if (currentField != null && currentBlock[currentField] !== undefined) {
      // Append with newline preserved
      currentBlock[currentField] = (currentBlock[currentField] + '\n' + trimmed).trim();
      continue;
    }

    // Otherwise: a line we couldn't slot anywhere
    if (trimmed.length > 0 && trimmed.length < 80) {
      // Heuristic: maybe a "Field Value" without a colon — we won't try to
      // guess. Just record once.
      // (Skip noisy logging.)
    }
  }

  const briefType = detectBriefType(overview, creatives);

  // Special: parser stripped variation type from creatives if the field was
  // recognized but the user wrote markdown like "Variation Type: Lead" inline.
  // No additional cleanup needed.

  return { briefType, overview, creatives, unrecognizedFields: unrecognized };
}
