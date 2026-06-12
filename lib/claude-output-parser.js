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

import { TEMPLATES, DROPDOWN_OPTIONS, normalizeDropdownValue } from './templates-config.js';

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
    .replace(/\s*\/\s*/g, '/')// collapse spaces around slashes ("Net New / Iteration" -> "net new/iteration")
    .replace(/\s+/g, ' ')     // collapse runs of whitespace
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

/**
 * If a line is JUST a known field name (no colon, no value), return the
 * field info. Used by the "field name on one line, value on next line"
 * pasted format that comes from copy-pasting a Google Doc table.
 */
function tryMatchBareFieldName(line) {
  const cleaned = line.replace(/[\*_`]/g, '').trim();
  // Strip a single trailing colon if present (e.g. "Idea Name:" with no value)
  const noColon = cleaned.replace(/[:?]?\s*$/, (m) => m.includes('?') ? m : '');
  // Don't match if there's any internal colon or whitespace pattern that suggests
  // free text — but DO allow short multi-word names ("Idea Name", "Brand Voice").
  if (cleaned.length === 0 || cleaned.length > 50) return null;
  if (cleaned.includes(':')) return null; // those go through the colon-line parser
  // Skip lines that look like values not labels (e.g. start with http, contain numbers
  // ending in a unit, or are clearly sentence text)
  if (/^https?:\/\//i.test(cleaned)) return null;
  if (/[.!?]$/.test(cleaned) && cleaned.length > 25) return null; // probably a sentence
  // The whole line must look like a field name
  const field = lookupField(cleaned);
  if (!field) return null;
  return { field: field.canonical, scope: field.scope };
}

// Match variation header lines: "Creative 1", "Variation 1", "Video 1", "Copy 1"
// (with optional markdown #, **, : trailing). Also "Creative #1".
// Also matches a bare number on its own line ("1", "2", ...) — common when
// the user copy-pastes from a Google Doc where each variation table starts
// with a number cell.
const VARIATION_HEADER_RE = /^[\s>#*]*\**\s*(?:Creative|Variation|Video|Copy|Static|Ad)\s*#?\s*(\d+)\b(?:\s*[:\-\u2014].*)?\s*\**\s*$/i;
const BARE_NUMBER_RE = /^\s*(\d{1,2})\s*$/;

// Section-break lines that separate top-level brief sections (not fields and not
// variation headers). Used to terminate multi-line value runs and the shared
// locked-body capture so stray section text doesn't bleed into the prior field.
const SECTION_BREAK_RE = /^[\s>#*]*\**\s*(overview|videos?|deliverables?|locked body|locked body conversation|batch|total|hooks?|hook variants?|repurpose)\b/i;

function isVariationHeader(line) {
  if (VARIATION_HEADER_RE.test(line.trim())) return true;
  // Also treat a bare "1" / "2" / ... on its own line as a variation marker.
  const bare = BARE_NUMBER_RE.exec(line);
  if (bare) {
    const n = parseInt(bare[1], 10);
    if (n >= 1 && n <= 20) return true;
  }
  return false;
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

// -------- Google-Doc-table ("packed") preprocessing --------
//
// When a user copy-pastes a 2-column brief table out of Google Docs, each table
// collapses onto ONE line: cells separated by runs of 3+ spaces, and inside a
// cell the format is "Field Name<space>Value" with NO colon. The header row
// ("Field" | "Details") shows up as a leading "Field Details" cell. Each table
// block (overview + every video/creative) starts with that header.
//
// The line-based parser below can't read that. So we detect packed lines and
// expand them back into the canonical "Field: value" lines it understands.

// Find the LONGEST known field-name prefix of a single table cell and split it
// into { field, scope, value }. Returns null if the cell starts with no field
// name (e.g. the "Field Details" header cell) so the caller can drop it.
function matchFieldPrefix(cell) {
  const cleaned = cell.replace(/^[\s>#*\-]+/, '').trim();
  if (!cleaned) return null;
  const words = cleaned.split(/\s+/);
  let best = null; // { field, scope, n }
  const maxN = Math.min(words.length, 6); // longest real field name is 4 words
  for (let n = 1; n <= maxN; n++) {
    const candidate = words.slice(0, n).join(' ').replace(/:$/, '');
    const f = lookupField(candidate);
    if (f) best = { field: f.canonical, scope: f.scope, n };
  }
  if (!best) return null;
  const value = words.slice(best.n).join(' ').trim();
  return { field: best.field, scope: best.scope, value };
}

// Is this physical line a packed Google-Doc table row (the whole table on one
// line)? True when it starts with the "Field Details" header and has cells, or
// when it has many 3-space-separated cells that resolve to known field names.
function isPackedTableLine(line) {
  const cells = line.split(/\s{3,}/).map(s => s.trim()).filter(Boolean);
  if (cells.length < 2) return false;
  if (/^\s*Field\s+Details\b/i.test(line)) return true;
  let hits = 0;
  for (const c of cells) if (matchFieldPrefix(c)) hits++;
  return hits >= 2;
}

// Drop leading and trailing blank lines from an array of lines, keeping any
// blank lines in the middle (so internal paragraph breaks survive).
function trimBlankEdges(arr) {
  let start = 0;
  let end = arr.length;
  while (start < end && arr[start].trim() === '') start++;
  while (end > start && arr[end - 1].trim() === '') end--;
  return arr.slice(start, end);
}

// Expand a packed line into ["Field: value", ...] lines. Cells with no known
// field prefix (the header cell) are dropped.
function expandPackedLine(line) {
  const out = [];
  const cells = line.split(/\s{3,}/).map(s => s.trim()).filter(Boolean);
  for (const cell of cells) {
    const m = matchFieldPrefix(cell);
    if (!m) continue; // header cell ("Field Details") or junk
    out.push(m.value ? `${m.field}: ${m.value}` : `${m.field}:`);
  }
  return out;
}

// If the pasted text contains any packed table lines, rebuild it into clean
// per-line text. Also lifts the shared "BODY SCRIPT" section and fans it out to
// every video variation (the template stores Body Script per-creative). When no
// packed line is present, the original text is returned untouched.
function expandGoogleDocTable(text) {
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  if (!lines.some(isPackedTableLine)) return text;

  const out = [];
  let mode = 'normal'; // 'normal' | 'bodyscript'
  let sharedBody = [];

  for (const line of lines) {
    const trimmed = line.trim();

    if (isPackedTableLine(line)) {
      mode = 'normal';
      for (const expanded of expandPackedLine(line)) out.push(expanded);
      continue;
    }

    // "BODY SCRIPT (Shared across all 5 variations)" — start collecting the
    // shared body script; don't emit the header itself.
    if (/^body script\b/i.test(trimmed)) {
      mode = 'bodyscript';
      sharedBody = [];
      continue;
    }

    // Variation header ("VIDEO 1", "Creative 2", ...) — emit it, then inject the
    // shared body script into this creative so all variations get it. Internal
    // blank lines are kept (paragraph breaks); only leading/trailing ones drop.
    if (isVariationHeader(line)) {
      mode = 'normal';
      out.push(line);
      const body = trimBlankEdges(sharedBody);
      if (body.length) {
        out.push(`Body Script: ${body[0].trim()}`);
        for (let i = 1; i < body.length; i++) out.push(body[i]);
      }
      continue;
    }

    // Inside the shared body script: collect every line (including blanks) until
    // the next variation header / packed table line terminates the section.
    if (mode === 'bodyscript') {
      sharedBody.push(line);
      continue;
    }

    out.push(line);
  }

  return out.join('\n');
}

// -------- Tab-separated + locked-body preprocessing --------
//
// Our generated testing-batch briefs come as a tab-separated 2-column table
// ("Field<TAB>Value") plus a single shared "Locked Body Conversation" block that
// every video reuses. The two helpers below normalize that shape into the
// canonical "Field: value" lines the line parser understands, and fan the shared
// body out to each video so the form fills completely from one paste.

// A value that points at the shared locked body instead of supplying its own,
// e.g. "[same as locked body]" or "(same body)".
const SAME_BODY_REF_RE = /\bsame\b[\s\S]*\bbody\b/i;

// Convert tab-separated "Field<TAB>Value" rows into "Field: value" lines.
// Handles 2-column and wider rows (Field/Value pairs across the row). Pure
// bracket/paren placeholders ("[fill in]") are blanked so they don't land in the
// form as literal text — EXCEPT same-as-body references, which are preserved so
// expandLockedBody can fan the shared body into them.
function normalizeTabFields(text) {
  if (!text.includes('\t')) return text;
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  const out = [];
  for (const line of lines) {
    if (!line.includes('\t')) { out.push(line); continue; }
    const cells = line.split(/\t+/).map(c => c.trim());
    let emittedAny = false;
    let i = 0;
    while (i < cells.length) {
      const labelCell = cells[i];
      const f = labelCell ? lookupField(labelCell.replace(/[\*_`]/g, '').trim()) : null;
      if (f) {
        const valueCell = (cells[i + 1] || '').trim();
        const isBracket = /^\[.*\]$/.test(valueCell) || /^\(.*\)$/.test(valueCell);
        const isSameRef = SAME_BODY_REF_RE.test(valueCell);
        if (valueCell && (!isBracket || isSameRef)) {
          out.push(`${f.canonical}: ${valueCell}`);
        } else {
          out.push(`${f.canonical}:`);
        }
        emittedAny = true;
        i += 2;
      } else {
        i += 1;
      }
    }
    if (!emittedAny) out.push(line);
  }
  return out.join('\n');
}

// Lift the shared "Locked Body Conversation" block and fan its Body Script out to
// every video whose Body Script is empty, a bracket placeholder, or a same-as-body
// reference. Removes the standalone locked-body section so it doesn't parse as a
// stray creative. No-op when there's no locked-body header.
function expandLockedBody(text) {
  const lines = text.replace(/\r\n/g, '\n').split('\n');
  let headerIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    if (/^[\s>#*]*\**\s*locked body\b/i.test(lines[i].trim())) { headerIdx = i; break; }
  }
  if (headerIdx === -1) return text;

  // Capture the shared body from after the header until a boundary line.
  const bodyLines = [];
  let endIdx = lines.length;
  for (let i = headerIdx + 1; i < lines.length; i++) {
    const ln = lines[i];
    const t = ln.trim();
    const isLockedHeader = /^[\s>#*]*\**\s*locked body\b/i.test(t);
    if (isVariationHeader(ln) || (SECTION_BREAK_RE.test(t) && !isLockedHeader)) { endIdx = i; break; }
    const fv = tryParseFieldLine(ln);
    if (fv && fv.field !== 'Body Script') { endIdx = i; break; }
    const bare = tryMatchBareFieldName(ln);
    if (bare && bare.field !== 'Body Script') { endIdx = i; break; }
    bodyLines.push(ln);
  }

  // Strip a leading "Body Script:" label, keeping just the value text.
  let captured = bodyLines.slice();
  while (captured.length && captured[0].trim() === '') captured.shift();
  if (captured.length) {
    const fv0 = tryParseFieldLine(captured[0]);
    if (fv0 && fv0.field === 'Body Script') captured[0] = fv0.value;
  }
  captured = trimBlankEdges(captured);
  const sharedBody = captured.map(l => (l.trim() === '' ? '' : l.trim()));
  if (!sharedBody.length) return text;

  // Rebuild without the locked-body section, fanning the body into each video.
  const merged = [...lines.slice(0, headerIdx), ...lines.slice(endIdx)];
  const out = [];
  for (const ln of merged) {
    const fv = tryParseFieldLine(ln);
    if (fv && fv.field === 'Body Script') {
      const v = fv.value.trim();
      const isBracket = /^\[.*\]$/.test(v) || /^\(.*\)$/.test(v);
      const isSameRef = SAME_BODY_REF_RE.test(v);
      if (!v || isSameRef || isBracket) {
        out.push(`Body Script: ${sharedBody[0]}`);
        for (let i = 1; i < sharedBody.length; i++) out.push(sharedBody[i]);
        continue;
      }
    }
    out.push(ln);
  }
  return out.join('\n');
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

  text = normalizeTabFields(text);
  text = expandLockedBody(text);
  text = expandGoogleDocTable(text);

  const lines = text.replace(/\r\n/g, '\n').split('\n');

  const overview = {};
  const creatives = [];
  let currentBlock = overview;       // either overview or a creative {} object
  let currentField = null;           // the most-recent field name we wrote to
  const unrecognized = [];

  // Process header line ("Brief for Dan Henry | ...") if present
  // (not strictly needed; just skip lines until the first field/header)

  // We use a "pending field" for the no-colon format: when a line is JUST a
  // field name (no colon), we stash it as `pendingField` and treat the next
  // non-blank line as its value. Continuation lines append until we hit
  // another field or variation marker.
  let pendingField = null; // { field, scope } from a label-only line
  let pendingBlanks = 0;   // blank lines buffered inside a multi-line value

  for (const rawLine of lines) {
    const line = rawLine; // preserve internal whitespace; trim only as needed
    const trimmed = line.trim();

    // Variation header? -> start a new creative block
    if (isVariationHeader(line)) {
      const newCreative = {};
      creatives.push(newCreative);
      currentBlock = newCreative;
      currentField = null;
      pendingField = null;
      pendingBlanks = 0;
      continue;
    }

    // 1. Field line WITH colon: "Idea Name: value"
    const fv = tryParseFieldLine(line);
    if (fv) {
      // If pendingField was set but we hit a new field before a value,
      // discard the pending one (label with no value).
      pendingField = null;
      pendingBlanks = 0;
      if (fv.scope === 'creative' && currentBlock === overview) {
        const newCreative = {};
        creatives.push(newCreative);
        currentBlock = newCreative;
      }
      currentBlock[fv.field] = fv.value;
      currentField = fv.field;
      continue;
    }

    // 2. Bare field name (no colon, no value): stash as pendingField
    const bareField = tryMatchBareFieldName(line);
    if (bareField) {
      pendingField = bareField;
      currentField = null; // stop appending continuation lines until we get a value
      pendingBlanks = 0;
      continue;
    }

    // Blank line. Inside a multi-line value it's a potential paragraph break:
    // buffer it and only commit it if more value text follows (so trailing
    // blanks before the next section are dropped). Otherwise it just clears a
    // half-started field.
    if (trimmed === '') {
      if (currentField != null && currentBlock[currentField] !== undefined) {
        pendingBlanks++;
      } else {
        pendingField = null;
      }
      continue;
    }

    // Unknown bracket-leading or all-caps section header — skip cleanly
    const looksLikeHeader = /^[A-Z][A-Z\s\-=]{2,}$/.test(trimmed) && !trimmed.includes(' the ');
    const isOnlyMarkdown = /^[#*\->\s]+$/.test(trimmed);
    if (looksLikeHeader || isOnlyMarkdown || SECTION_BREAK_RE.test(trimmed)) {
      // End the current field's run (a header line breaks continuation)
      currentField = null;
      pendingField = null;
      pendingBlanks = 0;
      continue;
    }

    // 3. We have content. If pendingField is set, this is its value.
    if (pendingField) {
      // Strip placeholder values like "[as generated in the creative tracker]"
      const valueClean = trimmed;
      if (valueClean.startsWith('[') && valueClean.endsWith(']')) {
        // Bracket placeholder — skip the value, but consume pendingField so
        // we don't accidentally bind subsequent text to it.
        pendingField = null;
        currentField = null;
        continue;
      }
      if (pendingField.scope === 'creative' && currentBlock === overview) {
        const newCreative = {};
        creatives.push(newCreative);
        currentBlock = newCreative;
      }
      currentBlock[pendingField.field] = valueClean;
      currentField = pendingField.field;
      pendingField = null;
      pendingBlanks = 0;
      continue;
    }

    // 4. Continuation line for the most-recent field (multi-line value).
    // Commit any buffered blank lines as paragraph breaks, then append.
    if (currentField != null && currentBlock[currentField] !== undefined) {
      const sep = '\n'.repeat(pendingBlanks + 1);
      pendingBlanks = 0;
      currentBlock[currentField] = (currentBlock[currentField] + sep + trimmed).trim();
      continue;
    }

    // Otherwise: a line we couldn't slot anywhere — silently skip
  }

  const briefType = detectBriefType(overview, creatives);

  // Validate dropdown values: any 'kind: dropdown' field whose value doesn't
  // match a valid option goes into `warnings` with a suggested fix.
  const warnings = briefType
    ? collectDropdownWarnings(briefType, overview, creatives)
    : [];

  return { briefType, overview, creatives, warnings, unrecognizedFields: unrecognized };
}

// -------- Dropdown validation + fuzzy suggestions --------

/**
 * For each dropdown field that was filled in, check the value against the
 * valid options. If invalid, attempt a fuzzy match and produce a Warning
 * the UI can show for accept/reject.
 *
 * Warning shape:
 *   {
 *     scope: 'overview' | 'creative',
 *     creativeIndex: number | null,
 *     field: 'Lead Type',
 *     originalValue: 'Hook',
 *     suggestion: 'Hook' (the closest valid option) | null,
 *     options: ['Offer', 'Promise', ...],  // all valid options for this field
 *   }
 */
function collectDropdownWarnings(briefType, overview, creatives) {
  const cfg = TEMPLATES[briefType];
  const warnings = [];

  // Check one dropdown field. If the value already matches an option, do nothing.
  // Otherwise try alias/partial normalization and write the resolved value back
  // in place. Only surface a warning when it still can't be resolved.
  const check = (obj, def, scope, creativeIndex) => {
    if (def.kind !== 'dropdown') return;
    const val = obj[def.field];
    if (val == null || val === '') return;
    const opts = DROPDOWN_OPTIONS[def.options] || [];
    if (matchesOption(val, opts)) return;

    const normalized = normalizeDropdownValue(val, opts, def.field);
    if (matchesOption(normalized, opts)) {
      obj[def.field] = normalized; // auto-apply the resolved option
      return;
    }

    warnings.push({
      scope,
      creativeIndex,
      field: def.field,
      originalValue: val,
      suggestion: closestOption(val, opts),
      options: opts,
    });
  };

  for (const def of cfg.overview) check(overview, def, 'overview', null);
  creatives.forEach((c, idx) => {
    for (const def of cfg.creative) check(c, def, 'creative', idx);
  });

  return warnings;
}

function matchesOption(value, options) {
  const v = String(value).trim().toLowerCase();
  return options.some(o => o.toLowerCase() === v);
}

/**
 * Return the closest matching option using a normalized-text + Levenshtein
 * distance heuristic. Returns null if no option is reasonably close.
 */
function closestOption(value, options) {
  if (!options.length) return null;
  const v = normalizeForMatch(value);

  // Try common aliases first (mirrors brief-to-tracker mapping)
  // Lead -> Hook etc. is brief-to-tracker territory; here we focus on
  // "user mistyped the dropdown value".
  // First pass: any option that's a substring of the value or vice versa
  for (const opt of options) {
    const o = normalizeForMatch(opt);
    if (v === o) return opt;
    if (v.includes(o) || o.includes(v)) return opt;
  }
  // Second pass: Levenshtein distance
  let best = null;
  let bestDist = Infinity;
  for (const opt of options) {
    const d = levenshtein(v, normalizeForMatch(opt));
    if (d < bestDist) {
      bestDist = d;
      best = opt;
    }
  }
  // Only suggest if reasonably close: distance <= 30% of the longer string
  const len = Math.max(v.length, normalizeForMatch(best || '').length);
  if (best && bestDist <= Math.max(2, Math.ceil(len * 0.3))) return best;
  return null;
}

function normalizeForMatch(s) {
  return String(s).toLowerCase().replace(/[\s\-_]+/g, '').replace(/[?!.,]/g, '');
}

function levenshtein(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;
  const dp = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    let prev = dp[0];
    dp[0] = i;
    for (let j = 1; j <= b.length; j++) {
      const tmp = dp[j];
      dp[j] = a[i - 1] === b[j - 1]
        ? prev
        : Math.min(prev, dp[j], dp[j - 1]) + 1;
      prev = tmp;
    }
  }
  return dp[b.length];
}
