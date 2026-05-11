// Map a parsed brief → tracker rows (block 1 + block 2) using TRACKERS config.

import { TRACKERS, TRANSFORMS, extractCid, extractVariationName } from './tracker-config.js';

/**
 * @param {{briefType: string, overview: object, variations: object[]}} brief
 * @param {{
 *   requestDoc?: string,
 *   briefTypeOverride?: string,
 *   clientTrackerOverrides?: object  // optional per-client layout tweaks
 * }} options
 * @returns {{
 *   tracker: 'static'|'video'|'copy',
 *   trackerLabel: string,
 *   block1: { startCol: string, headers: string[], rows: string[][] },
 *   block2: { startCol: string, headers: string[], rows: string[][] },
 *   notes: string[],
 *   variationCount: number,
 * }}
 */
export function mapBriefToTracker(brief, options = {}) {
  const tracker = options.briefTypeOverride || brief.briefType;
  if (!tracker || !TRACKERS[tracker]) {
    throw new Error(`Unknown brief type: ${tracker}`);
  }
  const baseCfg = TRACKERS[tracker];
  // Apply per-client tracker overrides (e.g. Quintessa drops Winner Status).
  const cfg = applyTrackerOverride(baseCfg, options.clientTrackerOverrides?.[tracker]);

  const today = formatDate(new Date());
  const ctx = {
    today,
    overview: brief.overview || {},
    variations: brief.variations || [],
    requestDoc: options.requestDoc || '',
    clientFolderUrl: options.clientFolderUrl || '',
  };

  const buildBlock = (blockCfg) => {
    const headers = blockCfg.fields.map(f => `${f.col}: ${f.label}`);
    const rows = ctx.variations.map(variation => {
      return blockCfg.fields.map(field => resolveFieldValue(field, variation, ctx));
    });
    return { startCol: blockCfg.startCol, headers, rows };
  };

  const block1 = buildBlock(cfg.block1);
  const block2 = buildBlock(cfg.block2);

  // Collect unique notes from fields (only ones that are actually applicable)
  const noteSet = new Set();
  for (const f of [...cfg.block1.fields, ...cfg.block2.fields]) {
    if (f.note) noteSet.add(f.note);
  }
  const notes = [...noteSet];

  return {
    tracker,
    trackerLabel: cfg.label,
    block1,
    block2,
    notes,
    variationCount: ctx.variations.length,
  };
}

function resolveFieldValue(field, variation, ctx) {
  const raw = resolveSource(field.source, variation, ctx);
  if (field.transform && TRANSFORMS[field.transform]) {
    return TRANSFORMS[field.transform](raw);
  }
  return raw;
}

function resolveSource(source, variation, ctx) {
  if (!source || source === 'blank') return '';

  // Check fallback chain BEFORE prefix branches: 'overview:Photo Folder|fallback:clientFolderUrl'
  // tries the primary, then the fallback if the primary is empty.
  if (source.includes('|fallback:')) {
    const idx = source.indexOf('|fallback:');
    const primary = source.slice(0, idx);
    const fallbackPart = source.slice(idx + '|fallback:'.length);
    const primaryVal = resolveSource(primary, variation, ctx);
    if (primaryVal) return primaryVal;
    return resolveSource(fallbackPart, variation, ctx);
  }

  if (source === 'today') return ctx.today;
  if (source === 'requestDoc') return ctx.requestDoc;
  if (source === 'clientFolderUrl') return ctx.clientFolderUrl;

  if (source.startsWith('literal:')) {
    return source.slice('literal:'.length);
  }
  if (source.startsWith('overview:')) {
    const key = source.slice('overview:'.length);
    return ctx.overview[key] || '';
  }
  if (source.startsWith('variation:')) {
    const key = source.slice('variation:'.length);
    return variation[key] || '';
  }
  if (source === 'computed:cid') {
    return extractCid(variation['File Name'] || variation['Name'] || '');
  }
  if (source === 'computed:variationName') {
    const fn = variation['File Name'] || variation['Name'] || '';
    return extractVariationName(fn, ctx.overview['Angle Name']);
  }
  return '';
}

function formatDate(d) {
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

/**
 * Apply per-client overrides to a tracker config.
 *
 * Override schema (per tracker type):
 *   {
 *     block1: { drop?: ['Winner Status', ...], replaceFields?: [...] },
 *     block2: { drop?: [...], replaceFields?: [...] },
 *   }
 *
 * `drop` removes named fields from a block. After dropping, columns shift
 * left automatically (col letters get re-stamped A,B,C,...).
 *
 * `replaceFields` (rare) lets you provide a custom field array for that
 * block, used when the override is more than just deletions.
 */
function applyTrackerOverride(baseCfg, override) {
  if (!override) return baseCfg;
  const newCfg = { ...baseCfg };
  for (const blockName of ['block1', 'block2']) {
    const baseBlock = baseCfg[blockName];
    const overrideBlock = override[blockName];
    if (!overrideBlock) continue;
    let fields = baseBlock.fields;
    if (overrideBlock.replaceFields) {
      fields = overrideBlock.replaceFields;
    }
    if (overrideBlock.drop && overrideBlock.drop.length > 0) {
      const dropSet = new Set(overrideBlock.drop);
      fields = fields.filter(f => !dropSet.has(f.label));
    }
    // Re-stamp column letters starting from baseBlock.startCol
    const startCol = baseBlock.startCol;
    fields = fields.map((f, i) => ({ ...f, col: addColLetters(startCol, i) }));
    newCfg[blockName] = { ...baseBlock, fields };
  }
  return newCfg;
}

/**
 * Add an integer offset to a spreadsheet column letter.
 *   addColLetters('A', 0) -> 'A'
 *   addColLetters('A', 1) -> 'B'
 *   addColLetters('Z', 1) -> 'AA'
 *   addColLetters('W', 5) -> 'AB'
 */
function addColLetters(start, offset) {
  // Convert letters to 0-based index
  let idx = 0;
  for (const ch of start) {
    idx = idx * 26 + (ch.charCodeAt(0) - 'A'.charCodeAt(0) + 1);
  }
  idx = idx - 1 + offset;
  // Convert back
  let out = '';
  while (idx >= 0) {
    out = String.fromCharCode('A'.charCodeAt(0) + (idx % 26)) + out;
    idx = Math.floor(idx / 26) - 1;
  }
  return out;
}

/**
 * Serialize a block to tab-separated text suitable for clipboard paste into
 * Google Sheets. Each row becomes one line; cells are tab-separated. Newlines
 * inside cell values are converted to spaces (Sheets paste interprets newlines
 * as row breaks otherwise).
 */
export function blockToTSV(block) {
  return block.rows
    .map(row =>
      row
        .map(cell => String(cell ?? '').replace(/\r?\n+/g, ' ').replace(/\t/g, ' '))
        .join('\t')
    )
    .join('\n');
}
