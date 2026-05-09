// Map a parsed brief → tracker rows (block 1 + block 2) using TRACKERS config.

import { TRACKERS, TRANSFORMS, extractCid, extractVariationName } from './tracker-config.js';

/**
 * @param {{briefType: string, overview: object, variations: object[]}} brief
 * @param {{requestDoc?: string, briefTypeOverride?: string}} options
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
  const cfg = TRACKERS[tracker];

  const today = formatDate(new Date());
  const ctx = {
    today,
    overview: brief.overview || {},
    variations: brief.variations || [],
    requestDoc: options.requestDoc || '',
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
  if (source === 'today') return ctx.today;
  if (source === 'requestDoc') return ctx.requestDoc;

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
