// Parse a brief that was Cmd+A → Cmd+C → Cmd+V'd from a Google Doc / Word .docx.
//
// The pasted text is roughly:
//   <header line>
//   OVERVIEW\tOVERVIEW
//   AI Allowed?\t...
//   Idea Name\t...
//   ...
//   1\t1                    <- variation marker
//   Name\t...
//   File Name\t...
//   ...
//   2\t2
//   ...
//
// Multi-line cell values (e.g. Body Copy with paragraphs) come through as
// extra lines before the next field name. We use a known field-name set to
// detect "this line starts a new field" vs "this line continues the previous
// field's value".

// Field names that can appear in any brief overview (used to detect new fields)
const KNOWN_OVERVIEW_FIELDS = new Set([
  'OVERVIEW',
  'AI Allowed?',
  'Video Type',
  'Photo Folder',
  'Footage Folder',
  'Reference',
  'Idea Name',
  'Angle Name',
  'Style Name',
  'Copy Type',
  'Task',
  'General Notes',
  'Design Notes',
  'Editing Notes',
  'Link to Brand Assets',
  'Any Other Relevant Assets',
  'Ratio Format(s)',
  'Ad Platform',
  'Avatar',
  'Brand Voice',
  'Brand Voice ', // template has a trailing space sometimes
  'Net New/Iteration',
  'Landing Page URL',
  'Conversion Objective',
  'Copywriter',
]);

const KNOWN_VARIATION_FIELDS = new Set([
  'Name',
  'File Name',
  'File',
  'Video File',
  'Notes',
  'Design Notes',
  'Editing Notes',
  'Variation Type',
  'Awareness Level',
  'Lead Type',
  'Status',
  'Copy',
  'Headline',
  'Body Copy',
  'Lead Script',
  'Body Script',
]);

const ALL_KNOWN_FIELDS = new Set([...KNOWN_OVERVIEW_FIELDS, ...KNOWN_VARIATION_FIELDS]);

/**
 * Detect brief type from pasted text content.
 * @returns {'static'|'video'|'copy'|null}
 */
export function detectBriefType(text) {
  const lower = text.toLowerCase();
  // Header-line clues first (most reliable)
  if (/body copy brief/i.test(text)) return 'copy';
  if (/video (ad )?brief/i.test(text) && !/static/i.test(text.split('\n')[0] || '')) return 'video';
  if (/static (ad )?brief/i.test(text)) return 'static';
  // Field-presence fallbacks
  const hasFootage = lower.includes('footage folder');
  const hasLeadScript = lower.includes('lead script');
  if (hasFootage || hasLeadScript) return 'video';
  const hasCopyType = lower.includes('copy type');
  const hasHeadlineAndBody = lower.includes('headline') && lower.includes('body copy');
  if (hasCopyType && hasHeadlineAndBody) return 'copy';
  const hasPhoto = lower.includes('photo folder');
  if (hasPhoto) return 'static';
  return null;
}

/**
 * Parse a single field/value line. Returns { field, value } if the line
 * matches a known field, else null. Tab-separated; the first cell is the
 * label, the rest is the value.
 */
function parseFieldLine(line) {
  const tabIdx = line.indexOf('\t');
  if (tabIdx < 0) {
    // Maybe the line is just a field name with no value (template placeholder)
    const name = line.trim();
    if (ALL_KNOWN_FIELDS.has(name)) return { field: name, value: '' };
    return null;
  }
  const field = line.slice(0, tabIdx).trim();
  const rest = line.slice(tabIdx + 1);
  if (!ALL_KNOWN_FIELDS.has(field)) return null;
  return { field, value: rest };
}

/** Match a variation header row like "1\t1" or "1" alone. */
function isVariationMarker(line) {
  return /^\s*(\d+)(\t\1)?\s*$/.test(line);
}

/**
 * Parse a block of lines into a {field: value} object. Handles multi-line
 * cell values by joining continuation lines until the next known field.
 */
function parseFieldBlock(lines) {
  const result = {};
  let currentField = null;
  let currentValue = [];

  const flush = () => {
    if (currentField !== null) {
      // Trim trailing whitespace, keep internal newlines
      const v = currentValue.join('\n').replace(/\s+$/, '');
      result[currentField] = v;
    }
  };

  for (const line of lines) {
    const fv = parseFieldLine(line);
    if (fv) {
      flush();
      currentField = fv.field;
      currentValue = fv.value !== '' ? [fv.value] : [];
    } else {
      // Continuation line for the current field
      if (currentField !== null) currentValue.push(line);
    }
  }
  flush();
  return result;
}

/**
 * Parse a full brief paste.
 * @param {string} text
 * @returns {{briefType: string|null, header: string, overview: object, variations: object[]}}
 */
export function parseBrief(text) {
  const briefType = detectBriefType(text);

  // Normalize line endings
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');

  // Header is everything before the first OVERVIEW marker
  let overviewStart = -1;
  for (let i = 0; i < lines.length; i++) {
    if (/^\s*OVERVIEW\s*(\tOVERVIEW)?\s*$/.test(lines[i])) {
      overviewStart = i;
      break;
    }
  }
  if (overviewStart < 0) {
    return { briefType, header: '', overview: {}, variations: [] };
  }
  const header = lines.slice(0, overviewStart).filter(l => l.trim()).join('\n');

  // Find variation markers
  const variationStarts = [];
  for (let i = overviewStart + 1; i < lines.length; i++) {
    if (isVariationMarker(lines[i])) variationStarts.push(i);
  }

  // Overview = lines after OVERVIEW until first variation
  const ovEnd = variationStarts.length > 0 ? variationStarts[0] : lines.length;
  const overview = parseFieldBlock(lines.slice(overviewStart + 1, ovEnd));
  stripBracketPlaceholders(overview);

  // Each variation = lines from its marker (exclusive) to the next marker
  const variations = [];
  for (let vi = 0; vi < variationStarts.length; vi++) {
    const start = variationStarts[vi] + 1;
    const end = vi + 1 < variationStarts.length ? variationStarts[vi + 1] : lines.length;
    const data = parseFieldBlock(lines.slice(start, end));
    stripBracketPlaceholders(data);
    variations.push(data);
  }

  return { briefType, header, overview, variations };
}

/**
 * Remove leftover template placeholders like "[the master picture folder for
 * this brief, if any]" or "[Any general notes]". Treat any field value that
 * starts with "[" and ends with "]" as a placeholder. Mutates in place.
 */
function stripBracketPlaceholders(obj) {
  for (const [k, v] of Object.entries(obj)) {
    if (typeof v === 'string' && v.trim().startsWith('[') && v.trim().endsWith(']')) {
      obj[k] = '';
    }
  }
}
