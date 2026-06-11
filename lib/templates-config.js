// Per-template row→field mappings + dropdown options.
// Row indices and field names mirror the BAD Marketing templates exactly.
// Dropdown option lists were extracted from the actual SDTs in the .docx files.

export const DROPDOWN_OPTIONS = {
  // Static + Video share these
  staticVideoAdPlatform: ['All', 'Meta', 'Google', 'N/A'],
  // Body copy uses different ad-platform options
  bodyCopyAdPlatform: ['Meta', 'YouTube', 'PMAX/Search/Display', 'All'],

  aiAllowed: ['Yes, AI Is Allowed', 'NO NO NO NO'],
  netNewIteration: ['Net New', 'Iteration'],
  awarenessLevel: ['Most Aware', 'Product Aware', 'Solution Aware', 'Problem Aware', 'Unaware'],
  leadType: ['Offer', 'Promise', 'Problem-Solution', 'Secret', 'Proclamation', 'Story'],
  status: ['Ready For Internal', 'Changes Required', 'Needs Client Approval', 'Approved'],

  // Static only
  staticVariationType: ['Copy', 'Visual'],

  // Video only
  videoType: ['Intro Video', 'VSL', 'Webinar', 'Ad Video', 'Other Video'],
  videoVariationType: ['Lead', 'Pattern Interrupt', 'Body', 'CTA'],

  // Body copy only
  copyType: ['Primary Text', 'Headline', 'Asset Set'],
  copyVariationType: ['Lead', 'Body', 'CTA', 'Full'],

  // Used by multiselect fields (Ratio Format(s))
  ratioFormats: ['1:1 (Square)', '4:5 (Vertical)', '9:16 (Story/Reel)', '16:9 (Landscape)'],
};

// Aliases mirror generate_doc.py — accept common Claude/user variants
export const VALUE_ALIASES = {
  'Lead Type': {
    'proof': 'Proclamation',
    'pain': 'Problem-Solution',
    'curiosity': 'Secret',
    'problem': 'Problem-Solution',
    'solution': 'Problem-Solution',
    'problem solution': 'Problem-Solution',
    'story': 'Story',
  },
  'Variation Type': {
    'hook': 'Copy',
    'headline': 'Copy',
    'visual style': 'Visual',
    'visual': 'Visual',
  },
  'Status': {
    'ready': 'Ready For Internal',
    'ready for review': 'Ready For Internal',
    'internal': 'Ready For Internal',
    'changes': 'Changes Required',
    'needs approval': 'Needs Client Approval',
    'client approval': 'Needs Client Approval',
  },
};

// Field definitions for each brief type.
// `row` is the row index in the overview table (table 0) or per-creative table (tables 1-5).
// `kind`: 'text' (plain text cell) or 'dropdown' (SDT cell).
// `options`: dropdown option key when kind='dropdown'.
// `multiline`: true → renders as <textarea>, false → <input>.
// `required`: marks the field as required for form submission.

export const TEMPLATES = {
  static: {
    label: 'Static',
    file: 'templates/static.docx',
    overview: [
      { row: 1,  field: 'AI Allowed?',         kind: 'dropdown', options: 'aiAllowed' },
      { row: 2,  field: 'Photo Folder',        kind: 'text' },
      { row: 3,  field: 'Reference',           kind: 'text', multiline: true },
      { row: 4,  field: 'Idea Name',           kind: 'text'},
      { row: 5,  field: 'Angle Name',          kind: 'text' },
      { row: 6,  field: 'Style Name',          kind: 'text' },
      { row: 7,  field: 'Task',                kind: 'text', multiline: true },
      { row: 8,  field: 'General Notes',       kind: 'text', multiline: true },
      { row: 9,  field: 'Design Notes',        kind: 'text', multiline: true },
      { row: 10, field: 'Link to Brand Assets',kind: 'text' },
      { row: 11, field: 'Ratio Format(s)',     kind: 'multiselect', options: 'ratioFormats', default: '1:1 (Square), 9:16 (Story/Reel)', force: true },
      { row: 12, field: 'Ad Platform',         kind: 'dropdown', options: 'staticVideoAdPlatform', default: 'Meta' },
      { row: 13, field: 'Avatar',              kind: 'text', multiline: true },
      { row: 14, field: 'Brand Voice',         kind: 'text', multiline: true },
      { row: 15, field: 'Net New/Iteration',   kind: 'dropdown', options: 'netNewIteration' },
      { row: 16, field: 'Landing Page URL',    kind: 'text' },
      { row: 17, field: 'Conversion Objective',kind: 'text' },
      { row: 18, field: 'Copywriter',          kind: 'text' },
    ],
    creative: [
      { row: 1, field: 'File Name',      kind: 'text' },
      { row: 2, field: 'File',           kind: 'text' },
      { row: 3, field: 'Notes',          kind: 'text', multiline: true },
      { row: 4, field: 'Design Notes',   kind: 'text', multiline: true },
      { row: 5, field: 'Variation Type', kind: 'dropdown', options: 'staticVariationType' },
      { row: 6, field: 'Awareness Level',kind: 'dropdown', options: 'awarenessLevel' },
      { row: 7, field: 'Lead Type',      kind: 'dropdown', options: 'leadType' },
      { row: 8, field: 'Status',         kind: 'dropdown', options: 'status' },
      { row: 9, field: 'Copy',           kind: 'text', multiline: true },
    ],
  },

  video: {
    label: 'Video',
    file: 'templates/video.docx',
    overview: [
      { row: 1,  field: 'Video Type',          kind: 'dropdown', options: 'videoType' },
      { row: 2,  field: 'AI Allowed?',         kind: 'dropdown', options: 'aiAllowed' },
      { row: 3,  field: 'Footage Folder',      kind: 'text' },
      { row: 4,  field: 'Idea Name',           kind: 'text'},
      { row: 5,  field: 'Angle Name',          kind: 'text' },
      { row: 6,  field: 'Style Name',          kind: 'text' },
      { row: 7,  field: 'Task',                kind: 'text', multiline: true },
      { row: 8,  field: 'General Notes',       kind: 'text', multiline: true },
      { row: 9,  field: 'Editing Notes',       kind: 'text', multiline: true },
      { row: 10, field: 'Link to Brand Assets',kind: 'text' },
      { row: 11, field: 'Any Other Relevant Assets', kind: 'text', multiline: true },
      { row: 12, field: 'Ratio Format(s)',     kind: 'multiselect', options: 'ratioFormats', default: '9:16 (Story/Reel)', force: true },
      { row: 13, field: 'Ad Platform',         kind: 'dropdown', options: 'staticVideoAdPlatform', default: 'Meta' },
      { row: 14, field: 'Avatar',              kind: 'text', multiline: true },
      { row: 15, field: 'Brand Voice',         kind: 'text', multiline: true },
      { row: 16, field: 'Net New/Iteration',   kind: 'dropdown', options: 'netNewIteration' },
      { row: 17, field: 'Landing Page URL',    kind: 'text' },
      { row: 18, field: 'Conversion Objective',kind: 'text' },
      { row: 19, field: 'Copywriter',          kind: 'text' },
    ],
    creative: [
      { row: 1,  field: 'File Name',      kind: 'text' },
      { row: 2,  field: 'Video File',     kind: 'text' },
      { row: 3,  field: 'Notes',          kind: 'text', multiline: true },
      { row: 4,  field: 'Editing Notes',  kind: 'text', multiline: true },
      { row: 5,  field: 'Variation Type', kind: 'dropdown', options: 'videoVariationType' },
      { row: 6,  field: 'Awareness Level',kind: 'dropdown', options: 'awarenessLevel' },
      { row: 7,  field: 'Lead Type',      kind: 'dropdown', options: 'leadType' },
      { row: 8,  field: 'Status',         kind: 'dropdown', options: 'status' },
      { row: 9,  field: 'Lead Script',    kind: 'text', multiline: true },
      { row: 10, field: 'Body Script',    kind: 'text', multiline: true },
    ],
  },

  copy: {
    label: 'Body Copy',
    file: 'templates/body-copy.docx',
    overview: [
      { row: 1,  field: 'AI Allowed?',         kind: 'dropdown', options: 'aiAllowed' },
      { row: 2,  field: 'Idea Name',           kind: 'text'},
      { row: 3,  field: 'Angle Name',          kind: 'text' },
      { row: 4,  field: 'Copy Type',           kind: 'dropdown', options: 'copyType' },
      { row: 5,  field: 'Task',                kind: 'text', multiline: true },
      { row: 6,  field: 'General Notes',       kind: 'text', multiline: true },
      { row: 7,  field: 'Ad Platform',         kind: 'dropdown', options: 'bodyCopyAdPlatform', default: 'Meta' },
      { row: 8,  field: 'Net New/Iteration',   kind: 'dropdown', options: 'netNewIteration' },
      { row: 9,  field: 'Landing Page URL',    kind: 'text' },
      { row: 10, field: 'Conversion Objective',kind: 'text' },
      { row: 11, field: 'Copywriter',          kind: 'text' },
    ],
    creative: [
      { row: 1, field: 'Name',           kind: 'text' },
      { row: 2, field: 'Notes',          kind: 'text', multiline: true },
      { row: 3, field: 'Variation Type', kind: 'dropdown', options: 'copyVariationType' },
      { row: 4, field: 'Awareness Level',kind: 'dropdown', options: 'awarenessLevel' },
      { row: 5, field: 'Lead Type',      kind: 'dropdown', options: 'leadType' },
      { row: 6, field: 'Status',         kind: 'dropdown', options: 'status' },
      { row: 7, field: 'Headline',       kind: 'text', multiline: true },
      { row: 8, field: 'Body Copy',      kind: 'text', multiline: true },
    ],
  },
};

// Resolve a value against a dropdown's allowed options, applying field-specific aliases.
// Returns the matched option (case-corrected) or the original value if no match.
export function normalizeDropdownValue(value, options, fieldName = '') {
  if (!value) return value;
  const trimmed = String(value).trim();

  // Direct case-insensitive match
  for (const opt of options) {
    if (trimmed.toLowerCase() === opt.toLowerCase()) return opt;
  }

  // Try aliases
  const aliases = VALUE_ALIASES[fieldName] || {};
  const aliased = aliases[trimmed.toLowerCase()];
  if (aliased) {
    for (const opt of options) {
      if (aliased.toLowerCase() === opt.toLowerCase()) return opt;
    }
  }

  // Partial match (substring either way)
  const lower = trimmed.toLowerCase();
  for (const opt of options) {
    if (lower.includes(opt.toLowerCase()) || opt.toLowerCase().includes(lower)) return opt;
  }

  return trimmed;
}
