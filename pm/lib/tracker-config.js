// Column mappings for the Static / Video / Body Copy creative trackers.
//
// Each tracker has TWO paste blocks separated by performance columns we don't
// touch. The PM pastes Block 1 starting at column A and Block 2 starting at
// the second-block letter (W for Static + Copy, U for Video).
//
// Each "field" entry:
//   { col: 'A', label: 'Created Date', source: 'today' | 'overview:Idea Name' | 'variation:Lead Type' | 'literal:Meta' | 'blank' | 'computed:cid', note?: '...' }
// `source` is interpreted by tracker-mapper.js. Notes get surfaced in the UI
// so the PM knows why a column was blanked or defaulted.

const STATUS_NOTE = "Status (col E) left blank — brief states like 'Ready For Internal' don't map to tracker dropdowns (Not Launched / Testing / Winner / Loser). Set after launch.";
const PLATFORM_NOTE = "Platform defaulted to 'Meta' regardless of brief value.";

export const TRACKERS = {
  static: {
    label: 'Static Creative Tracker',
    block1: {
      startCol: 'A',
      fields: [
        { col: 'A', label: 'Created Date',         source: 'today' },
        { col: 'B', label: 'Launch Date',          source: 'blank' },
        { col: 'C', label: 'Static Creative Name', source: 'blank', note: 'Static Creative Name (col C) left blank — tracker formula auto-generates it from File Name / Lead Type / CID.' },
        { col: 'D', label: 'Type',                 source: 'overview:Net New/Iteration' },
        { col: 'E', label: 'Status',               source: 'blank', note: STATUS_NOTE },
        { col: 'F', label: 'Winner Status',        source: 'blank' },
        { col: 'G', label: 'Platform',             source: 'literal:Meta', note: PLATFORM_NOTE },
        { col: 'H', label: 'Static Format',        source: 'literal:Single Static', note: "Static Format defaulted to 'Single Static'. Change manually for carousels." },
        { col: 'I', label: 'Awareness Level',      source: 'variation:Awareness Level' },
        { col: 'J', label: 'Lead Type',            source: 'variation:Lead Type', transform: 'leadTypeNormalize' },
        { col: 'K', label: 'Variation Type',       source: 'variation:Variation Type', transform: 'staticVariationType', note: "Variation Type passes through Copy / Visual; other values left blank." },
        { col: 'L', label: 'CID',                  source: 'computed:cid' },
        { col: 'M', label: 'Request Doc',          source: 'requestDoc', renderAs: 'url' },
        { col: 'N', label: 'Folder Link',          source: 'overview:Photo Folder|fallback:clientFolderUrl', renderAs: 'url' },
      ],
    },
    block2: {
      startCol: 'W',
      fields: [
        { col: 'W',  label: 'Idea Name',      source: 'overview:Idea Name' },
        { col: 'X',  label: 'Angle Name',     source: 'overview:Angle Name' },
        { col: 'Y',  label: 'Image Style',    source: 'overview:Style Name' },
        { col: 'Z',  label: 'Slide Count',    source: 'blank' },
        { col: 'AA', label: 'Copywriter',     source: 'overview:Copywriter' },
        { col: 'AB', label: 'Designer',       source: 'blank' },
        { col: 'AC', label: 'Variation Name', source: 'computed:variationName' },
      ],
    },
  },

  video: {
    label: 'Video Creative Tracker',
    block1: {
      startCol: 'A',
      fields: [
        { col: 'A', label: 'Created Date',        source: 'today' },
        { col: 'B', label: 'Launch Date',         source: 'blank' },
        { col: 'C', label: 'Video Creative Name', source: 'blank', note: 'Video Creative Name (col C) left blank — tracker formula auto-generates it from File Name / Lead Type / CID.' },
        { col: 'D', label: 'Type',                source: 'overview:Net New/Iteration' },
        { col: 'E', label: 'Status',              source: 'blank', note: STATUS_NOTE },
        { col: 'F', label: 'Winner Status',       source: 'blank' },
        { col: 'G', label: 'Platform',            source: 'literal:Meta', note: PLATFORM_NOTE },
        { col: 'H', label: 'Awareness Level',     source: 'variation:Awareness Level' },
        { col: 'I', label: 'Lead Type',           source: 'variation:Lead Type', transform: 'leadTypeNormalize' },
        { col: 'J', label: 'Variation Type',      source: 'variation:Variation Type', transform: 'videoVariationType', note: "Variation Type: brief 'Lead' → tracker 'Hook'." },
        { col: 'K', label: 'CID',                 source: 'computed:cid' },
        { col: 'L', label: 'Request Doc',         source: 'requestDoc', renderAs: 'url' },
        { col: 'M', label: 'Folder Link',         source: 'overview:Footage Folder|fallback:clientFolderUrl', renderAs: 'url' },
      ],
    },
    block2: {
      startCol: 'U',
      fields: [
        { col: 'U',  label: 'Idea Name',      source: 'overview:Idea Name' },
        { col: 'V',  label: 'Angle Name',     source: 'overview:Angle Name' },
        { col: 'W',  label: 'Style Name',     source: 'overview:Style Name' },
        { col: 'X',  label: 'Talent',         source: 'blank' },
        { col: 'Y',  label: 'Length (in Sec)', source: 'blank' },
        { col: 'Z',  label: 'Copywriter',     source: 'overview:Copywriter' },
        { col: 'AA', label: 'Designer',       source: 'blank' },
        { col: 'AB', label: 'Variation Name', source: 'computed:variationName' },
      ],
    },
  },

  copy: {
    label: 'Copy Tracker',
    block1: {
      startCol: 'A',
      fields: [
        { col: 'A', label: 'Created Date',  source: 'today' },
        { col: 'B', label: 'Launch Date',   source: 'blank' },
        { col: 'C', label: 'Copy Name',     source: 'blank', note: 'Copy Name (col C) left blank — tracker formula auto-generates from Angle + Lead Type + CPID.' },
        { col: 'D', label: 'Type',          source: 'overview:Net New/Iteration' },
        { col: 'E', label: 'Status',        source: 'blank', note: STATUS_NOTE },
        { col: 'F', label: 'Winner Status', source: 'blank' },
        { col: 'G', label: 'Platform',      source: 'literal:Meta', note: PLATFORM_NOTE },
        { col: 'H', label: 'Copy Type',     source: 'overview:Copy Type' },
        { col: 'I', label: 'Awareness Level', source: 'variation:Awareness Level' },
        { col: 'J', label: 'Lead Type',     source: 'variation:Lead Type', transform: 'leadTypeNormalize' },
        { col: 'K', label: 'Variation Type', source: 'variation:Variation Type' },
        { col: 'L', label: 'CPID',          source: 'blank', note: 'CPID (col L) left blank — tracker formula auto-generates.' },
        { col: 'M', label: 'Request Doc',   source: 'requestDoc', renderAs: 'url' },
        { col: 'N', label: 'Folder Link',   source: 'clientFolderUrl', renderAs: 'url' },
      ],
    },
    block2: {
      startCol: 'W',
      fields: [
        { col: 'W',  label: 'Idea Name',          source: 'overview:Idea Name' },
        { col: 'X',  label: 'Angle Name',         source: 'overview:Angle Name' },
        { col: 'Y',  label: 'Linked Creative ID', source: 'blank' },
        { col: 'Z',  label: 'Copywriter',         source: 'overview:Copywriter' },
        { col: 'AA', label: 'Designer',           source: 'blank' },
        { col: 'AB', label: 'Variation Name',     source: 'blank', note: 'Variation Name (col AB) left blank — tracker formula auto-generates it from Lead Type / Variation Type, same as Copy Name and CPID.' },
      ],
    },
  },
};

// -------- Transform helpers --------

const LEAD_TYPE_VALID = new Set(['Offer', 'Promise', 'Problem-Solution', 'Secret', 'Proclamation', 'Story']);

export function leadTypeNormalize(value) {
  if (!value) return '';
  let v = String(value).trim();
  // Strip trailing variation number ("Story Lead 3" → "Story Lead")
  v = v.replace(/\s+\d+$/, '');
  // Strip trailing " Lead" ("Story Lead" → "Story")
  v = v.replace(/\s+Lead$/i, '');
  // Case-correct against known dropdown values
  for (const valid of LEAD_TYPE_VALID) {
    if (v.toLowerCase() === valid.toLowerCase()) return valid;
  }
  return v;
}

// Static tracker dropdown accepts: Copy, Visual (the brief uses the same set,
// so this is mostly a passthrough with case-correction).
const STATIC_VARIATION_VALID = ['Copy', 'Visual'];
// Video tracker dropdown accepts these — brief uses Lead which we map to Hook
// (the original brief→tracker convention).
const VIDEO_VARIATION_REMAP = {
  'lead': 'Hook',
  'hook': 'Hook',
  'pattern interrupt': 'Pattern Interupt',
  'pattern interupt': 'Pattern Interupt',
  'body': 'Body',
  'cta': 'CTA',
};

export function staticVariationType(value) {
  if (!value) return '';
  const v = String(value).trim().toLowerCase();
  for (const valid of STATIC_VARIATION_VALID) {
    if (v === valid.toLowerCase()) return valid;
  }
  return ''; // unknown value → leave blank
}

export function videoVariationType(value) {
  if (!value) return '';
  const v = String(value).trim().toLowerCase();
  return VIDEO_VARIATION_REMAP[v] || '';
}

export const TRANSFORMS = {
  leadTypeNormalize,
  staticVariationType,
  videoVariationType,
};

// -------- Computed value helpers --------

/**
 * Extract the CID/CPID code from a File Name.
 * "Quiet Overwhelm - Offer Lead 1 - CID65D8FUL" → "65D8FUL"
 * "Million Follower Sales Flop Comparison - Problem-Solution - CPIDKPK" → "KPK"
 * Returns '' if no recognizable code found.
 */
export function extractCid(fileName) {
  if (!fileName) return '';
  const m = String(fileName).match(/\bCP?ID([A-Za-z0-9]+)\s*$/);
  return m ? m[1] : '';
}

/**
 * Extract the Variation Name (descriptor + number) from a File Name.
 * "Done For You - Problem-Solution Lead 1 - CIDYHXHG5C" → "Problem-Solution Lead 1"
 * "Quiet Overwhelm | SV8_Gmail" (no CID, has visual code) → "SV8_Gmail" or fallback to whole minus angle prefix
 * Returns '' if File Name is empty.
 */
export function extractVariationName(fileName, angleName) {
  if (!fileName) return '';
  let s = String(fileName).trim();
  // Strip trailing " - CID..." or " - CPID..."
  s = s.replace(/\s*-\s*CP?ID[A-Za-z0-9]+\s*$/, '');
  // Strip leading "{angleName} - "
  if (angleName) {
    const prefix = String(angleName).trim() + ' -';
    if (s.toLowerCase().startsWith(prefix.toLowerCase())) {
      s = s.slice(prefix.length).trim();
    }
  }
  return s;
}
