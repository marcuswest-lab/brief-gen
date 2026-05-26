// Strategist status classifier.
//
// Given a Creative Tracker workbook + a client key, produces an array of
// statuses (one per data row) ready to paste into the tracker's Status column.
//
// Hard rule (applies to ALL clients): a row with < $1,500 spend is "Testing".
// Above that, per-client KPI thresholds decide Winner / On the fence / Out of KPI.
//
// Fence rule: when not explicitly specified, fence = 10% looser than winner.
//   - "under $X" metric (CPA, CPQ, CPR):   fenceMax = winnerMax * 1.10
//   - "at least Y" metric (ROAS):          fenceMin = winnerMin * 0.90

export const STATUS = {
  WINNER:    'Winner',
  FENCE:     'On the fence',
  OUT:       'Out of KPI',
  TESTING:   'Testing',
};

export const DEFAULT_MIN_SPEND_TO_QUALIFY = 1500;
const KPI_OVERRIDES_KEY = 'pbg.strategist.kpiOverrides.v1';

// metric: which column to read.
//   'roas'  → higher is better, compare against `winnerMin` / `fenceMin`
//   'cpa', 'cpq', 'cpr' → lower is better, compare against `winnerMax` / `fenceMax`
export const DEFAULT_KPI_RULES = {
  'vam': {
    label: 'Value Added Moving',
    metric: 'roas',
    winnerMin: 1.8,
    // fence: 10% looser → 1.62 ROAS
    notes: 'Winner: ≥$1,500 spend AND ≥1.8x ROAS.',
  },
  'ceo-lawyer': {
    label: 'CEO Lawyer',
    metric: 'cpa',
    winnerMax: 2000,
    notes: 'Winner: ≥$1,500 spend AND <$2,000 CPA.',
  },
  'tcc': {
    label: 'TCC',
    metric: 'cpq',
    winnerMax: 500,
    notes: 'Winner: ≥$1,500 spend AND <$500 Cost per Qualified.',
  },
  'nurse-coach': {
    label: 'Nurse Coach',
    metric: 'cpa',
    winnerMax: 1500,
    notes: 'Winner: ≥$1,500 spend AND <$1,500 CPA.',
  },
  'case-source': {
    label: 'Case Source',
    metric: 'cpa',
    winnerMax: 1800,
    notes: 'Winner: ≥$1,500 spend AND <$1,800 CPA.',
  },
  'quintessa': {
    label: 'Quintessa',
    metric: 'cpa',
    winnerMax: 1200,
    notes: 'Winner: ≥$1,500 spend AND <$1,200 CPA.',
  },
  'trusy': {
    label: 'Trusy',
    metric: 'cpa',
    winnerMax: 300,
    notes: 'Winner: ≥$1,500 spend AND <$300 CPA.',
  },
  'dan-henry-mdw': {
    label: 'Dan Henry — MDW',
    metric: 'cpa',
    winnerMax: 200,
    fenceMaxExplicit: 277,
    notes: 'Winner: <$200 CPA. On the fence: <$277 CPA. (Both require ≥$1,500 spend.)',
  },
  'dan-henry-pb': {
    label: 'Dan Henry — PB',
    metric: 'cpr',
    winnerMax: 12,
    fenceMaxExplicit: 24,
    notes: 'Winner: <$12 Cost per Reg. On the fence: <$24 CPR. (Both require ≥$1,500 spend.)',
  },
};

/**
 * Resolve the effective fence threshold for a rule. Returns the matching
 * comparison value for either the "lower is better" or "higher is better"
 * metric direction.
 */
export function fenceThreshold(rule) {
  if (rule.metric === 'roas') {
    // Higher is better.
    if (rule.fenceMinExplicit != null) return rule.fenceMinExplicit;
    return rule.winnerMin * 0.9;
  }
  // Lower is better (cpa / cpq / cpr).
  if (rule.fenceMaxExplicit != null) return rule.fenceMaxExplicit;
  return rule.winnerMax * 1.1;
}

// -------- Live overrides (editable in the Rules tab) --------
//
// The Rules tab lets strategists tweak thresholds without touching code.
// Overrides are stored per-client in localStorage as a sparse object — only
// the fields the user explicitly changed are persisted, so future tweaks to
// defaults still flow through for untouched fields.

function loadOverrides() {
  try {
    const raw = localStorage.getItem(KPI_OVERRIDES_KEY);
    return raw ? JSON.parse(raw) : {};
  } catch (e) { return {}; }
}

function saveOverrides(overrides) {
  try { localStorage.setItem(KPI_OVERRIDES_KEY, JSON.stringify(overrides)); }
  catch (e) { /* ignore quota errors */ }
}

/**
 * The current effective KPI rules: defaults merged with any user overrides.
 * Recomputed on every access so edits show up immediately.
 */
export function getKpiRules() {
  const overrides = loadOverrides();
  const merged = {};
  for (const [key, rule] of Object.entries(DEFAULT_KPI_RULES)) {
    merged[key] = { ...rule, ...(overrides[key] || {}) };
  }
  return merged;
}

/**
 * Effective min-spend floor (applies to all clients).
 */
export function getMinSpend() {
  const overrides = loadOverrides();
  const v = overrides.__minSpend;
  return (typeof v === 'number' && v >= 0) ? v : DEFAULT_MIN_SPEND_TO_QUALIFY;
}

/**
 * Patch a single client's rule fields. Pass `null` as value to clear a
 * specific override (reverting that field to the default).
 */
export function setRuleOverride(clientKey, patch) {
  const overrides = loadOverrides();
  const current = overrides[clientKey] || {};
  for (const [k, v] of Object.entries(patch)) {
    if (v == null) delete current[k];
    else current[k] = v;
  }
  if (Object.keys(current).length === 0) delete overrides[clientKey];
  else overrides[clientKey] = current;
  saveOverrides(overrides);
}

/**
 * Update (or clear with null) the global min-spend qualifier.
 */
export function setMinSpendOverride(value) {
  const overrides = loadOverrides();
  if (value == null) delete overrides.__minSpend;
  else overrides.__minSpend = value;
  saveOverrides(overrides);
}

/**
 * Wipe ALL overrides (Reset to defaults button).
 */
export function resetAllOverrides() {
  try { localStorage.removeItem(KPI_OVERRIDES_KEY); } catch (e) {}
}

/**
 * Returns true if `clientKey` (or '__minSpend') has a user override.
 * Used to render a "modified" badge in the UI.
 */
export function hasOverride(clientKey, field) {
  const overrides = loadOverrides();
  if (clientKey === '__minSpend') return overrides.__minSpend != null;
  const block = overrides[clientKey];
  if (!block) return false;
  return field ? block[field] != null : Object.keys(block).length > 0;
}

// -------- Workbook reading --------

// Candidate header names for the columns we care about. Lower-case, no punctuation.
const HEADER_CANDIDATES = {
  name: [
    'static creative name', 'video creative name', 'creative name', 'copy name',
    'ad name', 'name',
  ],
  spend: [
    'spend', 'amount spent', 'amount spent (usd)', 'total spend', 'ad spend', 'cost',
  ],
  roas: [
    'roas', 'roas (7d)', 'purchase roas', 'purchase roas (7d)', '7d roas',
    'website purchase roas', 'website purchase roas (7d)',
  ],
  cpa: [
    'cpa', 'cost per lead', 'cost per acquisition', 'cost per result',
    'cost per purchase', 'cpl',
  ],
  cpq: [
    'cpq', 'cpql', 'cost per qualified', 'cost per qualified lead',
    'cost per qualified call', 'cost per ql',
  ],
  cpr: [
    'cpr', 'cost per reg', 'cost per registration', 'cost per signup',
    'cost per sign up',
  ],
  status: ['status'],
  idea: ['idea name'],
  launch: ['launch date'],
};

const SECTIONS = [
  { sheet: 'Static Creative Tracker', label: 'Statics' },
  { sheet: 'Video Creative Tracker',  label: 'Videos'  },
];

function normHeader(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ').replace(/[:.]+$/, '');
}

function getCell(sheet, r, c) {
  const ref = XLSX.utils.encode_cell({ r, c });
  return sheet[ref] || null;
}
function getCellValue(sheet, r, c) {
  const cell = getCell(sheet, r, c);
  return cell ? cell.v : null;
}

/**
 * Scan the first ~5 rows of the sheet to find the header row (the row that
 * actually contains the column names — some clients have a title row above).
 * Returns { headerRow, headers: { [normalizedName]: colIndex } }.
 */
function findHeaderRow(sheet) {
  const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
  if (!range) return { headerRow: null, headers: {} };

  const maxScan = Math.min(range.s.r + 5, range.e.r);
  let best = { score: 0, row: range.s.r, headers: {} };

  for (let r = range.s.r; r <= maxScan; r++) {
    const headers = {};
    let score = 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const v = getCellValue(sheet, r, c);
      if (typeof v !== 'string') continue;
      const norm = normHeader(v);
      if (!norm) continue;
      headers[norm] = c;
      // Score: each candidate field we recognize bumps the score.
      for (const list of Object.values(HEADER_CANDIDATES)) {
        if (list.includes(norm)) { score += 1; break; }
      }
    }
    if (score > best.score) best = { score, row: r, headers };
  }
  return { headerRow: best.row, headers: best.headers };
}

function pickColumn(headers, candidateList) {
  for (const cand of candidateList) {
    if (headers[cand] != null) return headers[cand];
  }
  return null;
}

function detectColumns(sheet) {
  const { headerRow, headers } = findHeaderRow(sheet);
  if (headerRow == null) return null;
  const range = XLSX.utils.decode_range(sheet['!ref']);
  return {
    headerRow,
    range,
    name:   pickColumn(headers, HEADER_CANDIDATES.name),
    spend:  pickColumn(headers, HEADER_CANDIDATES.spend),
    roas:   pickColumn(headers, HEADER_CANDIDATES.roas),
    cpa:    pickColumn(headers, HEADER_CANDIDATES.cpa),
    cpq:    pickColumn(headers, HEADER_CANDIDATES.cpq),
    cpr:    pickColumn(headers, HEADER_CANDIDATES.cpr),
    status: pickColumn(headers, HEADER_CANDIDATES.status),
    idea:   pickColumn(headers, HEADER_CANDIDATES.idea),
    launch: pickColumn(headers, HEADER_CANDIDATES.launch),
    headers,
  };
}

/**
 * Coerce a cell value to a number, stripping common formatting artifacts
 * like $, commas, parentheses (negatives), and trailing 'x' (e.g. "1.8x").
 * Returns null if the value isn't numeric.
 */
function toNum(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  let s = String(v).trim();
  if (!s) return null;
  const neg = /^\(.*\)$/.test(s);
  s = s.replace(/[(),$\s]/g, '').replace(/x$/i, '').replace(/%$/, '');
  if (s === '' || s === '-' || s.toLowerCase() === 'n/a' || s.toLowerCase() === 'nan') return null;
  const n = parseFloat(s);
  if (!isFinite(n)) return null;
  return neg ? -n : n;
}

// -------- Classification --------

/**
 * Classify a single row's metrics against the rule.
 * Returns one of the STATUS.* strings.
 */
export function classifyRow({ spend, roas, cpa, cpq, cpr }, rule, minSpend) {
  if (minSpend == null) minSpend = getMinSpend();

  // Zero / missing spend → blank (no opinion). Some-spend-but-below-floor → Testing.
  if (spend == null || spend <= 0) return '';
  if (spend < minSpend) return STATUS.TESTING;

  if (rule.metric === 'roas') {
    if (roas == null) return STATUS.TESTING; // spent above floor but no ROAS yet
    if (roas >= rule.winnerMin) return STATUS.WINNER;
    if (roas >= fenceThreshold(rule)) return STATUS.FENCE;
    return STATUS.OUT;
  }

  // Lower-is-better metrics.
  const value = rule.metric === 'cpa' ? cpa
              : rule.metric === 'cpq' ? cpq
              : rule.metric === 'cpr' ? cpr
              : null;
  if (value == null || value <= 0) return STATUS.TESTING;
  if (value < rule.winnerMax) return STATUS.WINNER;
  if (value < fenceThreshold(rule)) return STATUS.FENCE;
  return STATUS.OUT;
}

/**
 * Classify all rows in a single workbook (covering both Static + Video sheets
 * if present). Returns an object with per-sheet results.
 *
 * Each sheet result: {
 *   sheetName, label, rowCount,
 *   columns: { spend, roas, cpa, cpq, cpr, name, status },
 *   missingMetric: bool,
 *   rows: [{ rowIndex, name, spend, roas, cpa, cpq, cpr, oldStatus, newStatus }]
 * }
 */
export function classifyWorkbook(workbook, ruleKey) {
  const rules = getKpiRules();
  const rule = rules[ruleKey];
  if (!rule) throw new Error(`Unknown client rule: ${ruleKey}`);
  const minSpend = getMinSpend();

  const sheets = [];
  for (const { sheet, label } of SECTIONS) {
    if (!workbook.SheetNames.includes(sheet)) continue;
    const ws = workbook.Sheets[sheet];
    const cols = detectColumns(ws);
    if (!cols || cols.name == null) continue;

    const metricCol = cols[rule.metric];
    const sheetResult = {
      sheetName: sheet,
      label,
      headerRow: cols.headerRow,
      columns: {
        name: cols.name, spend: cols.spend, roas: cols.roas,
        cpa: cols.cpa, cpq: cols.cpq, cpr: cols.cpr, status: cols.status,
      },
      missingMetric: metricCol == null,
      missingSpend: cols.spend == null,
      rows: [],
    };

    for (let r = cols.range.s.r + (cols.headerRow - cols.range.s.r) + 1; r <= cols.range.e.r; r++) {
      const nameVal = getCellValue(ws, r, cols.name);
      if (typeof nameVal !== 'string' || !nameVal.trim()) continue;
      const spend = cols.spend != null ? toNum(getCellValue(ws, r, cols.spend)) : null;
      const roas  = cols.roas  != null ? toNum(getCellValue(ws, r, cols.roas))  : null;
      const cpa   = cols.cpa   != null ? toNum(getCellValue(ws, r, cols.cpa))   : null;
      const cpq   = cols.cpq   != null ? toNum(getCellValue(ws, r, cols.cpq))   : null;
      const cpr   = cols.cpr   != null ? toNum(getCellValue(ws, r, cols.cpr))   : null;
      const oldStatus = cols.status != null
        ? String(getCellValue(ws, r, cols.status) || '').trim()
        : '';

      const newStatus = sheetResult.missingMetric || sheetResult.missingSpend
        ? '' // can't classify — leave blank rather than guess
        : classifyRow({ spend, roas, cpa, cpq, cpr }, rule, minSpend);

      sheetResult.rows.push({
        rowIndex: r + 1, // 1-based for user-facing reference
        name: nameVal.trim(),
        spend, roas, cpa, cpq, cpr,
        oldStatus, newStatus,
      });
    }

    sheets.push(sheetResult);
  }

  return { ruleKey, rule, sheets };
}

/**
 * Build a paste-ready single-column block of statuses (one value per line,
 * aligned to the data rows starting just below the header row). Caller picks
 * which sheet's output to copy.
 */
export function buildStatusColumn(sheetResult) {
  return sheetResult.rows.map(r => r.newStatus).join('\n');
}

/**
 * Summary counts for the UI legend.
 */
export function summarize(sheetResult) {
  const counts = {
    [STATUS.WINNER]: 0,
    [STATUS.FENCE]: 0,
    [STATUS.OUT]: 0,
    [STATUS.TESTING]: 0,
    blank: 0,
  };
  for (const r of sheetResult.rows) {
    if (!r.newStatus) counts.blank++;
    else counts[r.newStatus] = (counts[r.newStatus] || 0) + 1;
  }
  return counts;
}

// -------- Client detection from filename --------

const FILENAME_PATTERNS = [
  { key: 'ceo-lawyer',     re: /ceo\s*lawyer/i },
  { key: 'nurse-coach',    re: /nurse\s*coach/i },
  { key: 'vam',            re: /value\s*added\s*moving|\bvam\b/i },
  { key: 'case-source',    re: /case\s*source/i },
  { key: 'quintessa',      re: /quintessa/i },
  { key: 'trusy',          re: /trusy/i },
  { key: 'tcc',            re: /\btcc\b/i },
  // Dan Henry is split into two funnels — fall through to user picker if
  // filename doesn't disambiguate.
  { key: 'dan-henry-mdw',  re: /dan\s*henry.*\bmdw\b|\bmdw\b.*dan\s*henry/i },
  { key: 'dan-henry-pb',   re: /dan\s*henry.*\bpb\b|\bpb\b.*dan\s*henry/i },
];

export function detectClientFromFilename(filename) {
  if (!filename) return null;
  for (const p of FILENAME_PATTERNS) {
    if (p.re.test(filename)) return p.key;
  }
  // Bare "Dan Henry" with no funnel hint → return null; UI will prompt.
  if (/dan\s*henry/i.test(filename)) return 'dan-henry-mdw'; // sane default
  return null;
}

// -------- Workbook loader --------

export async function readWorkbookFromFile(file) {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { type: 'array', cellHyperlinks: true, cellDates: true });
}
