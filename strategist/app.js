// Strategist Tools — single-page app.
//
// Workflow: drop one or more Creative Tracker .xlsx files →
//   1. Get a paste-ready Status column (one value per data row) based on
//      per-client KPI rules.
//   2. Optionally trigger / open the GitHub-Pages dashboard for the matched
//      client (and view the local command to regenerate it).

import {
  DEFAULT_KPI_RULES, DEFAULT_MIN_SPEND_TO_QUALIFY, STATUS,
  getKpiRules, getMinSpend, setRuleOverride, setMinSpendOverride,
  resetAllOverrides, hasOverride,
  classifyWorkbook, buildStatusColumn, summarize,
  detectClientFromFilename, detectFileRole, FILE_ROLE_LABEL,
  readWorkbookFromFile, fenceThreshold,
} from './lib/status-classifier.js';

const STORAGE_KEY = 'pbg.strategist.state.v1';

// Dashboard config (URLs pulled from auto-memory: dashboard_deployments.md).
//
// `inputs` lists the file roles required to regenerate this client's
// dashboard. Older clients use the three-way split (OLD tracking sheet +
// Raw Data + NEW 2.0 tracker). Newer clients (Case Source, Quintessa,
// Dan Henry) only need Raw Data + the NEW 2.0 tracker.
const DASHBOARDS = {
  'tcc':          { name: 'TCC',             url: 'https://marcuswest-lab.github.io/tcc-creative-dashboard/',        repo: 'marcuswest-lab/tcc-creative-dashboard',        regenCmd: 'python build_tcc_combos.py',                              inputs: ['old-tracker', 'raw-data', 'new-tracker'] },
  'ceo-lawyer':   { name: 'CEO Lawyer',      url: 'https://marcuswest-lab.github.io/ceo-lawyer-creative-dashboard/', repo: 'marcuswest-lab/ceo-lawyer-creative-dashboard', regenCmd: 'python build_ceolawyer_nursecoach_dashboards.py',         inputs: ['old-tracker', 'raw-data', 'new-tracker'] },
  'nurse-coach':  { name: 'Nurse Coach',                                                                                                                                   regenCmd: 'python build_ceolawyer_nursecoach_dashboards.py',         inputs: ['old-tracker', 'raw-data', 'new-tracker'] },
  'vam':          { name: 'VAM',                                                                                                                                           regenCmd: 'python build_vam_dashboard.py',                           inputs: ['old-tracker', 'raw-data', 'new-tracker'] },
  'case-source':  { name: 'Case Source',                                                                                                                                   regenCmd: 'python build_casesource_dashboard.py',                    inputs: ['raw-data', 'new-tracker'] },
  'dan-henry-mdw':{ name: 'Dan Henry — MDW',                                                                                                                               regenCmd: 'python build_danhenry_dashboard.py',                      inputs: ['raw-data', 'new-tracker'] },
  'dan-henry-pb': { name: 'Dan Henry — PB',                                                                                                                                regenCmd: 'python build_danhenry_dashboard.py',                      inputs: ['raw-data', 'new-tracker'] },
  'quintessa':    { name: 'Quintessa',                                                                                                                                                                                                          inputs: ['raw-data', 'new-tracker'] },
  'trusy':        { name: 'Trusy',                                                                                                                                                                                                              inputs: ['raw-data', 'new-tracker'] },
};

// Human-readable hint for each file role — shown in the dashboard checklist.
const FILE_ROLE_HINTS = {
  'old-tracker': '[Client] - Creative & Copy Tracking Sheet … .xlsx',
  'raw-data':    '[Client] Raw Data (N).xlsx',
  'new-tracker': '[Client X 2.0] - Creative Tracker (N).xlsx',
};

const state = {
  tab: 'status',                  // 'status' | 'rules'
  files: [],                      // [{ name, size, workbook, clientKey, result }]
  selectedFileIdx: 0,
  selectedSheetIdx: 0,
  error: '',
};

// -------- Persistence --------

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      tab: state.tab,
      selectedFileIdx: state.selectedFileIdx,
      selectedSheetIdx: state.selectedSheetIdx,
    }));
  } catch (e) { /* ignore */ }
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    Object.assign(state, JSON.parse(raw));
  } catch (e) { /* ignore */ }
}

// -------- DOM helpers --------

function el(tag, props = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (k === 'class') node.className = v;
    else if (k === 'html') node.innerHTML = v;
    else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2).toLowerCase(), v);
    else if (v === false || v == null) continue;
    else if (v === true) node.setAttribute(k, '');
    else node.setAttribute(k, v);
  }
  for (const child of children) {
    if (child == null || child === false) continue;
    node.appendChild(typeof child === 'string' ? document.createTextNode(child) : child);
  }
  return node;
}

async function copyToClipboard(text) {
  try { await navigator.clipboard.writeText(text); return true; }
  catch (e) { return false; }
}

function flashStatus(msg, kind = 'success') {
  const status = document.getElementById('strat-status');
  if (!status) return;
  status.className = kind;
  status.textContent = msg;
  setTimeout(() => {
    if (status.textContent === msg) { status.textContent = ''; status.className = ''; }
  }, 3500);
}

function fmtBytes(n) {
  if (n > 1024 * 1024) return (n / 1024 / 1024).toFixed(1) + ' MB';
  if (n > 1024) return (n / 1024).toFixed(1) + ' KB';
  return n + ' B';
}

function fmtNum(n, opts = {}) {
  if (n == null || !isFinite(n)) return '—';
  const { currency, decimals = 2 } = opts;
  const fixed = n.toFixed(decimals);
  return currency ? `$${fixed}` : fixed;
}

// -------- Tabs --------

function renderTabs() {
  const wrap = document.getElementById('strat-tabs');
  wrap.innerHTML = '';
  const tabs = [
    { id: 'status', label: '🎯 Status & Dashboards' },
    { id: 'rules',  label: '📋 KPI Rules' },
  ];
  for (const t of tabs) {
    wrap.appendChild(el('button', {
      type: 'button',
      class: 'tab-btn' + (state.tab === t.id ? ' active' : ''),
      onclick: () => { state.tab = t.id; saveState(); renderAll(); },
    }, t.label));
  }
}

// -------- Main tab: Status & Dashboards --------

function renderStatusTab() {
  const wrap = document.getElementById('strat-content');
  wrap.innerHTML = '';

  // Card 1 — Upload
  const upCard = el('div', { class: 'card' });
  upCard.appendChild(el('h2', {}, 'Drop creative trackers'));
  upCard.appendChild(el('p', { class: 'muted', style: 'margin-top:0;' },
    'Drop one or more Creative Tracker .xlsx files. Each file is matched to a client by filename and processed against that client\'s KPI rule.'));

  const fileInput = el('input', {
    type: 'file', multiple: true, accept: '.xlsx,.xlsm',
    style: 'display:none;',
    onchange: e => { handleFiles(e.target.files); e.target.value = ''; },
  });
  const dropzone = el('div', {
    class: 'strat-dropzone',
    onclick: () => fileInput.click(),
    ondragover: e => { e.preventDefault(); dropzone.classList.add('drag-active'); },
    ondragleave: () => dropzone.classList.remove('drag-active'),
    ondrop: e => {
      e.preventDefault();
      dropzone.classList.remove('drag-active');
      handleFiles(e.dataTransfer?.files);
    },
  },
    el('div', { style: 'font-weight:500;' }, 'Drag & drop tracker .xlsx files here'),
    el('div', { class: 'muted', style: 'font-size:0.85em;margin-top:4px;' }, 'or click to choose files'),
    el('div', { class: 'muted', style: 'font-size:0.75em;margin-top:8px;' }, 'Allowed: .xlsx, .xlsm — drop the 2.0 Creative Tracker, Raw Data, and (for older clients) the Old Tracking Sheet.'),
    fileInput,
  );
  upCard.appendChild(dropzone);

  if (state.files.length) {
    const list = el('ul', { class: 'strat-filelist' });
    state.files.forEach((f, i) => {
      // Role badge — color-coded so it's obvious which file is which
      const roleClass = f.role === 'new-tracker' ? 'winner'
                      : f.role === 'old-tracker' ? 'fence'
                      : f.role === 'raw-data'    ? 'testing'
                      : 'out-of-kpi';
      const roleBadge = el('span', { class: `strat-pill ${roleClass}`, style: 'font-size:11px;' },
        FILE_ROLE_LABEL[f.role] || 'Unrecognized');

      const detected = el('span', { class: 'muted', style: 'font-size:12px;' },
        f.clientKey ? `→ ${getKpiRules()[f.clientKey]?.label || f.clientKey}` : '→ client not detected — pick below');

      // Client picker only matters for files we'll classify (2.0 trackers).
      // OLD/Raw files still need a client tag for the dashboard checklist,
      // but they don't get a KPI run.
      const picker = el('select', {
        onchange: e => {
          state.files[i].clientKey = e.target.value || null;
          reclassify(i);
          renderAll();
        },
      });
      picker.appendChild(el('option', { value: '' }, '— pick client —'));
      for (const [k, rule] of Object.entries(getKpiRules())) {
        const opt = el('option', { value: k }, rule.label);
        if (f.clientKey === k) opt.setAttribute('selected', '');
        picker.appendChild(opt);
      }

      list.appendChild(el('li', {},
        el('span', { style: 'font-weight:500;' }, f.name),
        el('span', { class: 'muted', style: 'font-size:0.85em;' }, fmtBytes(f.size)),
        roleBadge,
        detected,
        picker,
        el('button', {
          type: 'button',
          class: 'btn-link',
          style: 'margin-left:auto;',
          onclick: () => {
            state.files.splice(i, 1);
            if (state.selectedFileIdx >= state.files.length) state.selectedFileIdx = 0;
            renderAll();
          },
        }, '✕'),
      ));
    });
    upCard.appendChild(list);

    upCard.appendChild(el('div', { class: 'pm-actions', style: 'margin-top:10px;' },
      el('button', {
        type: 'button',
        class: 'btn-secondary',
        onclick: () => {
          state.files = [];
          state.selectedFileIdx = 0;
          state.selectedSheetIdx = 0;
          state.error = '';
          renderAll();
        },
      }, 'Clear all'),
    ));
  }

  if (state.error) {
    upCard.appendChild(el('div', { class: 'status-msg error', style: 'margin-top:12px;' }, state.error));
  }

  wrap.appendChild(upCard);

  // No files yet → done.
  if (state.files.length === 0) return;

  // File selector (when multiple files dropped).
  const file = state.files[state.selectedFileIdx];
  if (!file) return;

  if (state.files.length > 1) {
    const switchCard = el('div', { class: 'card' });
    switchCard.appendChild(el('h3', { style: 'margin:0 0 8px 0;' }, 'Viewing'));
    const row = el('div', { class: 'strat-client-row' });
    const fSel = el('select', {
      onchange: e => {
        state.selectedFileIdx = parseInt(e.target.value, 10) || 0;
        state.selectedSheetIdx = 0;
        saveState();
        renderAll();
      },
    });
    state.files.forEach((f, i) => {
      const opt = el('option', { value: String(i) }, `${f.name}${f.clientKey ? ` (${getKpiRules()[f.clientKey]?.label})` : ''}`);
      if (i === state.selectedFileIdx) opt.setAttribute('selected', '');
      fSel.appendChild(opt);
    });
    row.appendChild(fSel);
    switchCard.appendChild(row);
    wrap.appendChild(switchCard);
  }

  // Render the dual output for the active file.
  if (!file.clientKey) {
    wrap.appendChild(el('div', { class: 'strat-warn' },
      el('strong', {}, 'Pick a client above'),
      ' so I know which KPI rule to apply (and which dashboard checklist to show) for ', file.name, '.'));
    return;
  }

  // 2.0 tracker → full dual output (status + dashboard).
  if (file.role === 'new-tracker') {
    if (!file.result) {
      wrap.appendChild(el('div', { class: 'status-msg error' }, 'Classification failed — see console.'));
      return;
    }
    renderDualOutput(wrap, file);
    return;
  }

  // OLD tracker / Raw data → no status column, just the dashboard checklist
  // (the file still counts toward the required-inputs check).
  const rule = getKpiRules()[file.clientKey];
  const card = el('div', { class: 'card', style: 'margin-top:12px;' });
  card.appendChild(el('h2', { style: 'margin:0 0 6px 0;' }, `${rule.label} — Dashboard Inputs`));
  card.appendChild(el('p', { class: 'muted', style: 'font-size:13px;margin-top:0;' },
    `${file.name} is a ${FILE_ROLE_LABEL[file.role]}. Status-column classification only runs on the 2.0 Creative Tracker — drop that file too to generate one. This file still counts toward the dashboard checklist below.`));
  card.appendChild(buildDashboardPanel(file.clientKey, rule.label));
  wrap.appendChild(card);
}

function renderDualOutput(wrap, file) {
  const result = file.result;
  const rule = result.rule;

  if (result.sheets.length === 0) {
    wrap.appendChild(el('div', { class: 'strat-warn' },
      el('strong', {}, 'No tracker sheets found.'),
      ' Expected a sheet named "Static Creative Tracker" or "Video Creative Tracker".'));
    return;
  }

  // Sheet picker
  const sheetCard = el('div', { class: 'card', style: 'margin-top:12px;' });
  sheetCard.appendChild(el('h2', { style: 'margin:0 0 8px 0;' }, `${rule.label} — Status Column`));
  sheetCard.appendChild(el('div', { class: 'strat-summary' }, rule.notes));

  if (result.sheets.length > 1) {
    const tabs = el('div', { class: 'pm-source-toggle', style: 'margin-bottom:10px;' });
    result.sheets.forEach((s, i) => {
      tabs.appendChild(el('button', {
        type: 'button',
        class: 'pm-source-tab' + (i === state.selectedSheetIdx ? ' active' : ''),
        onclick: () => { state.selectedSheetIdx = i; saveState(); renderAll(); },
      }, `${s.label} (${s.rows.length})`));
    });
    sheetCard.appendChild(tabs);
  }

  const sheet = result.sheets[state.selectedSheetIdx] || result.sheets[0];

  // Warnings about missing columns
  if (sheet.missingSpend) {
    sheetCard.appendChild(el('div', { class: 'strat-warn' },
      el('strong', {}, 'No "Spend" column found in this sheet — '),
      'every row will fall back to "Testing". Add a Spend column to the tracker (header names tried: Spend, Amount Spent, Total Spend, Ad Spend, Cost).'));
  }
  if (sheet.missingMetric) {
    const metricLabel = rule.metric === 'roas' ? 'ROAS'
                      : rule.metric === 'cpa'  ? 'CPA'
                      : rule.metric === 'cpq'  ? 'Cost per Qualified'
                      : 'Cost per Reg';
    sheetCard.appendChild(el('div', { class: 'strat-warn' },
      el('strong', {}, `No "${metricLabel}" column found — `),
      'statuses will be blank. Add the column to the tracker so I can classify each ad.'));
  }

  // Legend + counts
  const counts = summarize(sheet);
  sheetCard.appendChild(el('div', { class: 'strat-legend' },
    el('span', { class: 'strat-pill winner' },     `Winner: ${counts[STATUS.WINNER]}`),
    el('span', { class: 'strat-pill fence' },      `On the fence: ${counts[STATUS.FENCE]}`),
    el('span', { class: 'strat-pill out-of-kpi' }, `Out of KPI: ${counts[STATUS.OUT]}`),
    el('span', { class: 'strat-pill testing' },    `Testing: ${counts[STATUS.TESTING]}`),
    counts.blank ? el('span', { class: 'strat-pill testing' }, `Blank: ${counts.blank}`) : null,
  ));

  // The dual grid
  const grid = el('div', { class: 'strat-results-grid' });

  // ---- LEFT: paste-ready status column ----
  const statusBlock = el('div', { class: 'strat-dash-card' });
  statusBlock.appendChild(el('h3', {}, '📋 Status column'));
  statusBlock.appendChild(el('p', { class: 'muted', style: 'font-size:13px;margin-top:0;' },
    `One value per row, ${sheet.rows.length} rows total. Click the first data cell of the Status column (row ${sheet.headerRow + 2}) in the tracker, then paste.`));

  const colText = buildStatusColumn(sheet);
  const ta = el('textarea', {
    class: 'strat-status-output',
    readonly: true,
    onclick: e => e.target.select(),
  });
  ta.value = colText;
  statusBlock.appendChild(ta);

  statusBlock.appendChild(el('div', { class: 'pm-actions', style: 'margin-top:10px;gap:8px;' },
    el('button', {
      type: 'button',
      class: 'btn-primary',
      onclick: async () => {
        const ok = await copyToClipboard(colText);
        flashStatus(ok ? `Copied ${sheet.rows.length} statuses` : 'Copy failed', ok ? 'success' : 'error');
      },
    }, 'Copy column'),
    el('button', {
      type: 'button',
      class: 'btn-secondary',
      onclick: () => showDetailTable(sheet, rule),
    }, 'View row-by-row'),
  ));

  grid.appendChild(statusBlock);

  // ---- RIGHT: dashboard panel (with required-inputs checklist) ----
  grid.appendChild(buildDashboardPanel(result.ruleKey, rule.label));

  sheetCard.appendChild(grid);
  wrap.appendChild(sheetCard);
}

/**
 * Build the GitHub-dashboard panel. Lists the input files required for
 * this client's dashboard, ticks them off as the user drops them in, and
 * exposes the regen command + deployed-dashboard link. Safe to call even
 * when no 2.0 tracker has been classified yet — it just reads state.files
 * to see what's been dropped.
 */
function buildDashboardPanel(clientKey, clientLabel) {
  const dashBlock = el('div', { class: 'strat-dash-card' });
  dashBlock.appendChild(el('h3', {}, '📊 Dashboard regeneration'));

  const dash = DASHBOARDS[clientKey];
  if (!dash) {
    dashBlock.appendChild(el('p', { class: 'muted', style: 'font-size:13px;' },
      `No dashboard configured for ${clientLabel} yet.`));
    return dashBlock;
  }

  // Required-inputs checklist
  const required = dash.inputs || ['new-tracker'];
  dashBlock.appendChild(el('div', { class: 'muted', style: 'font-size:13px;margin-bottom:6px;' },
    `For ${dash.name}, feed in the following files:`));

  const detectedRoles = new Set(
    state.files
      .filter(f => f.clientKey === clientKey)
      .map(f => f.role)
  );

  const checklist = el('ul', { style: 'list-style:none;padding:0;margin:0 0 10px 0;' });
  let missingCount = 0;
  for (const role of required) {
    const present = detectedRoles.has(role);
    if (!present) missingCount++;
    checklist.appendChild(el('li', { style: 'display:flex;gap:10px;align-items:baseline;padding:4px 0;border-bottom:1px solid #f3f4f6;' },
      el('span', { style: `font-size:14px;width:18px;color:${present ? '#059669' : '#9ca3af'};` }, present ? '✓' : '○'),
      el('div', { style: 'flex:1;' },
        el('div', { style: 'font-size:13px;font-weight:500;' }, FILE_ROLE_LABEL[role]),
        el('div', { class: 'muted', style: 'font-size:11px;font-family:SF Mono,Menlo,monospace;' }, FILE_ROLE_HINTS[role] || ''),
      ),
    ));
  }
  dashBlock.appendChild(checklist);

  if (missingCount === 0) {
    dashBlock.appendChild(el('div', { class: 'strat-pill winner', style: 'font-size:12px;display:inline-block;margin-bottom:10px;' },
      '✓ All inputs present — ready to regenerate'));
  } else {
    dashBlock.appendChild(el('div', { class: 'muted', style: 'font-size:12px;margin-bottom:10px;' },
      `${missingCount} of ${required.length} input${required.length === 1 ? '' : 's'} still needed.`));
  }

  // Deployed dashboard link
  if (dash.url) {
    dashBlock.appendChild(el('div', { class: 'dash-row' },
      el('a', { href: dash.url, target: '_blank', rel: 'noopener' }, `Open ${dash.name} dashboard ↗`),
      el('span', { class: 'muted', style: 'font-size:12px;' }, 'live'),
    ));
  } else {
    dashBlock.appendChild(el('p', { class: 'muted', style: 'font-size:13px;margin:6px 0;' },
      `Dashboard for ${dash.name} not deployed yet.`));
  }
  if (dash.repo) {
    dashBlock.appendChild(el('div', { class: 'dash-row' },
      el('a', { href: `https://github.com/${dash.repo}`, target: '_blank', rel: 'noopener' }, `Repo: ${dash.repo} ↗`),
    ));
  }

  // Regen command
  if (dash.regenCmd) {
    dashBlock.appendChild(el('div', { style: 'margin-top:10px;' },
      el('div', { class: 'muted', style: 'font-size:13px;margin-bottom:4px;' },
        'Once all inputs are in your local repo, run:'),
      el('code', { class: 'dash-cmd' }, dash.regenCmd),
      el('button', {
        type: 'button',
        class: 'btn-link',
        style: 'margin-top:4px;font-size:12px;',
        onclick: async () => {
          const ok = await copyToClipboard(dash.regenCmd);
          flashStatus(ok ? 'Command copied' : 'Copy failed', ok ? 'success' : 'error');
        },
      }, 'Copy command'),
    ));
  }

  dashBlock.appendChild(el('p', { class: 'muted', style: 'font-size:12px;margin-top:12px;' },
    'In-browser regeneration + auto-push is coming. For now, drop all required files here for the checklist, then run the command above in your local repo to rebuild and push.'));

  return dashBlock;
}

// Detail modal — row-by-row table so the strategist can sanity-check
// classifications before pasting.
function showDetailTable(sheet, rule) {
  const overlay = el('div', {
    style: 'position:fixed;inset:0;background:rgba(15,23,42,0.55);z-index:1000;display:flex;align-items:center;justify-content:center;padding:24px;',
    onclick: e => { if (e.target === overlay) overlay.remove(); },
  });
  const modal = el('div', {
    style: 'background:#fff;max-width:980px;width:100%;max-height:85vh;overflow:auto;border-radius:10px;padding:18px 22px;',
  });
  modal.appendChild(el('div', { style: 'display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;' },
    el('h2', { style: 'margin:0;' }, `${rule.label} — ${sheet.label} (${sheet.rows.length} rows)`),
    el('button', { type: 'button', class: 'btn-secondary btn-small', onclick: () => overlay.remove() }, 'Close'),
  ));

  const metricLabel = rule.metric === 'roas' ? 'ROAS'
                    : rule.metric === 'cpa'  ? 'CPA'
                    : rule.metric === 'cpq'  ? 'CPQ'
                    : 'CPR';

  const table = el('table', {
    style: 'width:100%;border-collapse:collapse;font-size:13px;',
  });
  const thead = el('thead', {},
    el('tr', {},
      el('th', { style: thStyle() }, 'Row'),
      el('th', { style: thStyle() }, 'Name'),
      el('th', { style: thStyle('right') }, 'Spend'),
      el('th', { style: thStyle('right') }, metricLabel),
      el('th', { style: thStyle() }, 'Old status'),
      el('th', { style: thStyle() }, 'New status'),
    ),
  );
  table.appendChild(thead);

  const tbody = el('tbody');
  for (const r of sheet.rows) {
    const metricVal = rule.metric === 'roas' ? r.roas
                    : rule.metric === 'cpa'  ? r.cpa
                    : rule.metric === 'cpq'  ? r.cpq
                    : r.cpr;
    const pillCls = r.newStatus === STATUS.WINNER  ? 'winner'
                  : r.newStatus === STATUS.FENCE   ? 'fence'
                  : r.newStatus === STATUS.OUT     ? 'out-of-kpi'
                  : 'testing';
    tbody.appendChild(el('tr', {},
      el('td', { style: tdStyle() }, String(r.rowIndex)),
      el('td', { style: tdStyle() }, r.name),
      el('td', { style: tdStyle('right') }, fmtNum(r.spend, { currency: true, decimals: 0 })),
      el('td', { style: tdStyle('right') }, fmtNum(metricVal, { currency: rule.metric !== 'roas', decimals: 2 })),
      el('td', { style: tdStyle() }, r.oldStatus || '—'),
      el('td', { style: tdStyle() }, el('span', { class: `strat-pill ${pillCls}` }, r.newStatus || '—')),
    ));
  }
  table.appendChild(tbody);
  modal.appendChild(table);

  overlay.appendChild(modal);
  document.body.appendChild(overlay);
}

function thStyle(align = 'left') {
  return `text-align:${align};padding:6px 8px;border-bottom:1px solid #e5e7eb;font-weight:600;font-size:12px;color:#374151;background:#f9fafb;position:sticky;top:0;`;
}
function tdStyle(align = 'left') {
  return `text-align:${align};padding:6px 8px;border-bottom:1px solid #f3f4f6;`;
}

// -------- File handling --------

async function handleFiles(fileList) {
  if (!fileList || fileList.length === 0) return;
  const rejected = [];
  for (const file of fileList) {
    const ext = file.name.toLowerCase().split('.').pop();
    if (ext !== 'xlsx' && ext !== 'xlsm') {
      rejected.push(`${file.name} (not .xlsx)`);
      continue;
    }
    if (state.files.some(f => f.name === file.name && f.size === file.size)) continue;
    try {
      const role = detectFileRole(file.name);
      const clientKey = detectClientFromFilename(file.name);
      // Only parse the workbook if we'll actually classify it. Raw Data
      // workbooks can be large — skipping the parse keeps the page fast.
      let workbook = null;
      if (role === 'new-tracker') {
        workbook = await readWorkbookFromFile(file);
      }
      const entry = { name: file.name, size: file.size, role, workbook, clientKey, result: null };
      if (workbook && clientKey) {
        try { entry.result = classifyWorkbook(workbook, clientKey); }
        catch (e) { console.error(e); }
      }
      state.files.push(entry);
    } catch (e) {
      rejected.push(`${file.name} (parse error: ${e.message})`);
    }
  }
  state.error = rejected.length ? `Skipped: ${rejected.join(', ')}` : '';
  renderAll();
}

function reclassify(idx) {
  const f = state.files[idx];
  if (!f) return;
  if (!f.workbook || !f.clientKey) { f.result = null; return; }
  try { f.result = classifyWorkbook(f.workbook, f.clientKey); }
  catch (e) { console.error(e); f.result = null; }
  state.selectedSheetIdx = 0;
}

// -------- Rules tab --------

function renderRulesTab() {
  const wrap = document.getElementById('strat-content');
  wrap.innerHTML = '';

  const card = el('div', { class: 'card' });
  const head = el('div', { style: 'display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;' });
  head.appendChild(el('h2', { style: 'margin:0;' }, 'KPI Rules'));
  head.appendChild(el('button', {
    type: 'button',
    class: 'btn-secondary btn-small',
    onclick: () => {
      if (!confirm('Reset every KPI rule back to the built-in defaults? Your custom values will be lost.')) return;
      resetAllOverrides();
      // Re-classify all open files with the restored defaults.
      state.files.forEach((_, i) => reclassify(i));
      renderAll();
      flashStatus('Rules reset to defaults.');
    },
  }, 'Reset all to defaults'));
  card.appendChild(head);

  card.appendChild(el('p', { class: 'muted', style: 'margin-top:8px;' },
    'Edit any threshold below. Changes save automatically to your browser and re-classify any open trackers immediately. The “Spend qualifier” at the top is the floor that applies to every client — rows with positive spend below this are “Testing”; rows with zero/blank spend are left blank.'));

  // Global min-spend editor
  const minSpend = getMinSpend();
  const minBlock = el('div', { class: 'strat-rule-card' });
  const minTitleRow = el('div', { style: 'display:flex;justify-content:space-between;align-items:center;' });
  minTitleRow.appendChild(el('div', { class: 'rule-title' },
    'Spend qualifier (applies to all clients)',
    hasOverride('__minSpend') ? el('span', { class: 'strat-pill fence', style: 'margin-left:8px;font-size:11px;' }, 'modified') : null,
  ));
  if (hasOverride('__minSpend')) {
    minTitleRow.appendChild(el('button', {
      type: 'button', class: 'btn-link', style: 'font-size:12px;',
      onclick: () => {
        setMinSpendOverride(null);
        state.files.forEach((_, i) => reclassify(i));
        renderAll();
      },
    }, `Revert to $${DEFAULT_MIN_SPEND_TO_QUALIFY.toLocaleString()}`));
  }
  minBlock.appendChild(minTitleRow);
  minBlock.appendChild(numericField({
    label: 'Minimum spend to leave “Testing”',
    prefix: '$',
    value: minSpend,
    step: 50,
    onCommit: (v) => {
      setMinSpendOverride(v === DEFAULT_MIN_SPEND_TO_QUALIFY ? null : v);
      state.files.forEach((_, i) => reclassify(i));
      renderAll();
    },
  }));
  card.appendChild(minBlock);

  // Per-client editors
  const rules = getKpiRules();
  for (const [key, rule] of Object.entries(rules)) {
    card.appendChild(renderRuleEditor(key, rule));
  }

  wrap.appendChild(card);
}

/**
 * Editor card for a single client's rule. Renders the winner threshold,
 * an optional explicit fence threshold (with toggle to enable/disable), the
 * computed effective fence value, and a per-client "Revert" link when any
 * field has been overridden.
 */
function renderRuleEditor(key, rule) {
  const block = el('div', { class: 'strat-rule-card' });
  const titleRow = el('div', { style: 'display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap;' });

  const titleLeft = el('div', { class: 'rule-title' }, rule.label,
    hasOverride(key) ? el('span', { class: 'strat-pill fence', style: 'margin-left:8px;font-size:11px;' }, 'modified') : null,
  );
  titleRow.appendChild(titleLeft);

  if (hasOverride(key)) {
    titleRow.appendChild(el('button', {
      type: 'button', class: 'btn-link', style: 'font-size:12px;',
      onclick: () => {
        // Clear every overrideable field for this client.
        setRuleOverride(key, { winnerMin: null, winnerMax: null, fenceMinExplicit: null, fenceMaxExplicit: null });
        state.files.forEach((_, i) => reclassify(i));
        renderAll();
      },
    }, 'Revert to defaults'));
  }
  block.appendChild(titleRow);

  const metricLabel = rule.metric === 'roas' ? 'ROAS'
                    : rule.metric === 'cpa'  ? 'CPA'
                    : rule.metric === 'cpq'  ? 'Cost per Qualified'
                    : 'Cost per Reg';
  const isRoas = rule.metric === 'roas';
  const direction = isRoas ? '(higher is better)' : '(lower is better)';
  block.appendChild(el('div', { class: 'muted', style: 'font-size:13px;margin:4px 0 10px 0;' },
    'Metric: ', el('code', {}, metricLabel), ` ${direction}`));

  const row = el('div', { style: 'display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px;' });

  // Winner threshold
  if (isRoas) {
    row.appendChild(numericField({
      label: 'Winner threshold (ROAS ≥)',
      suffix: 'x',
      value: rule.winnerMin,
      step: 0.05,
      onCommit: (v) => {
        setRuleOverride(key, { winnerMin: v === DEFAULT_KPI_RULES[key].winnerMin ? null : v });
        state.files.forEach((_, i) => reclassify(i));
        renderAll();
      },
    }));
  } else {
    row.appendChild(numericField({
      label: `Winner threshold (${metricLabel} <)`,
      prefix: '$',
      value: rule.winnerMax,
      step: 10,
      onCommit: (v) => {
        setRuleOverride(key, { winnerMax: v === DEFAULT_KPI_RULES[key].winnerMax ? null : v });
        state.files.forEach((_, i) => reclassify(i));
        renderAll();
      },
    }));
  }

  // Fence threshold — show as computed by default; allow override
  const fenceVal = fenceThreshold(rule);
  const defaultRule = DEFAULT_KPI_RULES[key];
  const fenceField = isRoas ? 'fenceMinExplicit' : 'fenceMaxExplicit';
  const hasExplicitFence = rule[fenceField] != null;

  row.appendChild(numericField({
    label: `On-the-fence threshold (${isRoas ? 'ROAS ≥' : metricLabel + ' <'})`,
    prefix: isRoas ? '' : '$',
    suffix: isRoas ? 'x' : '',
    value: fenceVal,
    step: isRoas ? 0.05 : 10,
    hint: hasExplicitFence ? 'Custom value' : 'Auto (10% looser than winner)',
    onCommit: (v) => {
      // If user typed the auto-derived value back, clear the override.
      const defaultRuleCopy = { ...defaultRule };
      const autoDefault = fenceThreshold(defaultRuleCopy);
      const patch = {};
      patch[fenceField] = (Math.abs(v - autoDefault) < 0.001) ? null : v;
      setRuleOverride(key, patch);
      state.files.forEach((_, i) => reclassify(i));
      renderAll();
    },
  }));

  block.appendChild(row);
  return block;
}

/**
 * Reusable numeric input with prefix ($) / suffix (x), commit-on-blur or
 * commit-on-enter. Live re-classification fires only after commit so typing
 * a multi-digit number doesn't trigger N intermediate re-renders.
 */
function numericField({ label, prefix, suffix, value, step, hint, onCommit }) {
  const field = el('div', {});
  field.appendChild(el('label', { style: 'display:block;font-size:12px;color:#374151;margin-bottom:4px;' }, label));
  const wrap = el('div', { style: 'display:flex;align-items:center;gap:6px;' });
  if (prefix) wrap.appendChild(el('span', { class: 'muted', style: 'font-size:13px;' }, prefix));
  const input = el('input', {
    type: 'number',
    step: String(step ?? 1),
    style: 'flex:1;padding:6px 8px;border:1px solid #cfd6e4;border-radius:6px;font-size:14px;',
    onblur: commit,
    onkeydown: (e) => { if (e.key === 'Enter') { e.preventDefault(); input.blur(); } },
  });
  input.value = String(value ?? '');
  wrap.appendChild(input);
  if (suffix) wrap.appendChild(el('span', { class: 'muted', style: 'font-size:13px;' }, suffix));
  field.appendChild(wrap);
  if (hint) field.appendChild(el('div', { class: 'muted', style: 'font-size:11px;margin-top:3px;' }, hint));

  function commit() {
    const raw = input.value.trim();
    if (raw === '') return;
    const n = parseFloat(raw);
    if (!isFinite(n) || n < 0) {
      input.value = String(value ?? '');
      return;
    }
    if (n === value) return; // unchanged
    onCommit(n);
  }
  return field;
}

// -------- Render --------

function renderAll() {
  renderTabs();
  if (state.tab === 'rules') renderRulesTab();
  else renderStatusTab();
}

function init() {
  loadState();
  renderAll();
}

init();
