// PM app — two-tab tool for moving data between briefs and the Creative Tracker.

import { parseBrief, detectBriefType } from './lib/brief-parser.js';
import { mapBriefToTracker, blockToTSV } from './lib/tracker-mapper.js';
import { TRACKERS } from './lib/tracker-config.js';
import { loadBriefs, getBrief, upsertBrief, sortBriefsRecent, briefDisplayLabel, buildBriefFilename } from '../lib/brief-store.js';
import { TEMPLATES } from '../lib/templates-config.js';
import { generateBrief } from '../lib/docx-filler.js';
import { parseAdFilename, groupAdFiles, pickRepresentativeFile } from './lib/filename-parser.js';
import { parseClaudeOutput } from '../lib/claude-output-parser.js';
import {
  getApiKey, setApiKey, getOpenAIKey, setOpenAIKey,
  getModel, setModel, getConcurrency, setConcurrency,
  testApiKey, analyzeBatch, estimateCost, isVideoFile,
  DEFAULT_MODEL, SONNET_MODEL, OPUS_MODEL,
} from './lib/ad-analyzer.js';
import {
  buildMeetingNotesHtml, suggestedMeetingFilename,
  buildWeeklyUpdatesHtml, suggestedWeeklyFilename,
  readWorkbookFromFile,
} from './lib/tracker-pipeline.js';

const STORAGE_KEY = 'pbg.pm.state.v1';
const CLIENTS_STORAGE_KEY = 'pbg.localClients.v1';
const TRACKER_URL_OVERRIDES_KEY = 'pbg.trackerUrlOverrides.v1';

const state = {
  tab: 'brief-to-tracker',  // | 'tracker-to-brief'
  // Brief → Tracker
  briefSource: 'saved',    // 'saved' | 'paste' — toggle in B→T tab
  selectedBriefId: null,   // when briefSource === 'saved'
  briefText: '',           // when briefSource === 'paste'
  requestDoc: '',
  briefTypeOverride: '',  // '' = auto-detect
  briefFilterClientId: '', // '' = all; otherwise narrows the saved-brief picker
  clientId: null,          // selected client (for tracker URL)
  clients: [],             // loaded from clients.json + localStorage
  result: null,            // mapBriefToTracker output
  // Tracker → Brief
  trackerNames: '',
  filledStates: {},        // { [index]: true } UI checkmarks
  trackerToBriefBriefId: null,  // selected brief to bake names into
  trackerToBriefFilterClientId: '', // separate filter for the regen tab

  // Ad Categorizer
  adCatMode: 'static',      // 'static' | 'video' — sub-tab inside Ad Categorizer
  adCatGroups: [],          // [{ key, files, parsed, ratios, status, result, error }]
  adCatRunning: false,
  adCatResult: null,        // { tracker, trackerLabel, block1, block2, notes, variationCount }
  // Brief context for video matching
  adCatBriefMode: 'none',   // 'none' | 'saved' | 'paste'
  adCatBriefId: null,       // when adCatBriefMode === 'saved'
  adCatBriefText: '',       // when adCatBriefMode === 'paste'
  // Optional URL overrides (populate Request Doc / Folder Link in tracker output)
  adCatRequestDoc: '',
  adCatFolderUrl: '',

  // Meeting Notes (Launched this week + Testing + Production)
  // Files & generated HTML are kept in-memory only (not persisted).
  meetingFiles: [],      // [{ name, size, workbook }]
  meetingHtml: '',
  meetingFilename: '',
  meetingError: '',

  // Weekly Creative Updates (Ready to Launch + In Production)
  weeklyFiles: [],
  weeklyHtml: '',
  weeklyFilename: '',
  weeklyError: '',
};

// -------- Persistence --------

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      tab: state.tab,
      briefSource: state.briefSource,
      selectedBriefId: state.selectedBriefId,
      briefText: state.briefText,
      requestDoc: state.requestDoc,
      briefTypeOverride: state.briefTypeOverride,
      briefFilterClientId: state.briefFilterClientId,
      clientId: state.clientId,
      trackerNames: state.trackerNames,
      filledStates: state.filledStates,
      trackerToBriefBriefId: state.trackerToBriefBriefId,
      trackerToBriefFilterClientId: state.trackerToBriefFilterClientId,
      adCatMode: state.adCatMode,
      adCatBriefMode: state.adCatBriefMode,
      adCatBriefId: state.adCatBriefId,
      adCatBriefText: state.adCatBriefText,
      adCatRequestDoc: state.adCatRequestDoc,
      adCatFolderUrl: state.adCatFolderUrl,
    }));
  } catch (e) {
    console.warn('Save failed:', e);
  }
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const p = JSON.parse(raw);
    Object.assign(state, p);
    if (!state.filledStates) state.filledStates = {};
  } catch (e) {
    console.warn('Load failed:', e);
  }
}

async function loadClients() {
  let builtin = [];
  try {
    const res = await fetch('../clients.json');
    if (res.ok) {
      const data = await res.json();
      builtin = data.clients || [];
    }
  } catch (e) {
    console.warn('clients.json load failed:', e);
  }
  let local = [];
  try {
    const raw = localStorage.getItem(CLIENTS_STORAGE_KEY);
    if (raw) local = JSON.parse(raw) || [];
  } catch (e) {
    console.warn('Local clients load failed:', e);
  }
  state.clients = [...builtin, ...local];

  // Apply tracker URL overrides (set in Briefgen)
  try {
    const raw = localStorage.getItem(TRACKER_URL_OVERRIDES_KEY);
    if (raw) {
      const overrides = JSON.parse(raw) || {};
      for (const c of state.clients) {
        if (overrides[c.id]) c.tracker_url = overrides[c.id];
      }
    }
  } catch (e) {}

  if (!state.clientId && state.clients.length > 0) {
    state.clientId = state.clients[0].id;
  }
}

function getCurrentClient() {
  return state.clients.find(c => c.id === state.clientId) || null;
}

/**
 * Extract the client name from a brief header line like:
 *   "Static Ad Brief for Value Added Moving Quiet Overwhelm | 05-08-2026"
 *   "Body Copy Brief for Dan Henry | Followers Aren't Buyers | 05-08-2026"
 *   "Video Brief for TCC Done For You | 05-08-2026"
 * The format is: "{Type} Brief for {CLIENT}{some separator + rest}"
 * We compare against the known client names list to identify the match.
 */
function detectClientFromBrief(text) {
  if (!text || state.clients.length === 0) return null;
  const firstLine = (text.split('\n').find(l => l.trim()) || '').trim();
  // The header looks like "{Something} Brief for {ClientName}{...}"
  const m = firstLine.match(/Brief for\s+(.+)$/i);
  if (!m) return null;
  const after = m[1];
  // Try matching each known client name as a prefix (case-insensitive)
  // Sort longest-first so "Value Added Moving" beats a hypothetical "Value"
  const sorted = [...state.clients].sort((a, b) => b.name.length - a.name.length);
  for (const c of sorted) {
    const cn = c.name.toLowerCase();
    if (after.toLowerCase().startsWith(cn)) return c;
  }
  return null;
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
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch (e) {
    console.warn('Clipboard write failed:', e);
    return false;
  }
}

function flashStatus(msg, kind = 'success') {
  const status = document.getElementById('pm-status');
  if (!status) return;
  status.className = kind;
  status.textContent = msg;
  setTimeout(() => {
    if (status.textContent === msg) {
      status.textContent = '';
      status.className = '';
    }
  }, 3500);
}

// -------- Tabs --------

function renderTabs() {
  const wrap = document.getElementById('pm-tabs');
  wrap.innerHTML = '';
  const tabs = [
    { id: 'brief-to-tracker', label: 'Brief → Tracker' },
    { id: 'tracker-to-brief', label: 'Tracker → Brief' },
    { id: 'ad-categorizer', label: '🤖 Ad Categorizer' },
    { id: 'meeting-notes', label: '📝 Meeting Notes' },
    { id: 'weekly-updates', label: '📅 Weekly Creative Updates' },
  ];
  for (const t of tabs) {
    wrap.appendChild(el('button', {
      type: 'button',
      class: 'tab-btn' + (state.tab === t.id ? ' active' : ''),
      onclick: () => {
        state.tab = t.id;
        saveState();
        renderAll();
      },
    }, t.label));
  }
}

// -------- Tab 1: Brief → Tracker --------

function renderBriefToTracker() {
  const wrap = document.getElementById('pm-content');
  wrap.innerHTML = '';

  const savedBriefs = sortBriefsRecent(loadBriefs());
  const hasSaved = savedBriefs.length > 0;

  // If user state says "saved" but no saved briefs exist, fall back to paste
  if (state.briefSource === 'saved' && !hasSaved) {
    state.briefSource = 'paste';
  }

  // Client selector card (always visible at top)
  wrap.appendChild(renderClientSelectorCard());

  const card = el('div', { class: 'card' });
  card.appendChild(el('h2', {}, 'Choose a brief'));

  // Source toggle
  const toggleRow = el('div', { class: 'pm-source-toggle' });
  toggleRow.appendChild(makeSourceTab('saved', 'Saved brief', !hasSaved));
  toggleRow.appendChild(makeSourceTab('paste', 'Paste manually', false));
  card.appendChild(toggleRow);

  if (state.briefSource === 'saved') {
    // Saved brief picker
    if (savedBriefs.length === 0) {
      card.appendChild(el('div', { class: 'pm-detect pm-detect-empty' },
        'No saved briefs yet. Generate a brief in Briefgen and it will appear here.'));
    } else {
      // Client filter
      card.appendChild(renderClientFilter('briefFilterClientId', savedBriefs, () => renderBriefToTracker()));

      // Filter the briefs by selected client
      const filtered = state.briefFilterClientId
        ? savedBriefs.filter(b => b.clientId === state.briefFilterClientId)
        : savedBriefs;

      const sel = el('select', {
        id: 'pm-saved-brief',
        class: 'pm-saved-select',
        onchange: (e) => {
          state.selectedBriefId = e.target.value || null;
          saveState();
        },
      });
      sel.appendChild(el('option', { value: '' }, filtered.length === 0
        ? '— no briefs for this client —'
        : '— pick a brief —'));
      for (const b of filtered) {
        const opt = el('option', { value: b.id }, briefDisplayLabel(b));
        if (b.id === state.selectedBriefId) opt.setAttribute('selected', '');
        sel.appendChild(opt);
      }
      if (filtered.length === 0) sel.setAttribute('disabled', '');
      card.appendChild(sel);
    }
  } else {
    // Paste-text mode (manual upload)
    if (hasSaved) {
      card.appendChild(el('div', { class: 'pm-warn-banner' },
        el('strong', {}, '⚠️ Heads up: '),
        'Google Docs strips dropdown values (Variation Type, Awareness Level, Lead Type, Status) when you copy-paste. ',
        'For Briefgen-generated briefs, switch to ',
        el('strong', {}, 'Saved brief'),
        ' instead — it reads the actual brief data and keeps every field.',
      ));
    }
    const detectedType = state.briefTypeOverride || (state.briefText ? detectBriefType(state.briefText) : null);
    const detectionLine = detectedType
      ? el('div', { class: 'pm-detect' }, `Detected brief type: `, el('strong', {}, TRACKERS[detectedType]?.label || detectedType))
      : el('div', { class: 'pm-detect pm-detect-empty' }, 'Paste a brief to auto-detect its type.');
    card.appendChild(detectionLine);

    const briefTa = el('textarea', {
      id: 'brief-text',
      placeholder: 'Open the brief in Google Docs (or Word), Cmd+A, Cmd+C, then paste here…',
      rows: 14,
      oninput: (e) => {
        state.briefText = e.target.value;
        saveState();
        const newDetected = state.briefTypeOverride || (state.briefText ? detectBriefType(state.briefText) : null);
        const newLine = newDetected
          ? el('div', { class: 'pm-detect' }, 'Detected brief type: ', el('strong', {}, TRACKERS[newDetected]?.label || newDetected))
          : el('div', { class: 'pm-detect pm-detect-empty' }, 'Paste a brief to auto-detect its type.');
        detectionLine.replaceWith(newLine);
      },
    });
    briefTa.value = state.briefText;
    card.appendChild(briefTa);
  }

  // Inputs row: Request Doc URL + (only for paste) brief type override
  const inputsRow = el('div', { class: 'pm-input-row' });

  const urlField = el('div', { class: 'pm-input-field' });
  urlField.appendChild(el('label', { for: 'request-doc' }, 'Request Doc URL', el('span', { class: 'pm-hint-inline' }, ' (goes into the Request Doc tracker column)')));
  const urlInput = el('input', {
    id: 'request-doc',
    type: 'text',
    placeholder: 'https://docs.google.com/document/d/...',
    oninput: (e) => { state.requestDoc = e.target.value; saveState(); },
  });
  urlInput.value = state.requestDoc;
  urlField.appendChild(urlInput);
  inputsRow.appendChild(urlField);

  if (state.briefSource === 'paste') {
    const overrideField = el('div', { class: 'pm-input-field pm-input-narrow' });
    overrideField.appendChild(el('label', { for: 'type-override' }, 'Brief type override'));
    const overrideSelect = el('select', {
      id: 'type-override',
      onchange: (e) => {
        state.briefTypeOverride = e.target.value;
        saveState();
        renderBriefToTracker();
      },
    },
      el('option', { value: '' }, 'Auto-detect'),
      el('option', { value: 'static' }, 'Static'),
      el('option', { value: 'video' }, 'Video'),
      el('option', { value: 'copy' }, 'Body Copy'),
    );
    overrideSelect.value = state.briefTypeOverride;
    overrideField.appendChild(overrideSelect);
    inputsRow.appendChild(overrideField);
  }

  card.appendChild(inputsRow);

  const actions = el('div', { class: 'pm-actions' },
    el('button', { type: 'button', class: 'btn-secondary', onclick: handleClearBrief }, 'Clear'),
    el('button', { type: 'button', class: 'btn-primary', onclick: handleGenerateRows }, 'Generate Tracker Rows'),
  );
  card.appendChild(actions);
  wrap.appendChild(card);

  if (state.result) {
    wrap.appendChild(renderResultPanel(state.result));
  }
}

/**
 * Top-level client selector card for Brief→Tracker. Sets the active client
 * for tracker overrides + folder URL. Surfaces the Open Tracker button
 * upfront so PMs can jump to the spreadsheet without generating first.
 */
function renderClientSelectorCard() {
  const card = el('div', { class: 'card pm-client-card' });
  card.appendChild(el('h2', {}, 'Client'));

  const row = el('div', { class: 'pm-client-row' });
  const sel = el('select', {
    id: 'pm-client-select',
    class: 'pm-client-select',
    onchange: (e) => {
      state.clientId = e.target.value || null;
      saveState();
      // Re-render whichever tab is active (renderAll dispatches by state.tab)
      renderAll();
    },
  });
  sel.appendChild(el('option', { value: '' }, '— auto-detect from brief —'));
  for (const c of state.clients) {
    const opt = el('option', { value: c.id }, c.name);
    if (c.id === state.clientId) opt.setAttribute('selected', '');
    sel.appendChild(opt);
  }
  row.appendChild(sel);

  // Resolve the effective client (explicit selection wins over auto-detect)
  const effectiveClient = getEffectiveClient();

  if (effectiveClient && effectiveClient.tracker_url) {
    row.appendChild(el('a', {
      href: effectiveClient.tracker_url,
      target: '_blank',
      rel: 'noopener',
      class: 'btn-primary btn-small tracker-open-btn',
      title: `Open ${effectiveClient.name}'s Creative Tracker`,
    }, `📊 Open ${effectiveClient.name} Tracker ↗`));
  } else if (effectiveClient && !effectiveClient.tracker_url) {
    row.appendChild(el('span', { class: 'pm-no-tracker' },
      `No tracker URL for ${effectiveClient.name}. Set one in Briefgen.`));
  }
  card.appendChild(row);

  // Show whether client was auto-detected or explicit
  if (!state.clientId && effectiveClient) {
    card.appendChild(el('div', { class: 'pm-detect' },
      `Auto-detected: `, el('strong', {}, effectiveClient.name),
      ` (override above if wrong)`));
  } else if (!state.clientId && !effectiveClient) {
    card.appendChild(el('div', { class: 'pm-detect pm-detect-empty' },
      'Pick a client above, or it will auto-detect when you pick/paste a brief.'));
  }

  return card;
}

/**
 * Get the active client for tracker generation. Priority:
 *   1. Explicit selection in the client picker (state.clientId)
 *   2. Saved brief's clientId (when 'saved' source)
 *   3. Detected from brief paste header (when 'paste' source)
 */
function getEffectiveClient() {
  if (state.clientId) {
    return state.clients.find(c => c.id === state.clientId) || null;
  }
  if (state.briefSource === 'saved' && state.selectedBriefId) {
    const stored = getBrief(state.selectedBriefId);
    if (stored) {
      return state.clients.find(c => c.id === stored.clientId)
        || state.clients.find(c => c.name === stored.clientName)
        || null;
    }
  }
  if (state.briefSource === 'paste' && state.briefText) {
    return detectClientFromBrief(state.briefText);
  }
  return null;
}

/**
 * Render a small client filter dropdown that narrows a brief picker to one
 * client. `stateKey` is the field name on `state` to read/write (e.g.
 * 'briefFilterClientId'). `briefs` is the source list (used to derive which
 * client options to offer + counts). `onChange` is called after the value
 * changes (typically the parent renderer).
 */
function renderClientFilter(stateKey, briefs, onChange) {
  const wrap = el('div', { class: 'pm-client-filter' });
  wrap.appendChild(el('label', { for: `filter-${stateKey}` }, 'Filter by client'));

  // Build client → count map from the brief list (so empty clients aren't shown)
  const counts = new Map();
  for (const b of briefs) {
    counts.set(b.clientId, (counts.get(b.clientId) || 0) + 1);
  }
  const clientsWithBriefs = state.clients
    .filter(c => counts.has(c.id))
    .sort((a, b) => a.name.localeCompare(b.name));

  const sel = el('select', {
    id: `filter-${stateKey}`,
    class: 'pm-filter-select',
    onchange: (e) => {
      state[stateKey] = e.target.value;
      saveState();
      if (onChange) onChange();
    },
  });
  sel.appendChild(el('option', { value: '' }, `All clients (${briefs.length})`));
  for (const c of clientsWithBriefs) {
    const opt = el('option', { value: c.id }, `${c.name} (${counts.get(c.id)})`);
    if (state[stateKey] === c.id) opt.setAttribute('selected', '');
    sel.appendChild(opt);
  }
  wrap.appendChild(sel);
  return wrap;
}

function makeSourceTab(value, label, disabled) {
  return el('button', {
    type: 'button',
    class: 'pm-source-tab' + (state.briefSource === value ? ' active' : '') + (disabled ? ' disabled' : ''),
    disabled: disabled || undefined,
    onclick: disabled ? null : () => {
      state.briefSource = value;
      state.result = null;
      saveState();
      renderBriefToTracker();
    },
  }, label);
}

function handleClearBrief() {
  state.briefText = '';
  state.requestDoc = '';
  state.briefTypeOverride = '';
  state.selectedBriefId = null;
  state.result = null;
  saveState();
  renderBriefToTracker();
}

function handleGenerateRows() {
  try {
    let parsedBrief = null;

    if (state.briefSource === 'saved') {
      if (!state.selectedBriefId) {
        flashStatus('Pick a saved brief first.', 'error');
        return;
      }
      const stored = getBrief(state.selectedBriefId);
      if (!stored) {
        flashStatus('Saved brief not found (was it deleted?).', 'error');
        return;
      }
      parsedBrief = {
        briefType: stored.briefType,
        overview: stored.overview,
        variations: stored.creatives,
      };
    } else {
      if (!state.briefText.trim()) {
        flashStatus('Paste a brief first.', 'error');
        return;
      }
      parsedBrief = parseBrief(state.briefText);
    }

    // Effective client: explicit selection wins; otherwise infer from brief
    const effectiveClient = getEffectiveClient();

    const result = mapBriefToTracker(parsedBrief, {
      requestDoc: state.requestDoc,
      briefTypeOverride: state.briefTypeOverride || undefined,
      clientTrackerOverrides: effectiveClient?.tracker_overrides,
      clientFolderUrl: effectiveClient?.creative_folder_url,
    });
    if (result.variationCount === 0) {
      flashStatus('No variations found. Check that the brief has filled creative tables.', 'error');
      return;
    }
    state.result = result;
    saveState();
    renderBriefToTracker();
  } catch (e) {
    console.error(e);
    flashStatus('Error: ' + e.message, 'error');
  }
}

function renderResultPanel(result) {
  const card = el('div', { class: 'card pm-result' });
  card.appendChild(el('h2', {}, `${result.trackerLabel} — ${result.variationCount} row${result.variationCount === 1 ? '' : 's'}`));

  const detectedClient = getEffectiveClient();
  if (detectedClient && detectedClient.tracker_url) {
    const banner = el('div', { class: 'pm-tracker-banner' });
    banner.appendChild(el('div', { class: 'pm-tracker-banner-text' },
      el('strong', {}, `Detected client: ${detectedClient.name}`),
      el('br'),
      `Open the tracker → ${result.trackerLabel} tab → click the next empty row → paste each block at the correct starting column.`,
    ));
    banner.appendChild(el('a', {
      href: detectedClient.tracker_url,
      target: '_blank',
      rel: 'noopener',
      class: 'btn-primary btn-small pm-tracker-btn',
    }, `📊 Open ${detectedClient.name} Tracker ↗`));
    card.appendChild(banner);
  } else if (detectedClient) {
    card.appendChild(el('p', { class: 'pm-banner' },
      `Detected client: `, el('strong', {}, detectedClient.name),
      `. No tracker URL on file — add one in Briefgen (Edit URL button next to the client picker). Then paste each block at the correct starting column.`,
    ));
  } else {
    card.appendChild(el('p', { class: 'pm-banner' },
      `Open your client's tracker, go to the `, el('strong', {}, result.trackerLabel),
      ` tab, click the next empty row, then paste each block at the correct starting column.`,
    ));
  }

  // Row → filename map (Ad Categorizer only — set on the result object)
  if (result.adCatFilenames && result.adCatFilenames.length > 0) {
    const mapPanel = el('div', { class: 'pm-filename-map' });
    mapPanel.appendChild(el('h3', {}, 'Row → Filename'));
    mapPanel.appendChild(el('p', { class: 'pm-help', style: 'margin: 0 0 8px; font-size: 12px' },
      'Below shows which filename produced each output row, in paste order. The filenames are NOT pasted into the tracker — this is just for reference.'
    ));
    const list = el('ol', { class: 'pm-filename-list' });
    for (const name of result.adCatFilenames) {
      list.appendChild(el('li', { title: name }, name));
    }
    mapPanel.appendChild(list);
    card.appendChild(mapPanel);
  }

  // Block 1
  card.appendChild(renderBlockPanel('Block 1', result.block1, `Paste at column ${result.block1.startCol}`));
  // Block 2
  card.appendChild(renderBlockPanel('Block 2', result.block2, `Paste at column ${result.block2.startCol}`));

  // Notes
  if (result.notes && result.notes.length > 0) {
    const notesEl = el('div', { class: 'pm-notes' });
    notesEl.appendChild(el('h3', {}, 'Notes'));
    const list = el('ul', {});
    for (const n of result.notes) list.appendChild(el('li', {}, n));
    notesEl.appendChild(list);
    card.appendChild(notesEl);
  }

  return card;
}

function renderBlockPanel(title, block, subtitle) {
  const wrap = el('div', { class: 'pm-block' });

  const header = el('div', { class: 'pm-block-header' });
  header.appendChild(el('div', {},
    el('h3', {}, title),
    el('span', { class: 'pm-block-sub' }, subtitle),
  ));
  const copyBtn = el('button', {
    type: 'button',
    class: 'btn-primary btn-small',
    onclick: async () => {
      const tsv = blockToTSV(block);
      const ok = await copyToClipboard(tsv);
      if (ok) {
        copyBtn.textContent = '✓ Copied';
        flashStatus(`${title} copied — paste at column ${block.startCol}.`);
        setTimeout(() => { copyBtn.textContent = '📋 Copy'; }, 1800);
      } else {
        flashStatus('Copy failed — your browser may have blocked clipboard access.', 'error');
      }
    },
  }, '📋 Copy');
  header.appendChild(copyBtn);
  wrap.appendChild(header);

  // Render preview table
  const tbl = el('table', { class: 'pm-preview' });
  const thead = el('thead');
  const headerRow = el('tr');
  for (const h of block.headers) headerRow.appendChild(el('th', {}, h));
  thead.appendChild(headerRow);
  tbl.appendChild(thead);
  const tbody = el('tbody');
  for (const row of block.rows) {
    const tr = el('tr');
    row.forEach((cell, colIdx) => {
      const hint = block.renderHints?.[colIdx];
      const isEmpty = cell == null || cell === '';
      const td = el('td', { title: cell || '' });
      if (isEmpty) {
        td.textContent = '—';
        td.className = 'pm-empty-cell';
      } else if (hint?.renderAs === 'url' && /^https?:\/\//i.test(String(cell))) {
        // Render URL as a compact link pill (text shows hostname + first segment)
        const url = String(cell);
        let label = url;
        try {
          const u = new URL(url);
          // E.g. "drive.google.com/drive/u/1/folders/..." → "drive.google.com/folders/…"
          const path = u.pathname.replace(/\/+$/, '');
          const seg = path.split('/').filter(Boolean)[0] || '';
          label = `${u.hostname}${seg ? '/' + seg : ''}`;
          if (label.length > 32) label = label.slice(0, 31) + '…';
        } catch {
          label = url.length > 32 ? url.slice(0, 31) + '…' : url;
        }
        const a = el('a', {
          href: url,
          target: '_blank',
          rel: 'noopener',
          class: 'pm-url-pill',
        }, '🔗 ' + label);
        td.appendChild(a);
      } else {
        td.textContent = String(cell);
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }
  tbl.appendChild(tbody);
  const tableScroll = el('div', { class: 'pm-table-scroll' }, tbl);
  wrap.appendChild(tableScroll);

  return wrap;
}

// -------- Tab 2: Tracker → Brief --------

function renderTrackerToBrief() {
  const wrap = document.getElementById('pm-content');
  wrap.innerHTML = '';

  const card = el('div', { class: 'card' });
  card.appendChild(el('h2', {}, 'Paste auto-generated names from the tracker'));
  card.appendChild(el('p', { class: 'pm-help' },
    'After Block 1 is pasted, the tracker formula auto-fills the Creative Name column (Static Creative Name, Video Creative Name, or Copy Name). Select that column for the rows you just added, copy it, then paste here. One name per line.',
  ));

  const ta = el('textarea', {
    id: 'tracker-names',
    placeholder: `Quiet Overwhelm - Offer Lead 1 - CID65D8FUL\nQuiet Overwhelm - Promise Lead 2 - CIDABC1234\n...`,
    rows: 8,
    oninput: (e) => {
      state.trackerNames = e.target.value;
      saveState();
      renderTrackerOutput();
    },
  });
  ta.value = state.trackerNames;
  card.appendChild(ta);

  card.appendChild(el('div', { class: 'pm-actions' },
    el('button', { type: 'button', class: 'btn-secondary', onclick: () => {
      state.trackerNames = '';
      state.filledStates = {};
      saveState();
      renderTrackerToBrief();
    }}, 'Clear'),
  ));
  wrap.appendChild(card);

  // Regen card: pick a saved brief, bake names in, download fresh .docx
  const savedBriefs = sortBriefsRecent(loadBriefs());
  if (savedBriefs.length > 0) {
    const regenCard = el('div', { class: 'card' });
    regenCard.appendChild(el('h2', {}, 'Regenerate brief with these names'));
    regenCard.appendChild(el('p', { class: 'pm-help' },
      'Pick the saved brief these names belong to. Clicking Regenerate fills File Name / Name for each creative in order, updates the saved brief, and downloads a fresh .docx — ready to upload back into Google Docs.',
    ));

    regenCard.appendChild(renderClientFilter('trackerToBriefFilterClientId', savedBriefs, () => renderTrackerToBrief()));

    const filteredBriefs = state.trackerToBriefFilterClientId
      ? savedBriefs.filter(b => b.clientId === state.trackerToBriefFilterClientId)
      : savedBriefs;

    const sel = el('select', {
      id: 'pm-regen-brief',
      class: 'pm-saved-select',
      onchange: (e) => {
        state.trackerToBriefBriefId = e.target.value || null;
        saveState();
      },
    });
    sel.appendChild(el('option', { value: '' }, filteredBriefs.length === 0
      ? '— no briefs for this client —'
      : '— pick the saved brief —'));
    for (const b of filteredBriefs) {
      const opt = el('option', { value: b.id },
        `${briefDisplayLabel(b)} — ${b.creatives.length} creative${b.creatives.length === 1 ? '' : 's'}`);
      if (b.id === state.trackerToBriefBriefId) opt.setAttribute('selected', '');
      sel.appendChild(opt);
    }
    if (filteredBriefs.length === 0) sel.setAttribute('disabled', '');
    regenCard.appendChild(sel);

    regenCard.appendChild(el('div', { class: 'pm-actions' },
      el('button', { type: 'button', class: 'btn-primary', onclick: handleRegenWithNames },
        'Regenerate brief .docx with names'),
    ));
    wrap.appendChild(regenCard);
  }

  // Output panel
  const outputCard = el('div', { id: 'tracker-output', class: 'card' });
  wrap.appendChild(outputCard);
  renderTrackerOutput();
}

async function handleRegenWithNames() {
  const names = (state.trackerNames || '')
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean);

  if (names.length === 0) {
    flashStatus('Paste tracker-generated names first.', 'error');
    return;
  }
  if (!state.trackerToBriefBriefId) {
    flashStatus('Pick a saved brief to bake the names into.', 'error');
    return;
  }
  const brief = getBrief(state.trackerToBriefBriefId);
  if (!brief) {
    flashStatus('Saved brief not found.', 'error');
    return;
  }

  const config = TEMPLATES[brief.briefType];
  if (!config) {
    flashStatus(`Unknown brief type: ${brief.briefType}`, 'error');
    return;
  }

  if (names.length > brief.creatives.length) {
    if (!confirm(`You have ${names.length} names but the brief has ${brief.creatives.length} creative${brief.creatives.length === 1 ? '' : 's'}. Extra names will be ignored. Continue?`)) {
      return;
    }
  }

  // For body copy the field is "Name"; for static/video it's "File Name"
  const nameField = brief.briefType === 'copy' ? 'Name' : 'File Name';

  // Merge names into creatives (deep clone to avoid mutating storage)
  const updatedCreatives = brief.creatives.map((c, i) => ({
    ...c,
    [nameField]: i < names.length ? names[i] : (c[nameField] || ''),
  }));

  // Persist the updated brief
  upsertBrief({
    ...brief,
    creatives: updatedCreatives,
  });

  // Generate the .docx
  try {
    const res = await fetch(`../${config.file}`);
    if (!res.ok) throw new Error(`Failed to load template: ${res.status}`);
    const templateBuffer = await res.arrayBuffer();

    const blob = await generateBrief({
      briefType: brief.briefType,
      clientName: brief.clientName,
      overview: brief.overview,
      creatives: updatedCreatives,
      templateBuffer,
    });

    const filename = buildBriefFilename({
      clientName: brief.clientName,
      briefType: brief.briefType,
      ideaName: brief.ideaName,
      suffix: 'with names',
    });

    saveAs(blob, filename);
    flashStatus(`✓ Regenerated ${filename} (${Math.min(names.length, brief.creatives.length)} names baked in).`);
  } catch (e) {
    console.error(e);
    flashStatus('Error: ' + e.message, 'error');
  }
}

function renderTrackerOutput() {
  const card = document.getElementById('tracker-output');
  if (!card) return;
  card.innerHTML = '';

  const names = (state.trackerNames || '')
    .split('\n')
    .map(l => l.trim())
    .filter(Boolean);

  if (names.length === 0) {
    card.appendChild(el('div', { class: 'pm-detect pm-detect-empty' },
      'Paste names above to see one-click copy buttons here.'));
    return;
  }

  const header = el('div', { class: 'pm-block-header' },
    el('h3', {}, `${names.length} name${names.length === 1 ? '' : 's'}`),
    el('button', { type: 'button', class: 'btn-secondary btn-small', onclick: async () => {
      const ok = await copyToClipboard(names.join('\n'));
      if (ok) flashStatus('All names copied (newline-separated).');
    }}, '📋 Copy all'),
  );
  card.appendChild(header);

  const list = el('div', { class: 'pm-name-list' });
  names.forEach((name, idx) => {
    const item = el('div', { class: 'pm-name-item' + (state.filledStates[idx] ? ' filled' : '') });

    const cb = el('input', {
      type: 'checkbox',
      class: 'pm-name-check',
      title: 'Mark as pasted into brief',
      onchange: (e) => {
        state.filledStates[idx] = e.target.checked;
        saveState();
        item.classList.toggle('filled', e.target.checked);
      },
    });
    if (state.filledStates[idx]) cb.setAttribute('checked', '');
    item.appendChild(cb);

    item.appendChild(el('span', { class: 'pm-name-num' }, `Creative ${idx + 1}:`));
    item.appendChild(el('span', { class: 'pm-name-text' }, name));

    const btn = el('button', {
      type: 'button',
      class: 'btn-secondary btn-small',
      onclick: async () => {
        const ok = await copyToClipboard(name);
        if (ok) {
          btn.textContent = '✓';
          state.filledStates[idx] = true;
          saveState();
          item.classList.add('filled');
          item.querySelector('.pm-name-check').checked = true;
          setTimeout(() => { btn.textContent = '📋 Copy'; }, 1500);
        }
      },
    }, '📋 Copy');
    item.appendChild(btn);

    list.appendChild(item);
  });
  card.appendChild(list);
}

// -------- Ad Categorizer --------

// Per-client clue: which clients usually supply Idea/Angle/Style themselves.
// We still let the AI fill them but flag them as "client typically supplies".
const CLIENTS_THAT_SUPPLY_NAMES = new Set(['tcc']);

function renderAdCategorizer() {
  const wrap = document.getElementById('pm-content');
  wrap.innerHTML = '';

  // Top-of-tab client selector (reused)
  wrap.appendChild(renderClientSelectorCard());

  const apiKey = getApiKey();

  // API key gate
  if (!apiKey) {
    const gate = el('div', { class: 'card' });
    gate.appendChild(el('h2', {}, '🤖 Ad Categorizer'));
    gate.appendChild(el('p', { class: 'pm-help' },
      'Drop in a folder of client-supplied ad images and the app uses Claude vision to categorize each one (Awareness Level, Lead Type, Idea Name, Angle Name, Image Style) and outputs paste-ready tracker rows.'
    ));
    gate.appendChild(el('div', { class: 'pm-warn-banner' },
      el('strong', {}, '⚠️ Setup required: '),
      'Add your Anthropic API key to enable categorization. Stored only in your browser.',
    ));
    gate.appendChild(el('div', { class: 'pm-actions' },
      el('button', {
        type: 'button',
        class: 'btn-primary',
        onclick: openSettingsModal,
      }, 'Open Settings'),
    ));
    wrap.appendChild(gate);
    return;
  }

  // Header card with Settings link
  const headerCard = el('div', { class: 'card' });
  const headerRow = el('div', { class: 'pm-block-header' },
    el('div', {},
      el('h2', { style: 'margin: 0' }, '🤖 Ad Categorizer'),
      el('div', { class: 'pm-help', style: 'margin-top: 4px' },
        state.adCatMode === 'video'
          ? 'Drop video files (.mp4, .mov, .webm) or .zip. Each video has 5 keyframes extracted and sent to Claude for categorization. Output goes into the Video Creative Tracker.'
          : 'Drop ad images, .zip, or pick files. Each image is sent to Claude vision. Output goes into the Static Creative Tracker.'
      ),
    ),
    el('button', {
      type: 'button',
      class: 'btn-secondary btn-small',
      onclick: openSettingsModal,
    }, '⚙️ Settings'),
  );
  headerCard.appendChild(headerRow);

  // Static / Video sub-tab toggle
  const subTabs = el('div', { class: 'pm-source-toggle', style: 'margin-top: 14px' });
  for (const mode of ['static', 'video']) {
    subTabs.appendChild(el('button', {
      type: 'button',
      class: 'pm-source-tab' + (state.adCatMode === mode ? ' active' : ''),
      onclick: () => {
        if (state.adCatMode === mode) return;
        if (state.adCatGroups.length > 0 || state.adCatResult) {
          if (!confirm('Switching mode will clear the current files and results. Continue?')) return;
        }
        state.adCatMode = mode;
        state.adCatGroups = [];
        state.adCatResult = null;
        saveState();
        renderAdCategorizer();
      },
    }, mode === 'video' ? '🎬 Videos' : '🖼 Statics'));
  }
  headerCard.appendChild(subTabs);

  wrap.appendChild(headerCard);

  // Brief context card (only meaningful for video mode where transcript matching works)
  if (state.adCatMode === 'video') {
    wrap.appendChild(renderAdCatBriefContextCard());
  }

  // URL inputs card (Request Doc + Folder Link override)
  wrap.appendChild(renderAdCatUrlInputsCard());

  // Drop zone + file picker
  const inputCard = el('div', { class: 'card' });
  const dropZone = el('div', {
    id: 'ad-cat-dropzone',
    class: 'ad-cat-dropzone',
    ondragover: (e) => { e.preventDefault(); dropZone.classList.add('dragging'); },
    ondragleave: () => dropZone.classList.remove('dragging'),
    ondrop: async (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragging');
      await handleAdCatFileDrop(e);
    },
  });
  dropZone.appendChild(el('div', { class: 'ad-cat-dropzone-text' },
    state.adCatMode === 'video'
      ? 'Drag a folder, .zip, or video files here'
      : 'Drag a folder, .zip, or images here'
  ));

  const pickerAccept = state.adCatMode === 'video'
    ? '.mp4,.mov,.webm,.m4v,.zip'
    : '.png,.jpg,.jpeg,.webp,.gif,.zip';
  const pickerInput = el('input', {
    type: 'file',
    id: 'ad-cat-picker',
    multiple: true,
    accept: pickerAccept,
    style: 'display: none',
    onchange: (e) => handleAdCatFiles(Array.from(e.target.files)),
  });
  const folderInput = el('input', {
    type: 'file',
    id: 'ad-cat-folder',
    multiple: true,
    style: 'display: none',
    onchange: (e) => handleAdCatFiles(Array.from(e.target.files)),
  });
  folderInput.setAttribute('webkitdirectory', '');
  folderInput.setAttribute('directory', '');

  const buttonRow = el('div', { class: 'ad-cat-buttons' },
    el('button', { type: 'button', class: 'btn-secondary', onclick: () => pickerInput.click() }, 'Choose files'),
    el('button', { type: 'button', class: 'btn-secondary', onclick: () => folderInput.click() }, 'Choose folder'),
    pickerInput,
    folderInput,
  );
  dropZone.appendChild(buttonRow);

  inputCard.appendChild(dropZone);
  wrap.appendChild(inputCard);

  // File preview grid
  if (state.adCatGroups.length > 0) {
    wrap.appendChild(renderAdCatPreview());
    wrap.appendChild(renderAdCatActions());
  }

  // Result panel
  if (state.adCatResult) {
    wrap.appendChild(renderResultPanel(state.adCatResult));
  }
}

function renderAdCatUrlInputsCard() {
  const card = el('div', { class: 'card' });
  card.appendChild(el('h3', { style: 'margin: 0 0 4px' }, 'Tracker URLs (optional)'));
  const client = getEffectiveClient();
  const folderHint = client?.creative_folder_url
    ? `Defaults to ${client.name}'s folder URL if blank.`
    : 'Leave blank if you don\'t have one.';
  card.appendChild(el('p', { class: 'pm-help', style: 'margin: 0 0 12px' },
    `Pre-fill the Request Doc and Folder Link columns in the tracker output. ${folderHint}`
  ));

  const row = el('div', { class: 'pm-input-row' });

  const reqField = el('div', { class: 'pm-input-field' });
  reqField.appendChild(el('label', { for: 'ad-cat-request-doc' },
    'Request Doc URL', el('span', { class: 'pm-hint-inline' }, ' (goes into Request Doc column)')));
  const reqInput = el('input', {
    id: 'ad-cat-request-doc',
    type: 'text',
    placeholder: 'https://docs.google.com/document/d/...',
    oninput: (e) => {
      state.adCatRequestDoc = e.target.value;
      saveState();
      // If a result is already shown, rebuild it so the new URL flows in
      if (state.adCatResult) rebuildAdCatResult();
    },
  });
  reqInput.value = state.adCatRequestDoc;
  reqField.appendChild(reqInput);
  row.appendChild(reqField);

  const folderField = el('div', { class: 'pm-input-field' });
  folderField.appendChild(el('label', { for: 'ad-cat-folder-url' },
    'Folder Link URL', el('span', { class: 'pm-hint-inline' }, ' (goes into Folder Link column; overrides client default)')));
  const folderInput = el('input', {
    id: 'ad-cat-folder-url',
    type: 'text',
    placeholder: client?.creative_folder_url || 'https://drive.google.com/drive/folders/...',
    oninput: (e) => {
      state.adCatFolderUrl = e.target.value;
      saveState();
      if (state.adCatResult) rebuildAdCatResult();
    },
  });
  folderInput.value = state.adCatFolderUrl;
  folderField.appendChild(folderInput);
  row.appendChild(folderField);

  card.appendChild(row);
  return card;
}

function renderAdCatBriefContextCard() {
  const card = el('div', { class: 'card' });
  card.appendChild(el('h3', { style: 'margin: 0 0 4px' }, 'Does a brief exist for these videos?'));
  card.appendChild(el('p', { class: 'pm-help', style: 'margin: 0 0 12px' },
    'If the videos are recordings of scripts you already wrote in a brief, the AI can match each video to its source brief variation and inherit Awareness Level, Lead Type, Idea Name, etc. from the brief instead of guessing.'
  ));

  // Mode toggle
  const toggle = el('div', { class: 'pm-source-toggle', style: 'margin-bottom: 12px' });
  const setMode = (mode) => {
    state.adCatBriefMode = mode;
    saveState();
    renderAdCategorizer();
  };
  for (const m of [['none', 'No brief'], ['saved', 'Pick saved brief'], ['paste', 'Paste brief text']]) {
    toggle.appendChild(el('button', {
      type: 'button',
      class: 'pm-source-tab' + (state.adCatBriefMode === m[0] ? ' active' : ''),
      onclick: () => setMode(m[0]),
    }, m[1]));
  }
  card.appendChild(toggle);

  // Model recommendation when brief is loaded
  if (state.adCatBriefMode !== 'none' && getModel().includes('haiku')) {
    card.appendChild(el('div', { class: 'pm-warn-banner', style: 'margin-bottom: 12px' },
      el('strong', {}, '⚠️ For brief matching, switch to Sonnet 4.6 or Opus 4.6 in Settings.'),
      ' Haiku 4.5 is too weak at fuzzy script matching across multiple candidate variations \u2014 you\'ll get wrong/missed matches frequently.',
    ));
  }

  if (state.adCatBriefMode === 'saved') {
    // Show video saved-briefs filtered to current client (or all if no client)
    const allBriefs = sortBriefsRecent(loadBriefs()).filter(b => b.briefType === 'video');
    const client = getEffectiveClient();
    const briefs = client ? allBriefs.filter(b => b.clientId === client.id) : allBriefs;

    if (briefs.length === 0) {
      card.appendChild(el('div', { class: 'pm-detect pm-detect-empty' },
        client
          ? `No saved video briefs for ${client.name}. Switch client or paste text instead.`
          : 'No saved video briefs. Generate one in Briefgen first, or paste brief text.'
      ));
    } else {
      const sel = el('select', {
        class: 'pm-saved-select',
        onchange: (e) => {
          state.adCatBriefId = e.target.value || null;
          saveState();
        },
      });
      sel.appendChild(el('option', { value: '' }, '— pick a brief —'));
      for (const b of briefs) {
        const opt = el('option', { value: b.id }, briefDisplayLabel(b) + ` — ${b.creatives.length} variation${b.creatives.length === 1 ? '' : 's'}`);
        if (b.id === state.adCatBriefId) opt.setAttribute('selected', '');
        sel.appendChild(opt);
      }
      card.appendChild(sel);

      // Show preview of variations
      if (state.adCatBriefId) {
        const brief = briefs.find(b => b.id === state.adCatBriefId);
        if (brief) {
          const summary = el('div', { class: 'pm-detect', style: 'margin-top: 8px' },
            `Will match against ${brief.creatives.length} variation${brief.creatives.length === 1 ? '' : 's'} from this brief.`
          );
          card.appendChild(summary);
        }
      }
    }
  } else if (state.adCatBriefMode === 'paste') {
    const ta = el('textarea', {
      placeholder: 'Paste the full brief text here (same format as Briefgen Quick Fill).',
      rows: 8,
      style: 'width: 100%; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size: 12px; padding: 10px; border: 1px solid var(--border); border-radius: 6px;',
      oninput: (e) => {
        state.adCatBriefText = e.target.value;
        saveState();
        // Live preview of how many variations parsed
        const previewEl = card.querySelector('.pm-paste-preview');
        if (previewEl) previewEl.textContent = previewParsedBriefText(state.adCatBriefText);
      },
    });
    ta.value = state.adCatBriefText;
    card.appendChild(ta);
    card.appendChild(el('div', { class: 'pm-detect pm-paste-preview', style: 'margin-top: 8px' },
      previewParsedBriefText(state.adCatBriefText)
    ));
  }

  return card;
}

function previewParsedBriefText(text) {
  if (!text || !text.trim()) return 'Paste brief text to see how many variations are detected.';
  try {
    const parsed = parseClaudeOutput(text);
    return `Parsed ${parsed.creatives.length} variation${parsed.creatives.length === 1 ? '' : 's'} from the pasted brief.`;
  } catch (e) {
    return 'Parse error: ' + e.message;
  }
}

/**
 * Returns the candidate brief variations to send to Claude for matching, or
 * null if no brief is configured. Each candidate has briefId, briefDisplayLabel,
 * variationIndex, leadScript, bodyScript + the full overview/variation data
 * we'll merge into the result.
 */
function getAdCatBriefCandidates() {
  if (state.adCatMode !== 'video' || state.adCatBriefMode === 'none') return null;

  let briefSource = null;
  if (state.adCatBriefMode === 'saved' && state.adCatBriefId) {
    const stored = getBrief(state.adCatBriefId);
    if (!stored || stored.briefType !== 'video') return null;
    briefSource = {
      label: briefDisplayLabel(stored),
      overview: stored.overview,
      variations: stored.creatives,
    };
  } else if (state.adCatBriefMode === 'paste' && state.adCatBriefText.trim()) {
    try {
      const parsed = parseClaudeOutput(state.adCatBriefText);
      briefSource = {
        label: '(pasted brief)',
        overview: parsed.overview,
        variations: parsed.creatives,
      };
    } catch (e) {
      return null;
    }
  }
  if (!briefSource || briefSource.variations.length === 0) return null;

  return briefSource.variations.map((v, i) => ({
    briefDisplayLabel: briefSource.label,
    variationIndex: i,
    leadScript: v['Lead Script'] || '',
    bodyScript: v['Body Script'] || '',
    // Embedded full data to inherit when matched
    _overview: briefSource.overview,
    _variation: v,
  }));
}

function renderAdCatPreview() {
  const card = el('div', { class: 'card' });
  const header = el('div', { class: 'pm-block-header' },
    el('div', {},
      el('h3', { style: 'margin: 0' }, `${state.adCatGroups.length} unique creative${state.adCatGroups.length === 1 ? '' : 's'} (${state.adCatGroups.reduce((a, g) => a + g.files.length, 0)} files)`),
      el('div', { class: 'pm-help', style: 'margin-top: 4px; font-size: 12px' },
        state.adCatRunning
          ? 'Categorizing… please wait.'
          : 'Drag tiles to reorder. The output rows will follow this order.',
      ),
    ),
    el('div', { style: 'display: flex; gap: 8px' },
      el('button', {
        type: 'button',
        class: 'btn-secondary btn-small',
        onclick: () => {
          if (state.adCatRunning) return;
          // Sort by filename A→Z (parsed baseName)
          state.adCatGroups.sort((a, b) => a.parsed.baseName.localeCompare(b.parsed.baseName));
          renderAdCategorizer();
        },
        disabled: state.adCatRunning ? '' : null,
        title: 'Sort tiles alphabetically by filename',
      }, 'Sort A→Z'),
      el('button', {
        type: 'button',
        class: 'btn-secondary btn-small',
        onclick: () => {
          if (state.adCatRunning) return;
          state.adCatGroups = [];
          state.adCatResult = null;
          renderAdCategorizer();
        },
        disabled: state.adCatRunning ? '' : null,
      }, 'Clear'),
    ),
  );
  card.appendChild(header);

  const grid = el('div', { class: 'ad-cat-grid' });
  state.adCatGroups.forEach((g, idx) => {
    const tile = el('div', {
      class: 'ad-cat-tile ad-cat-tile-' + (g.status || 'pending'),
      draggable: state.adCatRunning ? null : 'true',
      'data-index': String(idx),
    });

    // Drag handlers
    tile.addEventListener('dragstart', (e) => {
      if (state.adCatRunning) return;
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', String(idx));
      tile.classList.add('ad-cat-tile-dragging');
    });
    tile.addEventListener('dragend', () => {
      tile.classList.remove('ad-cat-tile-dragging');
      // Clean up any leftover dragover styling
      grid.querySelectorAll('.ad-cat-tile-dropbefore').forEach(el => el.classList.remove('ad-cat-tile-dropbefore'));
    });
    tile.addEventListener('dragover', (e) => {
      if (state.adCatRunning) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      tile.classList.add('ad-cat-tile-dropbefore');
    });
    tile.addEventListener('dragleave', () => {
      tile.classList.remove('ad-cat-tile-dropbefore');
    });
    tile.addEventListener('drop', (e) => {
      if (state.adCatRunning) return;
      e.preventDefault();
      tile.classList.remove('ad-cat-tile-dropbefore');
      const fromIdx = parseInt(e.dataTransfer.getData('text/plain'), 10);
      const toIdx = idx;
      if (Number.isNaN(fromIdx) || fromIdx === toIdx) return;
      const arr = state.adCatGroups;
      const [moved] = arr.splice(fromIdx, 1);
      arr.splice(toIdx, 0, moved);
      renderAdCategorizer();
    });

    // Position indicator (Row N) — useful when planning paste order
    const posBadge = el('div', { class: 'ad-cat-pos' }, `Row ${idx + 1}`);
    tile.appendChild(posBadge);

    // Thumbnail (use first file from group as representative).
    // For video files, render a <video> element (poster = first frame, no controls).
    const repFile = pickRepresentativeFile(g);
    if (isVideoFile(repFile)) {
      const vid = el('video', {
        class: 'ad-cat-thumb',
        muted: true,
        playsinline: true,
        preload: 'metadata',
        disablePictureInPicture: true,
      });
      // Explicitly disable browser controls (they sometimes show even when
      // 'controls' attr is absent if the user has accessibility settings on)
      vid.controls = false;
      vid.removeAttribute('controls');
      vid.src = URL.createObjectURL(repFile);
      vid.onloadedmetadata = () => {
        // Seek to first frame for the poster
        vid.currentTime = 0.1;
      };
      tile.appendChild(vid);
    } else {
      const img = el('img', { class: 'ad-cat-thumb' });
      img.src = URL.createObjectURL(repFile);
      img.onload = () => URL.revokeObjectURL(img.src);
      tile.appendChild(img);
    }

    const meta = el('div', { class: 'ad-cat-meta' });
    // Show the full filename of the representative file (with extension)
    // so the PM can verify which file is which.
    const displayName = repFile.name;
    meta.appendChild(el('div', { class: 'ad-cat-filename', title: displayName }, displayName));
    const tags = el('div', { class: 'ad-cat-tags' });
    if (g.parsed.cid) tags.appendChild(el('span', { class: 'ad-cat-tag ad-cat-tag-cid' }, g.parsed.cid));
    if (g.ratios.length > 0) tags.appendChild(el('span', { class: 'ad-cat-tag' }, g.ratios.join(' + ')));
    if (g.parsed.leadType) tags.appendChild(el('span', { class: 'ad-cat-tag ad-cat-tag-lead' }, g.parsed.leadType));
    if (g.briefMatch) {
      tags.appendChild(el('span', {
        class: 'ad-cat-tag ad-cat-tag-match',
        title: `Matched to ${g.briefMatch.briefDisplayLabel} variation ${g.briefMatch.variationIndex + 1}`,
      }, `✅ Brief #${g.briefMatch.variationIndex + 1}`));
    }
    meta.appendChild(tags);

    // Status indicator
    const statusEl = el('div', { class: 'ad-cat-status' });
    if (g.status === 'analyzing') statusEl.textContent = '⏳ analyzing…';
    else if (g.status === 'done') statusEl.textContent = '✅ done';
    else if (g.status === 'error') statusEl.textContent = '❌ ' + (g.error || 'error');
    else statusEl.textContent = '';
    meta.appendChild(statusEl);

    // Retry button for failed tiles
    if (g.status === 'error' && !state.adCatRunning) {
      meta.appendChild(el('button', {
        type: 'button',
        class: 'btn-secondary btn-small',
        style: 'margin-top: 6px; width: 100%',
        onclick: () => handleAdCatRetry(idx),
      }, '🔁 Retry this'));
    }

    // After analysis, show diagnostics + manual override
    if (g.status === 'done' && g.result) {
      const r = g.result;

      // Manual brief-match override (only for video + when brief candidates exist)
      if (state.adCatMode === 'video' && state.adCatBriefMode !== 'none') {
        const candidates = getAdCatBriefCandidates();
        if (candidates && candidates.length > 0) {
          const overrideRow = el('div', { class: 'ad-cat-override' });
          overrideRow.appendChild(el('label', { style: 'font-size: 10px; color: var(--muted); display: block; margin-bottom: 2px' }, 'Brief variation:'));
          const sel = el('select', {
            class: 'ad-cat-override-select',
            onchange: (e) => {
              const newIdx = e.target.value === '' ? null : parseInt(e.target.value, 10);
              g.result.matchedCandidateIndex = newIdx;
              g.briefMatchManual = true; // PM said so — sticky over filename auto-match
              // Re-run mapping to apply the override
              rebuildAdCatResult();
              // Also flash a status so it's clear the change took effect
              flashStatus(newIdx === null
                ? 'Match cleared. Tracker rows updated.'
                : `Matched to brief variation #${newIdx + 1}. Tracker rows updated.`,
                'success');
            },
          });
          sel.appendChild(el('option', { value: '' }, '— no match (use AI fields) —'));
          candidates.forEach((c, i) => {
            const opt = el('option', { value: String(i) },
              `#${c.variationIndex + 1}: ${(c.leadScript || '(blank)').slice(0, 50)}…`);
            if (i === r.matchedCandidateIndex) opt.setAttribute('selected', '');
            sel.appendChild(opt);
          });
          overrideRow.appendChild(sel);
          meta.appendChild(overrideRow);
        }
      }

      // Match rationale (clickable to expand)
      if (r.matchRationale) {
        const rationaleEl = el('details', { class: 'ad-cat-diag' });
        rationaleEl.appendChild(el('summary', {}, '🔍 Why this match?'));
        rationaleEl.appendChild(el('div', { class: 'ad-cat-diag-body' }, r.matchRationale));
        meta.appendChild(rationaleEl);
      }

      // Transcript (clickable to expand) + transcript status
      if (r.transcript) {
        const trEl = el('details', { class: 'ad-cat-diag' });
        trEl.appendChild(el('summary', {}, '📝 Transcript'));
        trEl.appendChild(el('div', { class: 'ad-cat-diag-body ad-cat-diag-transcript' }, r.transcript));
        meta.appendChild(trEl);
      } else if (r.transcriptStatus && r.transcriptStatus !== 'ok') {
        // Transcript failed — make this loud so the PM knows brief matching can't work properly
        const labels = {
          'no-key': '⚠️ No OpenAI key — add one in Settings to enable transcription',
          'too-big': '⚠️ Audio still over 25 MB after extraction (very long video?)',
          'error': '⚠️ Transcription failed',
        };
        const msg = labels[r.transcriptStatus] || '⚠️ No transcript';
        const errEl = el('div', { class: 'ad-cat-transcript-warn', title: r.transcriptError || '' }, msg);
        meta.appendChild(errEl);
      }
    }

    tile.appendChild(meta);
    grid.appendChild(tile);
  });
  card.appendChild(grid);
  return card;
}

function renderAdCatActions() {
  const card = el('div', { class: 'card' });
  const numItems = state.adCatGroups.length;
  const isVideo = state.adCatMode === 'video';
  const cost = estimateCost(numItems, undefined, isVideo ? 'video' : 'static');
  const costStr = `≈$${cost.toFixed(2)}`;

  const hasOpenAI = !!getOpenAIKey();
  const transcriptHint = isVideo
    ? (hasOpenAI
        ? ' Audio will be transcribed via Whisper (~$0.006/min) for better Lead Type accuracy.'
        : ' ⚠️ No OpenAI key set — Lead Type accuracy will be lower without audio transcription. Add a key in Settings.')
    : '';
  const detail = isVideo
    ? `Each video has 5 keyframes extracted in your browser, then sent together to Claude.${transcriptHint}`
    : `Each ad gets a separate API call to Claude vision.`;

  const info = el('p', { class: 'pm-help' },
    `Will analyze ${numItems} unique ${isVideo ? 'video' : 'creative'}${numItems === 1 ? '' : 's'} (${costStr}). ${detail}`
  );
  card.appendChild(info);

  const actions = el('div', { class: 'pm-actions' },
    el('button', {
      type: 'button',
      class: 'btn-primary',
      disabled: state.adCatRunning ? '' : null,
      onclick: handleAdCatCategorize,
    }, state.adCatRunning ? 'Categorizing…' : `Categorize ${numItems} ${isVideo ? 'video' : 'ad'}${numItems === 1 ? '' : 's'}`),
  );
  card.appendChild(actions);
  return card;
}

async function handleAdCatFileDrop(e) {
  const items = e.dataTransfer.items;
  const files = [];
  if (items && items.length > 0) {
    // Use webkitGetAsEntry for folder support
    const entries = [];
    for (const item of items) {
      const entry = item.webkitGetAsEntry?.();
      if (entry) entries.push(entry);
      else if (item.kind === 'file') files.push(item.getAsFile());
    }
    for (const entry of entries) {
      await collectEntryFiles(entry, files);
    }
  } else {
    for (const f of e.dataTransfer.files) files.push(f);
  }
  await handleAdCatFiles(files);
}

async function collectEntryFiles(entry, out) {
  if (entry.isFile) {
    await new Promise((resolve, reject) => {
      entry.file(f => { out.push(f); resolve(); }, reject);
    });
  } else if (entry.isDirectory) {
    const reader = entry.createReader();
    const readBatch = () => new Promise((resolve, reject) => {
      reader.readEntries(resolve, reject);
    });
    while (true) {
      const batch = await readBatch();
      if (batch.length === 0) break;
      for (const e of batch) await collectEntryFiles(e, out);
    }
  }
}

async function handleAdCatFiles(rawFiles) {
  const isVideoMode = state.adCatMode === 'video';
  const acceptedRe = isVideoMode
    ? /\.(mp4|mov|webm|m4v)$/i
    : /\.(png|jpe?g|webp|gif)$/i;

  // Filter to accepted media + zips, expand zips
  const filtered = [];
  for (const f of rawFiles) {
    if (!f) continue;
    const lower = (f.name || '').toLowerCase();
    if (acceptedRe.test(lower)) {
      filtered.push(f);
    } else if (lower.endsWith('.zip') && typeof JSZip !== 'undefined') {
      try {
        const zip = await JSZip.loadAsync(f);
        for (const [name, entry] of Object.entries(zip.files)) {
          if (entry.dir) continue;
          if (!acceptedRe.test(name)) continue;
          if (name.includes('__MACOSX/') || name.endsWith('.DS_Store')) continue;
          const blob = await entry.async('blob');
          const base = name.split('/').pop();
          const file = new File([blob], base, { type: blob.type || (isVideoMode ? 'video/mp4' : 'image/png') });
          filtered.push(file);
        }
      } catch (e) {
        flashStatus(`Could not unzip ${f.name}: ${e.message}`, 'error');
      }
    }
  }

  if (filtered.length === 0) {
    flashStatus(
      isVideoMode
        ? 'No video files found. Accepted: .mp4, .mov, .webm, .m4v, .zip'
        : 'No image files found. Accepted: .png, .jpg, .webp, .gif, .zip',
      'error'
    );
    return;
  }

  // Group by CID/basename
  const groupsMap = groupAdFiles(filtered);
  const groups = [...groupsMap.values()].map(g => ({ ...g, status: 'pending', result: null, error: null }));
  state.adCatGroups = groups;
  state.adCatResult = null;
  renderAdCategorizer();
  flashStatus(`Loaded ${filtered.length} files → ${groups.length} unique creative${groups.length === 1 ? '' : 's'}.`);
}

async function handleAdCatCategorize() {
  if (state.adCatGroups.length === 0) return;
  const apiKey = getApiKey();
  if (!apiKey) {
    flashStatus('No API key set. Open Settings to add one.', 'error');
    return;
  }

  state.adCatRunning = true;
  state.adCatResult = null;
  // Reset per-tile state from any prior run so reruns start clean
  for (const g of state.adCatGroups) {
    g.status = 'pending';
    g.result = null;
    g.error = null;
    g.briefMatch = null;
    // NOTE: g.briefMatchManual is intentionally NOT reset \u2014 the PM's manual
    // overrides survive re-categorization.
  }
  renderAdCategorizer();

  const jobs = state.adCatGroups.map(g => ({
    file: pickRepresentativeFile(g),
    filenameHints: {
      ideaName: g.parsed.ideaName,
      leadType: g.parsed.leadType,
      cid: g.parsed.cid,
    },
  }));

  // Brief-match candidates (video mode only). Indexed list — Claude returns
  // the matching candidate's index per video, and we look it up here.
  const briefCandidates = getAdCatBriefCandidates();

  try {
    await analyzeBatch(jobs, (idx, status, payload) => {
      state.adCatGroups[idx].status = status;
      if (status === 'done') state.adCatGroups[idx].result = payload;
      if (status === 'error') state.adCatGroups[idx].error = payload?.message || 'error';
      // Re-render preview only (keep cost actions)
      const oldPreview = document.querySelector('.ad-cat-grid')?.parentElement;
      if (oldPreview) oldPreview.replaceWith(renderAdCatPreview());
    }, { briefCandidates });
  } catch (e) {
    console.error('analyzeBatch failed:', e);
    flashStatus('Categorization failed: ' + e.message, 'error');
  } finally {
    // Always clear the running flag so the UI doesn't lock up forever on errors
    state.adCatRunning = false;
  }

  rebuildAdCatResult();
}

async function handleAdCatRetry(idx) {
  const g = state.adCatGroups[idx];
  if (!g || state.adCatRunning) return;
  state.adCatRunning = true;
  // Reset just this tile
  g.status = 'pending';
  g.result = null;
  g.error = null;
  g.briefMatch = null;
  renderAdCategorizer();

  const job = {
    file: pickRepresentativeFile(g),
    filenameHints: {
      ideaName: g.parsed.ideaName,
      leadType: g.parsed.leadType,
      cid: g.parsed.cid,
    },
  };
  const briefCandidates = getAdCatBriefCandidates();

  try {
    const results = await analyzeBatch([job], (i, status, payload) => {
      g.status = status;
      if (status === 'done') g.result = payload;
      if (status === 'error') g.error = payload?.message || 'error';
      const oldPreview = document.querySelector('.ad-cat-grid')?.parentElement;
      if (oldPreview) oldPreview.replaceWith(renderAdCatPreview());
    }, { briefCandidates });
    if (results[0]?.ok) flashStatus('Retry succeeded.');
  } catch (e) {
    flashStatus('Retry failed: ' + e.message, 'error');
  } finally {
    state.adCatRunning = false;
  }
  rebuildAdCatResult();
}

function rebuildAdCatResult() {
  const successful = state.adCatGroups.filter(g => g.status === 'done' && g.result);
  if (successful.length === 0) {
    flashStatus('No ads were successfully analyzed.', 'error');
    renderAdCategorizer();
    return;
  }

  const effectiveClient = getEffectiveClient();
  const isVideo = state.adCatMode === 'video';
  const briefCandidates = getAdCatBriefCandidates();

  const variations = successful.map(g => {
    const r = g.result;
    // Reset prior briefMatch state so re-runs reflect current matchedCandidateIndex
    g.briefMatch = null;

    // FILENAME-BASED MATCH PRIORITY: if the filename explicitly contains a
    // variation number (e.g. "Office Ads - Lead 3.mp4") and the brief has
    // that many variations, prefer that over whatever Claude guessed.
    // Only do this when the PM hasn't manually overridden (we detect manual
    // override by g.briefMatchManual being set).
    if (isVideo && briefCandidates && !g.briefMatchManual) {
      const vn = g.parsed.variationNumber;
      if (Number.isInteger(vn) && vn >= 1 && vn <= briefCandidates.length) {
        const candIdx = vn - 1; // candidates are 0-indexed in array order
        if (r.matchedCandidateIndex !== candIdx) {
          // Override Claude's guess with the filename-based match
          r.matchedCandidateIndex = candIdx;
          r._matchSource = 'filename';
        }
      }
    }

    // Step 1: AI + filename baseline
    const leadType = g.parsed.leadType || r.leadType || '';
    const ideaName = g.parsed.ideaName || r.ideaName || '';
    const styleField = isVideo ? (r.styleName || '') : (r.imageStyle || '');
    const v = {
      'File Name': '',
      'Variation Type': r.variationType || (isVideo ? 'Lead' : 'Copy'),
      'Awareness Level': r.awarenessLevel || '',
      'Lead Type': leadType,
      'Status': '',
      'Idea Name': ideaName,
      'Angle Name': r.angleName || '',
      'Style Name': styleField,
      'Image Style': styleField,
    };
    if (isVideo && isFinite(r.durationSec) && r.durationSec > 0) {
      v['Length (in Sec)'] = String(Math.round(r.durationSec));
    }

    // Step 2: if Claude matched this video to a brief variation, override
    // every brief-set field with the brief's value.
    if (isVideo && briefCandidates && Number.isInteger(r.matchedCandidateIndex)) {
      const cand = briefCandidates[r.matchedCandidateIndex];
      if (cand) {
        const overview = cand._overview || {};
        const variation = cand._variation || {};
        // Brief is source of truth — only overwrite if brief actually has the value
        const inheritFromBrief = (key) => {
          if (variation[key] != null && String(variation[key]).trim() !== '') v[key] = variation[key];
          else if (overview[key] != null && String(overview[key]).trim() !== '') v[key] = overview[key];
        };
        inheritFromBrief('Variation Type');
        inheritFromBrief('Awareness Level');
        inheritFromBrief('Lead Type');
        inheritFromBrief('Status');
        inheritFromBrief('Idea Name');
        inheritFromBrief('Angle Name');
        inheritFromBrief('Style Name');
        // Style Name in overview maps to Image Style for static; kept in sync
        if (variation['Style Name']) v['Image Style'] = variation['Style Name'];
        else if (overview['Style Name']) v['Image Style'] = overview['Style Name'];
        // Tag the group with the match info for the tile badge
        g.briefMatch = {
          briefDisplayLabel: cand.briefDisplayLabel,
          variationIndex: cand.variationIndex,
        };
      }
    }
    return v;
  });

  const fakeBrief = {
    briefType: isVideo ? 'video' : 'static',
    overview: { 'Net New/Iteration': 'Net New' },
    variations,
  };
  try {
    const result = mapBriefToTracker(fakeBrief, {
      requestDoc: state.adCatRequestDoc || '',
      clientTrackerOverrides: effectiveClient?.tracker_overrides,
      // PM-supplied folder URL takes priority over the client default
      clientFolderUrl: state.adCatFolderUrl || effectiveClient?.creative_folder_url,
    });

    // Append a "client typically supplies" note where relevant
    const notes = [...(result.notes || [])];
    if (effectiveClient && CLIENTS_THAT_SUPPLY_NAMES.has(effectiveClient.id)) {
      notes.push(`${effectiveClient.name} typically supplies Idea Name, Angle Name, and Image Style themselves — review and clear those columns before pasting if you don't want to overwrite the client's data.`);
    }
    // Brief-match summary
    if (briefCandidates) {
      const matched = successful.filter(g => g.briefMatch).length;
      const total = successful.length;
      if (matched > 0) {
        notes.push(`✅ Brief matching: ${matched} of ${total} video${total === 1 ? '' : 's'} matched to a brief variation. Matched videos inherit Awareness Level, Lead Type, Idea Name, Angle Name, and Style Name from the brief (overriding AI guesses).`);
      } else {
        notes.push(`⚠️ Brief matching: 0 of ${total} videos matched any brief variation confidently. AI-inferred fields used as-is. Check that the brief you provided matches these videos.`);
      }
    }
    result.notes = notes;
    // Attach the row-by-row filename mapping so the result panel can show it
    result.adCatFilenames = successful.map(g => g.parsed.baseName);
    state.adCatResult = result;
  } catch (e) {
    flashStatus('Mapping error: ' + e.message, 'error');
  }
  renderAdCategorizer();
}

function openSettingsModal() {
  const root = document.getElementById('modal-root');
  root.innerHTML = '';
  const close = () => {
    root.innerHTML = '';
    document.removeEventListener('keydown', onKey);
  };
  const onKey = (e) => { if (e.key === 'Escape') close(); };
  document.addEventListener('keydown', onKey);

  const overlay = el('div', { class: 'modal-overlay', onclick: (e) => { if (e.target === overlay) close(); } });
  const dialog = el('div', { class: 'modal-dialog', style: 'max-width: 540px' });

  dialog.appendChild(el('div', { class: 'modal-header' },
    el('h2', {}, '⚙️ Ad Categorizer Settings'),
    el('button', { type: 'button', class: 'modal-close', onclick: close, title: 'Close' }, '×'),
  ));

  const body = el('div', { class: 'modal-body' });

  body.appendChild(el('p', { class: 'modal-hint' },
    'Your Anthropic API key is stored only in this browser (localStorage) and sent directly to api.anthropic.com from your browser. Get a key at ',
    el('a', { href: 'https://console.anthropic.com/settings/keys', target: '_blank', rel: 'noopener' }, 'console.anthropic.com'),
    '.'
  ));

  // API key input
  const keyField = el('div', { class: 'pm-input-field' });
  keyField.appendChild(el('label', { for: 'settings-api-key' }, 'Anthropic API Key'));
  const keyInput = el('input', {
    id: 'settings-api-key',
    type: 'password',
    placeholder: 'sk-ant-...',
  });
  keyInput.value = getApiKey();
  keyField.appendChild(keyInput);

  const showHide = el('button', {
    type: 'button',
    class: 'btn-secondary btn-small',
    style: 'margin-top: 4px',
    onclick: () => {
      keyInput.type = keyInput.type === 'password' ? 'text' : 'password';
      showHide.textContent = keyInput.type === 'password' ? '👁 Show' : '🙈 Hide';
    },
  }, '👁 Show');
  keyField.appendChild(showHide);
  body.appendChild(keyField);

  // OpenAI key (for video transcription via Whisper)
  const openaiField = el('div', { class: 'pm-input-field', style: 'margin-top: 14px' });
  openaiField.appendChild(el('label', { for: 'settings-openai-key' },
    'OpenAI API Key',
    el('span', { class: 'pm-hint-inline' }, ' (optional — used to transcribe video audio for better Lead Type accuracy)')));
  const openaiInput = el('input', {
    id: 'settings-openai-key',
    type: 'password',
    placeholder: 'sk-...',
  });
  openaiInput.value = getOpenAIKey();
  openaiField.appendChild(openaiInput);
  const openaiShowHide = el('button', {
    type: 'button',
    class: 'btn-secondary btn-small',
    style: 'margin-top: 4px',
    onclick: () => {
      openaiInput.type = openaiInput.type === 'password' ? 'text' : 'password';
      openaiShowHide.textContent = openaiInput.type === 'password' ? '👁 Show' : '🙈 Hide';
    },
  }, '👁 Show');
  openaiField.appendChild(openaiShowHide);
  body.appendChild(openaiField);

  // Model picker
  const modelField = el('div', { class: 'pm-input-field', style: 'margin-top: 14px' });
  modelField.appendChild(el('label', { for: 'settings-model' }, 'Model'));
  const modelSel = el('select', { id: 'settings-model' },
    el('option', { value: DEFAULT_MODEL }, 'Haiku 4.5 — fast & cheap (~$0.25 per 100 ads)'),
    el('option', { value: SONNET_MODEL }, 'Sonnet 4.6 — balanced (~$0.75 per 100 ads)'),
    el('option', { value: OPUS_MODEL }, 'Opus 4.6 — highest quality, slowest (~$3.75 per 100 ads)'),
  );
  modelSel.value = getModel();
  modelField.appendChild(modelSel);
  body.appendChild(modelField);

  // Concurrency
  const concField = el('div', { class: 'pm-input-field', style: 'margin-top: 14px' });
  concField.appendChild(el('label', { for: 'settings-concurrency' }, 'Concurrency (parallel API calls)'));
  const concInput = el('input', {
    id: 'settings-concurrency',
    type: 'number',
    min: 1,
    max: 5,
    value: getConcurrency(),
  });
  concField.appendChild(concInput);
  body.appendChild(concField);

  // Test connection
  const testStatus = el('div', { class: 'pm-detect', style: 'margin-top: 10px' });
  body.appendChild(el('button', {
    type: 'button',
    class: 'btn-secondary btn-small',
    style: 'margin-top: 14px',
    onclick: async () => {
      testStatus.textContent = 'Testing…';
      const result = await testApiKey(keyInput.value);
      if (result.ok) {
        testStatus.className = 'pm-detect';
        testStatus.textContent = '✅ Connection successful.';
      } else {
        testStatus.className = 'pm-detect qf-detected-error';
        testStatus.textContent = '❌ ' + result.error;
      }
    },
  }, 'Test connection'));
  body.appendChild(testStatus);

  dialog.appendChild(body);

  dialog.appendChild(el('div', { class: 'modal-footer' },
    el('button', { type: 'button', class: 'btn-secondary', onclick: close }, 'Cancel'),
    el('button', {
      type: 'button',
      class: 'btn-primary',
      onclick: () => {
        setApiKey(keyInput.value.trim());
        setOpenAIKey(openaiInput.value.trim());
        setModel(modelSel.value);
        setConcurrency(parseInt(concInput.value, 10) || 3);
        close();
        if (state.tab === 'ad-categorizer') renderAdCategorizer();
        flashStatus('Settings saved.');
      },
    }, 'Save'),
  ));

  overlay.appendChild(dialog);
  root.appendChild(overlay);
}

function truncate(s, n) {
  return s.length > n ? s.slice(0, n - 1) + '…' : s;
}

// -------- Tab 4 & 5: Tracker report tabs (Meeting Notes + Weekly Updates) --------

/**
 * Render a generic "drop .xlsx → HTML report" tab. Both Meeting Notes and
 * Weekly Creative Updates use this same UI pattern, parameterized by the
 * builder function and a label.
 *
 * cfg = {
 *   stateKey: 'meeting' | 'weekly',
 *   title, intro, accept, buttonLabel,
 *   buildHtml(workbooksWithNames) -> html string,
 *   suggestedFilename() -> string,
 * }
 */
function renderTrackerReportTab(cfg) {
  const wrap = document.getElementById('pm-content');
  wrap.innerHTML = '';

  const filesKey = `${cfg.stateKey}Files`;
  const htmlKey = `${cfg.stateKey}Html`;
  const filenameKey = `${cfg.stateKey}Filename`;
  const errorKey = `${cfg.stateKey}Error`;

  const card = el('div', { class: 'card' });
  card.appendChild(el('h2', {}, cfg.title));
  card.appendChild(el('p', { class: 'muted', style: 'margin-top:0;' }, cfg.intro));

  // Dropzone
  const dropzone = el('div', {
    class: 'tracker-report-dropzone',
    onclick: () => fileInput.click(),
    ondragover: e => { e.preventDefault(); dropzone.classList.add('drag-active'); },
    ondragleave: () => dropzone.classList.remove('drag-active'),
    ondrop: e => {
      e.preventDefault();
      dropzone.classList.remove('drag-active');
      handleFiles(e.dataTransfer?.files);
    },
  });
  dropzone.appendChild(el('div', { style: 'font-weight:500;' }, 'Drag & drop tracker .xlsx files here'));
  dropzone.appendChild(el('div', { class: 'muted', style: 'font-size:0.85em;margin-top:4px;' }, 'or click to choose files'));
  dropzone.appendChild(el('div', { class: 'muted', style: 'font-size:0.75em;margin-top:8px;' }, 'Allowed: .xlsx, .xlsm'));
  const fileInput = el('input', {
    type: 'file', multiple: true, accept: '.xlsx,.xlsm',
    style: 'display:none;',
    onchange: e => { handleFiles(e.target.files); e.target.value = ''; },
  });
  dropzone.appendChild(fileInput);
  card.appendChild(dropzone);

  // File list
  if (state[filesKey].length) {
    const list = el('ul', { class: 'tracker-report-filelist' });
    state[filesKey].forEach((f, i) => {
      list.appendChild(el('li', {},
        el('span', {}, f.name),
        el('span', { class: 'muted', style: 'margin-left:auto;font-size:0.85em;' }, fmtBytes(f.size)),
        el('button', {
          type: 'button',
          class: 'btn-link',
          style: 'margin-left:8px;',
          onclick: () => {
            state[filesKey].splice(i, 1);
            state[htmlKey] = '';
            state[errorKey] = '';
            renderAll();
          },
        }, '✕'),
      ));
    });
    card.appendChild(list);
  }

  // Action row
  const actions = el('div', { class: 'form-row', style: 'margin-top:12px;gap:8px;' });
  const genBtn = el('button', {
    type: 'button',
    class: 'btn-primary',
    onclick: generate,
    disabled: state[filesKey].length === 0,
  }, cfg.buttonLabel);
  actions.appendChild(genBtn);
  if (state[filesKey].length || state[htmlKey]) {
    actions.appendChild(el('button', {
      type: 'button',
      class: 'btn-secondary',
      onclick: () => {
        state[filesKey] = [];
        state[htmlKey] = '';
        state[filenameKey] = '';
        state[errorKey] = '';
        renderAll();
      },
    }, 'Clear'));
  }
  card.appendChild(actions);

  if (state[errorKey]) {
    card.appendChild(el('div', { class: 'status-msg error', style: 'margin-top:12px;' }, state[errorKey]));
  }

  wrap.appendChild(card);

  // Result cards — render inline so the PM can select-and-copy directly, or
  // click "Copy for Google Docs". The Pipeline Summary widget is split out of
  // the generated HTML into its own preview window below the main notes, so it
  // can be read and copied separately.
  if (state[htmlKey]) {
    const tmp = el('div');
    tmp.innerHTML = extractBodyHtml(state[htmlKey]);

    // Detach the Pipeline Summary (and its preceding divider / empty wrapper).
    let summaryHtml = '';
    const summaryNode = tmp.querySelector('.pipeline-summary');
    if (summaryNode) {
      summaryHtml = summaryNode.outerHTML;
      let top = summaryNode;
      if (top.parentElement && top.parentElement !== tmp && top.parentElement.tagName === 'SECTION') {
        top = top.parentElement;
      }
      const prev = top.previousElementSibling;
      top.remove();
      if (prev && prev.tagName === 'HR') prev.remove();
    }

    const makePreviewCard = ({ title, innerHtml, hint, sourceHtml }) => {
      const resCard = el('div', { class: 'card', style: 'margin-top:16px;' });
      const head = el('div', { class: 'form-row', style: 'align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;' });
      head.appendChild(el('h3', { style: 'margin:0;' }, title));
      const headActions = el('div', { class: 'form-row', style: 'gap:8px;' });

      const previewBody = el('div', { class: 'tracker-report-preview' });
      previewBody.innerHTML = innerHtml;

      headActions.appendChild(el('button', {
        type: 'button',
        class: 'btn-primary',
        onclick: async () => {
          const ok = await copyHtmlAsRichText(previewBody);
          flashStatus(ok ? 'Copied — paste into Google Docs' : 'Copy failed', ok ? 'success' : 'error');
        },
      }, 'Copy for Google Docs'));
      if (sourceHtml) {
        headActions.appendChild(el('button', {
          type: 'button',
          class: 'btn-secondary',
          onclick: async () => {
            const ok = await copyToClipboard(sourceHtml);
            flashStatus(ok ? 'HTML source copied' : 'Copy failed', ok ? 'success' : 'error');
          },
        }, 'Copy HTML source'));
      }
      head.appendChild(headActions);
      resCard.appendChild(head);

      if (hint) {
        resCard.appendChild(el('p', { class: 'muted', style: 'margin:8px 0 12px 0;font-size:13px;' }, hint));
      }
      resCard.appendChild(previewBody);
      return resCard;
    };

    wrap.appendChild(makePreviewCard({
      title: 'Preview',
      innerHtml: tmp.innerHTML,
      hint: 'Select the content below and copy directly, or click "Copy for Google Docs" to grab the formatted version.',
      sourceHtml: state[htmlKey],
    }));

    if (summaryHtml) {
      wrap.appendChild(makePreviewCard({
        title: '📊 Pipeline Summary',
        innerHtml: summaryHtml,
        hint: 'Cross-client batch & ad counts — copy this into the doc separately.',
        sourceHtml: '',
      }));
    }
  }

  async function handleFiles(fileList) {
    if (!fileList || fileList.length === 0) return;
    const rejected = [];
    for (const file of fileList) {
      const ext = file.name.toLowerCase().split('.').pop();
      if (ext !== 'xlsx' && ext !== 'xlsm') {
        rejected.push(`${file.name} (not .xlsx)`);
        continue;
      }
      if (state[filesKey].some(f => f.name === file.name && f.size === file.size)) continue;
      try {
        const workbook = await readWorkbookFromFile(file);
        state[filesKey].push({ name: file.name, size: file.size, workbook });
      } catch (e) {
        rejected.push(`${file.name} (parse error: ${e.message})`);
      }
    }
    state[errorKey] = rejected.length ? `Skipped: ${rejected.join(', ')}` : '';
    state[htmlKey] = '';
    renderAll();
  }

  function generate() {
    if (!state[filesKey].length) return;
    try {
      const workbooksWithNames = state[filesKey].map(f => ({ workbook: f.workbook, filename: f.name }));
      state[htmlKey] = cfg.buildHtml(workbooksWithNames);
      state[filenameKey] = cfg.suggestedFilename();
      state[errorKey] = '';
      flashStatus(`${cfg.title} generated.`);
    } catch (e) {
      console.error(e);
      state[errorKey] = `Generation failed: ${e.message}`;
      state[htmlKey] = '';
    }
    renderAll();
  }
}

function renderMeetingNotes() {
  renderTrackerReportTab({
    stateKey: 'meeting',
    title: '📝 Meeting Notes',
    intro: 'Drop one or more Creative Tracker .xlsx files. Generates an HTML report with three sections per client — Launched this week, Testing pipeline, Production pipeline.',
    buttonLabel: 'Generate Meeting Notes',
    buildHtml: buildMeetingNotesHtml,
    suggestedFilename: suggestedMeetingFilename,
  });
}

function renderWeeklyUpdates() {
  renderTrackerReportTab({
    stateKey: 'weekly',
    title: '📅 Weekly Creative Updates',
    intro: 'Drop one or more Creative Tracker .xlsx files. Generates an HTML report with two sections per client — Ready to Launch and In Production.',
    buttonLabel: 'Generate Weekly Updates',
    buildHtml: buildWeeklyUpdatesHtml,
    suggestedFilename: suggestedWeeklyFilename,
  });
}

function fmtBytes(n) {
  if (n > 1024 * 1024) return (n / 1024 / 1024).toFixed(1) + ' MB';
  if (n > 1024) return (n / 1024).toFixed(1) + ' KB';
  return n + ' B';
}

/**
 * Pull the inner HTML of <body> from a full HTML document string so we can
 * render it inline in the page without dragging in the document-level <style>
 * (which would bleed into the rest of the app).
 */
function extractBodyHtml(fullHtml) {
  const m = fullHtml.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  return m ? m[1] : fullHtml;
}

/**
 * Copy a DOM element's rendered HTML to the clipboard as rich text, so
 * pasting into Google Docs preserves bullets, links, headings, etc.
 * Falls back to plain HTML text if the rich-clipboard API is unavailable.
 */
async function copyHtmlAsRichText(node) {
  const html = node.innerHTML;
  const plain = node.innerText || node.textContent || '';
  try {
    if (typeof ClipboardItem !== 'undefined' && navigator.clipboard?.write) {
      const item = new ClipboardItem({
        'text/html': new Blob([html], { type: 'text/html' }),
        'text/plain': new Blob([plain], { type: 'text/plain' }),
      });
      await navigator.clipboard.write([item]);
      return true;
    }
  } catch (e) {
    console.warn('Rich clipboard write failed, trying execCommand fallback:', e);
  }
  // Fallback: select the node and execCommand('copy'). Works in older browsers
  // and Safari when the page is the active document.
  try {
    const range = document.createRange();
    range.selectNodeContents(node);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    const ok = document.execCommand('copy');
    sel.removeAllRanges();
    return ok;
  } catch (e) {
    console.warn('execCommand copy fallback failed:', e);
    return false;
  }
}

// -------- Render --------

function renderAll() {
  renderTabs();
  if (state.tab === 'ad-categorizer') {
    renderAdCategorizer();
  } else if (state.tab === 'tracker-to-brief') {
    renderTrackerToBrief();
  } else if (state.tab === 'meeting-notes') {
    renderMeetingNotes();
  } else if (state.tab === 'weekly-updates') {
    renderWeeklyUpdates();
  } else {
    renderBriefToTracker();
  }
}

// -------- Init --------

async function init() {
  loadState();
  await loadClients();
  renderAll();
}

document.addEventListener('DOMContentLoaded', init);
