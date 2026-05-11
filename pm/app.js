// PM app — two-tab tool for moving data between briefs and the Creative Tracker.

import { parseBrief, detectBriefType } from './lib/brief-parser.js';
import { mapBriefToTracker, blockToTSV } from './lib/tracker-mapper.js';
import { TRACKERS } from './lib/tracker-config.js';
import { loadBriefs, getBrief, upsertBrief, sortBriefsRecent, briefDisplayLabel } from '../lib/brief-store.js';
import { TEMPLATES } from '../lib/templates-config.js';
import { generateBrief } from '../lib/docx-filler.js';

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
  clientId: null,          // selected client (for tracker URL)
  clients: [],             // loaded from clients.json + localStorage
  result: null,            // mapBriefToTracker output
  // Tracker → Brief
  trackerNames: '',
  filledStates: {},        // { [index]: true } UI checkmarks
  trackerToBriefBriefId: null,  // selected brief to bake names into
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
      clientId: state.clientId,
      trackerNames: state.trackerNames,
      filledStates: state.filledStates,
      trackerToBriefBriefId: state.trackerToBriefBriefId,
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
      const sel = el('select', {
        id: 'pm-saved-brief',
        class: 'pm-saved-select',
        onchange: (e) => {
          state.selectedBriefId = e.target.value || null;
          saveState();
        },
      });
      sel.appendChild(el('option', { value: '' }, '— pick a brief —'));
      for (const b of savedBriefs) {
        const opt = el('option', { value: b.id }, briefDisplayLabel(b));
        if (b.id === state.selectedBriefId) opt.setAttribute('selected', '');
        sel.appendChild(opt);
      }
      card.appendChild(sel);
    }
  } else {
    // Paste-text mode (manual upload)
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

    const result = mapBriefToTracker(parsedBrief, {
      requestDoc: state.requestDoc,
      briefTypeOverride: state.briefTypeOverride || undefined,
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

  // Detect client: from selected saved brief if available, else from pasted text
  let detectedClient = null;
  if (state.briefSource === 'saved' && state.selectedBriefId) {
    const stored = getBrief(state.selectedBriefId);
    if (stored) {
      detectedClient = state.clients.find(c => c.id === stored.clientId)
        || state.clients.find(c => c.name === stored.clientName)
        || null;
    }
  }
  if (!detectedClient) {
    detectedClient = detectClientFromBrief(state.briefText);
  }
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
    for (const cell of row) {
      const td = el('td', { title: cell || '' });
      td.textContent = cell == null || cell === '' ? '—' : String(cell);
      if (cell === '' || cell == null) td.className = 'pm-empty-cell';
      tbody.appendChild(tr);
      tr.appendChild(td);
    }
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

    const sel = el('select', {
      id: 'pm-regen-brief',
      class: 'pm-saved-select',
      onchange: (e) => {
        state.trackerToBriefBriefId = e.target.value || null;
        saveState();
      },
    });
    sel.appendChild(el('option', { value: '' }, '— pick the saved brief —'));
    for (const b of savedBriefs) {
      const opt = el('option', { value: b.id },
        `${briefDisplayLabel(b)} — ${b.creatives.length} creative${b.creatives.length === 1 ? '' : 's'}`);
      if (b.id === state.trackerToBriefBriefId) opt.setAttribute('selected', '');
      sel.appendChild(opt);
    }
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

    const safe = (s) => String(s || '').replace(/[^\w\s-]/g, '').trim().replace(/\s+/g, '_').slice(0, 60);
    const today = new Date().toISOString().slice(0, 10);
    const filename = [
      safe(brief.clientName) || 'Client',
      config.label.replace(/\s+/g, ''),
      safe(brief.ideaName) || 'Brief',
      'with-names',
      today,
    ].join('_') + '.docx';

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

// -------- Render --------

function renderAll() {
  renderTabs();
  if (state.tab === 'tracker-to-brief') {
    renderTrackerToBrief();
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
