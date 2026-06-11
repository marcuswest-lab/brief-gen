// Form state, UI rendering, and generate orchestration.

import { TEMPLATES, DROPDOWN_OPTIONS } from './lib/templates-config.js';
import { generateBrief } from './lib/docx-filler.js';
import { loadBriefs, upsertBrief, deleteBrief, getBrief, sortBriefsRecent, briefDisplayLabel, buildBriefFilename } from './lib/brief-store.js';
import { parseClaudeOutput } from './lib/claude-output-parser.js';
import {
  docxBlobToHtml,
  wrapHtmlForPaste,
  copyHtmlToClipboard,
  downloadHtmlFile,
  docxFilenameToHtml,
} from './lib/html-exporter.js';

const MAX_CREATIVES = 10;
const STORAGE_KEY = 'pbg.formState.v1';
const CLIENTS_STORAGE_KEY = 'pbg.localClients.v1';
const PRESETS_STORAGE_KEY = 'pbg.presets.v1';

// -------- App state --------
const state = {
  briefType: 'static',          // 'static' | 'video' | 'copy'
  quickFillOpen: false,         // when true, shows the Quick Fill paste panel above the form
  quickFillText: '',
  clientId: null,
  clients: [],                  // loaded from clients.json + localStorage
  // Per-client overview-field presets. Shape:
  //   { [clientId]: [ { id, name, values: { fieldName: value, ... } }, ... ] }
  // Presets are global per client (not per brief type) — only fields present in
  // the current brief's overview are applied; unrelated fields are ignored.
  presets: {},
  // Form data is keyed by briefType so switching tabs preserves work.
  // `activePresetId` is the preset that was last applied for this brief type +
  // current client. Cleared when the user picks the placeholder option.
  // `loadedBriefId` is non-null when the current form is editing a saved
  // brief — Generate updates that record instead of creating a new one.
  forms: {
    static: { overview: {}, creatives: [{}], activePresetId: null, loadedBriefId: null },
    video:  { overview: {}, creatives: [{}], activePresetId: null, loadedBriefId: null },
    copy:   { overview: {}, creatives: [{}], activePresetId: null, loadedBriefId: null },
  },
};

// -------- Persistence --------

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      briefType: state.briefType,
      quickFillOpen: state.quickFillOpen,
      quickFillText: state.quickFillText,
      clientId: state.clientId,
      forms: state.forms,
    }));
  } catch (e) {
    console.warn('Autosave failed:', e);
  }
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (parsed.briefType) state.briefType = parsed.briefType;
    if (parsed.clientId) state.clientId = parsed.clientId;
    if (parsed.quickFillOpen !== undefined) state.quickFillOpen = parsed.quickFillOpen;
    if (parsed.quickFillText) state.quickFillText = parsed.quickFillText;
    if (parsed.forms) {
      for (const k of ['static', 'video', 'copy']) {
        if (parsed.forms[k]) state.forms[k] = parsed.forms[k];
        // Make sure each form has at least one creative
        if (!state.forms[k].creatives || state.forms[k].creatives.length === 0) {
          state.forms[k].creatives = [{}];
        }
        // Backfill activePresetId for old saved state
        if (state.forms[k].activePresetId === undefined) {
          state.forms[k].activePresetId = null;
        }
        if (state.forms[k].loadedBriefId === undefined) {
          state.forms[k].loadedBriefId = null;
        }
      }
    }
  } catch (e) {
    console.warn('Loading saved state failed:', e);
  }
}

async function loadClients() {
  // Built-in clients (clients.json)
  let builtin = [];
  try {
    const res = await fetch('clients.json');
    if (res.ok) {
      const data = await res.json();
      builtin = data.clients || [];
    }
  } catch (e) {
    console.warn('clients.json load failed:', e);
  }

  // User-added clients (localStorage)
  let local = [];
  try {
    const raw = localStorage.getItem(CLIENTS_STORAGE_KEY);
    if (raw) local = JSON.parse(raw) || [];
  } catch (e) {
    console.warn('Local clients load failed:', e);
  }

  state.clients = [...builtin, ...local];
  if (!state.clientId && state.clients.length > 0) {
    state.clientId = state.clients[0].id;
  }
}

function saveLocalClients() {
  // Save only the clients that aren't from clients.json (we tag local ones with __local: true)
  const local = state.clients.filter(c => c.__local);
  localStorage.setItem(CLIENTS_STORAGE_KEY, JSON.stringify(local));
}

// -------- Presets --------

function loadPresets() {
  try {
    const raw = localStorage.getItem(PRESETS_STORAGE_KEY);
    if (raw) state.presets = JSON.parse(raw) || {};
  } catch (e) {
    console.warn('Preset load failed:', e);
    state.presets = {};
  }
}

function savePresets() {
  try {
    localStorage.setItem(PRESETS_STORAGE_KEY, JSON.stringify(state.presets));
  } catch (e) {
    console.warn('Preset save failed:', e);
  }
}

function getPresetsForCurrentClient() {
  if (!state.clientId) return [];
  return state.presets[state.clientId] || [];
}

function makePresetId() {
  return 'p_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 7);
}

// -------- Helpers --------

function getCurrentClient() {
  return state.clients.find(c => c.id === state.clientId) || null;
}

// A value is "placeholder" if it's an unfilled bracket token like "[insert link]".
// These should be treated as empty so auto-fill can overwrite them.
function isPlaceholder(v) {
  if (v == null) return false;
  const t = String(v).trim().toLowerCase();
  return /^\[.*\]$/.test(t) || t.includes('insert link');
}

// A field is fillable (auto-fill may set it) when it's empty or a placeholder.
function isFillable(v) {
  return v == null || v === '' || isPlaceholder(v);
}

function applyClientDefaults() {
  // Pre-fill overview fields. Two layers:
  //   1. Template-level defaults (e.g. video Ratio Format(s) → "9:16 (Story/Reel)")
  //   2. Client-level defaults from clients.json
  // Empty/placeholder fields get filled. Defs marked `force` always overwrite
  // (e.g. ratio formats are fixed company standards).
  const config = TEMPLATES[state.briefType];
  const ov = state.forms[state.briefType].overview;

  // Template defaults
  for (const def of config.overview) {
    if (def.default == null) continue;
    if (def.force || isFillable(ov[def.field])) {
      ov[def.field] = def.default;
    }
  }

  // Client defaults
  const client = getCurrentClient();
  if (client && client.defaults) {
    for (const def of config.overview) {
      if (client.defaults[def.field] != null && isFillable(ov[def.field])) {
        ov[def.field] = client.defaults[def.field];
      }
    }
  }
}

// Fields that are client-specific and should re-fill when the client changes.
const CLIENT_AUTOFILL_FIELDS = ['Link to Brand Assets', 'Landing Page URL'];

// All known default values any client defines for a field (used to detect
// auto-filled vs. manually-edited values).
function allClientDefaultValues(field) {
  const vals = new Set();
  for (const c of state.clients) {
    const v = c.defaults && c.defaults[field];
    if (v != null && v !== '') vals.add(v);
  }
  return vals;
}

// When the client changes, overwrite client-specific fields with the new
// client's value — but only if the current value looks auto-filled (empty or
// equal to some client's known default). Never clobber a value the user typed.
function applyClientSwitch() {
  const client = getCurrentClient();
  for (const briefType of Object.keys(state.forms)) {
    const config = TEMPLATES[briefType];
    const ov = state.forms[briefType].overview;
    for (const field of CLIENT_AUTOFILL_FIELDS) {
      // Skip fields this brief type doesn't have.
      if (!config.overview.some(d => d.field === field)) continue;
      const cur = ov[field];
      const knownDefaults = allClientDefaultValues(field);
      const isAutoFilled = isFillable(cur) || knownDefaults.has(cur);
      if (!isAutoFilled) continue; // user edited this — leave it alone
      const next = client && client.defaults ? client.defaults[field] : undefined;
      ov[field] = next != null ? next : '';
    }
  }
}

// -------- Rendering --------

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

function renderClientPicker() {
  const wrap = document.getElementById('client-picker');
  wrap.innerHTML = '';

  // -- Client row --
  const clientRow = el('div', { class: 'picker-row' });
  clientRow.appendChild(el('label', { for: 'client-select' }, 'Client'));

  const select = el('select', {
    id: 'client-select',
    onchange: (e) => {
      state.clientId = e.target.value;
      applyClientSwitch();
      applyClientDefaults();
      saveState();
      renderClientPicker();
      renderForm();
    },
  });
  for (const c of state.clients) {
    const opt = el('option', { value: c.id }, c.name);
    if (c.id === state.clientId) opt.setAttribute('selected', '');
    select.appendChild(opt);
  }
  clientRow.appendChild(select);

  // My Briefs button — opens the saved-briefs modal
  const briefCount = loadBriefs().length;
  clientRow.appendChild(el('button', {
    type: 'button',
    class: 'btn-secondary btn-small',
    onclick: handleOpenMyBriefs,
    title: 'View, load, edit, or regenerate previously saved briefs',
  }, briefCount > 0 ? `My Briefs (${briefCount})` : 'My Briefs'));

  clientRow.appendChild(el('button', {
    type: 'button',
    class: 'btn-secondary btn-small',
    onclick: handleAddClient,
  }, '+ Add client'));

  // Manage button — only show when there are user-added clients to clean up
  const localClients = state.clients.filter(c => c.__local);
  if (localClients.length > 0) {
    clientRow.appendChild(el('button', {
      type: 'button',
      class: 'btn-secondary btn-small',
      onclick: handleManageClients,
      title: 'Delete user-added clients (e.g. duplicates of built-in clients)',
    }, `Manage (${localClients.length})`));
  }

  // Tracker controls: when the client has a URL, show a one-click "Open
  // Tracker" link button + a small ✏️ to edit. Otherwise just show "+ Tracker URL".
  const currentClient = getCurrentClient();
  if (currentClient) {
    if (currentClient.tracker_url) {
      clientRow.appendChild(el('a', {
        href: currentClient.tracker_url,
        target: '_blank',
        rel: 'noopener',
        class: 'btn-primary btn-small tracker-open-btn',
        title: `Open ${currentClient.name} Creative Tracker in a new tab`,
      }, 'Open Tracker ↗'));
      clientRow.appendChild(el('button', {
        type: 'button',
        class: 'btn-secondary btn-small tracker-edit-btn',
        onclick: () => handleEditTrackerUrl(currentClient),
        title: 'Edit Creative Tracker URL',
      }, '✏️'));
    } else {
      clientRow.appendChild(el('button', {
        type: 'button',
        class: 'btn-secondary btn-small',
        onclick: () => handleEditTrackerUrl(currentClient),
        title: 'Set this client\'s Creative Tracker URL',
      }, '+ Tracker URL'));
    }
  }
  wrap.appendChild(clientRow);

  // Show "editing" indicator if a brief is loaded
  const loadedBriefId = state.forms[state.briefType].loadedBriefId;
  if (loadedBriefId) {
    const loadedBrief = getBrief(loadedBriefId);
    if (loadedBrief) {
      const editingRow = el('div', { class: 'picker-row editing-row' });
      editingRow.appendChild(el('span', { class: 'editing-tag' },
        `Editing: ${briefDisplayLabel(loadedBrief)}`,
      ));
      editingRow.appendChild(el('button', {
        type: 'button',
        class: 'btn-secondary btn-small',
        onclick: () => {
          state.forms[state.briefType].loadedBriefId = null;
          saveState();
          renderClientPicker();
        },
        title: 'Stop editing this saved brief (next Generate creates a new record)',
      }, 'New brief'));
      wrap.appendChild(editingRow);
    }
  }

  // -- Preset row --
  const presets = getPresetsForCurrentClient();
  const activePresetId = state.forms[state.briefType].activePresetId;
  // If the previously-active preset no longer exists (e.g. user switched
  // clients or deleted it), clear it so the placeholder shows.
  const activeStillValid = activePresetId && presets.some(p => p.id === activePresetId);
  if (activePresetId && !activeStillValid) {
    state.forms[state.briefType].activePresetId = null;
  }

  const presetRow = el('div', { class: 'picker-row' });
  presetRow.appendChild(el('label', { for: 'preset-select' }, 'Preset'));

  const presetSelect = el('select', {
    id: 'preset-select',
    onchange: (e) => {
      const id = e.target.value;
      if (!id) {
        // User picked the placeholder — clear the active preset (does NOT
        // unfill any fields, just stops claiming a preset is active).
        state.forms[state.briefType].activePresetId = null;
        saveState();
        return;
      }
      handleApplyPreset(id);
    },
  });
  presetSelect.appendChild(el('option', { value: '' }, presets.length === 0 ? '— no presets —' : '— none —'));
  for (const p of presets) {
    const opt = el('option', { value: p.id }, p.name);
    if (p.id === activePresetId && activeStillValid) opt.setAttribute('selected', '');
    presetSelect.appendChild(opt);
  }
  if (presets.length === 0) presetSelect.setAttribute('disabled', '');
  presetRow.appendChild(presetSelect);

  presetRow.appendChild(el('button', {
    type: 'button',
    class: 'btn-secondary btn-small',
    onclick: handleSavePreset,
    title: 'Snapshot the current overview field values as a NEW preset for this client',
  }, '+ Save new preset'));

  // Update button — only when a preset is currently active
  if (activeStillValid) {
    const activePreset = presets.find(p => p.id === activePresetId);
    presetRow.appendChild(el('button', {
      type: 'button',
      class: 'btn-secondary btn-small',
      onclick: () => handleUpdatePreset(activePreset),
      title: `Overwrite "${activePreset.name}" with the current overview field values`,
    }, 'Update preset'));
  }

  if (presets.length > 0) {
    presetRow.appendChild(el('button', {
      type: 'button',
      class: 'btn-secondary btn-small',
      onclick: handleManagePresets,
      title: 'Rename or delete presets for this client',
    }, 'Manage'));
  }

  wrap.appendChild(presetRow);
}

function handleAddClient() {
  const name = prompt('Client name (e.g. "TCC", "CEO Lawyer"):');
  if (!name || !name.trim()) return;
  const trackerUrl = prompt('Creative Tracker Google Sheet URL (optional — leave blank to skip):') || '';
  const id = name.trim().toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') + '-' + Date.now().toString(36);
  const newClient = { id, name: name.trim(), defaults: {}, __local: true };
  if (trackerUrl.trim()) newClient.tracker_url = trackerUrl.trim();
  state.clients.push(newClient);
  state.clientId = id;
  applyClientSwitch();
  applyClientDefaults();
  saveLocalClients();
  saveState();
  renderClientPicker();
  renderForm();
}

function handleManageClients() {
  const root = document.getElementById('modal-root');
  root.innerHTML = '';

  const close = () => {
    root.innerHTML = '';
    document.removeEventListener('keydown', onKey);
  };
  const onKey = (e) => { if (e.key === 'Escape') close(); };
  document.addEventListener('keydown', onKey);

  const renderList = () => {
    listEl.innerHTML = '';
    const builtinNames = new Set(state.clients.filter(c => !c.__local).map(c => c.name.toLowerCase()));
    const localClients = state.clients.filter(c => c.__local);

    if (localClients.length === 0) {
      listEl.appendChild(el('div', { class: 'pm-empty' },
        'No user-added clients. The list comes entirely from clients.json.'
      ));
      return;
    }

    for (const client of localClients) {
      const isDupe = builtinNames.has(client.name.toLowerCase());
      const item = el('div', { class: 'pm-item briefs-item' + (isDupe ? ' briefs-item-dupe' : '') });

      const label = el('div', { class: 'briefs-label' });
      label.appendChild(el('div', { class: 'briefs-label-main' },
        client.name,
        isDupe ? el('span', { class: 'briefs-current-tag dupe-tag' }, 'duplicate') : null,
      ));
      label.appendChild(el('div', { class: 'briefs-label-sub' },
        client.tracker_url ? `tracker: ${client.tracker_url.slice(0, 60)}…` : 'no tracker URL',
      ));
      item.appendChild(label);

      item.appendChild(el('button', {
        type: 'button',
        class: 'btn-secondary btn-small pm-delete',
        onclick: () => {
          if (!confirm(`Delete user-added client "${client.name}"?\n\nIf there's a built-in client with the same name in clients.json, it will still be available.`)) return;
          // Remove from state
          const idx = state.clients.findIndex(c => c.id === client.id);
          if (idx >= 0) state.clients.splice(idx, 1);
          // If this was the selected client, fall back to the first available
          if (state.clientId === client.id) {
            state.clientId = state.clients[0]?.id || null;
            applyClientSwitch();
            applyClientDefaults();
          }
          saveLocalClients();
          saveState();
          renderClientPicker();
          renderForm();
          renderList();
        },
      }, 'Delete'));

      listEl.appendChild(item);
    }

    if (localClients.some(c => builtinNames.has(c.name.toLowerCase()))) {
      const removeAllDupesBtn = el('button', {
        type: 'button',
        class: 'btn-secondary btn-small',
        style: 'margin-top: 12px;',
        onclick: () => {
          const dupes = localClients.filter(c => builtinNames.has(c.name.toLowerCase()));
          if (!confirm(`Delete all ${dupes.length} duplicate(s)?\n\n${dupes.map(c => '• ' + c.name).join('\n')}`)) return;
          for (const dupe of dupes) {
            const idx = state.clients.findIndex(c => c.id === dupe.id);
            if (idx >= 0) state.clients.splice(idx, 1);
            if (state.clientId === dupe.id) {
              state.clientId = state.clients[0]?.id || null;
              applyClientSwitch();
              applyClientDefaults();
            }
          }
          saveLocalClients();
          saveState();
          renderClientPicker();
          renderForm();
          renderList();
        },
      }, 'Remove all duplicates');
      listEl.appendChild(removeAllDupesBtn);
    }
  };

  const overlay = el('div', { class: 'modal-overlay', onclick: (e) => { if (e.target === overlay) close(); } });
  const dialog = el('div', { class: 'modal-dialog briefs-dialog' });

  dialog.appendChild(el('div', { class: 'modal-header' },
    el('h2', {}, 'Manage clients'),
    el('button', { type: 'button', class: 'modal-close', onclick: close, title: 'Close' }, '×'),
  ));

  const body = el('div', { class: 'modal-body' });
  body.appendChild(el('p', { class: 'modal-hint' },
    'Built-in clients (from clients.json) are not shown here — they can\'t be deleted from the app. Only user-added clients (saved to your browser) are listed below.'
  ));
  const listEl = el('div', { class: 'pm-list' });
  body.appendChild(listEl);
  dialog.appendChild(body);

  dialog.appendChild(el('div', { class: 'modal-footer' },
    el('button', { type: 'button', class: 'btn-primary', onclick: close }, 'Done'),
  ));

  overlay.appendChild(dialog);
  root.appendChild(overlay);
  renderList();
}

function handleEditTrackerUrl(client) {
  const current = client.tracker_url || '';
  const next = prompt(`Creative Tracker URL for ${client.name}:\n\n(Leave blank to clear.)`, current);
  if (next === null) return; // user cancelled
  if (next.trim()) {
    client.tracker_url = next.trim();
  } else {
    delete client.tracker_url;
  }
  // Persist: built-in clients are overridden in the local clients store using a
  // per-client URL override map so we don't lose customizations on page reload.
  if (client.__local) {
    saveLocalClients();
  } else {
    // Built-in client: stash the URL override separately
    saveTrackerUrlOverride(client.id, client.tracker_url);
  }
  renderClientPicker();
}

const TRACKER_URL_OVERRIDES_KEY = 'pbg.trackerUrlOverrides.v1';

function loadTrackerUrlOverrides() {
  try {
    const raw = localStorage.getItem(TRACKER_URL_OVERRIDES_KEY);
    if (!raw) return {};
    return JSON.parse(raw) || {};
  } catch (e) {
    return {};
  }
}

function saveTrackerUrlOverride(clientId, url) {
  const overrides = loadTrackerUrlOverrides();
  if (url) overrides[clientId] = url;
  else delete overrides[clientId];
  localStorage.setItem(TRACKER_URL_OVERRIDES_KEY, JSON.stringify(overrides));
}

function applyTrackerUrlOverrides() {
  const overrides = loadTrackerUrlOverrides();
  for (const c of state.clients) {
    if (overrides[c.id]) c.tracker_url = overrides[c.id];
  }
}

// -------- Preset handlers --------

function handleSavePreset() {
  const client = getCurrentClient();
  if (!client) {
    alert('Select a client first.');
    return;
  }
  // Snapshot ALL non-empty overview fields from the current brief type
  const overview = state.forms[state.briefType].overview;
  const values = {};
  for (const [k, v] of Object.entries(overview)) {
    if (v != null && String(v).trim() !== '') values[k] = v;
  }
  if (Object.keys(values).length === 0) {
    alert('No overview fields are filled in. Fill in some fields first, then save them as a preset.');
    return;
  }

  const name = prompt(`Preset name (e.g. "Long Distance Lander", "Local Movers"):\n\nWill save these fields:\n${Object.keys(values).join(', ')}`);
  if (!name || !name.trim()) return;

  if (!state.presets[client.id]) state.presets[client.id] = [];

  // Replace existing preset of same name, or append new
  const existing = state.presets[client.id].findIndex(p => p.name.toLowerCase() === name.trim().toLowerCase());
  const preset = { id: makePresetId(), name: name.trim(), values };
  if (existing >= 0) {
    if (!confirm(`A preset named "${name.trim()}" already exists. Overwrite it?`)) return;
    preset.id = state.presets[client.id][existing].id;
    state.presets[client.id][existing] = preset;
  } else {
    state.presets[client.id].push(preset);
  }
  savePresets();
  renderClientPicker();
}

function handleUpdatePreset(preset) {
  const client = getCurrentClient();
  if (!client || !preset) return;

  // Snapshot the current overview fields
  const overview = state.forms[state.briefType].overview;
  const values = {};
  for (const [k, v] of Object.entries(overview)) {
    if (v != null && String(v).trim() !== '') values[k] = v;
  }
  if (Object.keys(values).length === 0) {
    alert('No overview fields are filled in. Nothing to save into the preset.');
    return;
  }

  if (!confirm(`Overwrite preset "${preset.name}" with the current overview field values?\n\nFields to save:\n${Object.keys(values).join(', ')}`)) {
    return;
  }

  // Find + replace by id
  const list = state.presets[client.id] || [];
  const idx = list.findIndex(p => p.id === preset.id);
  if (idx >= 0) {
    list[idx] = { ...preset, values };
    savePresets();
    renderClientPicker();

    const status = document.getElementById('status');
    if (status) {
      status.className = 'success';
      status.textContent = `Updated preset "${preset.name}" — ${Object.keys(values).length} field${Object.keys(values).length === 1 ? '' : 's'} saved.`;
      setTimeout(() => {
        if (status.textContent.startsWith('Updated preset')) {
          status.textContent = '';
          status.className = '';
        }
      }, 4000);
    }
  }
}

function handleApplyPreset(presetId) {
  const client = getCurrentClient();
  if (!client) return;
  const preset = (state.presets[client.id] || []).find(p => p.id === presetId);
  if (!preset) return;

  const config = TEMPLATES[state.briefType];
  const ov = state.forms[state.briefType].overview;
  const fieldNames = new Set(config.overview.map(d => d.field));

  // Identify fields that would be overwritten (currently filled with a different value)
  const wouldOverwrite = [];
  for (const [k, v] of Object.entries(preset.values)) {
    if (!fieldNames.has(k)) continue; // not relevant to this brief type
    const current = ov[k];
    if (current != null && String(current).trim() !== '' && current !== v) {
      wouldOverwrite.push(k);
    }
  }

  if (wouldOverwrite.length > 0) {
    const ok = confirm(
      `Applying preset "${preset.name}" will overwrite these fields:\n\n` +
      wouldOverwrite.map(f => `• ${f}`).join('\n') +
      `\n\nContinue?`
    );
    if (!ok) {
      // Cancel — re-render so the dropdown reflects the still-active preset
      renderClientPicker();
      return;
    }
  }

  // Apply: set every preset field that's relevant to this brief type
  let applied = 0;
  for (const [k, v] of Object.entries(preset.values)) {
    if (!fieldNames.has(k)) continue;
    ov[k] = v;
    applied++;
  }

  // Mark this preset as the active selection for the current brief type
  state.forms[state.briefType].activePresetId = preset.id;

  saveState();
  renderClientPicker();
  renderForm();

  const status = document.getElementById('status');
  if (status) {
    status.className = 'success';
    status.textContent = `Applied preset "${preset.name}" — ${applied} field${applied === 1 ? '' : 's'} updated.`;
    setTimeout(() => { if (status.textContent.startsWith('Applied preset')) { status.textContent = ''; status.className = ''; } }, 4000);
  }
}

function handleManagePresets() {
  const client = getCurrentClient();
  if (!client) return;
  if (!state.presets[client.id]) state.presets[client.id] = [];
  openPresetManager(client);
}

// -------- Preset manager modal --------

function openPresetManager(client) {
  const root = document.getElementById('modal-root');
  root.innerHTML = '';

  const close = () => {
    root.innerHTML = '';
    document.removeEventListener('keydown', onKey);
  };
  const onKey = (e) => { if (e.key === 'Escape') close(); };
  document.addEventListener('keydown', onKey);

  const renderList = () => {
    listEl.innerHTML = '';
    const presets = state.presets[client.id] || [];

    if (presets.length === 0) {
      listEl.appendChild(el('div', { class: 'pm-empty' },
        'No presets yet. Use "+ Save preset" to create one from the current overview fields.'
      ));
      return;
    }

    for (const preset of presets) {
      const fieldList = Object.keys(preset.values);
      const item = el('div', { class: 'pm-item' });

      // Name
      const nameRow = el('div', { class: 'pm-row' });
      const nameInput = el('input', {
        type: 'text',
        value: preset.name,
        class: 'pm-name-input',
        oninput: (e) => {
          preset.name = e.target.value;
          savePresets();
          renderClientPicker();
        },
      });
      nameInput.value = preset.name;
      nameRow.appendChild(nameInput);

      const deleteBtn = el('button', {
        type: 'button',
        class: 'btn-secondary btn-small pm-delete',
        title: 'Delete preset',
        onclick: () => {
          if (!confirm(`Delete preset "${preset.name}"?`)) return;
          const presets = state.presets[client.id];
          const idx = presets.findIndex(p => p.id === preset.id);
          if (idx >= 0) presets.splice(idx, 1);
          // Clear activePresetId if it pointed to this preset
          for (const k of ['static', 'video', 'copy']) {
            if (state.forms[k].activePresetId === preset.id) {
              state.forms[k].activePresetId = null;
            }
          }
          savePresets();
          saveState();
          renderClientPicker();
          renderList();
        },
      }, 'Delete');
      nameRow.appendChild(deleteBtn);
      item.appendChild(nameRow);

      // Field chips
      const fieldsBox = el('div', { class: 'pm-fields' });
      for (const fname of fieldList) {
        const chip = el('span', { class: 'pm-chip', title: `Click × to remove ${fname} from this preset` });
        chip.appendChild(el('span', { class: 'pm-chip-name' }, fname));
        chip.appendChild(el('span', { class: 'pm-chip-val' }, ': ' + truncate(String(preset.values[fname]), 40)));
        chip.appendChild(el('button', {
          type: 'button',
          class: 'pm-chip-x',
          title: 'Remove this field from preset',
          onclick: () => {
            delete preset.values[fname];
            // If preset has no fields left, remove it entirely
            if (Object.keys(preset.values).length === 0) {
              const presets = state.presets[client.id];
              const idx = presets.findIndex(p => p.id === preset.id);
              if (idx >= 0) presets.splice(idx, 1);
              for (const k of ['static', 'video', 'copy']) {
                if (state.forms[k].activePresetId === preset.id) {
                  state.forms[k].activePresetId = null;
                }
              }
              saveState();
            }
            savePresets();
            renderClientPicker();
            renderList();
          },
        }, '×'));
        fieldsBox.appendChild(chip);
      }
      item.appendChild(fieldsBox);

      listEl.appendChild(item);
    }
  };

  const overlay = el('div', { class: 'modal-overlay', onclick: (e) => { if (e.target === overlay) close(); } });
  const dialog = el('div', { class: 'modal-dialog' });

  // Header
  const header = el('div', { class: 'modal-header' },
    el('h2', {}, `Presets for ${client.name}`),
    el('button', { type: 'button', class: 'modal-close', onclick: close, title: 'Close' }, '×'),
  );
  dialog.appendChild(header);

  // Body
  const body = el('div', { class: 'modal-body' });
  body.appendChild(el('p', { class: 'modal-hint' },
    'Edit names inline. Delete a preset, or remove individual fields with their × button.'
  ));
  const listEl = el('div', { class: 'pm-list' });
  body.appendChild(listEl);
  dialog.appendChild(body);

  // Footer
  const footer = el('div', { class: 'modal-footer' },
    el('button', { type: 'button', class: 'btn-primary', onclick: close }, 'Done'),
  );
  dialog.appendChild(footer);

  overlay.appendChild(dialog);
  root.appendChild(overlay);
  renderList();
}

function truncate(s, n) {
  return s.length > n ? s.slice(0, n - 1) + '…' : s;
}

function renderTabs() {
  const wrap = document.getElementById('brief-tabs');
  wrap.innerHTML = '';
  for (const [key, cfg] of Object.entries(TEMPLATES)) {
    const btn = el('button', {
      type: 'button',
      class: 'tab-btn' + (state.briefType === key ? ' active' : ''),
      onclick: () => {
        state.briefType = key;
        state.quickFillOpen = false;
        applyClientDefaults();
        saveState();
        renderForm();
      },
    }, cfg.label);
    wrap.appendChild(btn);
  }
  // Quick Fill tab — opens an inline panel above the form
  const qfBtn = el('button', {
    type: 'button',
    class: 'tab-btn tab-quick-fill' + (state.quickFillOpen ? ' active' : ''),
    onclick: () => {
      state.quickFillOpen = !state.quickFillOpen;
      saveState();
      renderForm();
    },
    title: 'Paste Claude/AI output and have it auto-fill the form',
  }, '⚡ Quick Fill');
  wrap.appendChild(qfBtn);
}

function renderField(def, value, onChange) {
  const id = `field-${def.field.replace(/[^a-z0-9]+/gi, '-')}-${Math.random().toString(36).slice(2, 7)}`;
  const label = el('label', { for: id }, def.field, def.required ? el('span', { class: 'req' }, ' *') : null);

  let input;
  if (def.kind === 'dropdown') {
    const opts = DROPDOWN_OPTIONS[def.options] || [];
    input = el('select', { id, onchange: (e) => onChange(e.target.value) },
      el('option', { value: '' }, '— select —'),
      ...opts.map(o => {
        const opt = el('option', { value: o }, o);
        if (o === value) opt.setAttribute('selected', '');
        return opt;
      }),
    );
  } else if (def.kind === 'multiselect') {
    // Stored value is a comma-joined string for backward compat with text fields
    // and so it serializes cleanly into the .docx cell.
    const opts = DROPDOWN_OPTIONS[def.options] || [];
    const selected = new Set(
      (value || '').split(',').map(s => s.trim()).filter(Boolean)
    );
    input = el('div', { id, class: 'multiselect' });
    for (const o of opts) {
      const cbId = `${id}-${o.replace(/[^a-z0-9]+/gi, '-')}`;
      const cb = el('input', {
        type: 'checkbox',
        id: cbId,
        value: o,
        onchange: (e) => {
          if (e.target.checked) selected.add(o);
          else selected.delete(o);
          // Preserve original option order
          const ordered = opts.filter(x => selected.has(x));
          onChange(ordered.join(', '));
        },
      });
      if (selected.has(o)) cb.setAttribute('checked', '');
      input.appendChild(el('label', { class: 'multiselect-opt', for: cbId }, cb, el('span', {}, o)));
    }
  } else if (def.multiline) {
    input = el('textarea', {
      id,
      rows: 3,
      oninput: (e) => onChange(e.target.value),
    });
    input.value = value || '';
  } else {
    input = el('input', {
      id,
      type: 'text',
      oninput: (e) => onChange(e.target.value),
    });
    input.value = value || '';
  }

  return el('div', { class: 'field' + (def.required ? ' required' : '') }, label, input);
}

function renderOverview() {
  const wrap = document.getElementById('overview-section');
  wrap.innerHTML = '';
  const config = TEMPLATES[state.briefType];
  const data = state.forms[state.briefType].overview;

  wrap.appendChild(el('h2', {}, 'Overview'));

  const grid = el('div', { class: 'fields-grid' });
  for (const def of config.overview) {
    grid.appendChild(renderField(def, data[def.field], (v) => {
      data[def.field] = v;
      saveState();
    }));
  }
  wrap.appendChild(grid);
}

function renderCreatives() {
  const wrap = document.getElementById('creatives-section');
  wrap.innerHTML = '';
  const config = TEMPLATES[state.briefType];
  const creatives = state.forms[state.briefType].creatives;

  const header = el('div', { class: 'section-header' },
    el('h2', {}, 'Creatives'),
    el('span', { class: 'count' }, `${creatives.length} of ${MAX_CREATIVES}`),
  );
  wrap.appendChild(header);

  creatives.forEach((cData, idx) => {
    const block = el('div', { class: 'creative-block' });

    const blockHeader = el('div', { class: 'creative-header' });
    blockHeader.appendChild(el('h3', {}, `Creative ${idx + 1}`));

    if (creatives.length > 1) {
      blockHeader.appendChild(el('button', {
        type: 'button',
        class: 'btn-remove',
        title: 'Remove this creative',
        onclick: () => {
          creatives.splice(idx, 1);
          saveState();
          renderCreatives();
        },
      }, '×'));
    }
    block.appendChild(blockHeader);

    const grid = el('div', { class: 'fields-grid' });
    for (const def of config.creative) {
      grid.appendChild(renderField(def, cData[def.field], (v) => {
        cData[def.field] = v;
        saveState();
      }));
    }
    block.appendChild(grid);
    wrap.appendChild(block);
  });

  if (creatives.length < MAX_CREATIVES) {
    wrap.appendChild(el('button', {
      type: 'button',
      class: 'btn-secondary',
      onclick: () => {
        creatives.push({});
        saveState();
        renderCreatives();
      },
    }, '+ Add Creative'));
  }
}

function renderForm() {
  renderTabs();
  renderQuickFillPanel();
  renderOverview();
  renderCreatives();
}

function renderQuickFillPanel() {
  // Render into a container right above the overview section. Create the
  // container if it doesn't exist yet.
  let wrap = document.getElementById('quick-fill-section');
  if (!wrap) {
    wrap = document.createElement('section');
    wrap.id = 'quick-fill-section';
    const overviewSection = document.getElementById('overview-section');
    overviewSection.parentNode.insertBefore(wrap, overviewSection);
  }
  wrap.innerHTML = '';
  if (!state.quickFillOpen) {
    wrap.style.display = 'none';
    return;
  }
  wrap.style.display = '';
  wrap.className = 'card quick-fill-card';

  wrap.appendChild(el('h2', {}, '⚡ Quick Fill'));
  wrap.appendChild(el('p', { class: 'qf-help' },
    'Paste a fully-written brief from Claude (or anywhere). The parser detects field labels like "Idea Name:", "Lead Type:", "Creative 1", etc., auto-detects whether it\'s a Static / Video / Body Copy brief, and fills the form below.'
  ));

  const ta = el('textarea', {
    id: 'quick-fill-text',
    class: 'qf-textarea',
    rows: 14,
    placeholder: 'Paste Claude\'s brief output here…\n\nExample:\nIdea Name: Quiet Overwhelm\nAngle Name: Cost Anchor\nTask: ...\n\nCreative 1\nFile Name: Quiet Overwhelm | SV8_Gmail\nVariation Type: Copy\nLead Type: Offer\nAwareness Level: Most Aware\nStatus: Ready For Internal\nCopy: ...',
    oninput: (e) => {
      state.quickFillText = e.target.value;
      saveState();
      // Live-preview the detected brief type without re-rendering the form
      updateQuickFillDetectedLine();
    },
  });
  ta.value = state.quickFillText;
  wrap.appendChild(ta);

  const detectedLine = el('div', { id: 'qf-detected-line', class: 'qf-detected' });
  wrap.appendChild(detectedLine);

  const actions = el('div', { class: 'qf-actions' });
  actions.appendChild(el('button', {
    type: 'button',
    class: 'btn-secondary',
    onclick: () => {
      state.quickFillText = '';
      saveState();
      renderQuickFillPanel();
    },
  }, 'Clear'));
  actions.appendChild(el('button', {
    type: 'button',
    class: 'btn-primary',
    onclick: handleQuickFillParse,
  }, 'Parse & fill form'));
  wrap.appendChild(actions);

  // Initial detection line
  updateQuickFillDetectedLine();
}

function updateQuickFillDetectedLine() {
  const lineEl = document.getElementById('qf-detected-line');
  if (!lineEl) return;
  lineEl.innerHTML = '';
  if (!state.quickFillText.trim()) {
    lineEl.className = 'qf-detected qf-detected-empty';
    lineEl.textContent = 'Paste text to see what gets detected.';
    return;
  }
  try {
    const parsed = parseClaudeOutput(state.quickFillText);
    const overviewCount = Object.keys(parsed.overview).length;
    const creativeCount = parsed.creatives.length;
    const typeLabel = parsed.briefType ? TEMPLATES[parsed.briefType].label : 'unknown';
    lineEl.className = 'qf-detected';
    lineEl.textContent = `Detected: ${typeLabel} brief · ${overviewCount} overview field${overviewCount === 1 ? '' : 's'} · ${creativeCount} creative${creativeCount === 1 ? '' : 's'}`;
  } catch (e) {
    lineEl.className = 'qf-detected qf-detected-error';
    lineEl.textContent = 'Parse error: ' + e.message;
  }
}

function handleQuickFillParse() {
  if (!state.quickFillText.trim()) {
    alert('Paste some brief text first.');
    return;
  }

  let parsed;
  try {
    parsed = parseClaudeOutput(state.quickFillText);
  } catch (e) {
    alert('Parse error: ' + e.message);
    return;
  }

  if (!parsed.briefType) {
    if (!confirm("Couldn't detect brief type from the text. Apply to the currently-active brief type tab?")) return;
    parsed.briefType = state.briefType;
  }

  if (parsed.creatives.length === 0 && Object.keys(parsed.overview).length === 0) {
    alert('No recognized fields found in the pasted text. Make sure each field looks like "Idea Name: ..." on its own line.');
    return;
  }

  // Check if target form already has data; confirm overwrite
  const targetForm = state.forms[parsed.briefType];
  const existingOverviewFilled = Object.values(targetForm.overview).filter(v => v != null && String(v).trim() !== '').length;
  const existingCreativesFilled = targetForm.creatives.filter(c => Object.values(c).some(v => v != null && String(v).trim() !== '')).length;
  if (existingOverviewFilled > 0 || existingCreativesFilled > 0) {
    const ok = confirm(
      `This will replace your current ${TEMPLATES[parsed.briefType].label} brief:\n\n` +
      `- ${existingOverviewFilled} overview field${existingOverviewFilled === 1 ? '' : 's'} filled\n` +
      `- ${existingCreativesFilled} creative${existingCreativesFilled === 1 ? '' : 's'} filled\n\n` +
      `Continue?`
    );
    if (!ok) return;
  }

  // Flag if the pasted Landing Page URL differs from the selected client's
  // pre-loaded default — confirm before changing it.
  const client = getCurrentClient();
  const clientLP = client && client.defaults ? client.defaults['Landing Page URL'] : null;
  const parsedLP = parsed.overview['Landing Page URL'];
  if (clientLP && parsedLP && !isPlaceholder(parsedLP) &&
      String(parsedLP).trim() !== String(clientLP).trim()) {
    const useParsed = confirm(
      `The pasted brief has a different Landing Page URL than ${client.name}'s default:\n\n` +
      `Pre-loaded:  ${clientLP}\n` +
      `Pasted:      ${parsedLP}\n\n` +
      `OK = use the pasted URL\nCancel = keep ${client.name}'s pre-loaded URL`
    );
    if (!useParsed) parsed.overview['Landing Page URL'] = clientLP;
  }

  // If there are dropdown-value warnings, open the review modal first.
  // Otherwise, apply directly.
  if (parsed.warnings && parsed.warnings.length > 0) {
    openQuickFillReviewModal(parsed);
  } else {
    applyParsedBrief(parsed);
  }
}

function applyParsedBrief(parsed) {
  state.briefType = parsed.briefType;
  state.forms[parsed.briefType] = {
    overview: { ...parsed.overview },
    creatives: parsed.creatives.length > 0 ? parsed.creatives : [{}],
    activePresetId: null,
    loadedBriefId: null,
  };
  applyClientDefaults();
  state.quickFillOpen = false;
  saveState();
  renderClientPicker();
  renderForm();

  const status = document.getElementById('status');
  if (status) {
    const ovCount = Object.keys(parsed.overview).length;
    const crCount = parsed.creatives.length;
    status.className = 'success';
    status.textContent = `✓ Filled ${TEMPLATES[parsed.briefType].label} brief: ${ovCount} overview field${ovCount === 1 ? '' : 's'}, ${crCount} creative${crCount === 1 ? '' : 's'}.`;
    setTimeout(() => {
      if (status.textContent.startsWith('✓ Filled')) {
        status.textContent = '';
        status.className = '';
      }
    }, 5000);
  }
}

/**
 * Show a modal listing every dropdown value that didn't match a valid
 * option. For each warning, the user picks the corrected value (suggested
 * fuzzy match pre-selected) or chooses "Leave blank" to skip the field.
 * On Apply, the warnings are resolved into the parsed object and the form
 * is filled.
 */
function openQuickFillReviewModal(parsed) {
  const root = document.getElementById('modal-root');
  root.innerHTML = '';

  const close = () => {
    root.innerHTML = '';
    document.removeEventListener('keydown', onKey);
  };
  const onKey = (e) => { if (e.key === 'Escape') close(); };
  document.addEventListener('keydown', onKey);

  // Decisions: { warningIndex: 'suggestion-value' | '' (leave blank) | 'keep-as-is' (keep original) }
  // Default: pick the suggestion if present; otherwise leave blank.
  const decisions = parsed.warnings.map(w => w.suggestion || '');

  const overlay = el('div', { class: 'modal-overlay', onclick: (e) => { if (e.target === overlay) close(); } });
  const dialog = el('div', { class: 'modal-dialog quick-fill-review-dialog' });

  dialog.appendChild(el('div', { class: 'modal-header' },
    el('h2', {}, `Review ${parsed.warnings.length} dropdown value${parsed.warnings.length === 1 ? '' : 's'}`),
    el('button', { type: 'button', class: 'modal-close', onclick: close, title: 'Cancel' }, '×'),
  ));

  const body = el('div', { class: 'modal-body' });
  body.appendChild(el('p', { class: 'modal-hint' },
    `Some pasted values don't match valid dropdown options. Pick the correct option for each (suggested closest match is pre-selected), or choose "Leave blank" to clear the field.`
  ));

  const list = el('div', { class: 'qfr-list' });
  parsed.warnings.forEach((w, i) => {
    const item = el('div', { class: 'qfr-item' });

    // Title row: field + scope
    const scopeLabel = w.scope === 'creative' ? `Creative ${w.creativeIndex + 1}` : 'Overview';
    item.appendChild(el('div', { class: 'qfr-title' },
      el('span', { class: 'qfr-scope-tag' }, scopeLabel),
      el('strong', {}, w.field),
    ));

    // Original value
    item.appendChild(el('div', { class: 'qfr-original' },
      'Pasted value: ',
      el('code', {}, w.originalValue),
    ));

    // Picker
    const sel = el('select', {
      class: 'qfr-select',
      onchange: (e) => { decisions[i] = e.target.value; },
    });
    sel.appendChild(el('option', { value: '' }, '— leave blank —'));
    for (const opt of w.options) {
      const optEl = el('option', { value: opt }, opt + (opt === w.suggestion ? ' (suggested)' : ''));
      if (opt === decisions[i]) optEl.setAttribute('selected', '');
      sel.appendChild(optEl);
    }
    item.appendChild(sel);

    list.appendChild(item);
  });
  body.appendChild(list);
  dialog.appendChild(body);

  dialog.appendChild(el('div', { class: 'modal-footer' },
    el('button', { type: 'button', class: 'btn-secondary', onclick: close }, 'Cancel'),
    el('button', { type: 'button', class: 'btn-primary', onclick: () => {
      // Resolve each warning into the parsed object
      parsed.warnings.forEach((w, i) => {
        const picked = decisions[i] || '';  // empty = clear field
        if (w.scope === 'overview') {
          if (picked) parsed.overview[w.field] = picked;
          else delete parsed.overview[w.field];
        } else {
          const c = parsed.creatives[w.creativeIndex];
          if (c) {
            if (picked) c[w.field] = picked;
            else delete c[w.field];
          }
        }
      });
      close();
      applyParsedBrief(parsed);
    }}, 'Apply corrections & fill form'),
  ));

  overlay.appendChild(dialog);
  root.appendChild(overlay);
}

// -------- Validation + Generate --------

function validate() {
  const config = TEMPLATES[state.briefType];
  const ov = state.forms[state.briefType].overview;
  const creatives = state.forms[state.briefType].creatives;
  const errors = [];

  for (const def of config.overview) {
    if (def.required && !ov[def.field]) {
      errors.push(`Overview: ${def.field} is required`);
    }
  }
  creatives.forEach((c, i) => {
    for (const def of config.creative) {
      if (def.required && !c[def.field]) {
        errors.push(`Creative ${i + 1}: ${def.field} is required`);
      }
    }
  });

  if (!getCurrentClient()) errors.push('Please select a client');
  return errors;
}

/**
 * Validate + build a filled docx Blob from the current form state. Shared by
 * `handleGenerate` (downloads the blob) and the HTML export handlers (convert
 * the blob via mammoth before delivering to the user).
 *
 * Returns null if validation fails. Errors are written to the #status div so
 * the user sees the same feedback regardless of which button they clicked.
 *
 * @returns {Promise<null | {blob: Blob, filename: string, overview: object, creatives: object[], briefType: string, client: object}>}
 */
async function buildBriefBlob() {
  const status = document.getElementById('status');
  status.className = '';
  status.textContent = '';

  const errors = validate();
  if (errors.length > 0) {
    status.className = 'error';
    status.textContent = errors.join(' • ');
    return null;
  }

  const config = TEMPLATES[state.briefType];
  const client = getCurrentClient();
  const ov = state.forms[state.briefType].overview;
  const creatives = state.forms[state.briefType].creatives;

  const res = await fetch(config.file);
  if (!res.ok) throw new Error(`Failed to load template: ${res.status}`);
  const templateBuffer = await res.arrayBuffer();

  const blob = await generateBrief({
    briefType: state.briefType,
    clientName: client.name,
    overview: ov,
    creatives,
    templateBuffer,
  });

  const filename = buildBriefFilename({
    clientName: client.name,
    briefType: state.briefType,
    ideaName: ov['Idea Name'],
  });

  return { blob, filename, overview: ov, creatives, briefType: state.briefType, client };
}

/**
 * Persist a built brief into the shared store. If we were editing a loaded
 * brief, update it in place; otherwise create a new record. Shared by
 * `handleGenerate` and `handlePreviewHtml` so both persist to "My Briefs".
 */
function saveBuiltBrief(built) {
  const { overview: ov, creatives, client } = built;
  const form = state.forms[state.briefType];
  const stored = upsertBrief({
    id: form.loadedBriefId || undefined,
    clientId: client.id,
    clientName: client.name,
    briefType: state.briefType,
    ideaName: ov['Idea Name'] || '',
    angleName: ov['Angle Name'] || '',
    overview: { ...ov },
    creatives: creatives.map(c => ({ ...c })),
  });
  form.loadedBriefId = stored.id;
  saveState();
  renderClientPicker();
  return stored;
}

async function handleGenerate() {
  const status = document.getElementById('status');
  const btn = document.getElementById('generate-btn');
  btn.disabled = true;
  btn.textContent = 'Generating…';

  try {
    const built = await buildBriefBlob();
    if (!built) return; // validation failure already surfaced in #status

    const { blob, filename } = built;

    saveAs(blob, filename);
    saveBuiltBrief(built);

    status.className = 'success';
    status.textContent = `✓ Generated ${filename} (saved to My Briefs)`;
  } catch (err) {
    console.error(err);
    status.className = 'error';
    status.textContent = `Error: ${err.message}`;
  } finally {
    btn.disabled = false;
    btn.textContent = 'Generate Brief';
  }
}

async function handlePreviewHtml() {
  const status = document.getElementById('status');
  const btn = document.getElementById('preview-html-btn');
  btn.disabled = true;
  btn.textContent = 'Generating preview…';

  try {
    const built = await buildBriefBlob();
    if (!built) return;

    const { blob, filename } = built;
    const inner = await docxBlobToHtml(blob);
    const fullHtml = wrapHtmlForPaste(inner, { title: filename });

    // Save the brief just like Generate does, so previewing also persists it.
    saveBuiltBrief(built);

    openHtmlPreviewModal({ fullHtml, filename });

    status.className = 'success';
    status.textContent = '✓ Saved to My Briefs';
    setTimeout(() => {
      if (status.textContent === '✓ Saved to My Briefs') {
        status.className = '';
        status.textContent = '';
      }
    }, 5000);
  } catch (err) {
    console.error(err);
    status.className = 'error';
    status.textContent = `Error: ${err.message}`;
  } finally {
    btn.disabled = false;
    btn.textContent = 'Preview HTML';
  }
}

/**
 * Render the HTML preview modal. Uses the existing .modal-* CSS classes.
 * The preview itself lives in a sandboxed iframe so the inline styles in the
 * generated HTML cannot leak into the app's chrome.
 */
function openHtmlPreviewModal({ fullHtml, filename }) {
  const root = document.getElementById('modal-root');
  root.innerHTML = '';

  const close = () => {
    root.innerHTML = '';
    document.removeEventListener('keydown', onKey);
  };
  const onKey = (e) => { if (e.key === 'Escape') close(); };
  document.addEventListener('keydown', onKey);

  // Sandboxed iframe so the pasted styles don't escape into the app.
  const iframe = el('iframe', {
    class: 'html-preview-frame',
    sandbox: 'allow-same-origin',
    title: 'HTML preview',
  });

  const copyBtn = el('button', {
    type: 'button',
    class: 'btn-primary btn-small',
    onclick: async () => {
      try {
        await copyHtmlToClipboard(fullHtml);
        copyBtn.textContent = '✓ Copied';
        setTimeout(() => { copyBtn.textContent = 'Copy to Clipboard'; }, 2000);
      } catch (err) {
        console.error(err);
        copyBtn.textContent = 'Copy failed';
        setTimeout(() => { copyBtn.textContent = 'Copy to Clipboard'; }, 2000);
      }
    },
  }, 'Copy to Clipboard');

  const downloadBtn = el('button', {
    type: 'button',
    class: 'btn-secondary btn-small',
    onclick: () => {
      downloadHtmlFile(fullHtml, docxFilenameToHtml(filename));
    },
  }, 'Download .html');

  const closeBtn = el('button', {
    type: 'button',
    class: 'btn-secondary btn-small',
    onclick: close,
  }, 'Close');

  const overlay = el('div', {
    class: 'modal-overlay',
    onclick: (e) => { if (e.target === overlay) close(); },
  },
    el('div', { class: 'modal-dialog html-preview-dialog' },
      el('div', { class: 'modal-header' },
        el('h2', {}, 'HTML Preview'),
        el('button', { class: 'modal-close', type: 'button', onclick: close, 'aria-label': 'Close' }, '×'),
      ),
      el('div', { class: 'modal-body' },
        el('p', { class: 'modal-hint' },
          'This is exactly what will be pasted into Google Docs. Click "Copy to Clipboard" and then Cmd+V into a Google Doc.',
        ),
        iframe,
      ),
      el('div', { class: 'modal-footer html-preview-footer' },
        downloadBtn,
        closeBtn,
        copyBtn,
      ),
    ),
  );

  root.appendChild(overlay);

  // Write the preview HTML after the iframe is in the DOM.
  const doc = iframe.contentDocument || iframe.contentWindow.document;
  doc.open();
  doc.write(fullHtml);
  doc.close();
}

function handleClear() {
  if (!confirm(`Clear all fields for the current ${TEMPLATES[state.briefType].label} brief?`)) return;
  state.forms[state.briefType] = { overview: {}, creatives: [{}], activePresetId: null, loadedBriefId: null };
  applyClientDefaults();
  saveState();
  renderClientPicker();
  renderForm();
}

// -------- My Briefs --------

function handleOpenMyBriefs() {
  const root = document.getElementById('modal-root');
  root.innerHTML = '';

  const close = () => {
    root.innerHTML = '';
    document.removeEventListener('keydown', onKey);
  };
  const onKey = (e) => { if (e.key === 'Escape') close(); };
  document.addEventListener('keydown', onKey);

  const renderList = () => {
    listEl.innerHTML = '';
    const briefs = sortBriefsRecent(loadBriefs());

    if (briefs.length === 0) {
      listEl.appendChild(el('div', { class: 'pm-empty' },
        'No saved briefs yet. Generate a brief and it will appear here.'
      ));
      return;
    }

    for (const brief of briefs) {
      const isCurrent = state.forms[brief.briefType]?.loadedBriefId === brief.id;
      const item = el('div', { class: 'pm-item briefs-item' + (isCurrent ? ' briefs-item-current' : '') });

      const label = el('div', { class: 'briefs-label' });
      label.appendChild(el('div', { class: 'briefs-label-main' },
        briefDisplayLabel(brief),
        isCurrent ? el('span', { class: 'briefs-current-tag' }, 'editing') : null,
      ));
      label.appendChild(el('div', { class: 'briefs-label-sub' },
        `${brief.creatives.length} creative${brief.creatives.length === 1 ? '' : 's'} · last edited ${new Date(brief.updatedAt || brief.createdAt).toLocaleString()}`,
      ));
      item.appendChild(label);

      const actions = el('div', { class: 'briefs-actions' });
      actions.appendChild(el('button', {
        type: 'button',
        class: 'btn-primary btn-small',
        onclick: () => {
          handleLoadBrief(brief);
          close();
        },
      }, isCurrent ? 'Editing' : 'Load'));

      actions.appendChild(el('button', {
        type: 'button',
        class: 'btn-secondary btn-small',
        onclick: async () => {
          handleLoadBrief(brief);
          close();
          // Trigger Generate after a tick so the form has rendered
          await new Promise(r => setTimeout(r, 50));
          handleGenerate();
        },
      }, 'Regenerate'));

      actions.appendChild(el('button', {
        type: 'button',
        class: 'btn-secondary btn-small pm-delete',
        onclick: () => {
          if (!confirm(`Delete "${brief.ideaName || '(untitled)'}"? This can't be undone.`)) return;
          deleteBrief(brief.id);
          // Clear loadedBriefId from any form that pointed here
          for (const k of ['static', 'video', 'copy']) {
            if (state.forms[k].loadedBriefId === brief.id) state.forms[k].loadedBriefId = null;
          }
          saveState();
          renderList();
        },
      }, 'Delete'));
      item.appendChild(actions);

      listEl.appendChild(item);
    }
  };

  const overlay = el('div', { class: 'modal-overlay', onclick: (e) => { if (e.target === overlay) close(); } });
  const dialog = el('div', { class: 'modal-dialog briefs-dialog' });

  dialog.appendChild(el('div', { class: 'modal-header' },
    el('h2', {}, 'My Briefs'),
    el('button', { type: 'button', class: 'modal-close', onclick: close, title: 'Close' }, '×'),
  ));

  const body = el('div', { class: 'modal-body' });
  body.appendChild(el('p', { class: 'modal-hint' },
    'Every brief you generate is saved here. Click Load to edit it; Regenerate to load + immediately download a fresh .docx (useful after filling in tracker-generated names via the PM app).'
  ));
  const listEl = el('div', { class: 'pm-list' });
  body.appendChild(listEl);
  dialog.appendChild(body);

  dialog.appendChild(el('div', { class: 'modal-footer' },
    el('button', { type: 'button', class: 'btn-primary', onclick: close }, 'Done'),
  ));

  overlay.appendChild(dialog);
  root.appendChild(overlay);
  renderList();
}

function handleLoadBrief(brief) {
  // Switch to the brief's type, set the client, hydrate form with stored data
  state.briefType = brief.briefType;
  // Try to switch client; if the client no longer exists, keep current
  if (state.clients.some(c => c.id === brief.clientId)) {
    state.clientId = brief.clientId;
  }
  state.forms[brief.briefType] = {
    overview: { ...brief.overview },
    creatives: brief.creatives.length > 0 ? brief.creatives.map(c => ({ ...c })) : [{}],
    activePresetId: null,
    loadedBriefId: brief.id,
  };
  saveState();
  renderClientPicker();
  renderForm();
}

// -------- Init --------

async function init() {
  loadState();
  loadPresets();
  await loadClients();
  applyTrackerUrlOverrides();
  applyClientDefaults();

  renderClientPicker();
  renderForm();

  document.getElementById('generate-btn').addEventListener('click', handleGenerate);
  document.getElementById('clear-btn').addEventListener('click', handleClear);
  document.getElementById('preview-html-btn').addEventListener('click', handlePreviewHtml);
}

document.addEventListener('DOMContentLoaded', init);
