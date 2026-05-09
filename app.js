// Form state, UI rendering, and generate orchestration.

import { TEMPLATES, DROPDOWN_OPTIONS } from './lib/templates-config.js';
import { generateBrief } from './lib/docx-filler.js';

const MAX_CREATIVES = 5;
const STORAGE_KEY = 'pbg.formState.v1';
const CLIENTS_STORAGE_KEY = 'pbg.localClients.v1';
const PRESETS_STORAGE_KEY = 'pbg.presets.v1';

// -------- App state --------
const state = {
  briefType: 'static',          // 'static' | 'video' | 'copy'
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
  forms: {
    static: { overview: {}, creatives: [{}], activePresetId: null },
    video:  { overview: {}, creatives: [{}], activePresetId: null },
    copy:   { overview: {}, creatives: [{}], activePresetId: null },
  },
};

// -------- Persistence --------

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({
      briefType: state.briefType,
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

function applyClientDefaults() {
  // Pre-fill empty overview fields with the current client's defaults
  // for the current brief type. Doesn't overwrite existing user values.
  const client = getCurrentClient();
  if (!client || !client.defaults) return;
  const config = TEMPLATES[state.briefType];
  const ov = state.forms[state.briefType].overview;
  for (const def of config.overview) {
    if (client.defaults[def.field] != null && (ov[def.field] == null || ov[def.field] === '')) {
      ov[def.field] = client.defaults[def.field];
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

  clientRow.appendChild(el('button', {
    type: 'button',
    class: 'btn-secondary btn-small',
    onclick: handleAddClient,
  }, '+ Add client'));

  // Tracker URL edit (used by the PM app — kept here so a copywriter who
  // adds a new client can also set its tracker URL once)
  const currentClient = getCurrentClient();
  if (currentClient) {
    clientRow.appendChild(el('button', {
      type: 'button',
      class: 'btn-secondary btn-small',
      onclick: () => handleEditTrackerUrl(currentClient),
      title: 'Set or edit this client\'s Creative Tracker URL (used by PM Tools)',
    }, currentClient.tracker_url ? 'Tracker URL ✓' : '+ Tracker URL'));
  }
  wrap.appendChild(clientRow);

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
    title: 'Save current overview field values as a preset for this client',
  }, '+ Save preset'));

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
  saveLocalClients();
  saveState();
  renderClientPicker();
  renderForm();
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
        applyClientDefaults();
        saveState();
        renderForm();
      },
    }, cfg.label);
    wrap.appendChild(btn);
  }
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
  renderOverview();
  renderCreatives();
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

async function handleGenerate() {
  const errors = validate();
  const status = document.getElementById('status');
  status.className = '';
  status.textContent = '';

  if (errors.length > 0) {
    status.className = 'error';
    status.textContent = errors.join(' • ');
    return;
  }

  const btn = document.getElementById('generate-btn');
  btn.disabled = true;
  btn.textContent = 'Generating…';

  try {
    const config = TEMPLATES[state.briefType];
    const client = getCurrentClient();
    const ov = state.forms[state.briefType].overview;
    const creatives = state.forms[state.briefType].creatives;

    // Fetch the template
    const res = await fetch(config.file);
    if (!res.ok) throw new Error(`Failed to load template: ${res.status}`);
    const templateBuffer = await res.arrayBuffer();

    // Generate
    const blob = await generateBrief({
      briefType: state.briefType,
      clientName: client.name,
      overview: ov,
      creatives,
      templateBuffer,
    });

    // Filename: {Client}_{BriefType}_{Idea Name}_{YYYY-MM-DD}.docx
    const safe = (s) => String(s || '').replace(/[^\w\s-]/g, '').trim().replace(/\s+/g, '_').slice(0, 60);
    const today = new Date().toISOString().slice(0, 10);
    const filename = [
      safe(client.name) || 'Client',
      config.label.replace(/\s+/g, ''),
      safe(ov['Idea Name']) || 'Brief',
      today,
    ].join('_') + '.docx';

    saveAs(blob, filename);

    status.className = 'success';
    status.textContent = `✓ Generated ${filename}`;
  } catch (err) {
    console.error(err);
    status.className = 'error';
    status.textContent = `Error: ${err.message}`;
  } finally {
    btn.disabled = false;
    btn.textContent = 'Generate Brief';
  }
}

function handleClear() {
  if (!confirm(`Clear all fields for the current ${TEMPLATES[state.briefType].label} brief?`)) return;
  state.forms[state.briefType] = { overview: {}, creatives: [{}] };
  applyClientDefaults();
  saveState();
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
}

document.addEventListener('DOMContentLoaded', init);
