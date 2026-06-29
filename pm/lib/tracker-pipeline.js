// Tracker pipeline reports — pure browser-side .xlsx → HTML.
//
// Two exports:
//   buildMeetingNotesHtml(workbook, filename) — Launched this week + Testing + Production
//   buildWeeklyUpdatesHtml(workbook, filename) — Ready to Launch + In Production
//
// Both read the same Static Creative Tracker / Video Creative Tracker tabs and
// auto-detect columns by header name (row 1). Hyperlinks come from cell.l.Target
// (SheetJS) — make sure XLSX.read was called with { cellHyperlinks: true, cellDates: true }.

const TEMPLATE_IDEAS = ['wife shame', 'educationfailedwomen', "wife won't kiss you", 'heart3heart'];

const SECTIONS = [
  { sheet: 'Static Creative Tracker', label: 'Statics', nameHeader: 'Static Creative Name' },
  { sheet: 'Video Creative Tracker',  label: 'Videos',  nameHeader: 'Video Creative Name' },
];

// -------- Utilities --------

function escapeHtml(s) {
  if (s == null) return '';
  return String(s).replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

function stripQuery(url) {
  if (!url) return url;
  return String(url).split(/[?#]/)[0];
}

function hasValue(v) {
  if (v == null) return false;
  const s = String(v).trim();
  return s !== '' && s !== '0' && s !== 'None' && s !== 'nan' && s !== 'NaN';
}

function isTemplate(name, idea) {
  for (const candidate of [idea, name]) {
    if (!candidate) continue;
    const n = String(candidate).trim().toLowerCase();
    if (!n) continue;
    for (const t of TEMPLATE_IDEAS) {
      if (n.includes(t)) return true;
    }
  }
  return false;
}

function normalizeDate(v) {
  if (v == null) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === 'number') {
    // Excel serial date — SheetJS usually converts to Date with cellDates:true, but guard.
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof v === 'string') {
    const s = v.trim();
    if (!s) return null;
    // Try a few common formats.
    const m1 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m1) return new Date(+m1[1], +m1[2] - 1, +m1[3]);
    const m2 = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})/);
    if (m2) {
      let y = +m2[3];
      if (y < 100) y += 2000;
      return new Date(y, +m2[1] - 1, +m2[2]);
    }
    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function withinLastNDays(dt, n, now) {
  if (!dt) return false;
  now = now || new Date();
  const diffMs = now.getTime() - dt.getTime();
  return diffMs >= 0 && diffMs <= n * 86400 * 1000;
}

// -------- Sheet reading --------

function getCell(sheet, r, c) {
  const ref = XLSX.utils.encode_cell({ r, c });
  return sheet[ref] || null;
}

function getCellValue(sheet, r, c) {
  const cell = getCell(sheet, r, c);
  return cell ? cell.v : null;
}

function getCellHyperlink(sheet, r, c) {
  const cell = getCell(sheet, r, c);
  if (!cell || !cell.l) return null;
  return cell.l.Target || null;
}

function detectColumns(sheet, nameHeader) {
  const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
  if (!range) return {};
  const headers = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const v = getCellValue(sheet, range.s.r, c);
    if (typeof v === 'string') headers[v.trim().toLowerCase()] = c;
  }
  let name = headers[nameHeader.toLowerCase()];
  if (name == null) {
    for (const fb of ['static creative name', 'video creative name', 'creative name', 'copy name']) {
      if (headers[fb] != null) { name = headers[fb]; break; }
    }
  }
  return {
    range,
    name,
    status: headers['status'],
    idea: headers['idea name'],
    launch: headers['launch date'],
    req: headers['request doc'],
    folder: headers['folder link'],
  };
}

function readRows(sheet, cols, kind) {
  if (cols.name == null) return [];
  const out = [];
  for (let r = cols.range.s.r + 1; r <= cols.range.e.r; r++) {
    const nameVal = getCellValue(sheet, r, cols.name);
    if (typeof nameVal !== 'string' || !nameVal.trim()) continue;
    const ideaVal = cols.idea != null ? getCellValue(sheet, r, cols.idea) : null;
    if (isTemplate(nameVal, ideaVal)) continue;
    const statusVal = cols.status != null ? getCellValue(sheet, r, cols.status) : null;
    const launchRaw = cols.launch != null ? getCellValue(sheet, r, cols.launch) : null;
    const launchDt = normalizeDate(launchRaw);

    const reqVal = cols.req != null ? getCellValue(sheet, r, cols.req) : null;
    const folderVal = cols.folder != null ? getCellValue(sheet, r, cols.folder) : null;
    const reqUrl = cols.req != null ? getCellHyperlink(sheet, r, cols.req) : null;
    const folderUrl = cols.folder != null ? getCellHyperlink(sheet, r, cols.folder) : null;

    out.push({
      kind,
      name: nameVal.trim(),
      status: statusVal ? String(statusVal).trim() : '',
      idea: ideaVal ? String(ideaVal).trim() : '',
      launchRaw,
      launchDt,
      reqText: reqVal ? String(reqVal).trim() : '',
      reqUrl,
      folderText: folderVal ? String(folderVal).trim() : '',
      folderUrl,
    });
  }
  return out;
}

function loadRecords(workbook) {
  const all = [];
  for (const { sheet, label, nameHeader } of SECTIONS) {
    if (!workbook.SheetNames.includes(sheet)) continue;
    const ws = workbook.Sheets[sheet];
    const cols = detectColumns(ws, nameHeader);
    all.push(...readRows(ws, cols, label));
  }
  return all;
}

// -------- Client detection from filename --------

export function detectClientName(filename) {
  const base = (filename || '').replace(/\.[^./\\]+$/, '');
  const n = base.toUpperCase();
  if (n.includes('CEO LAWYER')) return 'CEO Lawyer';
  if (n.includes('NURSE COACH')) return 'Nurse Coach';
  if (n.includes('VALUE ADDED MOVING') || /\bVAM\b/.test(n)) return 'VAM';
  if (n.includes('CASE SOURCE')) return 'Case Source';
  if (n.includes('DAN HENRY')) return 'Dan Henry';
  if (n.includes('QUINTESSA')) return 'Quintessa';
  if (/\bTCC\b/.test(n)) return 'TCC';
  const m = base.match(/\[([^\]]+)\]/);
  return m ? m[1].trim() : base;
}

// -------- Classifiers --------

function isNsr(rec) {
  const hay = [rec.name, rec.idea, rec.reqText, rec.folderText]
    .filter(Boolean).join(' ').toUpperCase();
  if (/(^| )NSR( |$|\|)|NRS( |$|\|)/.test(hay)) return true;
  return /^NSR[- |]/.test(rec.name.toUpperCase());
}

function isSpanish(rec) {
  const n = rec.name.replace(/^\s+/, '').toUpperCase();
  return n.startsWith('SP -') || n.startsWith('SP-');
}

// Dan Henry split: ads belong to the MDW, PB, or 5MS funnel (very different
// KPIs — $200 CPA for MDW vs $12 CPR for PB) so the notes show them as
// separate sub-sections (mirrors the VAM / NSR split pattern).
//
// The funnel marker is applied INCONSISTENTLY: PB and 5MS ads reliably carry
// their token as a creative-name prefix, but MDW ads sometimes carry "MDW"
// only in the Folder Link text (e.g. "6/26 | MDW | Teacher's Brain Confession")
// and sometimes omit it entirely (the "Chlamidya Comparision" videos have no
// MDW token anywhere). So 5MS / PB are matched as delimited tokens across the
// name + folder text + idea + request, and everything else falls back to the
// primary MDW funnel — which keeps untagged MDW creatives out of "Other".
function funnelHasToken(rec, token) {
  const hay = [rec.name, rec.folderText, rec.idea, rec.reqText]
    .filter(Boolean).join(' ').toUpperCase();
  return new RegExp(`(^|[\\s|/\\-])${token}([\\s|/\\-]|$)`).test(hay);
}
function isDanHenry5ms(rec) { return funnelHasToken(rec, '5MS'); }
function isDanHenryPb(rec)  { return funnelHasToken(rec, 'PB'); }
function isDanHenryMdw(rec) { return !isDanHenry5ms(rec) && !isDanHenryPb(rec); }

function isLaunched(rec) {
  return rec.launchDt != null || hasValue(rec.launchRaw);
}

function isPipelineStatus(rec) {
  const s = (rec.status || '').toLowerCase();
  return s === '' || s === 'not launched';
}

// As of June 2026 a single Request Doc can hold MANY batches, so the pipeline
// can't be keyed on the Request Doc alone — that collapses distinct batches
// into one line. We combine the Request Doc ("request stack") with the actual
// batch name. The batch name lives in the Folder Link display text
// (e.g. "Qualification Check | Batch 2"); rows that don't have a folder yet
// (production) fall back to the creative-name concept prefix.
function conceptPrefix(name) {
  return String(name || '').split(' - ')[0].trim();
}

function batchName(rec) {
  // Use the Folder Link text as the batch label only when it's a human batch
  // name (e.g. "Qualification Check | Batch 2"). Some clients (CEO Lawyer)
  // paste a raw Dropbox URL into that cell, which is unique per variation and
  // would explode every ad into its own one-off "batch" — fall back to the
  // Idea Name (then concept prefix) so those variations roll up correctly.
  const ft = rec.folderText;
  if (hasValue(ft) && !/^https?:\/\//i.test(ft.trim()) && !/^0+\s*-\s*copy submission/i.test(ft)) {
    return ft;
  }
  return (hasValue(rec.idea) ? rec.idea : null) || conceptPrefix(rec.name) || rec.name;
}

function batchKey(rec) {
  const stack = stripQuery(rec.reqUrl) || rec.reqText || '';
  return `${stack}||${batchName(rec)}`;
}

// -------- Meeting Notes (Launched + Testing + Production) --------

function groupByBatch(items) {
  const batches = new Map();
  for (const it of items) {
    const key = batchKey(it);
    if (!batches.has(key)) batches.set(key, { label: batchName(it), url: it.reqUrl || null, folders: [] });
    const b = batches.get(key);
    if (!b.url && it.reqUrl) b.url = it.reqUrl;
    if (it.folderUrl) {
      const entry = [it.folderText || it.folderUrl, it.folderUrl];
      if (!b.folders.some(([t, u]) => t === entry[0] && u === entry[1])) b.folders.push(entry);
    }
  }
  return batches;
}

function renderBatches(batches, includeFolders) {
  if (!batches || batches.size === 0) return '<li><em>(none currently)</em></li>';
  const out = [];
  let i = 1;
  for (const b of batches.values()) {
    const label = escapeHtml(b.label);
    const heading = b.url
      ? `<strong>Batch ${i} — <a href="${escapeHtml(b.url)}">${label}</a></strong>`
      : `<strong>Batch ${i} — ${label}</strong>`;
    if (includeFolders && b.folders.length) {
      const folderLis = b.folders.map(([t, u]) =>
        `<li>📁 <a href="${escapeHtml(u)}">${escapeHtml(t)}</a></li>`
      ).join('');
      out.push(`<li>${heading}<ul>${folderLis}</ul></li>`);
    } else {
      out.push(`<li>${heading}</li>`);
    }
    i++;
  }
  return out.join('');
}

function collapseSpanish(items, sectionLabel) {
  const sp = items.filter(isSpanish);
  if (sp.length === 0) return [items, null];
  const nonSp = items.filter(it => !isSpanish(it));
  const folders = [];
  const seen = new Set();
  for (const it of sp) {
    if (it.folderUrl && !seen.has(it.folderUrl)) {
      seen.add(it.folderUrl);
      folders.push([it.folderText || it.folderUrl, it.folderUrl]);
    }
  }
  return [nonSp, { label: `Top Spanish ${sectionLabel}`, url: null, folders }];
}

function splitByKind(records) {
  return {
    Statics: records.filter(r => r.kind === 'Statics'),
    Videos: records.filter(r => r.kind === 'Videos'),
  };
}

function renderLaunchedSection(launchedByKind) {
  const parts = ['<p><strong>Launched this week:</strong></p><ul>'];
  for (const label of ['Statics', 'Videos']) {
    const items = launchedByKind[label] || [];
    if (items.length === 0) {
      parts.push(`<li>${label}:<ul><li><em>(none this week)</em></li></ul></li>`);
      continue;
    }

    // Group ads into batches by the combined Request Doc + batch name (the
    // Folder Link display text, falling back to the creative-name concept).
    // Sorted by most recent launch date within the batch so the most recent
    // batch surfaces first.
    const batches = new Map();
    for (const it of items) {
      const key = batchKey(it);
      if (!batches.has(key)) batches.set(key, { label: batchName(it), ads: [] });
      batches.get(key).ads.push(it);
    }

    // Sort batches by the most recent launch date inside each.
    const sortedBatches = [...batches.values()].map(({ label, ads }) => {
      ads.sort((a, b) => (b.launchDt?.getTime() || 0) - (a.launchDt?.getTime() || 0));
      const mostRecent = ads[0]?.launchDt?.getTime() || 0;
      const link = ads.find(a => a.reqUrl)?.reqUrl
        || ads.find(a => a.folderUrl)?.folderUrl
        || '';
      return { label, ads, link, mostRecent };
    }).sort((a, b) => b.mostRecent - a.mostRecent);

    const sub = sortedBatches.map(({ label, ads, link }) => {
      const name = escapeHtml(label);
      const countSuffix = ads.length > 1 ? ` (${ads.length} variations)` : '';
      const launchDt = ads[0]?.launchDt;
      const dateStr = launchDt
        ? ` <span style="color:#888;">(${launchDt.getMonth() + 1}/${launchDt.getDate()})</span>`
        : '';
      return link
        ? `<li><a href="${escapeHtml(link)}">${name}</a>${countSuffix}${dateStr}</li>`
        : `<li>${name}${countSuffix}${dateStr}</li>`;
    }).join('');
    parts.push(`<li>${label}:<ul>${sub}</ul></li>`);
  }
  parts.push('</ul>');
  return parts.join('');
}

function buildPipeline(records, clientName, filter) {
  const pipelines = {
    testing: { Statics: new Map(), Videos: new Map() },
    production: { Statics: new Map(), Videos: new Map() },
  };
  const byKind = splitByKind(records);
  for (const label of ['Statics', 'Videos']) {
    let kindRecs = byKind[label].filter(r => isPipelineStatus(r) && !isLaunched(r));
    if (filter) kindRecs = kindRecs.filter(filter);

    let testing = kindRecs.filter(r => r.folderUrl);
    let production = kindRecs.filter(r => !r.folderUrl && (r.reqUrl || r.reqText));

    let spTesting = null, spProduction = null;
    if (clientName === 'Case Source') {
      [testing, spTesting] = collapseSpanish(testing, label);
      [production, spProduction] = collapseSpanish(production, label);
    }

    const testingBatches = groupByBatch(testing);
    const productionBatches = groupByBatch(production);
    if (spTesting) testingBatches.set(`__sp__${label}`, spTesting);
    if (spProduction) productionBatches.set(`__sp__${label}`, spProduction);
    pipelines.testing[label] = testingBatches;
    pipelines.production[label] = productionBatches;
  }
  return pipelines;
}

function renderPipelines(pipelines) {
  const parts = ['<p><strong>Testing pipeline:</strong></p><ul>'];
  for (const label of ['Statics', 'Videos']) {
    parts.push(`<li>${label}:<ul>${renderBatches(pipelines.testing[label], true)}</ul></li>`);
  }
  parts.push('</ul><p><strong>Production pipeline:</strong></p><ul>');
  for (const label of ['Statics', 'Videos']) {
    parts.push(`<li>${label}:<ul>${renderBatches(pipelines.production[label], false)}</ul></li>`);
  }
  parts.push('</ul>');
  return parts.join('');
}

function buildClientSectionMeeting(clientName, records, now) {
  const launched = records.filter(r => isLaunched(r) && withinLastNDays(r.launchDt, 7, now));
  const parts = [`<section><h2>${escapeHtml(clientName)}</h2>`];

  if (clientName === 'VAM') {
    const groups = [
      ['VAM Funnel', r => !isNsr(r)],
      ['NSR Funnel', isNsr],
    ];
    for (const [subLabel, filt] of groups) {
      const subLaunched = launched.filter(filt);
      const subRecs = records.filter(filt);
      parts.push(`<h3 style="margin-top:18px;color:#444;">${escapeHtml(subLabel)}</h3>`);
      parts.push(renderLaunchedSection(splitByKind(subLaunched)));
      parts.push(renderPipelines(buildPipeline(subRecs, clientName, filt)));
    }
  } else if (clientName === 'Dan Henry') {
    const groups = [
      ['MDW Funnel', isDanHenryMdw],
      ['5MS Funnel', isDanHenry5ms],
      ['PB Funnel',  isDanHenryPb],
    ];
    for (const [subLabel, filt] of groups) {
      const subLaunched = launched.filter(filt);
      const subRecs = records.filter(filt);
      parts.push(`<h3 style="margin-top:18px;color:#444;">${escapeHtml(subLabel)}</h3>`);
      parts.push(renderLaunchedSection(splitByKind(subLaunched)));
      parts.push(renderPipelines(buildPipeline(subRecs, clientName, filt)));
    }
  } else {
    parts.push(renderLaunchedSection(splitByKind(launched)));
    parts.push(renderPipelines(buildPipeline(records, clientName)));
  }
  parts.push('</section>');
  return parts.join('');
}

const PAGE_STYLE_MEETING = `body { font-family: -apple-system, system-ui, "Helvetica Neue", Arial, sans-serif; font-size: 11pt; color: #222; max-width: 980px; margin: 24px auto; padding: 0 16px; }
h1 { font-size: 18pt; margin-bottom: 4px; }
h2 { font-size: 14pt; margin-top: 28px; color: #1a3d7a; border-bottom: 2px solid #e3e6ef; padding-bottom: 4px; }
h3 { font-size: 12pt; }
ul { margin: 4px 0 10px 0; }
li { margin: 3px 0; }
a { color: #1a73e8; text-decoration: none; }
a:hover { text-decoration: underline; }
hr { border: none; border-top: 1px dashed #ccc; margin: 24px 0; }
section p { margin: 8px 0 4px 0; }`;

export function buildMeetingNotesHtml(workbooksWithNames, now) {
  now = now || new Date();
  const today = now.toISOString().slice(0, 10);

  const sections = [];
  for (const { workbook, filename } of workbooksWithNames) {
    try {
      const clientName = detectClientName(filename);
      const records = loadRecords(workbook);
      sections.push(buildClientSectionMeeting(clientName, records, now));
    } catch (e) {
      sections.push(
        `<section><h2>${escapeHtml(detectClientName(filename))}</h2>` +
        `<p style="color:#b00">Could not read tracker: ${escapeHtml(String(e))}</p></section>`
      );
    }
  }

  const body = sections.length ? sections.join('\n<hr/>\n') : '<p><em>No trackers uploaded.</em></p>';
  return `<!doctype html>
<html><head><meta charset="utf-8"/>
<title>Creative Meeting Notes — ${today}</title>
<style>${PAGE_STYLE_MEETING}</style></head>
<body>
<h1>Weekly Creative Meeting Notes</h1>
<p style="color:#666">Generated ${today}</p>
${body}
</body></html>
`;
}

export function suggestedMeetingFilename(now) {
  now = now || new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `creative_meeting_notes_${y}${m}${d}.html`;
}

// -------- Weekly Updates (Ready to Launch + In Production) --------

function stage(rec) {
  const hasFolder = !!(rec.folderUrl || rec.folderText);
  const hasBrief = !!(rec.reqUrl || rec.reqText);
  if (hasFolder && hasBrief) return 'ready';
  if (hasBrief && !hasFolder) return 'production';
  if (hasFolder && !hasBrief) return 'ready';
  return null;
}

function groupByIdea(records) {
  const buckets = new Map();
  for (const r of records) {
    const key = batchKey(r);
    if (!buckets.has(key)) buckets.set(key, { label: batchName(r), items: [] });
    buckets.get(key).items.push(r);
  }
  const out = [];
  for (const { label, items } of buckets.values()) {
    const link = items.find(i => i.folderUrl)?.folderUrl
      || items.find(i => i.reqUrl)?.reqUrl
      || items.find(i => i.folderText)?.folderText
      || items.find(i => i.reqText)?.reqText
      || '';
    out.push({ idea: label, count: items.length, link });
  }
  return out;
}

function renderBatchListWeekly(items) {
  if (!items.length) return '<li><i>(none currently)</i></li>';
  return items.map((it, i) => {
    const idea = escapeHtml(it.idea);
    const suffix = it.count > 1 ? ` (${it.count} variations)` : '';
    return it.link
      ? `<li>Batch ${i + 1} — <a href="${escapeHtml(it.link)}">${idea}</a>${suffix}</li>`
      : `<li>Batch ${i + 1} — ${idea}${suffix}</li>`;
  }).join('\n');
}

function byStageAndKind(records, stageName, kind) {
  return records.filter(r => r.kind === kind && stage(r) === stageName);
}

function renderFunnel(label, records) {
  const out = [];
  if (label) out.push(`<p><b>${escapeHtml(label)}</b></p>`);
  const readyS = groupByIdea(byStageAndKind(records, 'ready', 'Statics'));
  const readyV = groupByIdea(byStageAndKind(records, 'ready', 'Videos'));
  const prodS  = groupByIdea(byStageAndKind(records, 'production', 'Statics'));
  const prodV  = groupByIdea(byStageAndKind(records, 'production', 'Videos'));
  out.push('<p><b>Ready to launch</b></p>');
  out.push('<p><i>Statics:</i></p><ul>'); out.push(renderBatchListWeekly(readyS)); out.push('</ul>');
  out.push('<p><i>Videos:</i></p><ul>');  out.push(renderBatchListWeekly(readyV)); out.push('</ul>');
  out.push('<p><b>In Production</b></p>');
  out.push('<p><i>Statics:</i></p><ul>'); out.push(renderBatchListWeekly(prodS)); out.push('</ul>');
  out.push('<p><i>Videos:</i></p><ul>');  out.push(renderBatchListWeekly(prodV)); out.push('</ul>');
  return out.join('\n');
}

function buildClientSectionWeekly(clientName, records) {
  const pipelineRecs = records.filter(r => isPipelineStatus(r) && !isLaunched(r));
  const parts = [`<p><b>${escapeHtml(clientName)}</b></p>`];
  if (clientName === 'VAM') {
    parts.push(renderFunnel('VAM Funnel', pipelineRecs.filter(r => !isNsr(r))));
    parts.push(renderFunnel('NSR Funnel', pipelineRecs.filter(isNsr)));
  } else if (clientName === 'Dan Henry') {
    // Split by funnel — MDW vs PB. Each has different KPIs and ad inventory.
    parts.push(renderFunnel('MDW Funnel', pipelineRecs.filter(isDanHenryMdw)));
    parts.push(renderFunnel('PB Funnel',  pipelineRecs.filter(isDanHenryPb)));
    parts.push(renderFunnel('5MS Funnel', pipelineRecs.filter(isDanHenry5ms)));
    // Anything that doesn't start with MDW/PB/5MS → fall back to "Other"
    const other = pipelineRecs.filter(r => !isDanHenryMdw(r) && !isDanHenryPb(r) && !isDanHenry5ms(r));
    if (other.length) parts.push(renderFunnel('Other', other));
  } else {
    parts.push(renderFunnel('', pipelineRecs));
  }
  return parts.join('\n');
}

const PAGE_STYLE_WEEKLY = `body { font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.5; padding: 20px; max-width: 800px; }
a { color: #1155cc; }
h1 { font-size: 14pt; }
ul { margin-top: 4px; margin-bottom: 8px; }
hr { border: none; border-top: 1px dashed #ccc; margin: 24px 0; }`;

export function buildWeeklyUpdatesHtml(workbooksWithNames, now) {
  now = now || new Date();
  const todayDisp = `${now.getMonth() + 1}/${now.getDate()}/${String(now.getFullYear()).slice(-2)}`;

  const sections = [];
  for (const { workbook, filename } of workbooksWithNames) {
    try {
      const clientName = detectClientName(filename);
      const records = loadRecords(workbook);
      sections.push(buildClientSectionWeekly(clientName, records));
    } catch (e) {
      sections.push(
        `<p><b>${escapeHtml(detectClientName(filename))}</b></p>` +
        `<p style="color:#b00">Could not read tracker: ${escapeHtml(String(e))}</p>`
      );
    }
  }

  const body = sections.length ? sections.join('\n<hr>\n') : '<p><em>No trackers uploaded.</em></p>';
  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<title>Weekly Creative Updates — ${todayDisp}</title>
<style>${PAGE_STYLE_WEEKLY}</style></head>
<body>
<h1>${todayDisp}</h1>
${body}
</body></html>
`;
}

export function suggestedWeeklyFilename(now) {
  now = now || new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `weekly_creative_notes_${y}${m}${d}.html`;
}

// -------- Workbook loader (shared helper for the tabs) --------

export async function readWorkbookFromFile(file) {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { type: 'array', cellHyperlinks: true, cellDates: true });
}
