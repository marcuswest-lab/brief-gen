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
  const isUrl = (t) => /^https?:\/\//i.test(String(t || '').trim());
  for (const b of batches.values()) {
    const label = escapeHtml(b.label);
    // Some clients (CEO Lawyer) paste a raw per-variation URL into the Folder
    // Link cell, so listing every folder spills a stack of identical-looking
    // raw links under one batch. Collapse those URL-valued folders into a
    // single "(N variations)" count linking to the Request Doc (or the first
    // link); human-named folders still list out individually as before.
    const urlFolders = b.folders.filter(([t]) => isUrl(t));
    const namedFolders = b.folders.filter(([t]) => !isUrl(t));
    const countSuffix = urlFolders.length
      ? ` (${urlFolders.length} variation${urlFolders.length > 1 ? 's' : ''})`
      : '';
    const headLink = b.url || (urlFolders[0] && urlFolders[0][1]) || null;
    const heading = headLink
      ? `<strong>Batch ${i} — <a href="${escapeHtml(headLink)}">${label}</a>${countSuffix}</strong>`
      : `<strong>Batch ${i} — ${label}${countSuffix}</strong>`;
    if (includeFolders && namedFolders.length) {
      const folderLis = namedFolders.map(([t, u]) =>
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
  const summaries = [];
  for (const { workbook, filename } of workbooksWithNames) {
    const clientName = detectClientName(filename);
    try {
      const records = loadRecords(workbook);
      sections.push(buildClientSectionMeeting(clientName, records, now));
      summaries.push(...summarizeWorkbook(clientName, records));
    } catch (e) {
      sections.push(
        `<section><h2>${escapeHtml(clientName)}</h2>` +
        `<p style="color:#b00">Could not read tracker: ${escapeHtml(String(e))}</p></section>`
      );
    }
  }

  const body = sections.length ? sections.join('\n<hr/>\n') : '<p><em>No trackers uploaded.</em></p>';
  const summaryWidget = renderSummaryWidget(summaries, today);
  return `<!doctype html>
<html><head><meta charset="utf-8"/>
<title>Creative Meeting Notes — ${today}</title>
<style>${PAGE_STYLE_MEETING}
${SUMMARY_WIDGET_STYLE}</style></head>
<body>
<h1>Weekly Creative Meeting Notes</h1>
<p style="color:#666">Generated ${today}</p>
${body}
${summaryWidget ? `<hr/>\n<section>${summaryWidget}</section>` : ''}
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

function byStageAndKind(records, stageName, kind) {
  return records.filter(r => r.kind === kind && stage(r) === stageName);
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

const DASHBOARD_URLS = {
  'TCC': 'https://marcuswest-lab.github.io/tcc-creative-dashboard/',
  'CEO Lawyer': 'https://marcuswest-lab.github.io/ceolawyer-creative-dashboard/',
  'Nurse Coach': 'https://marcuswest-lab.github.io/nursecoach-creative-dashboard/',
  'VAM': 'https://marcuswest-lab.github.io/vam-creative-dashboard/',
  'Case Source': 'https://marcuswest-lab.github.io/casesource-creative-dashboard/',
  'Dan Henry': 'https://marcuswest-lab.github.io/danhenry-creative-dashboard/',
  'Quintessa': 'https://marcuswest-lab.github.io/quintessa-creative-dashboard/',
};

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
  const summaries = [];
  for (const { workbook, filename } of workbooksWithNames) {
    const clientName = detectClientName(filename);
    try {
      const records = loadRecords(workbook);
      sections.push(buildClientSectionWeekly(clientName, records));
      summaries.push(...summarizeWorkbook(clientName, records));
    } catch (e) {
      sections.push(
        `<p><b>${escapeHtml(clientName)}</b></p>` +
        `<p style="color:#b00">Could not read tracker: ${escapeHtml(String(e))}</p>`
      );
    }
  }

  const body = sections.length ? sections.join('\n<hr>\n') : '<p><em>No trackers uploaded.</em></p>';
  const summaryWidget = renderSummaryWidget(summaries, todayDisp);
  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<title>Weekly Creative Updates — ${todayDisp}</title>
<style>${PAGE_STYLE_WEEKLY}
${SUMMARY_WIDGET_STYLE}</style></head>
<body>
<h1>${todayDisp}</h1>
${body}
${summaryWidget ? `<hr>\n${summaryWidget}` : ''}
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

// -------- Pipeline Summary (cross-client batch + ad counts) --------

// For a set of pipeline records, count batches and ads per stage × kind.
// "batches" = distinct (Request Doc + batch name) groups; "ads" = the sum of
// variations across those batches (uses the same grouping as Weekly Updates so
// the numbers line up with that tab).
function pipelineCounts(records) {
  const calc = (stageName, kind) => {
    const batches = groupByIdea(byStageAndKind(records, stageName, kind));
    return { batches: batches.length, ads: batches.reduce((s, b) => s + b.count, 0) };
  };
  return {
    ready: { Statics: calc('ready', 'Statics'), Videos: calc('ready', 'Videos') },
    production: { Statics: calc('production', 'Statics'), Videos: calc('production', 'Videos') },
  };
}

// Returns one or more summary entries per workbook. VAM is split into its
// VAM Funnel + NSR Funnel sub-funnels and Dan Henry into its MDW/PB/5MS
// products (same splits as the pipelines above); every other client is a
// single entry.
function summarizeWorkbook(clientName, records) {
  const pipelineRecs = records.filter(r => isPipelineStatus(r) && !isLaunched(r));
  if (clientName === 'VAM') {
    return [
      { client: 'VAM', sub: 'VAM Funnel', counts: pipelineCounts(pipelineRecs.filter(r => !isNsr(r))) },
      { client: 'VAM', sub: 'NSR Funnel', counts: pipelineCounts(pipelineRecs.filter(isNsr)) },
    ];
  }
  if (clientName === 'Dan Henry') {
    return [
      { client: 'Dan Henry', sub: 'MDW Funnel', counts: pipelineCounts(pipelineRecs.filter(isDanHenryMdw)) },
      { client: 'Dan Henry', sub: 'PB Funnel',  counts: pipelineCounts(pipelineRecs.filter(isDanHenryPb)) },
      { client: 'Dan Henry', sub: '5MS Funnel', counts: pipelineCounts(pipelineRecs.filter(isDanHenry5ms)) },
    ];
  }
  return [{ client: clientName, sub: null, counts: pipelineCounts(pipelineRecs) }];
}

// Total / Videos / Statics bullets for one stage, or a single "none" bullet
// when the stage is empty.
function stageBullets(c) {
  const sB = c.Statics.batches, sA = c.Statics.ads;
  const vB = c.Videos.batches, vA = c.Videos.ads;
  const tB = sB + vB, tA = sA + vA;
  if (!tB) return '<ul>\n  <li>None</li>\n</ul>';
  const b = (name, batches, ads) =>
    `  <li>${name}: <b>${batches}</b> batch${batches === 1 ? '' : 'es'} (${ads} ad${ads === 1 ? '' : 's'})</li>`;
  return `<ul>
${b('Total', tB, tA)}
${b('Videos', vB, vA)}
${b('Statics', sB, sA)}
</ul>`;
}

function summaryStage(counts) {
  return `<p><i>Ready to launch</i></p>
${stageBullets(counts.ready)}
<p><i>In Production</i></p>
${stageBullets(counts.production)}`;
}

// Self-contained styles for the summary widget, scoped under .pipeline-summary
// so they don't collide with the surrounding notes styles. Bulleted/heading
// form (no table) so it pastes cleanly into Google Docs.
const SUMMARY_WIDGET_STYLE = `.pipeline-summary { margin-top: 8px; }
.pipeline-summary h2 { font-size: 14pt; margin-bottom: 2px; }
.pipeline-summary h3 { font-size: 12pt; margin: 14px 0 2px; }
.pipeline-summary .headline { margin: 6px 0 14px; }
.pipeline-summary .headline b { font-size: 12pt; }
.pipeline-summary ul { margin: 2px 0 10px 0; }
.pipeline-summary li { margin: 2px 0; }`;

// TCC pipeline is tracked in a dedicated Google Sheet rather than the tracker,
// so its summary entry links to that sheet instead of showing batch counts.
const TCC_PIPELINE_SHEET = 'https://docs.google.com/spreadsheets/d/1a0a5HyeTO18IGtvBvvy-Bk25pNAIvHl9pP3psO6qfOk/edit?gid=0#gid=0';

// Inner HTML for the pipeline summary widget (no page wrapper). `entries`
// is a flat array of { client, sub, counts } from summarizeWorkbook();
// `dateDisp` is the report date rendered as the section's H2 heading.
function renderSummaryWidget(entries, dateDisp) {
  if (!entries.length) return '';

  const zero = () => ({ batches: 0, ads: 0 });
  const totals = {
    ready: { Statics: zero(), Videos: zero() },
    production: { Statics: zero(), Videos: zero() },
  };
  for (const e of entries) {
    for (const st of ['ready', 'production']) {
      for (const k of ['Statics', 'Videos']) {
        totals[st][k].batches += e.counts[st][k].batches;
        totals[st][k].ads += e.counts[st][k].ads;
      }
    }
  }
  const totalReadyBatches = totals.ready.Statics.batches + totals.ready.Videos.batches;
  const totalReadyAds = totals.ready.Statics.ads + totals.ready.Videos.ads;
  const totalProdBatches = totals.production.Statics.batches + totals.production.Videos.batches;

  // Group the flat entries by client so each client shows its Creative
  // Dashboard link once at the top and Notes once at the bottom, with any
  // sub-funnel splits (VAM, Dan Henry) rendered as inner sections between.
  const groups = [];
  const groupIdx = {};
  for (const e of entries) {
    if (!(e.client in groupIdx)) {
      groupIdx[e.client] = groups.length;
      groups.push({ client: e.client, subs: [] });
    }
    groups[groupIdx[e.client]].subs.push(e);
  }

  const body = groups.map(g => {
    const dashUrl = DASHBOARD_URLS[g.client];
    const parts = [`<h3>${escapeHtml(g.client)}</h3>`];
    parts.push(`<p><i>Creative Dashboard:</i>${dashUrl ? ` <a href="${dashUrl}">${escapeHtml(dashUrl)}</a>` : ''}</p>`);
    if (g.client === 'TCC') {
      parts.push(`<p><i>Ready to launch &amp; In Production:</i> <a href="${TCC_PIPELINE_SHEET}">${escapeHtml(TCC_PIPELINE_SHEET)}</a></p>`);
    } else {
      for (const s of g.subs) {
        if (s.sub) parts.push(`<p><b>${escapeHtml(s.sub)}</b></p>`);
        parts.push(summaryStage(s.counts));
      }
    }
    parts.push('<p><i>Notes</i></p>\n<ul>\n  <li></li>\n</ul>');
    return parts.join('\n');
  }).join('\n');

  return `<div class="pipeline-summary">
<h2>${escapeHtml(dateDisp || 'Pipeline Summary')}</h2>
<p class="headline"><b>${totalReadyBatches} batch${totalReadyBatches === 1 ? '' : 'es'} ready to launch</b> (${totalReadyAds} ad${totalReadyAds === 1 ? '' : 's'}) · ${totalProdBatches} batch${totalProdBatches === 1 ? '' : 'es'} still in production</p>
${body}
<p style="color:#888;font-size:9.5pt;margin-top:6px">Counts use the same batch grouping as the pipelines above (batches, with total ad variations).</p>
</div>`;
}

// -------- Workbook loader (shared helper for the tabs) --------

export async function readWorkbookFromFile(file) {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, { type: 'array', cellHyperlinks: true, cellDates: true });
}
