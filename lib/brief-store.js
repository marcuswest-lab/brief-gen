// Shared brief store: persists finalized briefs to localStorage so Briefgen can
// list/edit/regenerate them and the PM app can pick one in Brief→Tracker /
// Tracker→Brief flows.
//
// Storage key: pbg.briefs.v1
//
// Schema:
//   {
//     id: 'brief_<timestamp>_<rand>',
//     clientId: 'dan-henry',
//     clientName: 'Dan Henry',        // snapshot at save time (in case clients.json changes)
//     briefType: 'static' | 'video' | 'copy',
//     ideaName: 'Quiet Overwhelm',    // mirrored from overview for picker display
//     createdAt: ISO string,
//     updatedAt: ISO string,
//     overview: { [fieldName]: string, ... },
//     creatives: [ { [fieldName]: string, ... }, ... ],
//   }

const STORAGE_KEY = 'pbg.briefs.v1';

export function loadBriefs() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    console.warn('brief-store load failed:', e);
    return [];
  }
}

export function saveAllBriefs(briefs) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(briefs));
  } catch (e) {
    console.warn('brief-store save failed:', e);
  }
}

export function getBrief(id) {
  return loadBriefs().find(b => b.id === id) || null;
}

/**
 * Upsert a brief. If `brief.id` is null/undefined, creates a new record with
 * a fresh ID; otherwise updates the matching record in place. Returns the
 * stored record (with id + timestamps).
 */
export function upsertBrief(brief) {
  const briefs = loadBriefs();
  const now = new Date().toISOString();

  if (brief.id) {
    const idx = briefs.findIndex(b => b.id === brief.id);
    if (idx >= 0) {
      briefs[idx] = { ...brief, updatedAt: now };
      saveAllBriefs(briefs);
      return briefs[idx];
    }
  }

  const newBrief = {
    ...brief,
    id: makeId(),
    createdAt: now,
    updatedAt: now,
  };
  briefs.push(newBrief);
  saveAllBriefs(briefs);
  return newBrief;
}

export function deleteBrief(id) {
  const briefs = loadBriefs().filter(b => b.id !== id);
  saveAllBriefs(briefs);
}

function makeId() {
  return 'brief_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 7);
}

/**
 * Sort briefs newest-first by updatedAt (falling back to createdAt).
 */
export function sortBriefsRecent(briefs) {
  return [...briefs].sort((a, b) => {
    const ta = new Date(a.updatedAt || a.createdAt || 0).getTime();
    const tb = new Date(b.updatedAt || b.createdAt || 0).getTime();
    return tb - ta;
  });
}

/**
 * Display label for a brief — used in pickers and modal lists.
 * Format: "{Client} — {Idea Name}{ - Angle Name} — {Type} — {date}"
 */
export function briefDisplayLabel(brief) {
  const type = brief.briefType === 'copy' ? 'Body Copy' : (brief.briefType[0].toUpperCase() + brief.briefType.slice(1));
  const idea = (brief.ideaName || '').trim();
  const angle = (brief.angleName || '').trim();
  const date = formatShortDate(brief.updatedAt || brief.createdAt);

  // Idea + angle joined when both present
  let title = '';
  if (idea && angle) title = `${idea} - ${angle}`;
  else if (idea) title = idea;
  else if (angle) title = angle;
  else title = '(untitled)';

  return `${brief.clientName} — ${title} — ${type} — ${date}`;
}

function formatShortDate(iso) {
  if (!iso) return '';
  try {
    const d = new Date(iso);
    return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
  } catch {
    return '';
  }
}

/**
 * Shared filename builder used by Briefgen + PM Tools regen.
 * Format:  "Dan Henry Video Brief | The Script Is the Business | 2026-05-12.docx"
 *          {Client} {Type} Brief | {Idea Name} | {YYYY-MM-DD}.docx
 *
 * Empty Idea Name -> just "{Client} {Type} Brief | {YYYY-MM-DD}.docx".
 *
 * @param {Object} args
 * @param {string} args.clientName
 * @param {string} args.briefType  'static' | 'video' | 'copy'
 * @param {string} [args.ideaName]
 * @param {string} [args.suffix]    optional extra label appended after the date
 *                                  (used by PM regen-with-names: "with-names")
 * @param {Date}   [args.date]      defaults to today
 */
export function buildBriefFilename({ clientName, briefType, ideaName, suffix, date }) {
  const typeLabel = briefType === 'copy'
    ? 'Body Copy'
    : (briefType[0].toUpperCase() + briefType.slice(1));
  const d = date || new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const datePart = `${yyyy}-${mm}-${dd}`;

  // Sanitize each segment: keep spaces and most safe punctuation, strip
  // characters that would break filesystems (/, \, :, *, ?, ", <, >).
  // Keep '|' since the user specifically requested it as a separator.
  const safe = (s) => String(s || '').replace(/[\/\\:*?"<>]/g, '').trim();

  const client = safe(clientName) || 'Client';
  const idea = safe(ideaName);
  const segments = [`${client} ${typeLabel} Brief`];
  if (idea) segments.push(idea);
  segments.push(suffix ? `${datePart} ${safe(suffix)}` : datePart);
  return segments.join(' | ') + '.docx';
}
