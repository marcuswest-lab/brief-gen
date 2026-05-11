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
 */
export function briefDisplayLabel(brief) {
  const type = brief.briefType === 'copy' ? 'Body Copy' : (brief.briefType[0].toUpperCase() + brief.briefType.slice(1));
  const idea = brief.ideaName || '(untitled)';
  const date = formatShortDate(brief.updatedAt || brief.createdAt);
  return `${brief.clientName} — ${type} — ${idea} — ${date}`;
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
