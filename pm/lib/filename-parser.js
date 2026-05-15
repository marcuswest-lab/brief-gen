// Extract metadata from ad filenames before sending to Claude vision API.
//
// Common patterns we recognize:
//   Dan Henry static:
//     "PB - Million Follower Sales Flop Comparison - Problem-Solution Lead - CIDVYJVTKW_1x1.png"
//     "PB - Million Follower Sales Flop Comparison - Secret Lead 2 - CIDGZQVTWA_9x16.png"
//
//   Generic with CID:
//     "<anything> CIDABCDEFG.<ext>"
//     "<anything>_CPID12345.<ext>"
//
//   With ratio suffix:
//     "<base>_1x1.png", "<base>_9x16.png", "<base>_4x5.jpg"
//
// Returns { cid, ratio, leadType, ideaName, baseName, ext } — all optional.

import { leadTypeNormalize } from './tracker-config.js';

const CID_RE = /\b(CP?ID)([A-Za-z0-9]+)/;            // CID or CPID
const RATIO_RE = /[_\s-](\d+)\s*[xX]\s*(\d+)\s*(?=\.[a-z]+$)/i;
const EXT_RE = /\.([a-z0-9]+)$/i;
const VALID_LEAD_TYPES = ['Offer', 'Promise', 'Problem-Solution', 'Secret', 'Proclamation', 'Story'];

export function parseAdFilename(filename) {
  const out = {
    filename,
    cid: '',
    cidPrefix: '',          // 'CID' or 'CPID' as found
    ratio: '',              // '1x1', '9x16', etc.
    leadType: '',
    ideaName: '',
    angleName: '',
    baseName: filename,     // filename minus extension and ratio suffix
    ext: '',
  };

  // Extract extension
  const extMatch = EXT_RE.exec(filename);
  if (extMatch) out.ext = extMatch[1].toLowerCase();

  // Extract CID
  const cidMatch = CID_RE.exec(filename);
  if (cidMatch) {
    out.cidPrefix = cidMatch[1];
    out.cid = cidMatch[2];
  }

  // Extract ratio
  const ratioMatch = RATIO_RE.exec(filename);
  if (ratioMatch) {
    out.ratio = `${ratioMatch[1]}x${ratioMatch[2]}`;
  }

  // Compute base name (strip ext, strip ratio suffix, strip CID-onward)
  let base = filename;
  if (out.ext) base = base.slice(0, -(out.ext.length + 1));
  if (out.ratio) {
    base = base.replace(new RegExp(`[_\\s-]${out.ratio}$`, 'i'), '');
  }
  // Strip trailing " - CIDxxx" segment for cleaner display
  base = base.replace(new RegExp(`\\s*[-_]\\s*${out.cidPrefix}${out.cid}.*$`, 'i'), '').trim();
  out.baseName = base;

  // Try Dan-Henry-style segment parse: "{prefix} - {Idea Name} - {Lead Type Lead [N]}"
  // After we've stripped the CID portion, base looks like:
  //   "PB - Million Follower Sales Flop Comparison - Problem-Solution Lead"
  //   "PB - Million Follower Sales Flop Comparison - Secret Lead 2"
  const segs = base.split(' - ').map(s => s.trim()).filter(Boolean);
  if (segs.length >= 3) {
    // Last segment is presumed to be "{LeadType} Lead [N]"
    const last = segs[segs.length - 1];
    const normalized = leadTypeNormalize(last);
    if (VALID_LEAD_TYPES.includes(normalized)) {
      out.leadType = normalized;
      // Idea Name = segs[1..N-1].join(' - ') (skip prefix at segs[0])
      out.ideaName = segs.slice(1, -1).join(' - ');
    }
  } else if (segs.length === 2) {
    // "{Idea Name} - {LeadType Lead}"
    const last = segs[segs.length - 1];
    const normalized = leadTypeNormalize(last);
    if (VALID_LEAD_TYPES.includes(normalized)) {
      out.leadType = normalized;
      out.ideaName = segs[0];
    }
  }

  return out;
}

/**
 * Group a list of files by their dedupe key. The key is the CID if present,
 * otherwise the base name (filename minus ratio + extension). Same key =
 * same logical creative, just rendered in different ratios.
 *
 * Returns Map<key, { key, files: File[], parsed: parseAdFilename of representative, ratios: string[] }>.
 */
export function groupAdFiles(files) {
  const groups = new Map();
  for (const file of files) {
    const parsed = parseAdFilename(file.name);
    const key = parsed.cid || parsed.baseName.toLowerCase();
    if (!groups.has(key)) {
      groups.set(key, {
        key,
        files: [],
        parsed,
        ratios: [],
      });
    }
    const g = groups.get(key);
    g.files.push(file);
    if (parsed.ratio && !g.ratios.includes(parsed.ratio)) g.ratios.push(parsed.ratio);
  }
  return groups;
}

/**
 * Pick the "best" file from a group to send to Claude — typically the
 * largest by byte size, since that's usually the highest-resolution variant.
 */
export function pickRepresentativeFile(group) {
  if (group.files.length === 1) return group.files[0];
  return group.files.slice().sort((a, b) => b.size - a.size)[0];
}
