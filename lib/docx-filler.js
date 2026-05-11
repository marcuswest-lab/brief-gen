// Core .docx filling logic. Mirrors generate_doc.py but in browser-side JS.
//
// Approach:
// 1. Load the .docx as a zip (via JSZip).
// 2. Parse word/document.xml with DOMParser (XML mode).
// 3. For each table+row+cell, write text or set SDT dropdown values.
// 4. For headers (word/header*.xml), do simple [PLACEHOLDER] replacement.
// 5. Re-serialize and write back into the zip.

import { TEMPLATES, DROPDOWN_OPTIONS, normalizeDropdownValue } from './templates-config.js';
import { stripComments } from './docx-cleaner.js';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

// -------- helpers --------

function tag(name) {
  return `{${W_NS}}${name}`;
}

// Find direct children with a w:* local name
function childrenW(parent, localName) {
  const out = [];
  for (const node of parent.childNodes) {
    if (node.nodeType === 1 && node.namespaceURI === W_NS && node.localName === localName) {
      out.push(node);
    }
  }
  return out;
}

// Find descendants with a w:* local name
function descendantsW(parent, localName) {
  const out = [];
  const walk = (node) => {
    for (const child of node.childNodes) {
      if (child.nodeType === 1 && child.namespaceURI === W_NS && child.localName === localName) {
        out.push(child);
      }
      if (child.nodeType === 1) walk(child);
    }
  };
  walk(parent);
  return out;
}

function getTables(doc) {
  // Skip tables nested inside other tables — only return top-level w:tbl elements
  // that sit directly under w:body.
  const body = descendantsW(doc.documentElement, 'body')[0];
  if (!body) return [];
  return childrenW(body, 'tbl');
}

function getRows(table) {
  return childrenW(table, 'tr');
}

function getCells(row) {
  return childrenW(row, 'tc');
}

// -------- text cell writing --------

/**
 * Replace the entire content of a cell with the given text.
 * Preserves the first paragraph's run formatting (rPr) where possible.
 * Multi-line text becomes multiple paragraphs.
 */
function setCellText(cell, text, doc) {
  if (text == null || text === '') return;

  const lines = String(text).split('\n');

  // Find existing paragraphs
  const paras = childrenW(cell, 'p');
  if (paras.length === 0) return;

  const firstP = paras[0];

  // Capture rPr template from first available run (even nested inside SDTs)
  let rPrTemplate = null;
  const allRuns = descendantsW(firstP, 'r');
  for (const r of allRuns) {
    const rPr = childrenW(r, 'rPr')[0];
    if (rPr) {
      rPrTemplate = rPr.cloneNode(true);
      break;
    }
  }

  // Capture pPr template
  const pPrTemplate = childrenW(firstP, 'pPr')[0]?.cloneNode(true) || null;

  // Clear first paragraph (keep pPr only)
  for (const child of Array.from(firstP.childNodes)) {
    if (!(child.nodeType === 1 && child.namespaceURI === W_NS && child.localName === 'pPr')) {
      firstP.removeChild(child);
    }
  }

  // Add first line as a run
  appendTextRun(firstP, lines[0], rPrTemplate, doc);

  // Remove extra paragraphs after the first
  for (let i = paras.length - 1; i > 0; i--) {
    paras[i].parentNode.removeChild(paras[i]);
  }

  // Add remaining lines as new paragraphs at the end of the cell
  for (let i = 1; i < lines.length; i++) {
    const p = doc.createElementNS(W_NS, 'w:p');
    if (pPrTemplate) p.appendChild(pPrTemplate.cloneNode(true));
    appendTextRun(p, lines[i], rPrTemplate, doc);
    cell.appendChild(p);
  }
}

function appendTextRun(parent, text, rPrTemplate, doc) {
  const r = doc.createElementNS(W_NS, 'w:r');
  if (rPrTemplate) r.appendChild(rPrTemplate.cloneNode(true));
  const t = doc.createElementNS(W_NS, 'w:t');
  t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  t.textContent = text;
  r.appendChild(t);
  parent.appendChild(r);
}

// -------- dropdown (SDT) cell writing --------

/**
 * Set a dropdown value in a cell. Walks SDTs, finds the dropDownList/comboBox,
 * matches the value to a listItem, and updates both the lastValue attribute
 * and the displayed text inside sdtContent.
 */
function setDropdownValue(cell, value, fieldName, doc) {
  if (value == null || value === '') return false;

  const sdts = descendantsW(cell, 'sdt');
  for (const sdt of sdts) {
    const sdtPr = childrenW(sdt, 'sdtPr')[0];
    if (!sdtPr) continue;

    const dd = childrenW(sdtPr, 'dropDownList')[0] || childrenW(sdtPr, 'comboBox')[0];
    if (!dd) continue;

    // Collect options
    const items = childrenW(dd, 'listItem');
    const options = items
      .map(it => it.getAttributeNS(W_NS, 'displayText') || it.getAttributeNS(W_NS, 'value'))
      .filter(Boolean);

    // Normalize value
    const normalized = normalizeDropdownValue(value, options, fieldName);

    // Find matching item
    let matchedDisplay = null;
    let matchedValue = null;
    for (const it of items) {
      const display = it.getAttributeNS(W_NS, 'displayText') || it.getAttributeNS(W_NS, 'value');
      const itVal = it.getAttributeNS(W_NS, 'value');
      if (display && normalized.toLowerCase() === display.toLowerCase()) {
        matchedDisplay = display;
        matchedValue = itVal || display;
        break;
      }
    }

    if (!matchedDisplay) {
      console.warn(`[docx-filler] '${value}' (normalized: '${normalized}') not in options ${JSON.stringify(options)} for ${fieldName}`);
      continue;
    }

    // Update lastValue attribute on the dropdown element
    dd.setAttributeNS(W_NS, 'w:lastValue', matchedValue);

    // Update display text inside sdtContent
    const sdtContent = childrenW(sdt, 'sdtContent')[0];
    if (!sdtContent) continue;

    // Try direct runs under sdtContent first
    const directRuns = childrenW(sdtContent, 'r');
    if (directRuns.length > 0) {
      for (const r of directRuns) {
        const ts = childrenW(r, 't');
        for (const t of ts) t.textContent = matchedDisplay;
      }
      return true;
    }

    // Try runs inside paragraphs under sdtContent
    const innerParas = childrenW(sdtContent, 'p');
    if (innerParas.length > 0) {
      for (const p of innerParas) {
        const runs = childrenW(p, 'r');
        for (const r of runs) {
          const ts = childrenW(r, 't');
          for (const t of ts) t.textContent = matchedDisplay;
        }
      }
      return true;
    }

    // Last resort: append a fresh run with the text
    const r = doc.createElementNS(W_NS, 'w:r');
    const t = doc.createElementNS(W_NS, 'w:t');
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
    t.textContent = matchedDisplay;
    r.appendChild(t);
    sdtContent.appendChild(r);
    return true;
  }

  // No SDT found — fall back to plain text
  setCellText(cell, value, doc);
  return false;
}

// -------- header substitution --------

function fillHeaderXml(xmlText, replacements) {
  // Simple [PLACEHOLDER] string replace — header template uses bracketed tokens
  let out = xmlText;
  for (const [placeholder, value] of Object.entries(replacements)) {
    // Escape XML special chars in the value
    const escaped = String(value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
    // Use split/join for global replace without regex escaping headaches
    out = out.split(placeholder).join(escaped);
  }
  return out;
}

// -------- removal of unused creative tables --------

/**
 * Remove creative tables for indices that weren't filled.
 * Templates have 5 creative tables (indices 1-5). If the user submitted N
 * creatives, remove tables N+1 through 5 entirely so the brief is clean.
 *
 * Also removes any paragraphs sitting between the last kept creative table
 * and the next (now-removed) table — typically a spacer paragraph.
 */
function removeUnusedCreativeTables(doc, numCreatives) {
  if (numCreatives >= 5) return;
  const tables = getTables(doc);
  // Tables 1..5 are the creative tables. We want to keep 1..numCreatives, drop the rest.
  for (let i = 5; i > numCreatives; i--) {
    if (!tables[i]) continue;
    const tbl = tables[i];

    // Remove any spacer paragraph immediately before this table
    let prev = tbl.previousSibling;
    while (prev) {
      if (prev.nodeType === 1) {
        if (prev.namespaceURI === W_NS && prev.localName === 'p') {
          // Only remove if it has no real text content (just a spacer)
          const text = descendantsW(prev, 't').map(t => t.textContent || '').join('').trim();
          if (text === '') {
            const toRemove = prev;
            prev = prev.previousSibling;
            toRemove.parentNode.removeChild(toRemove);
            continue;
          }
        }
        break;
      }
      prev = prev.previousSibling;
    }

    tbl.parentNode.removeChild(tbl);
  }
}

/**
 * Remove every w:p (and any other non-table, non-section-break element) that
 * sits at the start of <w:body> before the first <w:tbl>. The templates ship
 * with a "Batch 1" Title paragraph + empty spacer paragraph above OVERVIEW,
 * both of which we don't want in generated briefs.
 */
function removeCoverParagraphs(doc) {
  const body = descendantsW(doc.documentElement, 'body')[0];
  if (!body) return;

  // Walk children in order; remove everything until we hit the first w:tbl.
  // Preserve <w:sectPr> (page section properties) — those live at the END of
  // body, but be defensive in case any structural element appears.
  const toRemove = [];
  for (const node of Array.from(body.childNodes)) {
    if (node.nodeType !== 1) continue;
    if (node.namespaceURI !== W_NS) continue;
    if (node.localName === 'tbl') break;        // first table — stop
    if (node.localName === 'sectPr') continue;  // never strip section props
    toRemove.push(node);
  }
  for (const node of toRemove) {
    node.parentNode.removeChild(node);
  }
}

// -------- main entry point --------

/**
 * Generate a filled .docx blob.
 * @param {Object} params
 * @param {string} params.briefType - 'static' | 'video' | 'copy'
 * @param {string} params.clientName - displayed in [CLIENT] header placeholder
 * @param {Object} params.overview - { fieldName: value } for overview fields
 * @param {Array<Object>} params.creatives - per-creative data, each is { fieldName: value }
 * @param {ArrayBuffer} params.templateBuffer - raw .docx template bytes
 * @returns {Promise<Blob>}
 */
export async function generateBrief({ briefType, clientName, overview, creatives, templateBuffer }) {
  const config = TEMPLATES[briefType];
  if (!config) throw new Error(`Unknown brief type: ${briefType}`);

  // Load the zip
  const zip = await JSZip.loadAsync(templateBuffer);

  // -- Process word/document.xml --
  const docXmlText = await zip.file('word/document.xml').async('string');
  const parser = new DOMParser();
  const doc = parser.parseFromString(docXmlText, 'application/xml');

  // Bail loudly on parse errors
  const parseErr = doc.getElementsByTagName('parsererror')[0];
  if (parseErr) {
    throw new Error('Failed to parse document.xml: ' + parseErr.textContent);
  }

  // Strip cover-page leftovers: every w:p that sits in <w:body> BEFORE the
  // first <w:tbl>. Templates ship with a "Batch 1" title paragraph and an
  // empty spacer paragraph at the top — both should not appear in output.
  removeCoverParagraphs(doc);

  const tables = getTables(doc);
  if (tables.length < 6) {
    throw new Error(`Expected at least 6 tables (overview + 5 creatives), found ${tables.length}`);
  }

  // Fill overview (table 0)
  const overviewTable = tables[0];
  const overviewRows = getRows(overviewTable);
  for (const def of config.overview) {
    const value = overview[def.field];
    if (value == null || value === '') continue;
    const row = overviewRows[def.row];
    if (!row) continue;
    const cell = getCells(row)[1];
    if (!cell) continue;

    if (def.kind === 'dropdown') {
      setDropdownValue(cell, value, def.field, doc);
    } else {
      setCellText(cell, value, doc);
    }
  }

  // Fill creatives (tables 1..N)
  const numCreatives = Math.min(creatives.length, 5);
  for (let i = 0; i < numCreatives; i++) {
    const cTable = tables[i + 1];
    if (!cTable) break;
    const cRows = getRows(cTable);
    const cData = creatives[i];

    for (const def of config.creative) {
      const value = cData[def.field];
      if (value == null || value === '') continue;
      const row = cRows[def.row];
      if (!row) continue;
      const cell = getCells(row)[1];
      if (!cell) continue;

      if (def.kind === 'dropdown') {
        setDropdownValue(cell, value, def.field, doc);
      } else {
        setCellText(cell, value, doc);
      }
    }
  }

  // Remove unused creative tables (so a 2-creative brief doesn't carry 3 blank ones)
  removeUnusedCreativeTables(doc, numCreatives);

  // Serialize document.xml back
  const serializer = new XMLSerializer();
  const newDocXml = serializer.serializeToString(doc);
  zip.file('word/document.xml', newDocXml);

  // -- Process headers: [CLIENT], [CAMPAIGN/LAUNCH NAME], [DATE OF CREATION] --
  const dateStr = formatDate(new Date());
  const ideaName = (overview['Idea Name'] || '').toString().trim();
  const campaign = ideaName || `${config.label} Brief`;

  const replacements = {
    '[CLIENT]': clientName || '',
    '[CAMPAIGN/LAUNCH NAME]': campaign,
    '[DATE OF CREATION]': dateStr,
  };

  const headerNames = Object.keys(zip.files).filter(n => /^word\/header\d*\.xml$/.test(n));
  for (const name of headerNames) {
    const text = await zip.file(name).async('string');
    const newText = fillHeaderXml(text, replacements);
    if (newText !== text) zip.file(name, newText);
  }

  // -- Defense-in-depth: strip any review comments still present --
  await stripComments(zip);

  // -- Generate final blob --
  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
  });
}

function formatDate(d) {
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${mm}-${dd}-${yyyy}`;
}
