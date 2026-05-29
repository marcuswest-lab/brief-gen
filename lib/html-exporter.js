// HTML export for the generated docx Blob.
//
// We walk the filled docx XML directly (via JSZip + DOMParser) and emit HTML
// that mirrors the docx structure cell-for-cell: same column widths, same
// cell shading (e.g. the `#c9daf8` numbered creative headers), same row
// merges, same Proxima Nova / Arial font, same 10pt body / 14pt headings,
// same 1pt black borders. The output pastes into Google Docs as a real table
// matching the source docx visually.
//
// mammoth.js (used by the previous version) was too lossy — it dropped cell
// shading and column widths, so the result looked nothing like the original
// brief. This implementation is template-aware but generic enough to handle
// any filled docx the existing `generateBrief()` pipeline produces.

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

// Twentieths of a point → CSS px. 1pt = 1.3333px, so 1 dxa = 0.0667px.
// We mostly use dxa for column widths as percentages of the table width, so
// the absolute conversion rarely matters — but here it is anyway.
function dxaToPx(dxa) {
  return (Number(dxa) || 0) * 0.0667;
}
// Half-points → CSS pt. `<w:sz w:val="22"/>` means 11pt.
function halfPtToPt(sz) {
  return (Number(sz) || 0) / 2;
}

function getAttrNS(el, name) {
  // XML attributes in the w: namespace can be accessed by local name on
  // browsers' DOMParser output when the doc was parsed as XML.
  return el.getAttributeNS(W, name) || el.getAttribute('w:' + name) || el.getAttribute(name);
}

function childrenNS(parent, localName) {
  if (!parent) return [];
  const out = [];
  for (const c of parent.children) {
    if (c.localName === localName && (c.namespaceURI === W || c.tagName.startsWith('w:'))) {
      out.push(c);
    }
  }
  return out;
}

function firstChildNS(parent, localName) {
  const list = childrenNS(parent, localName);
  return list.length ? list[0] : null;
}

function getCellShadingFill(tc) {
  const tcPr = firstChildNS(tc, 'tcPr');
  if (!tcPr) return null;
  const shd = firstChildNS(tcPr, 'shd');
  if (!shd) return null;
  const fill = getAttrNS(shd, 'fill');
  if (!fill || fill.toLowerCase() === 'auto') return null;
  return '#' + fill.toLowerCase();
}

function getCellGridSpan(tc) {
  const tcPr = firstChildNS(tc, 'tcPr');
  if (!tcPr) return 1;
  const gs = firstChildNS(tcPr, 'gridSpan');
  if (!gs) return 1;
  return Number(getAttrNS(gs, 'val')) || 1;
}

function isCellMerged(tc) {
  // Returns true if this cell is a continuation of a vertical merge above.
  const tcPr = firstChildNS(tc, 'tcPr');
  if (!tcPr) return false;
  const vMerge = firstChildNS(tcPr, 'vMerge');
  if (!vMerge) return false;
  const val = getAttrNS(vMerge, 'val');
  return !val || val === 'continue'; // present-without-val OR explicit continue
}

function getParaAlignment(p) {
  const pPr = firstChildNS(p, 'pPr');
  if (!pPr) return null;
  const jc = firstChildNS(pPr, 'jc');
  if (!jc) return null;
  return getAttrNS(jc, 'val'); // left/center/right/both
}

function getParaStyle(p) {
  const pPr = firstChildNS(p, 'pPr');
  if (!pPr) return null;
  const ps = firstChildNS(pPr, 'pStyle');
  if (!ps) return null;
  return getAttrNS(ps, 'val'); // e.g. "Heading1"
}

// Heading sizes (pt) by style name. Matches the docx default Heading 1-3.
const HEADING_PT = {
  Heading1: 20,
  Heading2: 16,
  Heading3: 13,
};

function escapeHtml(s) {
  return String(s).replace(/[<>&"']/g, (c) => ({
    '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;', "'": '&#39;',
  }[c]));
}

// Render one paragraph (<w:p>) inside a cell as a single HTML <p>. Multiple
// runs are concatenated. Run-level bold/italic/underline are honoured. Empty
// paragraphs become a non-breaking space so spacing is preserved.
function renderParagraph(p) {
  const align = getParaAlignment(p);
  const style = getParaStyle(p);
  const headingPt = style && HEADING_PT[style];

  // Walk runs (w:r) and sdt content (w:sdt). SDTs (dropdowns) wrap the
  // rendered visible text inside w:sdtContent > w:r > w:t.
  const parts = [];
  for (const node of p.children) {
    if (node.localName === 'r' && (node.namespaceURI === W || node.tagName.startsWith('w:'))) {
      parts.push(renderRun(node));
    } else if (node.localName === 'sdt' && (node.namespaceURI === W || node.tagName.startsWith('w:'))) {
      const sdtContent = firstChildNS(node, 'sdtContent');
      if (sdtContent) {
        for (const child of sdtContent.children) {
          if (child.localName === 'r') parts.push(renderRun(child));
          else if (child.localName === 'p') parts.push(renderParagraph(child));
        }
      }
    } else if (node.localName === 'hyperlink') {
      // hyperlink wraps runs
      for (const r of childrenNS(node, 'r')) parts.push(renderRun(r));
    }
  }

  const inner = parts.join('') || '&nbsp;';

  // Build style for <p>: alignment, optional heading size, tight margins.
  const styles = ['margin:0', 'padding:0'];
  if (align === 'center') styles.push('text-align:center');
  else if (align === 'right') styles.push('text-align:right');
  else if (align === 'both') styles.push('text-align:justify');
  // Default left — no need to set.
  if (headingPt) styles.push(`font-size:${headingPt}pt`, 'font-weight:700');

  return `<p style="${styles.join(';')}">${inner}</p>`;
}

function renderRun(r) {
  const rPr = firstChildNS(r, 'rPr');
  let bold = false, italic = false, underline = false, sizePt = null;
  let colorHex = null, highlightHex = null;
  if (rPr) {
    const b = firstChildNS(rPr, 'b');
    if (b) {
      const v = getAttrNS(b, 'val');
      bold = v == null || v === '1' || v === 'true';
    }
    const i = firstChildNS(rPr, 'i');
    if (i) {
      const v = getAttrNS(i, 'val');
      italic = v == null || v === '1' || v === 'true';
    }
    const u = firstChildNS(rPr, 'u');
    if (u) {
      const v = getAttrNS(u, 'val');
      underline = !!(v && v !== 'none');
    }
    const sz = firstChildNS(rPr, 'sz');
    if (sz) sizePt = halfPtToPt(getAttrNS(sz, 'val'));
    const color = firstChildNS(rPr, 'color');
    if (color) {
      const v = getAttrNS(color, 'val');
      if (v && v.toLowerCase() !== 'auto') colorHex = '#' + v.toLowerCase();
    }
    // Run-level shading = text highlight (separate from cell shading).
    // Dropdown values like "Yes, AI Is Allowed" use this to mark the value
    // with a colored background behind the text.
    const shd = firstChildNS(rPr, 'shd');
    if (shd) {
      const fill = getAttrNS(shd, 'fill');
      if (fill && fill.toLowerCase() !== 'auto') highlightHex = '#' + fill.toLowerCase();
    }
    // Word also has a separate w:highlight element with named colors. Pick
    // it up too in case the source uses that instead of shd.
    const hl = firstChildNS(rPr, 'highlight');
    if (hl) {
      const v = getAttrNS(hl, 'val');
      if (v && v !== 'none') highlightHex = NAMED_HIGHLIGHTS[v] || highlightHex;
    }
  }

  // Collect text from w:t and w:br nodes.
  const buf = [];
  for (const node of r.children) {
    if (node.localName === 't') {
      buf.push(escapeHtml(node.textContent));
    } else if (node.localName === 'br') {
      buf.push('<br>');
    } else if (node.localName === 'tab') {
      buf.push('&nbsp;&nbsp;&nbsp;&nbsp;');
    }
  }
  const text = buf.join('');
  if (!text) return '';

  const styles = [];
  if (bold) styles.push('font-weight:700');
  if (italic) styles.push('font-style:italic');
  if (underline) styles.push('text-decoration:underline');
  if (sizePt) styles.push(`font-size:${sizePt}pt`);
  if (colorHex) styles.push(`color:${colorHex}`);
  if (highlightHex) {
    // Highlight: colored background + a touch of horizontal padding so the
    // text isn't squished against the highlight edges (matches docx render).
    styles.push(`background-color:${highlightHex}`);
    styles.push('padding:1px 4px');
    styles.push('border-radius:2px');
  }

  if (!styles.length) return text;
  return `<span style="${styles.join(';')}">${text}</span>`;
}

// Word's w:highlight uses named colors. Map to approximate hex.
const NAMED_HIGHLIGHTS = {
  yellow: '#ffff00',
  green: '#00ff00',
  cyan: '#00ffff',
  magenta: '#ff00ff',
  blue: '#0000ff',
  red: '#ff0000',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
  black: '#000000',
};

function renderCell(tc, { isFirstColumn }) {
  const colspan = getCellGridSpan(tc);
  const fill = getCellShadingFill(tc);

  // Cell content paragraphs.
  const paras = childrenNS(tc, 'p').map(renderParagraph).join('');

  const styles = [
    'border:1px solid #000',
    'padding:4pt 6pt',
    'vertical-align:top',
  ];
  if (fill) styles.push(`background:${fill}`);
  // The label column (first column on label/value rows) doesn't get an
  // explicit style — alignment is set on the paragraph from the docx. We
  // intentionally do NOT impose a gray label background since the original
  // docx uses white cells with right-aligned bold labels (per inspection).

  const colspanAttr = colspan > 1 ? ` colspan="${colspan}"` : '';
  return `<td${colspanAttr} style="${styles.join(';')}">${paras}</td>`;
}

function renderRow(tr) {
  const cells = [];
  for (const tc of childrenNS(tr, 'tc')) {
    if (isCellMerged(tc)) continue; // skip vMerge continuation cells
    cells.push(renderCell(tc, { isFirstColumn: cells.length === 0 }));
  }
  return `<tr>${cells.join('')}</tr>`;
}

function renderTable(tbl, { pageBreakBefore = false } = {}) {
  // Column widths from w:tblGrid > w:gridCol
  const tblGrid = firstChildNS(tbl, 'tblGrid');
  const cols = tblGrid ? childrenNS(tblGrid, 'gridCol') : [];
  const widthsDxa = cols.map((c) => Number(getAttrNS(c, 'w')) || 0);
  const totalDxa = widthsDxa.reduce((a, b) => a + b, 0) || 1;

  const colgroup = widthsDxa.length
    ? `<colgroup>${widthsDxa.map((w) => `<col style="width:${((w / totalDxa) * 100).toFixed(2)}%">`).join('')}</colgroup>`
    : '';

  const rows = childrenNS(tbl, 'tr').map(renderRow).join('');
  const tableStyles = ['border-collapse:collapse', 'width:100%', 'table-layout:fixed', 'margin:0 0 16pt 0'];
  if (pageBreakBefore) {
    // Both legacy and modern page-break properties for max paste-target
    // compatibility (Google Docs reads page-break-before; Word reads break-before).
    tableStyles.push('page-break-before:always', 'break-before:page');
  }
  return `<table style="${tableStyles.join(';')}">${colgroup}${rows}</table>`;
}

/**
 * Walk a filled docx Blob and produce HTML mirroring the docx structure.
 *
 * @param {Blob} blob filled .docx (output of generateBrief())
 * @returns {Promise<string>} HTML body content (no <html>/<head> wrapper)
 */
export async function docxBlobToHtml(blob) {
  if (typeof window === 'undefined' || !window.JSZip) {
    throw new Error('JSZip is not loaded — check the CDN script tag in index.html');
  }
  const arrayBuffer = await blob.arrayBuffer();
  const zip = await window.JSZip.loadAsync(arrayBuffer);
  const xmlText = await zip.file('word/document.xml').async('string');

  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlText, 'application/xml');

  // Body holds the top-level w:p (paragraphs) and w:tbl (tables) in document
  // order. We render headings (paragraphs) and tables interleaved as they
  // appear so the header line (e.g. "Brief for Case Source | Idea | Date")
  // shows up above the overview table.
  const body = xmlDoc.getElementsByTagNameNS(W, 'body')[0] || xmlDoc.documentElement;

  // Track table index so we can put each creative block on its own page.
  // Tables: 0 = Overview, 1+ = Creatives. Page-break before every creative.
  let tableIdx = 0;

  const parts = [];
  for (const child of body.children) {
    if (child.localName === 'p' && (child.namespaceURI === W || child.tagName.startsWith('w:'))) {
      // Skip empty section-break paragraphs at the end.
      const text = child.textContent.trim();
      if (!text) continue;
      parts.push(`<div style="margin:0 0 12pt 0">${renderParagraph(child)}</div>`);
    } else if (child.localName === 'tbl' && (child.namespaceURI === W || child.tagName.startsWith('w:'))) {
      parts.push(renderTable(child, { pageBreakBefore: tableIdx > 0 }));
      tableIdx += 1;
    }
  }

  return parts.join('\n');
}

/**
 * Wrap body HTML in a styled <div> with inline CSS that matches the docx
 * defaults (Proxima Nova → Arial fallback, 10pt body, single line height).
 *
 * @param {string} innerHtml body output of docxBlobToHtml
 * @param {{title?: string}} opts
 */
export function wrapHtmlForPaste(innerHtml, { title } = {}) {
  const safeTitle = (title || 'Brief').replace(/[<>&"]/g, (c) => ({
    '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;',
  }[c]));

  // Page-level styles mirror the docx defaults.
  // Font stack matches the inline font hint in the docx (Proxima Nova) with
  // Arial as fallback (also what the docx falls back to).
  const styles = `
    body { font-family: 'Proxima Nova', 'Source Sans 3', Arial, Helvetica, sans-serif; font-size: 10pt; color: #000; line-height: 1.2; margin: 24px; }
    table { border-collapse: collapse; }
    p { margin: 0; padding: 0; }
  `.trim();

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>${safeTitle}</title>
<style>
${styles}
</style>
</head>
<body>
${innerHtml}
</body>
</html>`;
}

function extractBodyAndStyle(fullHtml) {
  const styleMatch = fullHtml.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
  const bodyMatch = fullHtml.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  const style = styleMatch ? `<style>${styleMatch[1]}</style>` : '';
  const body = bodyMatch ? bodyMatch[1] : fullHtml;
  return `${style}\n${body}`;
}

function htmlToPlainText(html) {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<\/(p|tr|h\d|li|div)>/gi, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Copy HTML to the clipboard as rich text (`text/html`) + plain-text fallback.
 */
export async function copyHtmlToClipboard(fullHtml) {
  const richPayload = extractBodyAndStyle(fullHtml);
  const plainPayload = htmlToPlainText(richPayload);

  if (!navigator.clipboard || !window.ClipboardItem) {
    return fallbackCopyPlain(plainPayload);
  }

  const item = new ClipboardItem({
    'text/html': new Blob([richPayload], { type: 'text/html' }),
    'text/plain': new Blob([plainPayload], { type: 'text/plain' }),
  });
  await navigator.clipboard.write([item]);
}

function fallbackCopyPlain(text) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.style.position = 'fixed';
  ta.style.left = '-9999px';
  document.body.appendChild(ta);
  ta.select();
  try {
    document.execCommand('copy');
  } finally {
    document.body.removeChild(ta);
  }
}

/**
 * Download the full HTML document as a `.html` file via FileSaver.js.
 */
export function downloadHtmlFile(fullHtml, filename) {
  if (typeof window === 'undefined' || typeof window.saveAs !== 'function') {
    throw new Error('saveAs (FileSaver.js) is not loaded');
  }
  const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
  window.saveAs(blob, filename);
}

export function docxFilenameToHtml(docxFilename) {
  return String(docxFilename).replace(/\.docx$/i, '') + '.html';
}
