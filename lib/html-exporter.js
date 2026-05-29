// HTML export for the generated docx Blob.
//
// Pipeline: docx Blob -> ArrayBuffer -> mammoth.convertToHtml() -> raw HTML.
// We then wrap it in an inline <style> block tuned to make the tables look
// like the BAD Marketing brief template, so the rendered preview matches the
// docx visually AND the rich-text clipboard payload pastes into Google Docs
// with the same structure (label column / value column, table borders, bold
// field names).

/**
 * Convert a filled docx Blob to a raw HTML string (mammoth body output only).
 * Uses the globally-loaded `window.mammoth` from the CDN script tag in
 * index.html — we expect it to be available by the time the user clicks an
 * export button. Throws a clear error if it isn't.
 *
 * @param {Blob} blob filled .docx
 * @returns {Promise<string>} body HTML (no <html>, <head>, <style>)
 */
export async function docxBlobToHtml(blob) {
  if (typeof window === 'undefined' || !window.mammoth) {
    throw new Error('mammoth.js is not loaded — check the CDN script tag in index.html');
  }
  const arrayBuffer = await blob.arrayBuffer();
  const result = await window.mammoth.convertToHtml({ arrayBuffer });
  return result.value || '';
}

/**
 * Wrap mammoth output in a styled <div> + inline <style> block. The styles
 * target the table structure mammoth emits for the BAD Marketing brief
 * templates: 2-column tables (label / value). Inline CSS is preserved by
 * Google Docs on paste, so the visual hierarchy survives.
 *
 * @param {string} innerHtml the body HTML from docxBlobToHtml
 * @param {{title?: string}} opts
 * @returns {string} full HTML ready to copy or download
 */
export function wrapHtmlForPaste(innerHtml, { title } = {}) {
  const safeTitle = (title || 'Brief').replace(/[<>&"]/g, (c) => ({
    '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;',
  }[c]));

  // Styles tuned to match the docx template look:
  // - sans-serif body, 11pt
  // - tables with 1px borders, collapsed
  // - first cell in each row (the label) gets a light gray background and bold
  // - generous cell padding so the result is readable when pasted
  // - paragraph margins tightened so multi-line scripts don't get gaps
  const styles = `
    body { font-family: Arial, Helvetica, sans-serif; font-size: 11pt; color: #1a1a1a; line-height: 1.4; margin: 24px; }
    h1, h2, h3 { font-family: Arial, Helvetica, sans-serif; margin: 18px 0 8px; }
    h1 { font-size: 18pt; }
    h2 { font-size: 14pt; }
    h3 { font-size: 12pt; }
    p { margin: 4px 0; }
    table { border-collapse: collapse; width: 100%; margin: 12px 0 20px; }
    td, th { border: 1px solid #c8c8c8; padding: 8px 10px; vertical-align: top; }
    table tr td:first-child { background: #f3f3f3; font-weight: 600; width: 28%; }
    table tr:first-child td { background: #1a1a1a; color: #ffffff; font-weight: 700; text-align: center; }
    table tr:first-child td:first-child { background: #1a1a1a; color: #ffffff; }
    ul, ol { margin: 4px 0 4px 20px; }
    a { color: #1a73e8; text-decoration: underline; }
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

/**
 * Extract just the body inner HTML from a full HTML document. Useful when
 * writing to the clipboard — Google Docs reads either form, but body-only
 * tends to paste more cleanly without nested <html>/<body> in some browsers.
 *
 * @param {string} fullHtml output of wrapHtmlForPaste
 * @returns {string}
 */
function extractBodyAndStyle(fullHtml) {
  // Naive but robust enough: pull the <style>...</style> block and the
  // <body>...</body> contents. Concatenate them so the inline styles ride
  // along with the body content into the clipboard.
  const styleMatch = fullHtml.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
  const bodyMatch = fullHtml.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  const style = styleMatch ? `<style>${styleMatch[1]}</style>` : '';
  const body = bodyMatch ? bodyMatch[1] : fullHtml;
  return `${style}\n${body}`;
}

/**
 * Strip HTML tags for a plain-text clipboard fallback.
 */
function htmlToPlainText(html) {
  // Quick and dirty: drop tags, collapse whitespace, decode common entities.
  const tmp = html
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
    .replace(/\n{3,}/g, '\n\n');
  return tmp.trim();
}

/**
 * Copy HTML to the clipboard as rich text (`text/html`) with a plain-text
 * fallback (`text/plain`). Google Docs reads the HTML representation and
 * pastes formatted tables.
 *
 * @param {string} fullHtml output of wrapHtmlForPaste
 */
export async function copyHtmlToClipboard(fullHtml) {
  const richPayload = extractBodyAndStyle(fullHtml);
  const plainPayload = htmlToPlainText(richPayload);

  if (!navigator.clipboard || !window.ClipboardItem) {
    // Fallback: plain-text only via execCommand. Loses table structure.
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
 * Download the full HTML document as a `.html` file. Uses the saveAs from
 * FileSaver.js (loaded globally via CDN in index.html).
 *
 * @param {string} fullHtml
 * @param {string} filename should end in `.html`
 */
export function downloadHtmlFile(fullHtml, filename) {
  if (typeof window === 'undefined' || typeof window.saveAs !== 'function') {
    throw new Error('saveAs (FileSaver.js) is not loaded');
  }
  const blob = new Blob([fullHtml], { type: 'text/html;charset=utf-8' });
  window.saveAs(blob, filename);
}

/**
 * Swap a `.docx` filename for `.html` (keeps the rest of the name intact).
 */
export function docxFilenameToHtml(docxFilename) {
  return String(docxFilename).replace(/\.docx$/i, '') + '.html';
}
