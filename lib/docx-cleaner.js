// Runtime safety: strip review comments from a .docx zip object before saving.
// Templates committed in /templates/ are pre-cleaned by tools/strip_comments.py,
// but this is here as defense-in-depth in case a template is replaced with one
// that still has comments.

const COMMENT_PARTS = [
  'word/comments.xml',
  'word/commentsExtended.xml',
  'word/commentsIds.xml',
  'word/commentsExtensible.xml',
  'word/people.xml',
  'word/threadedComments.xml',
];

const COMMENT_TAG_RE = /<w:(?:commentRangeStart|commentRangeEnd|commentReference)\b[^/]*\/>/g;
const COMMENT_REL_RE = /<Relationship\b[^>]*Target="(?:[^"]*\/)?(?:comments|commentsExtended|commentsIds|commentsExtensible|people|threadedComments)\.xml"[^/]*\/>/g;
const COMMENT_OVERRIDE_RE = /<Override\b[^>]*PartName="\/word\/(?:comments|commentsExtended|commentsIds|commentsExtensible|people|threadedComments)\.xml"[^/]*\/>/g;

/**
 * Strip review comments from a JSZip docx object in place.
 * @param {JSZip} zip
 */
export async function stripComments(zip) {
  // Drop comment part files
  for (const part of COMMENT_PARTS) {
    if (zip.file(part)) zip.remove(part);
  }

  // Strip comment marker elements from document.xml + headers/footers
  const xmlNames = Object.keys(zip.files).filter(n =>
    n === 'word/document.xml' ||
    n.startsWith('word/header') ||
    n.startsWith('word/footer')
  );
  for (const name of xmlNames) {
    if (!name.endsWith('.xml')) continue;
    const text = await zip.file(name).async('string');
    const cleaned = text.replace(COMMENT_TAG_RE, '');
    if (cleaned !== text) zip.file(name, cleaned);
  }

  // Strip relationship entries
  const relsName = 'word/_rels/document.xml.rels';
  if (zip.file(relsName)) {
    const text = await zip.file(relsName).async('string');
    const cleaned = text.replace(COMMENT_REL_RE, '');
    if (cleaned !== text) zip.file(relsName, cleaned);
  }

  // Strip content-type Override entries
  const ctName = '[Content_Types].xml';
  if (zip.file(ctName)) {
    const text = await zip.file(ctName).async('string');
    const cleaned = text.replace(COMMENT_OVERRIDE_RE, '');
    if (cleaned !== text) zip.file(ctName, cleaned);
  }
}
