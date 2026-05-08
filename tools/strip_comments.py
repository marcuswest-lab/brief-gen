#!/usr/bin/env python3
"""
Strip review comments (Cayce's edits, etc.) from BAD Marketing .docx templates.

Usage:
  python3 strip_comments.py <input.docx> <output.docx>
  python3 strip_comments.py <input.docx>            # writes <input>.clean.docx

Removes:
- word/comments.xml, commentsExtended.xml, commentsIds.xml, commentsExtensible.xml, people.xml
- <w:commentRangeStart>, <w:commentRangeEnd>, <w:commentReference> elements in document.xml
- Relationship entries pointing to those parts in word/_rels/document.xml.rels
- <Override> entries for those content types in [Content_Types].xml
"""

import sys
import os
import re
import zipfile
import shutil

COMMENT_PARTS = {
    'word/comments.xml',
    'word/commentsExtended.xml',
    'word/commentsIds.xml',
    'word/commentsExtensible.xml',
    'word/people.xml',
    'word/threadedComments.xml',
}

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

# Match <w:commentRangeStart .../>, <w:commentRangeEnd .../>, <w:commentReference .../>
# Self-closing only (these elements have no children).
COMMENT_TAG_RE = re.compile(
    r'<w:(?:commentRangeStart|commentRangeEnd|commentReference)\b[^/]*/>',
)

# Match relationship lines pointing to comment parts
COMMENT_REL_RE = re.compile(
    r'<Relationship\b[^>]*Target="(?:[^"]*/)?(?:comments|commentsExtended|commentsIds|commentsExtensible|people|threadedComments)\.xml"[^/]*/>',
)

# Match content-type Override entries for the comment parts
COMMENT_OVERRIDE_RE = re.compile(
    r'<Override\b[^>]*PartName="/word/(?:comments|commentsExtended|commentsIds|commentsExtensible|people|threadedComments)\.xml"[^/]*/>',
)


def strip_comments(input_path, output_path):
    """Read input .docx, write output .docx with all comment data removed."""
    with zipfile.ZipFile(input_path, 'r') as zin:
        names = zin.namelist()
        members = {n: zin.read(n) for n in names}

    # Drop comment part files
    for part in COMMENT_PARTS:
        members.pop(part, None)

    # Strip comment marker elements from document.xml (and headers/footers, just in case)
    for name in list(members.keys()):
        if name.endswith('.xml') and (
            name == 'word/document.xml'
            or name.startswith('word/header')
            or name.startswith('word/footer')
        ):
            text = members[name].decode('utf-8')
            cleaned = COMMENT_TAG_RE.sub('', text)
            members[name] = cleaned.encode('utf-8')

    # Strip relationship entries pointing to comment parts
    rels_name = 'word/_rels/document.xml.rels'
    if rels_name in members:
        text = members[rels_name].decode('utf-8')
        members[rels_name] = COMMENT_REL_RE.sub('', text).encode('utf-8')

    # Strip content-type Override entries for comment parts
    ct_name = '[Content_Types].xml'
    if ct_name in members:
        text = members[ct_name].decode('utf-8')
        members[ct_name] = COMMENT_OVERRIDE_RE.sub('', text).encode('utf-8')

    # Write output
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_path = sys.argv[1]
    if not os.path.exists(input_path):
        print(f'ERROR: file not found: {input_path}')
        sys.exit(1)

    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        base, ext = os.path.splitext(input_path)
        output_path = f'{base}.clean{ext}'

    strip_comments(input_path, output_path)
    print(f'Stripped comments: {input_path} -> {output_path}')


if __name__ == '__main__':
    main()
