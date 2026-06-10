#!/usr/bin/env python3
"""
accept_changes.py — Accept all tracked changes in a .docx without Word.

Part of the word-agent XML fallback chain (used for clean-copy generation
when docx-mcp / word-mcp-live are unavailable).

Handled revision types:
  w:ins / w:moveTo          content kept (wrapper unwrapped)
  w:del / w:moveFrom        content removed
  deleted paragraph marks   paragraph merged with the following one
  w:*Change elements        removed (new formatting accepted)
  move range markers        removed

Processes document.xml plus footnotes, endnotes, headers and footers.

Usage:
    python3 accept_changes.py <input.docx> <output.docx>
    python3 accept_changes.py <input.docx>            # in-place (backup created)
"""

import argparse
import os
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime

from lxml import etree

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = f'{{{W_NS}}}'

CHANGE_TAGS = [
    'rPrChange', 'pPrChange', 'sectPrChange', 'tblPrChange',
    'tblGridChange', 'trPrChange', 'tcPrChange', 'numberingChange',
]
RANGE_MARKERS = [
    'moveFromRangeStart', 'moveFromRangeEnd',
    'moveToRangeStart', 'moveToRangeEnd',
]


def accept_revisions(root):
    """Accept all tracked changes under root. Returns number of revisions applied."""
    count = 0

    # 1. Deleted paragraph marks: merge paragraph into the following one.
    #    (The combined paragraph takes the following paragraph's properties,
    #    matching Word's accept behavior.)
    for p in root.findall(f'.//{W}p'):
        pPr = p.find(f'{W}pPr')
        if pPr is None:
            continue
        rPr = pPr.find(f'{W}rPr')
        if rPr is None or rPr.find(f'{W}del') is None:
            continue
        nxt = p.getnext()
        if nxt is not None and nxt.tag == f'{W}p':
            nxt_pPr = nxt.find(f'{W}pPr')
            insert_at = 1 if nxt_pPr is not None else 0
            for child in reversed([c for c in p if c.tag != f'{W}pPr']):
                nxt.insert(insert_at, child)
            p.getparent().remove(p)
        else:
            rPr.remove(rPr.find(f'{W}del'))
        count += 1

    # 2. Deleted content: drop w:del / w:moveFrom wrappers with their content.
    for tag in ('del', 'moveFrom'):
        for el in root.findall(f'.//{W}{tag}'):
            parent = el.getparent()
            if parent is None:
                continue
            parent.remove(el)
            count += 1

    # 3. Inserted content: unwrap w:ins / w:moveTo, keeping children in place.
    #    Paragraph-mark insertion markers (inside pPr/rPr) are empty — drop them.
    for tag in ('ins', 'moveTo'):
        for el in root.findall(f'.//{W}{tag}'):
            parent = el.getparent()
            if parent is None:
                continue
            idx = parent.index(el)
            for child in reversed(list(el)):
                parent.insert(idx, child)
            parent.remove(el)
            count += 1

    # 4. Formatting change records: accepting means keeping the new
    #    formatting, so the old-format snapshots are simply removed.
    for tag in CHANGE_TAGS + RANGE_MARKERS:
        for el in root.findall(f'.//{W}{tag}'):
            el.getparent().remove(el)
            count += 1

    return count


def process_docx(input_path, output_path):
    parser = etree.XMLParser(remove_blank_text=False)
    total = 0

    with zipfile.ZipFile(input_path, 'r') as zin:
        entries = zin.namelist()
        revisable = [e for e in entries if e == 'word/document.xml'
                     or e in ('word/footnotes.xml', 'word/endnotes.xml')
                     or (e.startswith('word/header') and e.endswith('.xml'))
                     or (e.startswith('word/footer') and e.endswith('.xml'))]

        out_dir = os.path.dirname(os.path.abspath(output_path)) or '.'
        fd, tmp = tempfile.mkstemp(suffix='.docx', dir=out_dir)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for entry in entries:
                    data = zin.read(entry)
                    if entry in revisable:
                        root = etree.fromstring(data, parser)
                        n = accept_revisions(root)
                        if n:
                            total += n
                            data = etree.tostring(
                                root, xml_declaration=True,
                                encoding='UTF-8', standalone=True)
                    zout.writestr(entry, data)
            with zipfile.ZipFile(tmp, 'r') as zf:
                if 'word/document.xml' not in zf.namelist():
                    raise ValueError('output verification failed')
            os.replace(tmp, output_path)
        except Exception:
            if os.path.exists(tmp):
                os.unlink(tmp)
            raise

    return total


def main():
    ap = argparse.ArgumentParser(description='Accept all tracked changes in a .docx')
    ap.add_argument('input', help='Input .docx file')
    ap.add_argument('output', nargs='?', help='Output .docx (default: in-place with backup)')
    ap.add_argument('--no-backup', action='store_true',
                    help='Skip backup when writing in-place')
    args = ap.parse_args()

    if not os.path.exists(args.input):
        print(f'ERROR: file not found: {args.input}', file=sys.stderr)
        sys.exit(1)

    output = args.output or args.input
    if output == args.input and not args.no_backup:
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        base, ext = os.path.splitext(args.input)
        backup = f'{base}_backup_{ts}{ext}'
        shutil.copy2(args.input, backup)
        print(f'Backup: {backup}')

    n = process_docx(args.input, output)
    print(f'Accepted {n} revision element(s)')
    print(f'Output: {output}')


if __name__ == '__main__':
    main()
