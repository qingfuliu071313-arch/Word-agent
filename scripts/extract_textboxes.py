#!/usr/bin/env python3
"""Extract text-box (w:txbxContent) text from ALL XML parts of a .docx.

Standard text extraction (get_document_text, docx2python) skips text boxes
entirely. Worse, text boxes frequently live OUTSIDE word/document.xml —
academic format templates often place formatting annotations in header
text boxes (e.g. word/header2.xml), so scanning document.xml alone misses
them. This script scans every XML part in the package: document.xml,
header*.xml, footer*.xml, footnotes.xml, endnotes.xml, etc.

Covers both DrawingML (wps) and VML text boxes — both wrap their text in
w:txbxContent. DrawingML boxes carry a duplicate VML copy inside
mc:Fallback; fallbacks are stripped first so each box is counted once.

Usage:
    python3 extract_textboxes.py FILE.docx          # JSON output
    python3 extract_textboxes.py FILE.docx --text   # human-readable
"""

import argparse
import json
import sys
import zipfile
import xml.etree.ElementTree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'


def strip_fallbacks(root):
    parent_map = {child: parent for parent in root.iter() for child in parent}
    for fb in list(root.iter(f'{{{MC}}}Fallback')):
        parent = parent_map.get(fb)
        if parent is not None:
            parent.remove(fb)


def extract(docx_path):
    results = []
    parts_with_boxes = []
    with zipfile.ZipFile(docx_path) as z:
        parts = sorted(
            n for n in z.namelist()
            if n.startswith('word/') and n.endswith('.xml')
        )
        for part in parts:
            data = z.read(part)
            if b'txbxContent' not in data:
                continue
            try:
                root = ET.fromstring(data)
            except ET.ParseError:
                continue
            strip_fallbacks(root)
            found_in_part = 0
            for txbx in root.iter(f'{{{W}}}txbxContent'):
                paragraphs = []
                for p in txbx.iter(f'{{{W}}}p'):
                    text = ''.join(t.text or '' for t in p.iter(f'{{{W}}}t'))
                    if text.strip():
                        paragraphs.append(text)
                if paragraphs:
                    found_in_part += 1
                    results.append({
                        'index': len(results) + 1,
                        'part': part,
                        'paragraphs': paragraphs,
                    })
            if found_in_part:
                parts_with_boxes.append({'part': part, 'count': found_in_part})
    return results, parts_with_boxes


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('docx', help='Path to .docx file')
    parser.add_argument('--text', action='store_true',
                        help='Human-readable output instead of JSON')
    args = parser.parse_args()

    try:
        results, parts_with_boxes = extract(args.docx)
    except (zipfile.BadZipFile, FileNotFoundError, KeyError) as e:
        print(json.dumps({'error': str(e)}, ensure_ascii=False))
        sys.exit(1)

    if args.text:
        if not results:
            print('No text boxes with content found.')
            return
        for box in results:
            print(f"[TextBox {box['index']}] ({box['part']})")
            for t in box['paragraphs']:
                print(f'  {t}')
            print()
    else:
        print(json.dumps({
            'textbox_count': len(results),
            'parts_with_textboxes': parts_with_boxes,
            'textboxes': results,
        }, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
