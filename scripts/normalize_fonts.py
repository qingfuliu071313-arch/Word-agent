#!/usr/bin/env python3
"""
Font Normalization Script for Word Documents
Scans all runs in a .docx and ensures w:rFonts eastAsia/ascii/hAnsi are consistently paired.
Also fixes theme font references and runs missing font attributes entirely.

Usage:
    python3 normalize_fonts.py input.docx [output.docx] [--cn 宋体] [--en "Times New Roman"] [--detect-only]

Examples:
    python3 normalize_fonts.py paper.docx                          # Fix in-place
    python3 normalize_fonts.py paper.docx paper_fixed.docx         # Fix to new file
    python3 normalize_fonts.py paper.docx --detect-only            # Only report issues
    python3 normalize_fonts.py paper.docx --cn 黑体 --en Arial     # Custom pairing
    python3 normalize_fonts.py paper.docx --unify                  # Force ALL to same pair
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import sys
import argparse
import tempfile
import json
import re

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = f'{{{W_NS}}}'

THEME_ATTRS = ['asciiTheme', 'eastAsiaTheme', 'hAnsiTheme', 'cstheme']
THEME_ATTRS_W = [f'{W}{a}' for a in THEME_ATTRS]

FONT_PAIRS_CN_TO_EN = {
    '宋体': 'Times New Roman',
    'SimSun': 'Times New Roman',
    '黑体': 'Arial',
    'SimHei': 'Arial',
    '楷体': 'Times New Roman',
    'KaiTi': 'Times New Roman',
    '仿宋': 'Times New Roman',
    'FangSong': 'Times New Roman',
    '微软雅黑': 'Arial',
    'Microsoft YaHei': 'Arial',
    '新宋体': 'Times New Roman',
    'NSimSun': 'Times New Roman',
}

FONT_PAIRS_EN_TO_CN = {
    'Times New Roman': '宋体',
    'Arial': '黑体',
    'Calibri': '宋体',
    'Cambria': '宋体',
    'Helvetica': '黑体',
    'Verdana': '黑体',
}

ALL_NS = {}


def collect_namespaces(xml_bytes):
    """Collect all namespace declarations from the XML to preserve them on write."""
    for match in re.finditer(rb'xmlns:(\w+)="([^"]+)"', xml_bytes):
        prefix = match.group(1).decode()
        uri = match.group(2).decode()
        if prefix not in ALL_NS:
            ALL_NS[prefix] = uri
            ET.register_namespace(prefix, uri)
    default_match = re.search(rb'xmlns="([^"]+)"', xml_bytes)
    if default_match:
        uri = default_match.group(1).decode()
        if uri not in ALL_NS.values():
            ET.register_namespace('', uri)


def serialize_xml(original_bytes, modified_root):
    """Serialize modified ET tree while preserving the original XML declaration and
    root element namespace declarations. ElementTree drops standalone='yes' and
    strips unused namespace prefixes, which causes Word to flag the file as corrupt."""
    original_str = original_bytes.decode('utf-8')

    xml_decl_match = re.match(r'<\?xml[^?]*\?>\s*', original_str)
    if xml_decl_match:
        original_decl = xml_decl_match.group(0)
    else:
        original_decl = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>\n"

    after_decl = original_str[len(original_decl):] if xml_decl_match else original_str
    root_tag_match = re.match(r'<[^>]+>', after_decl)
    original_root_tag = root_tag_match.group(0) if root_tag_match else None

    new_xml = ET.tostring(modified_root, encoding='unicode')

    if original_root_tag:
        new_root_tag_match = re.match(r'<[^>]+>', new_xml)
        if new_root_tag_match:
            new_xml = original_root_tag + new_xml[new_root_tag_match.end():]

    return original_decl + new_xml


def _strip_theme_attrs(rFonts):
    """Remove all *Theme attributes from a w:rFonts element. Returns True if any were removed."""
    removed = False
    for attr in THEME_ATTRS_W:
        if attr in rFonts.attrib:
            del rFonts.attrib[attr]
            removed = True
    return removed


def _has_theme_attrs(rFonts):
    """Check if rFonts has any theme-based attributes."""
    return any(attr in rFonts.attrib for attr in THEME_ATTRS_W)


def _run_has_text(run):
    """Check if a run contains any visible text."""
    for t in run.findall(f'{W}t'):
        if t.text and t.text.strip():
            return True
    return False


def detect_font_issues(file_path):
    """Scan document for ALL font issues: missing, mismatched, theme-based, and bare runs."""
    with zipfile.ZipFile(file_path, 'r') as z:
        xml_content = z.read('word/document.xml')
        styles_xml = z.read('word/styles.xml') if 'word/styles.xml' in z.namelist() else None

    root = ET.fromstring(xml_content)
    issues = []

    # --- Check styles.xml for theme font references ---
    if styles_xml:
        styles_root = ET.fromstring(styles_xml)
        doc_defaults = styles_root.find(f'{W}docDefaults')
        if doc_defaults is not None:
            rPr = doc_defaults.find(f'.//{W}rPr')
            if rPr is not None:
                rFonts = rPr.find(f'{W}rFonts')
                if rFonts is not None and _has_theme_attrs(rFonts):
                    issues.append({
                        'paragraph': -1,
                        'type': 'theme_in_docDefaults',
                        'detail': 'docDefaults uses theme font references instead of explicit fonts',
                        'preview': '[styles.xml docDefaults]'
                    })

        theme_style_count = 0
        for rFonts in styles_root.findall(f'.//{W}rFonts'):
            if _has_theme_attrs(rFonts):
                theme_style_count += 1
        if theme_style_count > 0:
            issues.append({
                'paragraph': -1,
                'type': 'theme_in_styles',
                'detail': f'{theme_style_count} rFonts element(s) in styles.xml use theme font references',
                'preview': '[styles.xml]'
            })

    # --- Check document.xml runs ---
    for i, para in enumerate(root.findall(f'.//{W}p')):
        fonts_ascii = set()
        fonts_east = set()
        fonts_hansi = set()
        missing_east = False
        missing_ascii = False
        has_theme = False
        bare_runs = 0

        for run in para.findall(f'{W}r'):
            if not _run_has_text(run):
                continue

            rPr = run.find(f'{W}rPr')
            if rPr is None:
                bare_runs += 1
                continue

            rFonts = rPr.find(f'{W}rFonts')
            if rFonts is None:
                bare_runs += 1
                continue

            if _has_theme_attrs(rFonts):
                has_theme = True

            a = rFonts.get(f'{W}ascii')
            e = rFonts.get(f'{W}eastAsia')
            h = rFonts.get(f'{W}hAnsi')
            if a:
                fonts_ascii.add(a)
            if e:
                fonts_east.add(e)
            if h:
                fonts_hansi.add(h)
            if a and not e:
                missing_east = True
            if e and not a:
                missing_ascii = True

        texts = para.findall(f'.//{W}t')
        preview = ''.join(t.text or '' for t in texts)[:60]
        if not preview.strip():
            continue

        if bare_runs > 0:
            issues.append({
                'paragraph': i,
                'type': 'no_font_spec',
                'detail': f'{bare_runs} run(s) with text but no w:rFonts — inherits theme/style fonts',
                'preview': preview
            })
        if has_theme:
            issues.append({
                'paragraph': i,
                'type': 'theme_font_ref',
                'detail': 'run uses *Theme attributes that override explicit font names',
                'preview': preview
            })
        if missing_east:
            issues.append({
                'paragraph': i,
                'type': 'missing_eastAsia',
                'detail': f'ascii={list(fonts_ascii)} but no eastAsia set',
                'preview': preview
            })
        if missing_ascii:
            issues.append({
                'paragraph': i,
                'type': 'missing_ascii',
                'detail': f'eastAsia={list(fonts_east)} but no ascii set',
                'preview': preview
            })
        if len(fonts_ascii) > 1:
            issues.append({
                'paragraph': i,
                'type': 'mixed_ascii',
                'detail': f'multiple ascii fonts: {list(fonts_ascii)}',
                'preview': preview
            })
        if len(fonts_east) > 1:
            issues.append({
                'paragraph': i,
                'type': 'mixed_eastAsia',
                'detail': f'multiple eastAsia fonts: {list(fonts_east)}',
                'preview': preview
            })
        if fonts_hansi and fonts_ascii and fonts_hansi != fonts_ascii:
            issues.append({
                'paragraph': i,
                'type': 'hansi_mismatch',
                'detail': f'ascii={list(fonts_ascii)} but hAnsi={list(fonts_hansi)}',
                'preview': preview
            })

    return issues


def normalize_fonts(file_path, output_path=None, cn_font='宋体', en_font='Times New Roman', unify=False):
    """Normalize all font references in the document.

    Fixes three layers:
    1. styles.xml — replace theme refs in docDefaults and styles with explicit fonts
    2. document.xml runs WITH w:rFonts — fix pairings, strip theme attrs
    3. document.xml runs WITHOUT w:rFonts — inject explicit font attributes
    """
    if output_path is None:
        output_path = file_path

    with zipfile.ZipFile(file_path, 'r') as z:
        xml_content = z.read('word/document.xml')
        styles_xml = z.read('word/styles.xml') if 'word/styles.xml' in z.namelist() else None
        all_entries = z.namelist()

    collect_namespaces(xml_content)
    if styles_xml:
        collect_namespaces(styles_xml)

    root = ET.fromstring(xml_content)
    fixed_count = 0

    # ── Layer 1: Fix styles.xml ──
    styles_root = None
    styles_changed = False
    if styles_xml:
        styles_root = ET.fromstring(styles_xml)

        # Fix docDefaults
        doc_defaults = styles_root.find(f'{W}docDefaults')
        if doc_defaults is not None:
            rPrDefault = doc_defaults.find(f'{W}rPrDefault')
            if rPrDefault is not None:
                rPr = rPrDefault.find(f'{W}rPr')
                if rPr is not None:
                    rFonts = rPr.find(f'{W}rFonts')
                    if rFonts is not None:
                        if _strip_theme_attrs(rFonts):
                            styles_changed = True
                        rFonts.set(f'{W}ascii', en_font)
                        rFonts.set(f'{W}eastAsia', cn_font)
                        rFonts.set(f'{W}hAnsi', en_font)
                        rFonts.set(f'{W}cs', en_font)
                        styles_changed = True
                        fixed_count += 1

        # Fix ALL rFonts in styles.xml (including nested tblStylePr elements)
        for rFonts in styles_root.findall(f'.//{W}rFonts'):
            if _has_theme_attrs(rFonts) or unify:
                _strip_theme_attrs(rFonts)
                if unify:
                    rFonts.set(f'{W}ascii', en_font)
                    rFonts.set(f'{W}eastAsia', cn_font)
                    rFonts.set(f'{W}hAnsi', en_font)
                else:
                    if not rFonts.get(f'{W}ascii'):
                        rFonts.set(f'{W}ascii', en_font)
                    if not rFonts.get(f'{W}eastAsia'):
                        rFonts.set(f'{W}eastAsia', cn_font)
                    if not rFonts.get(f'{W}hAnsi'):
                        rFonts.set(f'{W}hAnsi', rFonts.get(f'{W}ascii', en_font))

                styles_changed = True
                fixed_count += 1

    # ── Layer 2: Fix existing w:rFonts in document.xml ──
    for rFonts in root.findall(f'.//{W}rFonts'):
        east = rFonts.get(f'{W}eastAsia')
        ascii_font = rFonts.get(f'{W}ascii')
        hansi = rFonts.get(f'{W}hAnsi')
        changed = False

        # Always strip theme attrs — they override explicit values
        if _strip_theme_attrs(rFonts):
            changed = True

        if unify:
            if ascii_font != en_font:
                rFonts.set(f'{W}ascii', en_font)
                changed = True
            if east != cn_font:
                rFonts.set(f'{W}eastAsia', cn_font)
                changed = True
            if hansi != en_font:
                rFonts.set(f'{W}hAnsi', en_font)
                changed = True
        else:
            if ascii_font and not east:
                paired = FONT_PAIRS_EN_TO_CN.get(ascii_font, cn_font)
                rFonts.set(f'{W}eastAsia', paired)
                changed = True
            if east and not ascii_font:
                paired = FONT_PAIRS_CN_TO_EN.get(east, en_font)
                rFonts.set(f'{W}ascii', paired)
                rFonts.set(f'{W}hAnsi', paired)
                changed = True

            ascii_font = rFonts.get(f'{W}ascii')
            hansi = rFonts.get(f'{W}hAnsi')
            if ascii_font and hansi != ascii_font:
                rFonts.set(f'{W}hAnsi', ascii_font)
                changed = True

        if changed:
            fixed_count += 1

    # ── Layer 3: Inject w:rFonts into bare runs (no rPr or no rFonts) ──
    for run in root.findall(f'.//{W}r'):
        if not _run_has_text(run):
            continue

        rPr = run.find(f'{W}rPr')
        if rPr is None:
            rPr = ET.SubElement(run, f'{W}rPr')
            run.remove(rPr)
            run.insert(0, rPr)

        rFonts = rPr.find(f'{W}rFonts')
        if rFonts is None:
            rFonts = ET.SubElement(rPr, f'{W}rFonts')
            rFonts.set(f'{W}ascii', en_font)
            rFonts.set(f'{W}eastAsia', cn_font)
            rFonts.set(f'{W}hAnsi', en_font)
            fixed_count += 1

    # ── Write back (preserve original XML declarations and namespaces) ──
    xml_out = serialize_xml(xml_content, root)
    styles_out = serialize_xml(styles_xml, styles_root) if styles_changed and styles_root is not None else None

    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(file_path, 'r') as z:
            z.extractall(tmpdir)

        doc_xml_path = os.path.join(tmpdir, 'word', 'document.xml')
        with open(doc_xml_path, 'w', encoding='utf-8') as f:
            f.write(xml_out)

        if styles_out:
            styles_path = os.path.join(tmpdir, 'word', 'styles.xml')
            with open(styles_path, 'w', encoding='utf-8') as f:
                f.write(styles_out)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for entry in all_entries:
                full_path = os.path.join(tmpdir, entry)
                if os.path.isfile(full_path):
                    zout.write(full_path, entry)

    return fixed_count


def extract_textboxes(file_path):
    """Extract text content from all text boxes (w:txbxContent)."""
    with zipfile.ZipFile(file_path, 'r') as z:
        xml_content = z.read('word/document.xml')

    root = ET.fromstring(xml_content)
    textboxes = root.findall(f'.//{W}txbxContent')

    results = []
    for i, txbx in enumerate(textboxes):
        paragraphs = []
        for p in txbx.findall(f'.//{W}p'):
            runs = p.findall(f'.//{W}t')
            text = ''.join(r.text or '' for r in runs)
            if text.strip():
                paragraphs.append(text)
        if paragraphs:
            results.append({'index': i + 1, 'content': paragraphs})

    return results


def main():
    parser = argparse.ArgumentParser(description='Normalize CJK/Latin font pairing in Word documents')
    parser.add_argument('input', help='Input .docx file')
    parser.add_argument('output', nargs='?', help='Output .docx file (default: overwrite input)')
    parser.add_argument('--cn', default='宋体', help='Default Chinese font (default: 宋体)')
    parser.add_argument('--en', default='Times New Roman', help='Default English font (default: Times New Roman)')
    parser.add_argument('--detect-only', action='store_true', help='Only detect issues, do not fix')
    parser.add_argument('--unify', action='store_true', help='Force ALL runs to the same font pair (aggressive mode for academic papers)')
    parser.add_argument('--textboxes', action='store_true', help='Extract and display text box content')
    parser.add_argument('--json', action='store_true', help='Output results as JSON')

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    # Text box extraction
    if args.textboxes:
        tbs = extract_textboxes(args.input)
        if args.json:
            print(json.dumps({'textbox_count': len(tbs), 'textboxes': tbs}, ensure_ascii=False, indent=2))
        else:
            if not tbs:
                print("No text boxes found.")
            else:
                print(f"Found {len(tbs)} text box(es):\n")
                for tb in tbs:
                    print(f"[TextBox {tb['index']}]")
                    for line in tb['content']:
                        print(f"  {line}")
                    print()
        return

    # Detection
    issues = detect_font_issues(args.input)

    if args.json:
        if args.detect_only:
            print(json.dumps({'issue_count': len(issues), 'issues': issues}, ensure_ascii=False, indent=2))
            return
    else:
        if issues:
            print(f"Found {len(issues)} font issue(s):\n")
            for issue in issues:
                loc = f"P{issue['paragraph']}" if issue['paragraph'] >= 0 else "STYLE"
                print(f"  {loc}: [{issue['type']}] {issue['detail']}")
                print(f"    \"{issue['preview']}\"")
            print()
        else:
            print("No font issues detected.\n")

    if args.detect_only:
        return

    # Normalization
    output = args.output or args.input
    fixed = normalize_fonts(args.input, output, cn_font=args.cn, en_font=args.en, unify=args.unify)

    if args.json:
        result = {'issues_found': len(issues), 'runs_fixed': fixed, 'output': output}
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print(f"Fixed {fixed} font attribute(s) with pairing: {args.cn} + {args.en}")
        print(f"Output: {output}")

    # Verify
    remaining = detect_font_issues(output)
    if remaining:
        if not args.json:
            print(f"\nWarning: {len(remaining)} issue(s) remain after normalization:")
            for issue in remaining:
                loc = f"P{issue['paragraph']}" if issue['paragraph'] >= 0 else "STYLE"
                print(f"  {loc}: [{issue['type']}] {issue['detail']}")
    else:
        if not args.json:
            print("Verification: all font issues resolved.")


if __name__ == '__main__':
    main()
