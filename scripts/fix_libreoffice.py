#!/usr/bin/env python3
"""
Post-conversion fixer for LibreOffice .doc → .docx output.
Repairs all known artifacts that cause format corruption when opened in MS Word.

Usage:
    python3 fix_libreoffice.py input.docx [output.docx] [--detect-only] [--json]
    python3 fix_libreoffice.py input.docx --skip-font-normalize
    python3 fix_libreoffice.py input.docx --unify-fonts   # aggressive font unification

This script automatically calls normalize_fonts.py (pairing mode) as the final
step; pass --unify-fonts to force a single font pair on all runs instead.
"""

import zipfile
import xml.etree.ElementTree as ET
import os
import sys
import argparse
import tempfile
import json
import re
import subprocess

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
W = f'{{{W_NS}}}'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
R = f'{{{R_NS}}}'
WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'

FONT_MAP = {
    'Liberation Serif': 'Times New Roman',
    'Liberation Sans': 'Arial',
    'Liberation Mono': 'Courier New',
    'NSimSun': '宋体',
    '新宋体': '宋体',
}

FONT_SEMICOLONS = {
    '宋体;SimSun': '宋体',
    '黑体;SimHei': '黑体',
    '楷体_GB2312;KaiTi': '楷体',
    '楷体_GB2312': '楷体',
    '仿宋_GB2312;FangSong': '仿宋',
    '仿宋_GB2312': '仿宋',
    'SimSun;宋体': '宋体',
    'SimHei;黑体': '黑体',
}

ALL_NS = {}


def collect_namespaces(xml_bytes):
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
    """Serialize modified ET tree preserving original XML declaration and root namespaces.
    ElementTree drops standalone='yes' and strips unused namespace prefixes,
    which causes Word to flag the file as corrupt."""
    original_str = original_bytes.decode('utf-8') if isinstance(original_bytes, bytes) else original_bytes

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

    return (original_decl + new_xml).encode('utf-8')


def _emu_to_cm(emu):
    return emu / 914400 * 2.54


def _cm_to_emu(cm):
    return int(cm / 2.54 * 914400)


def _twips_to_pt(twips):
    return twips / 20


def _pt_to_twips(pt):
    return int(pt * 20)


# ── Fix 1: Settings (compatibilityMode, characterSpacingControl) ──

def fix_settings(settings_xml):
    """Fix compatibilityMode and add CJK spacing control."""
    fixes = []
    content = settings_xml.decode('utf-8')
    original = content

    compat_match = re.search(r'(compatibilityMode[^/]*?w:val=")(\d+)(")', content)
    if compat_match and int(compat_match.group(2)) < 15:
        content = content[:compat_match.start(2)] + '15' + content[compat_match.end(2):]
        fixes.append(f'compatibilityMode {compat_match.group(2)}→15')

    if 'characterSpacingControl' not in content:
        content = content.replace(
            '</w:settings>',
            '<w:characterSpacingControl w:val="compressPunctuation"/></w:settings>'
        )
        fixes.append('added characterSpacingControl=compressPunctuation')

    if content != original:
        return content.encode('utf-8'), fixes
    return None, fixes


# ── Fix 2: Font name mapping (Liberation fonts, semicolon fallbacks) ──

def fix_font_names_in_xml(xml_bytes):
    """Replace LibreOffice-specific font names with standard ones."""
    fixes = []
    content = xml_bytes.decode('utf-8')
    original = content

    for old, new in FONT_MAP.items():
        if old in content:
            count = content.count(old)
            content = content.replace(old, new)
            fixes.append(f'font: {old}→{new} ({count}x)')

    for old, new in FONT_SEMICOLONS.items():
        if old in content:
            count = content.count(old)
            content = content.replace(old, new)
            fixes.append(f'font fallback: {old}→{new} ({count}x)')

    # Catch any remaining semicolon patterns in font names
    def clean_semicolon_font(match):
        full = match.group(0)
        attr = match.group(1)
        val = match.group(2)
        if ';' in val:
            cleaned = val.split(';')[0]
            fixes.append(f'stripped fallback: {val}→{cleaned}')
            return f'{attr}"{cleaned}"'
        return full

    content = re.sub(r'(w:(?:ascii|hAnsi|eastAsia|cs)=")([^"]*;[^"]*)"', clean_semicolon_font, content)

    if content != original:
        return content.encode('utf-8'), fixes
    return None, fixes


# ── Fix 3: Style chain repair ──

def fix_styles(styles_xml):
    """Fix broken style inheritance and LibreOffice style issues."""
    fixes = []
    content = styles_xml.decode('utf-8')
    original = content

    # Heading basedOn TOC → Normal
    for heading_match in re.finditer(
        r'(<w:style[^>]*styleId="Heading(\d+)".*?<w:basedOn w:val=")TOC\d+(")',
        content, re.DOTALL
    ):
        old_fragment = heading_match.group(0)
        new_fragment = heading_match.group(1) + 'Normal' + heading_match.group(3)
        content = content.replace(old_fragment, new_fragment)
        fixes.append(f'Heading{heading_match.group(2)}: basedOn TOC→Normal')

    # Fix styles that inherit from non-existent styles
    style_ids = set(re.findall(r'w:styleId="([^"]+)"', content))
    for based_on in re.finditer(r'<w:basedOn w:val="([^"]+)"', content):
        parent = based_on.group(1)
        if parent not in style_ids and parent not in ('Normal', 'DefaultParagraphFont', 'TableNormal'):
            fixes.append(f'warning: style basedOn "{parent}" not found in document')

    if content != original:
        return content.encode('utf-8'), fixes
    return None, fixes


# ── Fix 4: Table normalization ──

def fix_tables(doc_xml_bytes):
    """Fix table widths, borders, and cell properties after conversion."""
    fixes = []
    collect_namespaces(doc_xml_bytes)
    root = ET.fromstring(doc_xml_bytes)

    for tbl_idx, tbl in enumerate(root.findall(f'.//{W}tbl')):
        tblPr = tbl.find(f'{W}tblPr')
        if tblPr is None:
            continue

        # Fix table width: if 0 or missing, set to 100% (5000 fifths of percent)
        tblW = tblPr.find(f'{W}tblW')
        if tblW is not None:
            w_val = tblW.get(f'{W}w', '0')
            w_type = tblW.get(f'{W}type', 'auto')
            if w_type == 'auto' or w_val == '0':
                tblW.set(f'{W}w', '5000')
                tblW.set(f'{W}type', 'pct')
                fixes.append(f'table[{tbl_idx}]: width auto/0→100%')

        # Ensure table has borders (LibreOffice sometimes drops them)
        tblBorders = tblPr.find(f'{W}tblBorders')
        if tblBorders is None:
            tblBorders = ET.SubElement(tblPr, f'{W}tblBorders')
            for edge in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
                border = ET.SubElement(tblBorders, f'{W}{edge}')
                border.set(f'{W}val', 'single')
                border.set(f'{W}sz', '4')
                border.set(f'{W}space', '0')
                border.set(f'{W}color', 'auto')
            fixes.append(f'table[{tbl_idx}]: added missing borders')

        # Fix zero-width columns
        tblGrid = tbl.find(f'{W}tblGrid')
        if tblGrid is not None:
            cols = tblGrid.findall(f'{W}gridCol')
            zero_cols = [c for c in cols if c.get(f'{W}w', '0') == '0']
            if zero_cols and cols:
                # Distribute evenly: assume A4 body width ~9072 twips (15.98cm)
                total_width = 9072
                non_zero = [c for c in cols if c.get(f'{W}w', '0') != '0']
                if non_zero:
                    used = sum(int(c.get(f'{W}w', '0')) for c in non_zero)
                    remaining = max(total_width - used, len(zero_cols) * 500)
                    per_col = remaining // len(zero_cols)
                else:
                    per_col = total_width // len(cols)
                for c in zero_cols:
                    c.set(f'{W}w', str(per_col))
                fixes.append(f'table[{tbl_idx}]: fixed {len(zero_cols)} zero-width column(s)')

        # Fix cell widths: ensure every tc has a tcW
        for row in tbl.findall(f'{W}tr'):
            for cell in row.findall(f'{W}tc'):
                tcPr = cell.find(f'{W}tcPr')
                if tcPr is not None:
                    tcW = tcPr.find(f'{W}tcW')
                    if tcW is not None:
                        cw = tcW.get(f'{W}w', '0')
                        if cw == '0':
                            tcW.set(f'{W}w', '1500')
                            tcW.set(f'{W}type', 'dxa')

    if fixes:
        return serialize_xml(doc_xml_bytes, root), fixes
    return None, fixes


# ── Fix 5: Image reference validation ──

def fix_images(doc_xml_bytes, rels_xml_bytes, media_files):
    """Validate image references and fix broken ones."""
    fixes = []
    if not rels_xml_bytes:
        return None, fixes

    rels_root = ET.fromstring(rels_xml_bytes)
    ns_rels = 'http://schemas.openxmlformats.org/package/2006/relationships'

    image_rels = {}
    for rel in rels_root.findall(f'{{{ns_rels}}}Relationship'):
        rel_type = rel.get('Type', '')
        if 'image' in rel_type:
            rid = rel.get('Id', '')
            target = rel.get('Target', '')
            image_rels[rid] = target

    for rid, target in image_rels.items():
        full_path = f'word/{target}' if not target.startswith('/') else target.lstrip('/')
        if full_path not in media_files and target not in media_files:
            # Try common renames
            basename = os.path.basename(target)
            candidates = [f for f in media_files if os.path.basename(f) == basename]
            if candidates:
                # Detection only — the rels file is not rewritten here.
                fixes.append(f'image {rid}: WARNING broken path {target} '
                             f'(possible match: {candidates[0]}, not auto-fixed)')
            else:
                fixes.append(f'image {rid}: WARNING broken reference to {target}')

    return None, fixes


# ── Fix 6: Spacing normalization ──

def fix_spacing(doc_xml_bytes):
    """Fix paragraph spacing issues from conversion."""
    fixes = []
    root = ET.fromstring(doc_xml_bytes)
    fix_count = 0

    for pPr in root.findall(f'.//{W}pPr'):
        spacing = pPr.find(f'{W}spacing')
        if spacing is None:
            continue

        # Fix abnormally large before/after spacing (>= 600 twips = 30pt)
        for attr in ['before', 'after']:
            val = spacing.get(f'{W}{attr}')
            if val and val.isdigit() and int(val) > 600:
                # Cap at 480 twips (24pt) which is the max common academic spacing
                spacing.set(f'{W}{attr}', '480')
                fix_count += 1

        # Fix line spacing: if lineRule="exact" and line < 200 twips (10pt),
        # it's likely a conversion error — switch to proportional
        line_rule = spacing.get(f'{W}lineRule', '')
        line_val = spacing.get(f'{W}line', '')
        if line_rule == 'exact' and line_val.isdigit():
            val = int(line_val)
            if val < 200:
                spacing.set(f'{W}lineRule', 'auto')
                spacing.set(f'{W}line', '360')  # 1.5x line spacing
                fix_count += 1

    if fix_count > 0:
        fixes.append(f'fixed {fix_count} paragraph spacing issue(s)')
        return serialize_xml(doc_xml_bytes, root), fixes
    return None, fixes


# ── Fix 7: Page setup validation ──

def fix_page_setup(doc_xml_bytes):
    """Ensure page setup is reasonable (not corrupted margins/size)."""
    fixes = []
    root = ET.fromstring(doc_xml_bytes)

    for sectPr in root.findall(f'.//{W}sectPr'):
        pgSz = sectPr.find(f'{W}pgSz')
        pgMar = sectPr.find(f'{W}pgMar')

        if pgSz is not None:
            w = int(pgSz.get(f'{W}w', '0'))
            h = int(pgSz.get(f'{W}h', '0'))
            # A4 = 11906x16838 twips, Letter = 12240x15840
            # If dimensions are 0 or wildly wrong, set to A4
            if w == 0 or h == 0 or w > 30000 or h > 30000 or w < 5000 or h < 5000:
                pgSz.set(f'{W}w', '11906')
                pgSz.set(f'{W}h', '16838')
                fixes.append(f'page size: {w}x{h}→A4 (11906x16838)')

        if pgMar is not None:
            # Check for zero or absurd margins
            for attr in ['top', 'bottom', 'left', 'right']:
                val = pgMar.get(f'{W}{attr}', '0')
                if val.lstrip('-').isdigit():
                    v = int(val)
                    # Margins should be between 360 (0.25in) and 4320 (3in) twips
                    if v < 360 or v > 4320:
                        default = '1440' if attr in ('top', 'bottom') else '1800'
                        pgMar.set(f'{W}{attr}', default)
                        fixes.append(f'margin {attr}: {val}→{default} twips')

    if fixes:
        return serialize_xml(doc_xml_bytes, root), fixes
    return None, fixes


# ── Main pipeline ──

def detect_issues(file_path):
    """Scan for all conversion artifacts without fixing."""
    all_issues = []

    with zipfile.ZipFile(file_path, 'r') as z:
        names = z.namelist()

        if 'word/settings.xml' in names:
            _, fixes = fix_settings(z.read('word/settings.xml'))
            all_issues.extend([{'type': 'settings', 'detail': f} for f in fixes])

        xml_files = [n for n in names if n.endswith('.xml') and n.startswith('word/')]
        for xf in xml_files:
            _, fixes = fix_font_names_in_xml(z.read(xf))
            all_issues.extend([{'type': 'font_name', 'file': xf, 'detail': f} for f in fixes])

        if 'word/styles.xml' in names:
            _, fixes = fix_styles(z.read('word/styles.xml'))
            all_issues.extend([{'type': 'style', 'detail': f} for f in fixes])

        if 'word/document.xml' in names:
            doc_xml = z.read('word/document.xml')
            _, fixes = fix_tables(doc_xml)
            all_issues.extend([{'type': 'table', 'detail': f} for f in fixes])

            _, fixes = fix_spacing(doc_xml)
            all_issues.extend([{'type': 'spacing', 'detail': f} for f in fixes])

            _, fixes = fix_page_setup(doc_xml)
            all_issues.extend([{'type': 'page_setup', 'detail': f} for f in fixes])

            rels_xml = z.read('word/_rels/document.xml.rels') if 'word/_rels/document.xml.rels' in names else None
            media = [n for n in names if n.startswith('word/media/')]
            _, fixes = fix_images(doc_xml, rels_xml, media)
            all_issues.extend([{'type': 'image', 'detail': f} for f in fixes])

    return all_issues


def fix_all(file_path, output_path=None, skip_font_normalize=False, unify_fonts=False):
    """Apply all fixes to the converted document."""
    if output_path is None:
        output_path = file_path

    all_fixes = []

    with zipfile.ZipFile(file_path, 'r') as z:
        all_entries = z.namelist()

    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(file_path, 'r') as z:
            z.extractall(tmpdir)

        # Collect all XML files to process
        xml_files = [n for n in all_entries if n.endswith('.xml') and n.startswith('word/')]

        # Fix 1: Settings
        settings_path = os.path.join(tmpdir, 'word', 'settings.xml')
        if os.path.exists(settings_path):
            with open(settings_path, 'rb') as f:
                data = f.read()
            fixed, fixes = fix_settings(data)
            all_fixes.extend(fixes)
            if fixed:
                with open(settings_path, 'wb') as f:
                    f.write(fixed if isinstance(fixed, bytes) else fixed.encode('utf-8'))

        # Fix 2: Font names in ALL XML files
        for xf in xml_files:
            full_path = os.path.join(tmpdir, xf)
            if not os.path.exists(full_path):
                continue
            with open(full_path, 'rb') as f:
                data = f.read()
            fixed, fixes = fix_font_names_in_xml(data)
            all_fixes.extend(fixes)
            if fixed:
                with open(full_path, 'wb') as f:
                    f.write(fixed)

        # Fix 3: Styles
        styles_path = os.path.join(tmpdir, 'word', 'styles.xml')
        if os.path.exists(styles_path):
            with open(styles_path, 'rb') as f:
                data = f.read()
            fixed, fixes = fix_styles(data)
            all_fixes.extend(fixes)
            if fixed:
                with open(styles_path, 'wb') as f:
                    f.write(fixed)

        # Fix 4-7: Document XML fixes (tables, images, spacing, page setup)
        doc_path = os.path.join(tmpdir, 'word', 'document.xml')
        if os.path.exists(doc_path):
            with open(doc_path, 'rb') as f:
                doc_data = f.read()

            collect_namespaces(doc_data)

            # Tables
            fixed, fixes = fix_tables(doc_data)
            all_fixes.extend(fixes)
            if fixed:
                doc_data = fixed
                with open(doc_path, 'wb') as f:
                    f.write(doc_data)

            # Spacing
            fixed, fixes = fix_spacing(doc_data)
            all_fixes.extend(fixes)
            if fixed:
                doc_data = fixed
                with open(doc_path, 'wb') as f:
                    f.write(doc_data)

            # Page setup
            fixed, fixes = fix_page_setup(doc_data)
            all_fixes.extend(fixes)
            if fixed:
                doc_data = fixed
                with open(doc_path, 'wb') as f:
                    f.write(doc_data)

            # Image validation
            rels_path = os.path.join(tmpdir, 'word', '_rels', 'document.xml.rels')
            rels_data = None
            if os.path.exists(rels_path):
                with open(rels_path, 'rb') as f:
                    rels_data = f.read()
            media = [n for n in all_entries if n.startswith('word/media/')]
            _, fixes = fix_images(doc_data, rels_data, media)
            all_fixes.extend(fixes)

        # Repack
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for entry in all_entries:
                full_path = os.path.join(tmpdir, entry)
                if os.path.isfile(full_path):
                    zout.write(full_path, entry)

    # Fix 8: Font normalization (calls normalize_fonts.py)
    # Default is pairing mode: it completes missing eastAsia/ascii pairs and
    # fixes cross-script mistakes without flattening deliberate font choices.
    # --unify-fonts opts into the aggressive force-single-pair mode.
    if not skip_font_normalize:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        normalize_script = os.path.join(script_dir, 'normalize_fonts.py')
        if os.path.exists(normalize_script):
            cmd = [sys.executable, normalize_script, output_path, '--json', '--no-backup']
            if unify_fonts:
                cmd.append('--unify')
            result = subprocess.run(cmd, capture_output=True, text=True)
            # Exit code 3 = normalized, but some issues remain (still JSON output).
            if result.returncode in (0, 3):
                try:
                    font_result = json.loads(result.stdout)
                    msg = f"font normalization: {font_result.get('runs_fixed', 0)} fixes"
                    remaining = font_result.get('remaining_issue_count', 0)
                    if remaining:
                        msg += f' ({remaining} issue(s) remain)'
                    all_fixes.append(msg)
                except json.JSONDecodeError:
                    all_fixes.append('font normalization: completed (non-JSON output)')
            else:
                all_fixes.append(f'font normalization: FAILED — {result.stderr[:200]}')
        else:
            all_fixes.append(f'font normalization: SKIPPED — normalize_fonts.py not found at {normalize_script}')

    return all_fixes


def main():
    parser = argparse.ArgumentParser(
        description='Fix LibreOffice .doc→.docx conversion artifacts'
    )
    parser.add_argument('input', help='Input .docx file (converted from .doc)')
    parser.add_argument('output', nargs='?', help='Output .docx file (default: overwrite input)')
    parser.add_argument('--detect-only', action='store_true', help='Only detect issues, do not fix')
    parser.add_argument('--skip-font-normalize', action='store_true', help='Skip font normalization step')
    parser.add_argument('--unify-fonts', action='store_true',
                        help='Use aggressive --unify font normalization (forces a single '
                             'font pair on all runs; default is the safer pairing mode)')
    parser.add_argument('--json', action='store_true', help='Output as JSON')

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if args.detect_only:
        issues = detect_issues(args.input)
        if args.json:
            print(json.dumps({'issue_count': len(issues), 'issues': issues}, ensure_ascii=False, indent=2))
        else:
            if issues:
                print(f"Found {len(issues)} conversion artifact(s):\n")
                for issue in issues:
                    file_info = f" [{issue['file']}]" if 'file' in issue else ''
                    print(f"  [{issue['type']}]{file_info} {issue['detail']}")
            else:
                print("No conversion artifacts detected.")
        return

    output = args.output or args.input
    fixes = fix_all(args.input, output, skip_font_normalize=args.skip_font_normalize,
                    unify_fonts=args.unify_fonts)

    if args.json:
        print(json.dumps({'fixes_applied': len(fixes), 'fixes': fixes, 'output': output},
                         ensure_ascii=False, indent=2))
    else:
        if fixes:
            print(f"Applied {len(fixes)} fix(es):\n")
            for fix in fixes:
                print(f"  {fix}")
            print(f"\nOutput: {output}")
        else:
            print("No fixes needed.")


if __name__ == '__main__':
    main()
