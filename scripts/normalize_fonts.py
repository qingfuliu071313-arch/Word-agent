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
import shutil
import sys
import argparse
import tempfile
import json
import re
from datetime import datetime

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
    '华文中宋': 'Times New Roman',
    'STZhongsong': 'Times New Roman',
    '华文楷体': 'Times New Roman',
    'STKaiti': 'Times New Roman',
    '华文宋体': 'Times New Roman',
    'STSong': 'Times New Roman',
    '华文仿宋': 'Times New Roman',
    'STFangsong': 'Times New Roman',
    '华文细黑': 'Arial',
    'STXihei': 'Arial',
    '华文黑体': 'Arial',
    '方正小标宋': 'Times New Roman',
    '方正仿宋': 'Times New Roman',
    '方正楷体': 'Times New Roman',
    '方正黑体': 'Arial',
    '方正书宋': 'Times New Roman',
    '等线': 'Calibri',
    'DengXian': 'Calibri',
    '隶书': 'Times New Roman',
    'LiSu': 'Times New Roman',
    '幼圆': 'Arial',
    'YouYuan': 'Arial',
}

FONT_PAIRS_EN_TO_CN = {
    'Times New Roman': '宋体',
    'Arial': '黑体',
    'Calibri': '宋体',
    'Cambria': '宋体',
    'Helvetica': '黑体',
    'Verdana': '黑体',
    'Georgia': '宋体',
    'Palatino Linotype': '宋体',
    'Tahoma': '黑体',
    'Segoe UI': '黑体',
    'Consolas': '宋体',
    'Courier New': '宋体',
}

# Fonts that must never be unified away: symbol fonts would turn their
# characters into garbage, monospace fonts are deliberate (code blocks).
PRESERVE_FONTS = {
    'Symbol', 'Wingdings', 'Wingdings 2', 'Wingdings 3', 'Webdings',
    'MT Extra', 'Cambria Math', 'Marlett', 'ZapfDingbats',
    'Courier New', 'Consolas', 'Menlo', 'Monaco',
}

# Tracked-change containers — runs inside them are revision content.
REV_TAGS = {f'{W}ins', f'{W}del', f'{W}moveFrom', f'{W}moveTo'}


def _inside_revision(el, parent_map):
    """True if the element sits inside a tracked-change container."""
    p = parent_map.get(el)
    while p is not None:
        if p.tag in REV_TAGS:
            return True
        p = parent_map.get(p)
    return False


ALL_NS = {}

OOXML_KNOWN_NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'v': 'urn:schemas-microsoft-com:vml',
    'o': 'urn:schemas-microsoft-com:office:office',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
    'w16cid': 'http://schemas.microsoft.com/office/word/2016/wordml/cid',
    'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
}


def collect_namespaces(xml_bytes):
    """Collect all namespace declarations from the XML to preserve them on write.
    Also injects missing well-known OOXML namespace declarations into the root element
    so that ET.fromstring can parse documents with undeclared prefixes."""
    for prefix, uri in OOXML_KNOWN_NS.items():
        if prefix not in ALL_NS:
            ALL_NS[prefix] = uri
            ET.register_namespace(prefix, uri)
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


def inject_missing_namespaces(xml_bytes):
    """Inject missing well-known OOXML namespace declarations into root element.
    Converted .doc files often omit declarations for prefixes like a:, wp:, etc."""
    declared = set(m.group(1).decode() for m in re.finditer(rb'xmlns:(\w+)=', xml_bytes))
    used = set(m.group(1).decode() for m in re.finditer(rb'<(\w+):', xml_bytes))
    used.update(m.group(1).decode() for m in re.finditer(rb' (\w+):', xml_bytes))
    missing = (used - declared) & set(OOXML_KNOWN_NS.keys())
    if not missing:
        return xml_bytes
    injection = b''
    for prefix in missing:
        uri = OOXML_KNOWN_NS[prefix]
        injection += f' xmlns:{prefix}="{uri}"'.encode()
    return re.sub(rb'(<\w+:?\w+)([ >])', lambda m: m.group(1) + injection + m.group(2),
                  xml_bytes, count=1)


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

    xml_content = inject_missing_namespaces(xml_content)
    collect_namespaces(xml_content)
    if styles_xml:
        styles_xml = inject_missing_namespaces(styles_xml)
        collect_namespaces(styles_xml)
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

        # .// so runs nested in hyperlinks/ins/del are also inspected —
        # keeps detect's scope consistent with what normalize actually fixes.
        for run in para.findall(f'.//{W}r'):
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


def normalize_fonts(file_path, output_path=None, cn_font='宋体', en_font='Times New Roman',
                    unify=False, skip_revisions=False, backup=True):
    """Normalize all font references in the document.

    Fixes four layers:
    1. styles.xml — replace theme refs in docDefaults and styles with explicit fonts
    2. document.xml runs WITH w:rFonts — fix pairings, strip theme attrs
    3. document.xml runs WITHOUT w:rFonts — inject explicit font attributes
    4. theme1.xml — East Asian (<a:ea>) slots pointing at non-CJK fonts

    skip_revisions: leave runs inside w:ins/w:del untouched, so a document
    with tracked changes is not silently altered outside the revision flow.
    Symbol and monospace fonts (PRESERVE_FONTS) are never unified.
    """
    if output_path is None:
        output_path = file_path

    with zipfile.ZipFile(file_path, 'r') as z:
        xml_content = z.read('word/document.xml')
        styles_xml = z.read('word/styles.xml') if 'word/styles.xml' in z.namelist() else None
        all_entries = z.namelist()

    xml_content = inject_missing_namespaces(xml_content)
    collect_namespaces(xml_content)
    if styles_xml:
        styles_xml = inject_missing_namespaces(styles_xml)
        collect_namespaces(styles_xml)

    root = ET.fromstring(xml_content)
    fixed_count = 0

    # Parent map (stdlib ET has no getparent()); used by Layer 2 for the
    # revision check and by Layer 3 for style-aware injection.
    parent_map = {child: parent for parent in root.iter() for child in parent}

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
            need_fix = _has_theme_attrs(rFonts) or unify
            _strip_theme_attrs(rFonts)
            if unify:
                rFonts.set(f'{W}ascii', en_font)
                rFonts.set(f'{W}eastAsia', cn_font)
                rFonts.set(f'{W}hAnsi', en_font)
                styles_changed = True
                fixed_count += 1
            else:
                s_asc = rFonts.get(f'{W}ascii')
                s_ea = rFonts.get(f'{W}eastAsia')
                if not s_asc:
                    rFonts.set(f'{W}ascii', en_font)
                    need_fix = True
                if not s_ea:
                    rFonts.set(f'{W}eastAsia', cn_font)
                    need_fix = True
                # Fix cross-script: CJK font as ascii, or Latin font as eastAsia
                s_asc = rFonts.get(f'{W}ascii')
                s_ea = rFonts.get(f'{W}eastAsia')
                if s_asc and s_asc in FONT_PAIRS_CN_TO_EN:
                    rFonts.set(f'{W}ascii', FONT_PAIRS_CN_TO_EN[s_asc])
                    need_fix = True
                if s_ea and s_ea in FONT_PAIRS_EN_TO_CN:
                    rFonts.set(f'{W}eastAsia', FONT_PAIRS_EN_TO_CN[s_ea])
                    need_fix = True
                if not rFonts.get(f'{W}hAnsi'):
                    rFonts.set(f'{W}hAnsi', rFonts.get(f'{W}ascii', en_font))
                    need_fix = True
                # Sync hAnsi with ascii
                s_asc = rFonts.get(f'{W}ascii')
                s_hansi = rFonts.get(f'{W}hAnsi')
                if s_asc and s_hansi != s_asc:
                    rFonts.set(f'{W}hAnsi', s_asc)
                    need_fix = True
                if need_fix:
                    styles_changed = True
                    fixed_count += 1

    # ── Layer 2: Fix existing w:rFonts in document.xml ──
    for rFonts in root.findall(f'.//{W}rFonts'):
        if skip_revisions and _inside_revision(rFonts, parent_map):
            continue

        east = rFonts.get(f'{W}eastAsia')
        ascii_font = rFonts.get(f'{W}ascii')
        hansi = rFonts.get(f'{W}hAnsi')
        changed = False

        # Always strip theme attrs — they override explicit values
        if _strip_theme_attrs(rFonts):
            changed = True

        if unify:
            # Never unify symbol/monospace fonts — that destroys their content.
            if ascii_font in PRESERVE_FONTS or east in PRESERVE_FONTS:
                if changed:
                    fixed_count += 1
                continue
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
            if not ascii_font and not east:
                # rFonts exists but carries no explicit fonts (e.g. only w:hint)
                rFonts.set(f'{W}ascii', en_font)
                rFonts.set(f'{W}eastAsia', cn_font)
                rFonts.set(f'{W}hAnsi', en_font)
                changed = True
            if ascii_font and not east:
                paired = FONT_PAIRS_EN_TO_CN.get(ascii_font, cn_font)
                rFonts.set(f'{W}eastAsia', paired)
                changed = True
            if east and not ascii_font:
                paired = FONT_PAIRS_CN_TO_EN.get(east, en_font)
                rFonts.set(f'{W}ascii', paired)
                rFonts.set(f'{W}hAnsi', paired)
                changed = True

            # Fix cross-script fonts: CJK font in ascii position or Latin font in eastAsia
            ascii_font = rFonts.get(f'{W}ascii')
            east = rFonts.get(f'{W}eastAsia')
            if ascii_font and ascii_font in FONT_PAIRS_CN_TO_EN:
                correct_en = FONT_PAIRS_CN_TO_EN[ascii_font]
                rFonts.set(f'{W}ascii', correct_en)
                rFonts.set(f'{W}hAnsi', correct_en)
                changed = True
            if east and east in FONT_PAIRS_EN_TO_CN:
                rFonts.set(f'{W}eastAsia', FONT_PAIRS_EN_TO_CN[east])
                changed = True

            ascii_font = rFonts.get(f'{W}ascii')
            hansi = rFonts.get(f'{W}hAnsi')
            if ascii_font and hansi != ascii_font:
                rFonts.set(f'{W}hAnsi', ascii_font)
                changed = True

        if changed:
            fixed_count += 1

    # ── Build style font map for style-aware Layer 3 ──
    style_fonts = {}
    if styles_root is not None:
        for style_el in styles_root.findall(f'{W}style'):
            sid = style_el.get(f'{W}styleId', '')
            rPr_el = style_el.find(f'{W}rPr')
            if rPr_el is not None:
                rf = rPr_el.find(f'{W}rFonts')
                if rf is not None:
                    s_ea = rf.get(f'{W}eastAsia')
                    s_asc = rf.get(f'{W}ascii')
                    if s_ea or s_asc:
                        s_en = s_asc or FONT_PAIRS_CN_TO_EN.get(s_ea, en_font)
                        s_cn = s_ea or FONT_PAIRS_EN_TO_CN.get(s_asc, cn_font)
                        style_fonts[sid] = (s_cn, s_en)

    # ── Layer 3: Inject w:rFonts into bare runs (style-aware) ──
    for run in root.findall(f'.//{W}r'):
        if not _run_has_text(run):
            continue
        if skip_revisions and _inside_revision(run, parent_map):
            continue

        rPr = run.find(f'{W}rPr')
        if rPr is None:
            rPr = ET.SubElement(run, f'{W}rPr')
            run.remove(rPr)
            run.insert(0, rPr)

        rFonts = rPr.find(f'{W}rFonts')
        if rFonts is None:
            # Determine font pair from paragraph style
            run_cn, run_en = cn_font, en_font
            para = parent_map.get(run)
            if para is not None and para.tag == f'{W}p':
                pPr = para.find(f'{W}pPr')
                if pPr is not None:
                    pStyle = pPr.find(f'{W}pStyle')
                    if pStyle is not None:
                        sid = pStyle.get(f'{W}val', '')
                        if sid in style_fonts:
                            run_cn, run_en = style_fonts[sid]

            rFonts = ET.SubElement(rPr, f'{W}rFonts')
            rFonts.set(f'{W}ascii', run_en)
            rFonts.set(f'{W}eastAsia', run_cn)
            rFonts.set(f'{W}hAnsi', run_en)
            fixed_count += 1

    # ── Layer 4: Fix theme.xml East Asian font slots ──
    # Only the <a:ea typeface="..."> slots are touched. A blanket
    # bytes.replace would also rewrite <a:latin> (making English text render
    # in a CJK font) and corrupt compound names ("Calibri Light" → "宋体 Light").
    theme_fixed = False
    theme_xml = None
    if 'word/theme/theme1.xml' in all_entries:
        with zipfile.ZipFile(file_path, 'r') as z:
            theme_xml = z.read('word/theme/theme1.xml')
        theme_xml, ea_fixed = _fix_theme_ea_fonts(theme_xml, cn_font)
        if ea_fixed:
            theme_fixed = True
            fixed_count += ea_fixed
            print(f"  Fixed {ea_fixed} theme.xml <a:ea> font slot(s) → {cn_font}")

    # ── Write back ──
    # Entries are streamed straight from the source zip (no extractall — that
    # silently drops/merges entries on case-insensitive filesystems). The new
    # archive is written to a temp file in the output directory, verified,
    # then atomically swapped in. In-place runs get a timestamped backup first.
    xml_out = serialize_xml(xml_content, root)
    styles_out = serialize_xml(styles_xml, styles_root) if styles_changed and styles_root is not None else None

    replacements = {'word/document.xml': xml_out.encode('utf-8')}
    if styles_out:
        replacements['word/styles.xml'] = styles_out.encode('utf-8')
    if theme_fixed and theme_xml is not None:
        replacements['word/theme/theme1.xml'] = theme_xml

    if backup and os.path.abspath(output_path) == os.path.abspath(file_path):
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        base, ext = os.path.splitext(file_path)
        backup_path = f"{base}_backup_{ts}{ext}"
        shutil.copy2(file_path, backup_path)
        print(f"  Backup: {backup_path}")

    out_dir = os.path.dirname(os.path.abspath(output_path)) or '.'
    fd, tmp = tempfile.mkstemp(suffix='.docx', dir=out_dir)
    os.close(fd)
    try:
        with zipfile.ZipFile(file_path, 'r') as zin, \
             zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for entry in zin.namelist():
                data = replacements.get(entry)
                zout.writestr(entry, data if data is not None else zin.read(entry))
        with zipfile.ZipFile(tmp, 'r') as zf:
            if 'word/document.xml' not in zf.namelist():
                raise ValueError('output verification failed: word/document.xml missing')
        os.replace(tmp, output_path)
    except Exception:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise

    return fixed_count


def _fix_theme_ea_fonts(theme_xml, cn_font):
    """Replace non-CJK fonts in <a:ea typeface=...> slots with cn_font.
    Returns (new_bytes, replacement_count)."""
    pat = re.compile(rb'(<a:ea typeface=")([^"]*)(")')
    count = 0

    def _is_cjk_font(name):
        return (name in FONT_PAIRS_CN_TO_EN
                or any('一' <= ch <= '鿿' for ch in name))

    def repl(m):
        nonlocal count
        val = m.group(2).decode('utf-8')
        if val and not _is_cjk_font(val):
            count += 1
            return m.group(1) + cn_font.encode('utf-8') + m.group(3)
        return m.group(0)

    return pat.sub(repl, theme_xml), count


def extract_textboxes(file_path):
    """Extract text content from all text boxes (w:txbxContent)."""
    with zipfile.ZipFile(file_path, 'r') as z:
        xml_content = z.read('word/document.xml')

    xml_content = inject_missing_namespaces(xml_content)
    collect_namespaces(xml_content)
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
    parser.add_argument('--unify', action='store_true', help='Force ALL runs to the same font pair (aggressive mode for academic papers; symbol/monospace fonts are preserved)')
    parser.add_argument('--skip-revisions', action='store_true',
                        help='Leave runs inside tracked changes (w:ins/w:del) untouched. '
                             'Use on documents with revision marks meant for reviewers.')
    parser.add_argument('--no-backup', action='store_true',
                        help='Skip the timestamped backup created before in-place writes')
    parser.add_argument('--textboxes', action='store_true', help='Extract and display text box content')
    parser.add_argument('--json', action='store_true', help='Output results as JSON')
    parser.add_argument('--check-lock', action='store_true',
                        help='Check if file is locked (e.g., open in Word) before processing. '
                             'If locked, print warning and exit with code 2 instead of failing.')

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    if args.check_lock:
        # Word's owner file is '~$' + the name with its first 1-2 chars removed;
        # check all variants. (An open() probe alone misses this on macOS,
        # where Word does not take a mandatory lock.)
        d = os.path.dirname(os.path.abspath(args.input))
        name = os.path.basename(args.input)
        locked = any(os.path.exists(os.path.join(d, f"~${name[i:]}")) for i in (0, 1, 2))
        if not locked:
            try:
                with open(args.input, 'r+b'):
                    pass
            except (PermissionError, OSError):
                locked = True
        if locked:
            msg = f"Warning: file is locked (likely open in Word): {args.input}"
            if args.json:
                print(json.dumps({"status": "locked", "message": msg}))
            else:
                print(msg, file=sys.stderr)
            sys.exit(2)

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
    fixed = normalize_fonts(args.input, output, cn_font=args.cn, en_font=args.en,
                            unify=args.unify, skip_revisions=args.skip_revisions,
                            backup=not args.no_backup)

    # Verify
    remaining = detect_font_issues(output)

    if args.json:
        result = {
            'issues_found': len(issues),
            'runs_fixed': fixed,
            'output': output,
            'remaining_issue_count': len(remaining),
            'remaining_issues': remaining,
        }
        print(json.dumps(result, ensure_ascii=False, indent=2))
        if remaining:
            sys.exit(3)
    else:
        print(f"Fixed {fixed} font attribute(s) with pairing: {args.cn} + {args.en}")
        print(f"Output: {output}")
        if remaining:
            print(f"\nWarning: {len(remaining)} issue(s) remain after normalization:")
            for issue in remaining:
                loc = f"P{issue['paragraph']}" if issue['paragraph'] >= 0 else "STYLE"
                print(f"  {loc}: [{issue['type']}] {issue['detail']}")
        else:
            print("Verification: all font issues resolved.")


if __name__ == '__main__':
    main()
