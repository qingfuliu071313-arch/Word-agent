# Font Normalization Reference / 字体归一化参考

## Problem Description / 问题描述

When Word documents are created or edited by CLI tools (e.g., MCP word-document-server) without Word Agent guidance, Chinese fonts become chaotic — a single sentence may contain multiple fonts. This is a known issue caused by how OpenXML handles font attributes.

当通过 CLI 工具（如 MCP word-document-server）直接创建或编辑 Word 文档时，中文字体会出现混乱——一句话中可能混入多种字体。这是 OpenXML 字体属性处理方式导致的已知问题。

## Root Causes / 根本原因

### 1. `w:rFonts` 属性不完整

OpenXML uses separate font attributes for different character categories:

| Attribute | Covers | Example |
|-----------|--------|---------|
| `w:ascii` | Basic Latin (A-Z, 0-9, punctuation) | Times New Roman |
| `w:hAnsi` | Extended Latin (accented chars) | Times New Roman |
| `w:eastAsia` | CJK characters (Chinese, Japanese, Korean) | 宋体 |
| `w:cs` | Complex scripts (Arabic, Hebrew) | — |

MCP tools typically only set `w:ascii`, leaving `w:eastAsia` to fall back to Word's default theme font. This means Chinese characters render in a different font than intended.

MCP 工具通常只设置 `w:ascii` 属性，`w:eastAsia` 则回退到 Word 的默认主题字体。这导致中文字符使用了与预期不同的字体。

### 2. Multiple `<w:r>` Runs with Inconsistent Formatting

Each write operation creates a new `<w:r>` (run) element. When multiple operations touch the same paragraph, the result is:

```xml
<w:p>
  <w:r>
    <w:rPr><w:rFonts w:ascii="Arial"/></w:rPr>
    <w:t>研究</w:t>
  </w:r>
  <w:r>
    <w:rPr><w:rFonts w:ascii="Times New Roman"/></w:rPr>
    <w:t>area</w:t>
  </w:r>
  <w:r>
    <w:rPr><w:rFonts w:ascii="SimSun"/></w:rPr>
    <w:t>的特征</w:t>
  </w:r>
</w:p>
```

每次写操作会创建新的 `<w:r>` 元素。多次操作后，同一段落中出现不同字体设置。

### 3. Direct Formatting Conflicts with Paragraph Styles

When `format_text` applies a font at the run level, it may conflict with the paragraph's style definition. The run-level `<w:rPr>` overrides style, but often only partially (setting `w:ascii` but not `w:eastAsia`).

当 `format_text` 在 run 级别应用字体时，会与段落样式定义冲突。run 级别的 `<w:rPr>` 会覆盖样式，但往往只覆盖部分属性。

## Common CJK Font Pairings / 常用中英文字体配对

| Chinese Font | English Font | Usage / 用途 |
|:---:|:---:|:---:|
| 宋体 (SimSun) | Times New Roman | 正文 — Body text (most common for Chinese academic papers) |
| 黑体 (SimHei) | Arial | 标题 — Headings |
| 楷体 (KaiTi) | Times New Roman | 引用 / 特殊强调 — Quotes / special emphasis |
| 仿宋 (FangSong) | Times New Roman | 公文 / 附录 — Official documents / appendices |
| 微软雅黑 (Microsoft YaHei) | Arial | 现代排版 — Modern layouts (rarely used in academic papers) |

**Default pairing rule:** Unless the user specifies otherwise, use 宋体 + Times New Roman for body, 黑体 + Arial for headings.

**默认配对规则：** 除非用户另有指定，正文使用宋体 + Times New Roman，标题使用黑体 + Arial。

## Font Inconsistency Detection / 字体不一致检测

Use this script to scan a document and report font inconsistencies — multiple fonts found within the same paragraph's runs:

```python
import zipfile
import xml.etree.ElementTree as ET
import json

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def detect_font_issues(file_path):
    """Scan document for font inconsistencies across runs within paragraphs."""
    with zipfile.ZipFile(file_path, 'r') as z:
        xml_content = z.read('word/document.xml')

    root = ET.fromstring(xml_content)
    issues = []

    for i, para in enumerate(root.findall('.//w:p', ns)):
        fonts_in_para = {'ascii': set(), 'eastAsia': set(), 'hAnsi': set()}
        missing_eastAsia = False

        for run in para.findall('.//w:r', ns):
            rFonts = run.find('.//w:rFonts', ns)
            if rFonts is not None:
                ascii_font = rFonts.get(f'{{{ns["w"]}}}ascii')
                east_font = rFonts.get(f'{{{ns["w"]}}}eastAsia')
                hansi_font = rFonts.get(f'{{{ns["w"]}}}hAnsi')

                if ascii_font:
                    fonts_in_para['ascii'].add(ascii_font)
                if east_font:
                    fonts_in_para['eastAsia'].add(east_font)
                if hansi_font:
                    fonts_in_para['hAnsi'].add(hansi_font)

                if ascii_font and not east_font:
                    missing_eastAsia = True

        # Get paragraph text preview
        texts = para.findall('.//w:t', ns)
        preview = ''.join(t.text or '' for t in texts)[:60]

        if not preview.strip():
            continue

        issue = None
        if len(fonts_in_para['ascii']) > 1:
            issue = f"P{i+1}: Multiple ascii fonts {fonts_in_para['ascii']}"
        elif len(fonts_in_para['eastAsia']) > 1:
            issue = f"P{i+1}: Multiple eastAsia fonts {fonts_in_para['eastAsia']}"
        elif missing_eastAsia:
            issue = f"P{i+1}: eastAsia font missing (ascii set but eastAsia not)"

        if issue:
            issues.append({"paragraph": i + 1, "issue": issue, "preview": preview})

    return issues
```

## Font Normalization Script / 字体归一化脚本

This script scans all runs in a `.docx` file and ensures `w:rFonts` attributes (`eastAsia`, `ascii`, `hAnsi`) are consistently set based on a configurable pairing map.

此脚本扫描 `.docx` 文件中的所有 run，确保 `w:rFonts` 属性（`eastAsia`、`ascii`、`hAnsi`）根据可配置的配对映射保持一致。

```python
import zipfile
import xml.etree.ElementTree as ET
import shutil
import os
import re

# Chinese → English font pairing map
FONT_PAIRS = {
    '宋体': 'Times New Roman',
    '黑体': 'Arial',
    '楷体': 'Times New Roman',
    '仿宋': 'Times New Roman',
    'SimSun': 'Times New Roman',
    'SimHei': 'Arial',
    'KaiTi': 'Times New Roman',
    'FangSong': 'Times New Roman',
    '微软雅黑': 'Arial',
    'Microsoft YaHei': 'Arial',
}
# English → Chinese reverse map
REVERSE_PAIRS = {v: k for k, v in FONT_PAIRS.items()}
# For REVERSE_PAIRS, prefer common fonts when multiple map to the same English font
REVERSE_PAIRS['Times New Roman'] = '宋体'
REVERSE_PAIRS['Arial'] = '黑体'

WML_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ns = {'w': WML_NS}

# All OOXML namespaces that may appear in document.xml — needed for faithful re-serialization
ALL_NS = {
    'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'cx': 'http://schemas.microsoft.com/office/drawing/2014/chartex',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o': 'urn:schemas-microsoft-com:office:office',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v': 'urn:schemas-microsoft-com:vml',
    'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'w': WML_NS,
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'w16se': 'http://schemas.microsoft.com/office/word/2015/wordml/symex',
    'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}


def normalize_fonts(file_path, output_path=None, cn_font='宋体', en_font='Times New Roman'):
    """
    Scan all runs in a .docx and ensure eastAsia + ascii + hAnsi are consistently paired.

    Parameters:
        file_path:   Path to the input .docx file
        output_path: Path for the output file (defaults to overwrite input)
        cn_font:     Default Chinese font when eastAsia is missing
        en_font:     Default English font when ascii/hAnsi is missing

    Returns:
        dict with keys: fixed_count (int), details (list of str)
    """
    if output_path is None:
        output_path = file_path

    temp_dir = file_path + '_unpack'
    try:
        # Unpack
        with zipfile.ZipFile(file_path, 'r') as z:
            z.extractall(temp_dir)

        doc_xml_path = os.path.join(temp_dir, 'word', 'document.xml')

        # Register all OOXML namespaces to preserve them during re-serialization
        for prefix, uri in ALL_NS.items():
            ET.register_namespace(prefix, uri)

        tree = ET.parse(doc_xml_path)
        root = tree.getroot()

        fixed_count = 0
        details = []

        for rFonts in root.findall('.//w:rFonts', ns):
            east = rFonts.get(f'{{{WML_NS}}}eastAsia')
            ascii_font = rFonts.get(f'{{{WML_NS}}}ascii')
            hansi = rFonts.get(f'{{{WML_NS}}}hAnsi')

            changes = []

            # Case 1: ascii is set but eastAsia is missing → fill eastAsia
            if ascii_font and not east:
                paired_cn = REVERSE_PAIRS.get(ascii_font, cn_font)
                rFonts.set(f'{{{WML_NS}}}eastAsia', paired_cn)
                changes.append(f'eastAsia←{paired_cn}')

            # Case 2: eastAsia is set but ascii is missing → fill ascii + hAnsi
            if east and not ascii_font:
                paired_en = FONT_PAIRS.get(east, en_font)
                rFonts.set(f'{{{WML_NS}}}ascii', paired_en)
                rFonts.set(f'{{{WML_NS}}}hAnsi', paired_en)
                changes.append(f'ascii/hAnsi←{paired_en}')

            # Case 3: ascii and hAnsi mismatch → align hAnsi to ascii
            ascii_font = rFonts.get(f'{{{WML_NS}}}ascii')  # re-read in case we just set it
            hansi = rFonts.get(f'{{{WML_NS}}}hAnsi')
            if ascii_font and hansi and hansi != ascii_font:
                rFonts.set(f'{{{WML_NS}}}hAnsi', ascii_font)
                changes.append(f'hAnsi: {hansi}→{ascii_font}')

            if changes:
                fixed_count += 1
                details.append('; '.join(changes))

        # Write back
        tree.write(doc_xml_path, xml_declaration=True, encoding='UTF-8')

        # Repack — preserve original entry order and compression
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for dirpath, dirnames, filenames in os.walk(temp_dir):
                for f in filenames:
                    full = os.path.join(dirpath, f)
                    arcname = os.path.relpath(full, temp_dir)
                    zout.write(full, arcname)

        return {'fixed_count': fixed_count, 'details': details}
    finally:
        # Always clean up temp directory
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
```

## Usage / 使用方法

### Standalone normalization / 独立归一化

```bash
python3 << 'PYEOF'
import zipfile, xml.etree.ElementTree as ET, shutil, os

# --- paste normalize_fonts() function here ---

result = normalize_fonts("paper.docx", cn_font="宋体", en_font="Times New Roman")
print(f"Fixed {result['fixed_count']} font attribute issues.")
for d in result['details'][:10]:
    print(f"  - {d}")
PYEOF
```

### Detection only / 仅检测

```bash
python3 << 'PYEOF'
import zipfile, xml.etree.ElementTree as ET, json

# --- paste detect_font_issues() function here ---

issues = detect_font_issues("paper.docx")
print(f"Found {len(issues)} paragraphs with font inconsistencies:")
for issue in issues[:20]:
    print(f"  P{issue['paragraph']}: {issue['issue']}")
    print(f"    Preview: {issue['preview']}")
PYEOF
```

### Custom pairing / 自定义配对

Users can specify their own font pairing when triggering normalization:

```
"字体归一化，正文用楷体+Times New Roman"
→ cn_font='楷体', en_font='Times New Roman'

"Fix fonts, use SimHei + Arial for everything"
→ cn_font='黑体', en_font='Arial'
```

## Integration Points / 集成点

| Module | How it uses this reference |
|--------|---------------------------|
| **word-read** | Runs `detect_font_issues()` during Document Map generation → reports in "Format Issues Detected" |
| **word-format** | Runs `normalize_fonts()` as post-processing after Phase 3, or on explicit trigger |
| **word-orchestrate** | Notes in pre-flight that documents created without Word Agent may need font normalization |
| **word-check** | Can flag font inconsistencies during compliance check |

## Notes / 注意事项

1. **Always back up before normalizing** — The script modifies XML directly; keep a copy of the original
2. **Style-level fonts are not touched** — This script only fixes run-level (`<w:rPr>`) font attributes. Style definitions in `styles.xml` are unchanged
3. **Theme fonts (`w:asciiTheme`, `w:eastAsiaTheme`) are preserved** — The script only acts on explicit font name attributes, not theme references
4. **Headers/footers are not covered** — The script operates on `word/document.xml` only. Extend to `word/header*.xml` and `word/footer*.xml` if needed
5. **Run after MCP operations, before final delivery** — This is a safety net, not a replacement for correct font specification during writing
