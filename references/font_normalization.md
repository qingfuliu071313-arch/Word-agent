# Font Normalization Reference / 字体归一化参考

## Problem Description / 问题描述

When Word documents are created or edited by CLI tools (e.g., MCP word-document-server) without Word Agent guidance, Chinese fonts become chaotic — a single sentence may contain multiple fonts. This is a known issue caused by how OpenXML handles font attributes at three levels.

当通过 CLI 工具（如 MCP word-document-server）直接创建或编辑 Word 文档时，中文字体会出现混乱——一句话中可能混入多种字体。这是 OpenXML 在三个层面处理字体属性的方式导致的已知问题。

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

### 2. Theme Font References Override Explicit Fonts / 主题字体引用覆盖显式字体

Word's `styles.xml` uses theme references like `asciiTheme="minorHAnsi"` and `eastAsiaTheme="minorEastAsia"` that resolve to theme-defined fonts (e.g., Calibri, Cambria). **Theme attributes (`*Theme`) override explicit font names in the same `w:rFonts` element.** This means even if `w:ascii="Times New Roman"` is set, `w:asciiTheme="minorHAnsi"` will make Word display Calibri instead.

Word 的 `styles.xml` 使用主题引用（如 `asciiTheme="minorHAnsi"`），这些引用解析为主题定义的字体（如 Calibri、Cambria）。**`*Theme` 属性会覆盖同一 `w:rFonts` 元素中的显式字体名称。** 因此即使设置了 `w:ascii="Times New Roman"`，`w:asciiTheme` 也会让 Word 显示为 Calibri。

Three places where theme refs hide:
- `docDefaults` → affects ALL runs without explicit fonts
- Style definitions (Heading 1, Title, Normal, etc.)
- Table style sub-elements (`tblStylePr`)

### 3. Bare Runs Without Any Font Specification / 无字体属性的裸 Run

Runs created without `<w:rPr>` or without `<w:rFonts>` inside `<w:rPr>` inherit their fonts entirely from the paragraph style → docDefaults → theme. If those ancestors use theme refs, the run displays in theme fonts (typically Calibri/Cambria for Latin, and a different CJK font than intended).

没有 `<w:rPr>` 或其中没有 `<w:rFonts>` 的 run，其字体完全继承自段落样式 → docDefaults → 主题。如果这些上级使用了主题引用，run 就会显示为主题字体。

### 4. Multiple `<w:r>` Runs with Inconsistent Formatting

Each write operation creates a new `<w:r>` (run) element. When multiple operations touch the same paragraph, the result is mixed fonts across runs.

每次写操作会创建新的 `<w:r>` 元素。多次操作后，同一段落中出现不同字体设置。

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

## Script: `scripts/normalize_fonts.py`

The canonical implementation lives at `scripts/normalize_fonts.py`. Do NOT use inline Python snippets — always call the script directly.

标准实现位于 `scripts/normalize_fonts.py`。不要使用内联 Python 代码片段——始终直接调用脚本。

### Detection / 检测

```bash
python3 scripts/normalize_fonts.py paper.docx --detect-only
python3 scripts/normalize_fonts.py paper.docx --detect-only --json   # machine-readable
```

Detects ALL font issues across three layers:
- **styles.xml**: theme font references in docDefaults and all style definitions (including nested table style elements)
- **document.xml runs with rFonts**: missing eastAsia/ascii, mixed fonts, hAnsi mismatches, theme attributes
- **document.xml bare runs**: text runs with no `w:rPr` or no `w:rFonts`

### Normalization / 归一化

```bash
# Gentle mode: fill in missing attributes using font pairing table
python3 scripts/normalize_fonts.py paper.docx

# Aggressive mode (recommended for academic papers): force ALL to same pair
python3 scripts/normalize_fonts.py paper.docx --unify

# Custom font pair
python3 scripts/normalize_fonts.py paper.docx --unify --cn 黑体 --en Arial

# Output to new file
python3 scripts/normalize_fonts.py paper.docx paper_fixed.docx --unify
```

The `--unify` flag is recommended for academic papers. It fixes three layers:

1. **styles.xml** — strips all `*Theme` attributes from every `w:rFonts` in docDefaults, style definitions, and nested table style elements; sets explicit font names
2. **document.xml existing rFonts** — strips theme attrs, sets uniform font pair
3. **document.xml bare runs** — injects `<w:rPr><w:rFonts>` with explicit fonts into runs that had no font specification

### Text Box Extraction / 文本框提取

```bash
python3 scripts/normalize_fonts.py paper.docx --textboxes
python3 scripts/normalize_fonts.py paper.docx --textboxes --json
```

### Custom Pairing / 自定义配对

Users can specify their own font pairing:

```
"字体归一化，正文用楷体+Times New Roman"
→ python3 scripts/normalize_fonts.py paper.docx --unify --cn 楷体 --en "Times New Roman"

"Fix fonts, use SimHei + Arial for everything"
→ python3 scripts/normalize_fonts.py paper.docx --unify --cn 黑体 --en Arial
```

## Integration Points / 集成点

| Module | How it uses this reference |
|--------|---------------------------|
| **word-read** | Runs `--detect-only` during Document Map generation → reports in "Format Issues Detected" |
| **word-format** | Runs `--unify` as post-processing after Phase 3, or on explicit trigger |
| **word-orchestrate** | Notes in pre-flight that documents created without Word Agent may need font normalization |
| **word-check** | Can flag font inconsistencies during compliance check |

## Notes / 注意事项

1. **Always back up before normalizing** — The script modifies XML directly; use the output argument to write to a new file
2. **`--unify` modifies styles.xml** — All heading styles, title, and table styles will get explicit fonts. This is intentional for academic papers where consistent fonts are required
3. **Theme fonts are stripped, not preserved** — This is by design. Theme references are the #1 cause of font chaos. The script replaces them with explicit font names
4. **Headers/footers are not covered** — The script operates on `word/document.xml` and `word/styles.xml`. Extend to `word/header*.xml` and `word/footer*.xml` if needed
5. **Run after MCP operations, before final delivery** — This is a safety net, not a replacement for correct font specification during writing
