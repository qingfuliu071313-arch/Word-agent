---
name: word-format
description: >-
  Academic paper formatting module. Parses user-provided format requirements
  into structured rules (Format Spec), then applies them to Word documents
  in batch. Covers page setup, fonts, spacing, heading styles, page numbers,
  headers/footers, TOC generation, and CJK mixed-text normalization.
  Triggers: 排版, 格式化, 改格式, 按要求排, format, style, apply formatting,
  目录, 生成目录, TOC, 中英文混排, 标点, 全角半角.
allowed-tools: Read Write Bash Glob Grep mcp__word-document-server__create_custom_style mcp__word-document-server__search_and_replace mcp__word-document-server__set_table_width mcp__word-document-server__set_table_column_widths mcp__word-document-server__format_table mcp__word-document-server__highlight_table_header mcp__word-document-server__add_heading mcp__word-document-server__add_page_break mcp__word-document-server__get_document_info mcp__word-document-server__get_document_outline mcp__word-document-server__get_document_text mcp__word-document-server__get_paragraph_text_from_document mcp__word-document-server__find_text_in_document mcp__word-document-server__get_document_xml
metadata:
    version: "1.1.0"
    category: formatting
    upstream-skills: [word-read]
    downstream-skills: [word-check]
---

# Word Formatter

## Overview

Word Formatter takes user-provided format requirements and applies them to a Word document. It operates in three phases: parse requirements → plan changes → execute in batch. It never guesses formatting rules — everything comes from the user's specification.

## When to Use

- Applying formatting rules from a requirement document to a paper
- Generating or inserting a Table of Contents
- Normalizing Chinese-English mixed text (spacing, punctuation)
- Adjusting page setup, fonts, line spacing, heading styles

## When NOT to Use

- Editing text content → use word-edit
- Checking format compliance → use word-check (run AFTER formatting)
- Managing references → use word-reference

## Prerequisites

- **Document Map** from word-reader (or will be requested via orchestrator)
- **Format requirements** from user (Word doc, PDF, or verbal description)

---

## Phase 0: Pre-flight Safety Checks (MANDATORY, before everything else)

1. **文件锁检测** — 确认文档没有被 Word 打开锁定：
   ```bash
   python3 scripts/normalize_fonts.py "{file_path}" --check-lock
   ```
   - exit code 0 → 未锁定，继续
   - exit code 2 → 文件被 Word 锁定 → 停止 file 模式写操作，提示用户：
     "文档正在 Word 中打开。请关闭 Word 后重试，或改用 live 编辑模式（word-mcp-live）。"

2. **已有修订检测** — 检查文档是否已含 tracked changes（来源：Document Map 的
   "Tracked Changes" 节 / `.word-agent/{name}.map.md` sidecar / `mcp__docx-mcp__get_tracked_changes`）：
   ```
   IF 文档已含 N 处修订 (N > 0):
     → 警告用户并等待确认：
       "文档已含 {N} 处修订，直接批量改格式会混入修订流（且可能改动修订标记内部的格式快照），
        是否改用修订模式处理，或先接受/拒绝现有修订后再排版？
        (A) 先处理修订再排版 (B) 仍然直接排版 (C) 取消"
     → 仅在用户明确选择 B 后才继续
   ```

---

## Strategy Selection (Before Phase 1)

Before parsing format requirements, determine which formatting strategy to use:

### Strategy A: Template Injection (Preferred when template .docx provided)

When the user provides a template .docx file with correct style definitions:

1. **Copy source document** as working file
2. **Transplant styles** — Copy the template's styles.xml into the working file:
   ```bash
   python3 scripts/format_document.py "{file_path}" transplant-styles \
     --template "{template_path}" --page-setup
   ```
   This replaces ALL style definitions (Normal, Heading 1-3, etc.) with the template's versions, and optionally copies page setup (margins, size).

3. **Clear direct formatting** — Remove run-level and paragraph-level overrides so content inherits from the transplanted styles:
   ```bash
   python3 scripts/format_document.py "{file_path}" clear-direct-format
   ```
   Use `--range "start,end"` to target specific sections (e.g., skip cover page).

4. **Verify** — Check that styles propagated correctly. If specific paragraphs need overrides (e.g., references with no indent), apply paragraph-level adjustments.

5. **Font normalization** — Run Phase 4 as usual.

**Why this is preferred:** The template's style definitions are authoritative — they were designed in Word with correct font pairing, spacing, indent, and visual tuning. Transplanting them avoids the error-prone process of manually extracting and re-encoding values from text box annotations.

**Limitations:** If source document uses custom style names not present in the template, those paragraphs lose formatting. Verify style name mapping before transplanting.

### Strategy B: Manual Spec (When no template .docx available)

When format requirements come from verbal description, PDF, or non-template Word doc:

1. Proceed to Phase 1 to parse requirements
2. Build Format Spec manually
3. Apply via Phase 3

### Strategy C: Hybrid

When a template provides SOME rules but not all:

1. Transplant styles from template (Strategy A steps 1-3)
2. Parse remaining rules from text/PDF (Phase 1)
3. Apply additional rules via `styles` or `paragraph` commands (Phase 3)

---

## Phase 1: Parse Format Requirements

### Input Sources

The user may provide format requirements as:
1. A Word/PDF document → read and extract rules
2. Verbal description in conversation → extract from text
3. A URL to journal guidelines → fetch and extract

### Extraction Process

Follow `../../references/format_spec_parser.md` to extract a structured Format Spec:

1. Read the format requirement document
2. **Extract text box content (critical)** — Format templates frequently place formatting instructions inside text boxes as annotations. `get_document_text` and `docx2python` completely skip text boxes. You MUST parse the document XML to extract `w:txbxContent` elements. See `format_spec_parser.md` Step 2 for the extraction script. Merge text box content with body text before searching for rules.
3. Search for keywords (in both body text AND text box content): 页边距, 字体, 字号, 行距, 编号, 页码, margins, font, spacing...
4. Extract specific values for each rule
5. For Chinese font sizes, convert using the size chart in `format_spec_parser.md`
6. Mark unspecified rules as `null`

### User Confirmation

Present the extracted Format Spec to the user:

```
我从您的格式要求中提取了以下规则：

📄 页面：A4，页边距上下2.54cm 左右3.17cm
📝 正文：宋体/Times New Roman，小四(12pt)，1.5倍行距
📌 一级标题：黑体/Arial，四号(14pt)，加粗
📌 二级标题：黑体/Arial，小四(12pt)，加粗
📌 三级标题：楷体/Arial，小四(12pt)
🖼 图表题注：宋体/Times New Roman，五号(10.5pt)
📚 参考文献：宋体/Times New Roman，五号(10.5pt)，单倍行距

以下未在要求中提及，将保持文档原样：
- 页眉内容
- 公式编号格式

请确认是否正确，或需要补充修改？
```

**Wait for user confirmation before proceeding to Phase 2.**

---

## Phase 2: Generate Change Plan

Compare Format Spec against current formatting (from Document Map).

### Diff Analysis

For each rule in Format Spec:
1. Check current value (from Document Map)
2. If current ≠ required → add to change list
3. If current = required → skip (no change needed)
4. If rule is `null` → skip (preserve current)

### Change Plan Output

```
格式修改计划：

需要修改 (8 项)：
1. 页面大小: Letter → A4
2. 左右边距: 2.54cm → 3.17cm
3. 正文字体: Arial → 宋体/Times New Roman (全文 142 段)
4. 正文字号: 14pt → 12pt (小四)
5. 一级标题样式: 需创建/更新 Heading 1 样式
6. 行距: 双倍 → 1.5倍 (P16-P165)
7. 参考文献行距: 1.5倍 → 单倍 (P166-P187)
8. 首行缩进: 无 → 2字符

无需修改 (4 项)：
- 页码位置: 已为底部居中 ✓
- 二级标题: 已符合 ✓
- 图题格式: 已符合 ✓
- 表题格式: 已符合 ✓

预计操作: 约 12 次 MCP 调用
确认执行？
```

**Wait for user confirmation before proceeding to Phase 3.**

---

## Phase 3: Execute Changes

### IMPORTANT: Style-First, No Direct Formatting

All formatting MUST go through style modifications, never through `format_text` direct formatting. See `../../references/document_creation_rules.md` Rule 1. The `format_document.py` script modifies style definitions in `styles.xml`, so changes propagate to all paragraphs using that style automatically.

### Preferred Path: One-Shot via apply-spec

If Phase 1 produced a complete Format Spec, convert it to JSON and apply everything in one call:

```bash
python3 scripts/format_document.py "{file_path}" apply-spec --spec format_spec.json
```

This single call handles: page setup, all style modifications (fonts, sizes, spacing, indent, alignment), and headers/footers. Follow with CJK normalization and TOC if needed.

### Alternative Path: Step-by-Step

If applying incrementally, execute in this order:

```
Step 0: Backup (automatic)
  → format_document.py creates a timestamped backup before any modification
  → Backup path printed to output for recovery if needed

Step 1: Page Setup
  → python3 scripts/format_document.py "{file_path}" page-setup \
      --size A4 --margins "2.54cm,2.54cm,3.17cm,3.17cm" --orientation portrait

Step 2: Style Definitions (handles fonts, sizes, spacing, indent, alignment)
  → python3 scripts/format_document.py "{file_path}" styles \
      --name Normal --cn-font 宋体 --en-font "Times New Roman" \
      --font-size 小四 --line-spacing 1.5 --first-indent 2char --alignment justify
  → python3 scripts/format_document.py "{file_path}" styles \
      --name "Heading 1" --cn-font 黑体 --en-font Arial \
      --font-size 三号 --bold --alignment center --space-before 24pt --space-after 12pt
  → Repeat for Heading 2, Heading 3, Caption, etc.
  → This replaces BOTH font setting AND paragraph formatting in one call per style

Step 2b: Clear Direct Formatting (if needed)
  → python3 scripts/format_document.py "{file_path}" clear-direct-format
  → Use --range to target specific sections: --range "44,250" (skip cover page)
  → IMPORTANT: Direct formatting on paragraphs OVERRIDES style definitions.
    If styles were modified but paragraphs still look wrong, direct formatting is the cause.
    Always clear direct formatting after transplanting or modifying styles.

Step 3: Paragraph Overrides (only for paragraphs needing different formatting from their style)
  → python3 scripts/format_document.py "{file_path}" paragraph \
      --style Normal --range "166,187" --line-spacing 1.0
  → Example: reference section paragraphs use Normal style but need single-spacing

Step 4: CJK Normalization (if requested)
  → Execute in order per references/chinese_standards.md:
     a. Punctuation normalization (search_and_replace, ~5-8 calls)
     b. CJK-English spacing (search_and_replace, ~3-4 calls)
     c. Number-unit spacing (search_and_replace, ~2-3 calls)

Step 5: Headers/Footers (if required)
  → python3 scripts/format_document.py "{file_path}" header-footer \
      --header "论文简称" --page-number --header-font "宋体,Times New Roman" --header-size 五号

Step 6: TOC (if requested)
  → python3 scripts/format_document.py "{file_path}" toc \
      --position "引言" --levels 3 --title 目录
  → See TOC Generation section below for details
```

### Error Handling

If any `format_document.py` call fails:
1. The script preserves the backup — original document is recoverable
2. **Stop execution** — do not proceed to next step
3. Report the error with specific details to the user
4. Suggest fix or alternative approach

If a `search_and_replace` MCP call fails:
1. Report the failed pattern
2. Continue with remaining patterns (CJK normalization patterns are independent)
3. Report all failures at the end

---

## TOC Generation

### When Triggered

User requests: "生成目录", "插入目录", "TOC", "table of contents"

### Process

1. **Verify headings** — From Document Map, confirm that headings use proper heading styles (Heading 1, Heading 2, etc.), not just bold/large text
   - If headings are manual formatting → warn user and offer to convert them to proper styles first

2. **Determine position** — Ask user where to insert TOC:
   - After title page (most common)
   - After abstract
   - Custom position

3. **Determine depth** — Ask or infer:
   - Default: show to heading level 3
   - User can specify: "只显示到二级标题"

4. **Insert TOC field code** — Via `format_document.py`:

```bash
python3 scripts/format_document.py "{file_path}" toc \
  --position "引言" --levels 3 --title 目录
```

The `--position` can be a paragraph index (0-based) or a text string to match. The script inserts a proper SDT+field code structure that Word recognizes as an auto-updating TOC.

5. **Inform user:**
   > 目录已插入。请在 Word 中打开文档后：
   > 1. 右键点击目录区域
   > 2. 选择"更新域"
   > 3. 选择"更新整个目录"
   > 即可显示正确的页码。

---

## CJK Normalization Sub-Module

### When Triggered

User requests: "中英文混排", "规范标点", "修复空格", "CJK normalization"

### Process

1. **Scan** — Use word-checker's cross-reference mode or direct text scanning to identify CJK issues
2. **Preview** — Show examples of what will be changed:
   ```
   发现 23 处中英文混排问题：
   - 缺少空格 (15处): "研究area" → "研究 area"
   - 标点错误 (5处): "包括soil,water" → "包括soil、water"
   - 数字间距 (3处): "30cm" → "30 cm"
   
   确认修正？
   ```
3. **Execute** — Run search_and_replace in the order specified in `references/chinese_standards.md`
4. **Verify** — Re-scan to confirm all issues resolved

---

## Phase 4: Font Normalization (MANDATORY)

### This phase ALWAYS runs after Phase 3. It is NOT optional.

Every MCP write operation risks introducing font inconsistencies. This phase is the final gate that guarantees the output document has consistent fonts. Skip it and the user gets font chaos.

### Why This Is Non-Negotiable

MCP tools (e.g., `format_text`) only set `w:ascii`, leaving `w:eastAsia` to fall back to theme fonts. Styles.xml uses theme references (`asciiTheme`, `eastAsiaTheme`) that override explicit font names. Bare runs inherit wrong fonts from themes. See `../../references/font_normalization.md` for the full technical explanation.

### Process

1. **Detect** — Run the font normalization script in detect-only mode:

```bash
python3 scripts/normalize_fonts.py "{file_path}" --detect-only --json
```

This detects ALL font issues including:
- Runs with missing `w:rFonts` (inherit theme/style fonts)
- Theme font references (`*Theme` attributes) in styles.xml and document.xml
- Missing `eastAsia` or `ascii` attributes
- Mixed fonts within a paragraph
- `hAnsi`/`ascii` mismatches

2. **Preview** — Show user what was found:

```
字体一致性检查：
发现 {n} 处字体问题：
- STYLE: styles.xml 中有 {x} 处使用了主题字体引用
- P12: eastAsia 字体缺失（仅设置了 ascii=Times New Roman）
- P25: 同一段落中混用多种 ascii 字体 {Arial, Times New Roman}
- P41: 文字 run 缺少字体属性（继承了主题字体）
...

将使用配对: 宋体 ↔ Times New Roman
确认修复？
```

If running automatically after Phase 3 and no issues found, skip silently.
If running automatically after Phase 3 and issues found, report and auto-fix (user already confirmed the formatting plan in Phase 2).

3. **Normalize** — Run the script with `--unify` for academic papers:

```bash
# For academic papers (recommended): force ALL runs to same font pair
python3 scripts/normalize_fonts.py "{file_path}" --unify --cn "{cn_font}" --en "{en_font}"

# For gentle fix: only fill in missing attributes using font pairing table
python3 scripts/normalize_fonts.py "{file_path}" --cn "{cn_font}" --en "{en_font}"
```

The script fixes three layers:
- **styles.xml**: replaces all theme font references with explicit fonts in docDefaults and style definitions
- **document.xml existing rFonts**: strips theme attrs, fixes pairings
- **document.xml bare runs**: injects `w:rPr`/`w:rFonts` into runs that had no font specification

4. **Report** — Show what was fixed:

```
字体归一化完成：
✅ 修复了 {n} 处字体属性
使用配对: {cn_font} ↔ {en_font}
验证: 所有字体问题已解决
```

### Custom Font Pairing

If the user specifies a font pairing in the Format Spec or via conversation:
- "正文用楷体" → `cn_font='楷体'`, `en_font='Times New Roman'`
- "标题用黑体+Arial" → use 黑体+Arial for heading runs
- Default: 宋体 + Times New Roman

---

## Post-Execution

1. **Report results:**
   ```
   格式化完成：
   ✅ 已修改 8 项格式设置
   ✅ 已处理 142 段正文格式
   ✅ 已修正 23 处中英文混排
   ✅ 字体归一化: 修复 15 处字体属性不一致
   ❌ 1 项失败: 页眉设置（需要手动在 Word 中调整）
   
   建议下一步：运行格式合规检查确认所有修改正确。
   使用 /word-agent:word-check
   ```

2. **Suggest word-checker** for verification

## Shared Resources

- `../../scripts/format_document.py` — Formatting engine: page setup, styles, headers, TOC, apply-spec
- `../../scripts/normalize_fonts.py` — Font normalization: detect and fix font inconsistencies
- `../../references/format_spec_parser.md` — Format Spec extraction rules
- `../../references/academic_formatting.md` — Academic formatting knowledge base
- `../../references/chinese_standards.md` — CJK normalization rules
- `../../references/tool_routing.md` — Tool selection priority
- `../../references/token_budget.md` — Token efficiency rules
- `../../references/font_normalization.md` — Font normalization detection + fix scripts, CJK font pairing map
- `../../references/document_creation_rules.md` — Style-first rules, anti-patterns (字体配对、内置样式、TOC)
