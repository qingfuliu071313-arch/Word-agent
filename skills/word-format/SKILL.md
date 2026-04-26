---
name: word-format
description: >-
  Academic paper formatting module. Parses user-provided format requirements
  into structured rules (Format Spec), then applies them to Word documents
  in batch. Covers page setup, fonts, spacing, heading styles, page numbers,
  headers/footers, TOC generation, and CJK mixed-text normalization.
  Triggers: 排版, 格式化, 改格式, 按要求排, format, style, apply formatting,
  目录, 生成目录, TOC, 中英文混排, 标点, 全角半角.
allowed-tools: Read Write Bash Glob Grep mcp__word-document-server__format_text mcp__word-document-server__create_custom_style mcp__word-document-server__search_and_replace mcp__word-document-server__set_table_width mcp__word-document-server__set_table_column_widths mcp__word-document-server__format_table mcp__word-document-server__highlight_table_header mcp__word-document-server__add_heading mcp__word-document-server__add_page_break mcp__word-document-server__get_document_info mcp__word-document-server__get_document_outline mcp__word-document-server__get_document_text mcp__word-document-server__get_paragraph_text_from_document mcp__word-document-server__find_text_in_document mcp__word-document-server__get_document_xml
metadata:
    version: "0.1.0"
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

## Phase 1: Parse Format Requirements

### Input Sources

The user may provide format requirements as:
1. A Word/PDF document → read and extract rules
2. Verbal description in conversation → extract from text
3. A URL to journal guidelines → fetch and extract

### Extraction Process

Follow `../../references/format_spec_parser.md` to extract a structured Format Spec:

1. Read the format requirement document
2. Search for keywords: 页边距, 字体, 字号, 行距, 编号, 页码, margins, font, spacing...
3. Extract specific values for each rule
4. For Chinese font sizes, convert using the size chart in `format_spec_parser.md`
5. Mark unspecified rules as `null`

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

### Execution Order

Group by operation type, not by location. This minimizes context switches and tool call overhead.

```
Step 1: Page Setup
  → If page size or margins need changing → XML edit or MCP
  → One operation covers the entire document

Step 2: Create/Update Styles
  → create_custom_style for each heading level that needs updating
  → This automatically applies to all paragraphs using that style

Step 3: Body Formatting
  → format_text for body paragraphs that need font/size/spacing changes
  → Batch by identifying paragraph ranges from Document Map

Step 4: Special Element Formatting
  → Figure/table captions, references, abstract
  → Targeted format_text calls using Document Map locations

Step 5: CJK Normalization (if requested)
  → Execute in order per references/chinese_standards.md:
     a. Punctuation normalization (search_and_replace, ~5-8 calls)
     b. CJK-English spacing (search_and_replace, ~3-4 calls)
     c. Number-unit spacing (search_and_replace, ~2-3 calls)

Step 6: Headers/Footers (if required)
  → XML manipulation for header/footer content and page numbers

Step 7: First-Line Indent (if required)
  → format_text or style update for body paragraphs

Step 8: TOC (if requested)
  → See TOC Generation section below
```

### Error Handling

If a format_text or search_and_replace call fails:
1. Log the failure with the specific paragraph/pattern
2. Continue with remaining operations (don't stop on first error)
3. Report all failures at the end
4. Suggest alternative approaches for failed operations

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

4. **Insert TOC field code** — Via XML manipulation:

```bash
# Unpack document
python scripts/office/unpack.py document.docx unpacked/

# Insert TOC field code at specified position in document.xml
# The TOC field code:
# <w:sdt>
#   <w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/></w:docPartObj></w:sdtPr>
#   <w:sdtContent>
#     <w:p>
#       <w:r>
#         <w:fldChar w:fldCharType="begin"/>
#       </w:r>
#       <w:r>
#         <w:instrText>TOC \o "1-3" \h \z \u</w:instrText>
#       </w:r>
#       <w:r>
#         <w:fldChar w:fldCharType="separate"/>
#       </w:r>
#       <w:r>
#         <w:t>（请在 Word 中更新域以显示目录）</w:t>
#       </w:r>
#       <w:r>
#         <w:fldChar w:fldCharType="end"/>
#       </w:r>
#     </w:p>
#   </w:sdtContent>
# </w:sdt>

# Repack
python scripts/office/pack.py unpacked/ output.docx --original document.docx
```

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

## Post-Execution

1. **Report results:**
   ```
   格式化完成：
   ✅ 已修改 8 项格式设置
   ✅ 已处理 142 段正文格式
   ✅ 已修正 23 处中英文混排
   ❌ 1 项失败: 页眉设置（需要手动在 Word 中调整）
   
   建议下一步：运行格式合规检查确认所有修改正确。
   使用 /word-agent:word-check
   ```

2. **Suggest word-checker** for verification

## Shared Resources

- `../../references/format_spec_parser.md` — Format Spec extraction rules
- `../../references/academic_formatting.md` — Academic formatting knowledge base
- `../../references/chinese_standards.md` — CJK normalization rules
- `../../references/tool_routing.md` — Tool selection priority
- `../../references/token_budget.md` — Token efficiency rules
- `../../references/document_creation_rules.md` — New document anti-pattern rules (字体配对、内置样式、TOC)
