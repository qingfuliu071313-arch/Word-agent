---
name: word-edit
description: >-
  Content editing module for academic paper Word documents. Handles text
  replacement, paragraph rewriting, section block replacement, insertion,
  deletion, and tracked changes mode. Auto-selects the most efficient tool
  path for each operation.
  Triggers: 修改内容, 改写, 替换, 把XX改成YY, change, rewrite, modify,
  replace, 添加段落, 删除段落, tracked changes, 修订模式.
allowed-tools: Read Write Edit Bash Glob Grep mcp__word-document-server__create_document mcp__word-document-server__create_custom_style mcp__word-document-server__copy_document mcp__word-document-server__add_heading mcp__word-document-server__format_text mcp__word-document-server__search_and_replace mcp__word-document-server__insert_line_or_paragraph_near_text mcp__word-document-server__replace_paragraph_block_below_header mcp__word-document-server__replace_block_between_manual_anchors mcp__word-document-server__delete_paragraph mcp__word-document-server__add_paragraph mcp__word-document-server__insert_header_near_text mcp__word-document-server__insert_numbered_list_near_text mcp__word-document-server__find_text_in_document mcp__word-document-server__get_paragraph_text_from_document mcp__word-document-server__get_document_outline
metadata:
    version: "0.1.0"
    category: editing
    upstream-skills: [word-read]
    downstream-skills: [word-check]
---

# Word Content Editor

## Overview

Word Content Editor modifies text content in Word documents while preserving formatting. It auto-selects the most efficient editing tool for each operation. For tracked changes, it falls back to XML manipulation via the docx skill.

## When to Use

- Replacing words, phrases, or sentences in the document
- Rewriting one or more paragraphs
- Inserting new paragraphs or sections
- Deleting paragraphs
- Making changes with tracked changes (revision marks)
- Creating a new document from scratch (with template or format spec)

## When NOT to Use

- Changing fonts, spacing, or styles → use word-format
- Modifying tables or figures → use word-table-figure
- Managing references or citations → use word-reference
- Reviewing/responding to reviewer comments → use word-review

## Prerequisites

- Document Map from word-reader (recommended but not required for simple replacements)
- For tracked changes: docx skill scripts available (`unpack.py`, `pack.py`)

---

## Phase 1: Understand the Edit

1. **Parse user request** — What text to change, where, and how
2. **Locate target** — Use Document Map or `find_text_in_document` to find the target
3. **Select mode** — Choose the most efficient editing path (see Mode Selection below)
4. **Confirm with user** — For large changes, show preview before executing

## Phase 2: Mode Selection

```
IF user says "从零写" or "新建文档" or "create document" or "write from scratch"
  → New Document Mode (create_document + add_heading + add_paragraph)

ELIF user says "修订模式" or "tracked changes" or "保留修改痕迹"
  → Tracked Changes Mode (XML)

ELIF change is word/phrase replacement ("把X改成Y")
  → Precise Replace Mode (search_and_replace)

ELIF change is rewriting a paragraph under a specific heading
  → Block Replace Mode (replace_paragraph_block_below_header)

ELIF change is replacing content between two known anchors
  → Anchor Replace Mode (replace_block_between_manual_anchors)

ELIF change is inserting new content
  → Insert Mode (insert_line_or_paragraph_near_text)

ELIF change is deleting a paragraph
  → Delete Mode (delete_paragraph)

ELIF change is adding a heading
  → Header Insert Mode (insert_header_near_text)
```

## Phase 3: Execute

### Precise Replace Mode

```
1. Call search_and_replace(file_path, old_text, new_text)
2. If fails → try python-docx-replace via Bash:
   python3 -c "
   from docx import Document
   from docx_replace import docx_replace
   doc = Document('{file_path}')
   docx_replace(doc, '{old_text}', '{new_text}')
   doc.save('{file_path}')
   "
3. If still fails → fall back to XML edit
```

### Block Replace Mode

```
1. Identify target heading from Document Map
2. Call replace_paragraph_block_below_header(file_path, header_text, new_content)
```

### Insert Mode

```
1. Identify anchor text (existing text near insertion point)
2. Call insert_line_or_paragraph_near_text(file_path, anchor_text, new_text, position="after")
```

### Delete Mode

```
1. Identify paragraph index from Document Map
2. Call delete_paragraph(file_path, paragraph_index)
```

### Tracked Changes Mode

For edits that must show revision marks in Word:

```
1. Unpack document:
   python scripts/office/unpack.py document.docx unpacked/

2. Read target location in unpacked/word/document.xml

3. Apply tracked change XML:
   - For replacement: wrap old text in <w:del>, add new text in <w:ins>
   - For insertion: add <w:ins> block
   - For deletion: wrap text in <w:del>, change <w:t> to <w:delText>
   - Author: "Claude", Date: current ISO timestamp

4. Repack:
   python scripts/office/pack.py unpacked/ output.docx --original document.docx
```

**Tracked Changes XML patterns:**

Replacement ("30 days" → "60 days"):
```xml
<w:del w:id="1" w:author="Claude" w:date="2026-04-26T00:00:00Z">
  <w:r><w:rPr>{copy original rPr}</w:rPr><w:delText>30</w:delText></w:r>
</w:del>
<w:ins w:id="2" w:author="Claude" w:date="2026-04-26T00:00:00Z">
  <w:r><w:rPr>{copy original rPr}</w:rPr><w:t>60</w:t></w:r>
</w:ins>
```

**Critical rules for tracked changes:**
- Replace entire `<w:r>` elements, never inject tags inside a run
- Copy the original `<w:rPr>` formatting block into tracked change runs
- Use `<w:delText>` inside `<w:del>`, never `<w:t>`
- Each `w:id` must be unique across the document

## Phase 4: Verify

After executing edits:
1. Read the modified section to confirm changes applied correctly
2. If tracked changes mode, suggest user open in Word to review
3. Suggest running word-checker if the edit affected figures/tables/references

### New Document Mode

#### Critical Rules (防止字体混乱、标题不一致、目录失效)

**Rule 1: 样式先行，禁止直接格式化**
- 所有格式通过段落样式 (style) 控制，NEVER 对单个段落调用 `format_text` 设字体字号
- 先定义好所有样式，再写入内容。写入时只指定 `style_name`

**Rule 2: 使用 Word 内置标题样式（Heading 1/2/3）**
- 标题必须用 Word 内置的 `Heading 1`, `Heading 2`, `Heading 3` 样式
- 可以修改内置样式的外观（字体、字号），但 NEVER 创建自定义标题样式
- 这是 TOC 自动目录能工作的前提条件

**Rule 3: 中英文字体必须配对设置**
- 每个样式必须同时设置 `eastAsia`（中文字体）和 `ascii`/`hAnsi`（英文字体）
- 常见配对：宋体 + Times New Roman, 黑体 + Arial, 楷体 + Times New Roman
- 绝不能只设一个，否则 Word 会用默认字体填充另一个

**Rule 4: 一段一次写入**
- 每个段落用一次 `add_paragraph` 写完，NEVER 分多次追加文本到同一段落
- 多次写入会创建多个 `<w:r>` run，每个 run 可能携带不同的 `<w:rPr>` 格式

**Rule 5: TOC 必须用 field code 插入**
- 通过 XML 插入 TOC 字段（见 word-format SKILL.md TOC 章节）
- 绝不能用纯文本模拟目录

---

Two paths depending on whether user provides a template or a format spec:

**Path A: Template-based (preferred, 最可靠)**

```
1. User provides template.docx
   → copy_document(template_path, output_path)
   → All styles, page setup, headers/footers, TOC field are inherited

2. word-reader analyzes template:
   → Identify existing heading styles and their formatting
   → Identify placeholder text / sections to replace

3. Write content section by section:
   → For each section heading already in template:
     replace_paragraph_block_below_header(output_path, heading, content)
   → For missing sections:
     insert_line_or_paragraph_near_text(output_path, anchor, content, "after")
   → Each paragraph written in ONE call with style_name matching template styles

4. Route to word-table-figure for tables/figures
5. Route to word-reference for reference list
6. Run word-checker as final gate
```

**Path B: Format spec-based (from blank)**

```
1. create_document(output_path)

2. Parse format spec into style rules (ask user confirm):
   - Page: A4, margins
   - Heading 1: 黑体/Arial, 16pt, bold, center
   - Heading 2: 黑体/Arial, 14pt, bold, left
   - Heading 3: 黑体/Arial, 13pt, left
   - Body: 宋体/Times New Roman, 12pt, first-line indent 2字符, 1.5倍行距
   - Caption (表/图): 宋体/Times New Roman, 10.5pt, center

3. Modify built-in heading styles via python-docx (NOT create_custom_style):
   python3 << 'PYEOF'
   from docx import Document
   from docx.shared import Pt, Cm
   from docx.enum.text import WD_ALIGN_PARAGRAPH

   doc = Document(output_path)

   # Modify built-in Heading 1
   h1 = doc.styles['Heading 1']
   h1.font.name = 'Arial'             # English font
   h1.font.size = Pt(16)
   h1.font.bold = True
   h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
   # Set Chinese font via XML
   from docx.oxml.ns import qn
   h1.font.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

   # Modify built-in Heading 2
   h2 = doc.styles['Heading 2']
   h2.font.name = 'Arial'
   h2.font.size = Pt(14)
   h2.font.bold = True
   h2.font.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

   # Modify built-in Heading 3
   h3 = doc.styles['Heading 3']
   h3.font.name = 'Arial'
   h3.font.size = Pt(13)
   h3.font.element.rPr.rFonts.set(qn('w:eastAsia'), '黑体')

   # Modify Normal (body text base)
   body = doc.styles['Normal']
   body.font.name = 'Times New Roman'
   body.font.size = Pt(12)
   body.font.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
   body.paragraph_format.first_line_indent = Cm(0.74)  # 2字符
   body.paragraph_format.line_spacing = 1.5

   # Page setup
   for section in doc.sections:
       section.page_width = Cm(21.0)
       section.page_height = Cm(29.7)
       section.top_margin = Cm(2.54)
       section.bottom_margin = Cm(2.54)
       section.left_margin = Cm(3.17)
       section.right_margin = Cm(3.17)

   doc.save(output_path)
   PYEOF

4. Build document structure top-down:
   → add_heading(output_path, "摘  要", level=1)
   → add_paragraph(output_path, abstract_text)  # inherits Normal style
   → add_heading(output_path, "第1章 绪论", level=1)
   → add_heading(output_path, "1.1 研究背景", level=2)
   → add_paragraph(output_path, section_text)
   → ... continue for all sections
   → Each add_paragraph inherits Normal style automatically (fonts correct)
   → Each add_heading inherits modified Heading N style (fonts correct)

5. Insert TOC field via word-format module (XML field code, NOT plain text)
6. Route to word-table-figure for tables/figures
7. Route to word-reference for reference list
8. Run word-checker as final gate
```

**User Confirmation Points:**
- After step 2 (Path B): "格式规则解析如下，是否正确？"
- After building structure: "文档大纲如下，确认后开始填写内容"
- After completion: Show word-checker results

**Common Anti-Patterns (NEVER DO):**
- ❌ `format_text(p_idx, font="宋体")` — 只设中文字体，英文变 Calibri
- ❌ `create_custom_style("MyHeading1", ...)` — 自定义标题样式，TOC 找不到
- ❌ 分两次写同一段落 — 产生两个 run，可能字体不同
- ❌ `add_paragraph(text); format_text(p_idx, ...)` — 先写后格式化，应该用样式
- ✅ 修改内置样式 → 写入内容时自动继承 → 全文一致

---

## Batch Editing

When user requests multiple changes:
1. List all changes and confirm with user
2. Execute in document order (top to bottom) to avoid position shifts
3. For multiple search_and_replace calls, execute sequentially
4. Report results for each change

## Shared Resources

- `../../references/tool_routing.md` — Tool selection priority and fallback rules
- `../../references/token_budget.md` — Token efficiency rules
- `../../references/document_creation_rules.md` — New document anti-pattern rules (字体配对、内置样式、TOC)
