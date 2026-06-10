---
name: word-read
description: >-
  Document analysis and comparison module. Reads Word documents and generates
  structured Document Maps with section structure, formatting summary, asset
  inventory, and format issue detection. Compares two document versions and
  produces structured diff reports. Triggers: 读取, 分析, 对比, 比较, read,
  analyze, compare, diff, 看看这个文档, 这两版有什么区别.
  Not for editing (word-edit) or formatting (word-format).
allowed-tools: Read Write Bash Glob Grep mcp__word-document-server__get_document_info mcp__word-document-server__get_document_outline mcp__word-document-server__get_document_text mcp__word-document-server__get_paragraph_text_from_document mcp__word-document-server__find_text_in_document mcp__word-document-server__get_all_comments mcp__word-document-server__get_comments_by_author mcp__word-document-server__validate_document_footnotes mcp__word-document-server__get_document_xml mcp__word-document-server__list_available_documents mcp__docx-mcp__get_tracked_changes
metadata:
    version: "1.1.0"
    category: analysis
    downstream-skills: [word-format, word-edit, word-reference, word-table-figure, word-review, word-check, word-submit]
---

# Word Reader

## Overview

Word Reader is the foundation module of word-agent. It reads documents once and produces a comprehensive Document Map that all downstream modules consume. It also compares two document versions to identify changes. It never modifies documents.

## When to Use

- Analyzing a document's structure and formatting before other operations
- Generating a Document Map for downstream modules
- Comparing two versions of a document
- Answering questions about document content or structure
- Detecting formatting inconsistencies

## When NOT to Use

- Editing content → use word-edit
- Fixing formatting → use word-format
- Any operation that modifies the document

## Mode Detection

```
IF user provides TWO documents or mentions "对比/compare/diff/区别/改了什么"
  → Compare Mode

ELIF user provides ONE document or mentions "读取/分析/看看/analyze/read"
  → Analysis Mode

ELIF downstream module needs specific paragraphs
  → Targeted Read Mode
```

---

## Analysis Mode: Document Map Generation

### Step 0: File Format Gate

Before any analysis, check if the input is a `.doc` file:

```
IF file_path ends with ".doc" (case-insensitive):
  → Abort analysis
  → Return to orchestrator: "文件为旧版 .doc 格式，需要先转换为 .docx。请通过 word-orchestrate 触发转换流程。"
  → Do NOT attempt to call docx2python or MCP tools on .doc files — they will fail
```

This ensures `.doc` files are always handled by the orchestrator's conversion pipeline before reaching the reader.

### Step 1: Basic Info (Cheap)

```
Call: get_document_info(file_path)
Extract: page count, file size, metadata
```

### Step 2: Structure (Cheap)

```
Call: get_document_outline(file_path)
Extract: heading hierarchy with levels
```

### Step 3: Deep Structure (One Call)

Run via Bash — `docx2python` extracts everything in one pass:

```bash
python3 << 'PYEOF'
import json
from docx2python import docx2python

doc = docx2python("{file_path}")

info = {
    "body_paragraphs": len(doc.body),
    "tables": [],
    "images": list(doc.images.keys()) if doc.images else [],
    "footnotes_count": len(doc.footnotes) if doc.footnotes else 0,
    "endnotes_count": len(doc.endnotes) if doc.endnotes else 0,
    "headers": doc.headers,
    "footers": doc.footers,
}

# Extract table locations and sizes
for i, item in enumerate(doc.body):
    text = str(item)
    if isinstance(item, list) and len(item) > 0 and isinstance(item[0], list):
        info["tables"].append({"index": i, "rows": len(item)})

print(json.dumps(info, ensure_ascii=False, indent=2))
PYEOF
```

### Step 3.5: Text Box Extraction (Critical)

**`get_document_text` and `docx2python` both skip text boxes entirely.** Text boxes (`w:txbxContent`) live in a separate XML layer (inside DrawingML or VML shapes) and are invisible to standard text extraction. They also frequently live **outside document.xml** — university thesis templates commonly place format annotations in header text boxes (e.g. `word/header2.xml`), so scanning document.xml alone misses them.

Run the dedicated script — it scans ALL XML parts (document, headers, footers, notes), covers both DrawingML and VML text boxes, and dedupes mc:Fallback copies:

```bash
python3 scripts/extract_textboxes.py "{file_path}"          # JSON
python3 scripts/extract_textboxes.py "{file_path}" --text   # human-readable
```

The JSON output includes `parts_with_textboxes`, attributing each text box to its source part. If text boxes are found, include their content (with source part) in the Document Map under the **Text Boxes** section. This is critical for format requirement templates where formatting rules are annotated inside text boxes.

### Step 4: Comment, Footnote & Tracked Changes Check (If Present)

```
Call: get_all_comments(file_path)           → if document has comments
Call: validate_document_footnotes(file_path) → if document has footnotes
Call: mcp__docx-mcp__get_tracked_changes(file_path) → if document may have revisions
  → Extract: total count, insertions vs deletions, author distribution, affected sections
  → If docx-mcp unavailable, scan XML for <w:ins>/<w:del> elements as fallback
```

### Step 5: Format Sampling

Read a few representative paragraphs to detect formatting patterns:
- First body paragraph (typical body formatting)
- First heading of each level
- First table caption / figure caption (if any)
- First reference entry

Use `get_paragraph_text_from_document` for targeted reads.

### Step 5.5: Font Consistency Check

Scan the document for font inconsistencies using the font normalization script. This detects missing `w:eastAsia` attributes, theme font references in styles.xml, bare runs without any font specification, and mixed fonts within paragraphs.

```bash
python3 scripts/normalize_fonts.py "{file_path}" --detect-only --json
```

The script checks three layers:
- **styles.xml**: theme font references (`*Theme` attributes) in docDefaults and style definitions
- **document.xml runs with rFonts**: missing eastAsia/ascii, mixed fonts, hAnsi mismatches
- **document.xml bare runs**: text runs with no `w:rPr` or no `w:rFonts` that silently inherit theme fonts

Include any detected font issues in the "Format Issues Detected" section of the Document Map. If issues are found, add a summary note recommending font normalization via word-format:

```
⚠ Font issues: {n} problems detected (theme refs, missing eastAsia, bare runs)
  → Recommend: python3 scripts/normalize_fonts.py "{file_path}" --unify
```

See `../../references/font_normalization.md` for technical background.

### Step 6: Assemble Document Map

Compile all gathered information into the standard Document Map format:

```markdown
## Document Map: {filename}

### Basic Info
- File: {absolute_path}
- Pages: {n}, Paragraphs: {n}, Language: {zh/en/mixed}
- Last modified: {date}

### Formatting Summary
- Body font: {font_name}, {size}
- Heading fonts: L1={font,size,bold}, L2={font,size}, ...
- Line spacing: body={n}, references={n}
- Margins: top={n} bottom={n} left={n} right={n}
- Page size: {A4/Letter/...}
- Page numbers: {position, starting page}
- Headers/Footers: {present/absent, content summary}

### Document Structure
- [P{start}-P{end}] {section_name}
  - [P{start}-P{end}] {subsection_name}
  ...

### Assets
- Tables: Table 1 (P{n}), Table 2 (P{n}), ...
- Figures: Fig 1 (P{n}), Fig 2 (P{n}), ...
- Text Boxes: {count} (content summary below if present)
- Footnotes: {count}
- Endnotes: {count}
- Comments: {count} (by: {author1} ×{n}, {author2} ×{n})
- Tracked Changes: {count} (ins: {n}, del: {n}, by: {author1} ×{n}, {author2} ×{n})

### Text Boxes (if any)
- [TextBox 1] ({source part, e.g. word/header2.xml}): "{first 100 chars of content}..."
- [TextBox 2] ({source part}): "{first 100 chars of content}..."
...
(omit this section if no text boxes found)

### Tracked Changes (if any)
- Total: {count} changes ({ins_count} insertions, {del_count} deletions)
- Authors: {author1} ×{n}, {author2} ×{n}
- Scope: {affected_sections_summary}
...
(omit this section if no tracked changes found)

### Format Issues Detected
- ⚠ P{n}: {description}
- ⚠ Font inconsistency: P{n} has multiple ascii fonts {Font1, Font2}
- ⚠ Font incomplete: P{n} eastAsia font missing (ascii set but eastAsia not)
...
(empty section if no issues)
```

### Step 7: Persist the Document Map to Disk (MANDATORY)

Write the assembled Document Map to a sidecar file next to the document (use the Write tool):

```
路径规则: {文档所在目录}/.word-agent/{文档名去扩展名}.map.md
示例:     /path/to/paper.docx → /path/to/.word-agent/paper.map.md
```

1. Write the full Document Map markdown to that path (the Write tool creates `.word-agent/` automatically)
2. **下游模块优先读此文件，不重读文档** — downstream modules (word-format, word-edit, word-check, ...) MUST Read this file first instead of re-reading the document
3. 失效判定靠 mtime：map 文件的 mtime 晚于 .docx 的 mtime 才有效。任何模块对 .docx 的写操作会让 docx 的 mtime 变新，map 自动视为过期（无需手动标记），由 orchestrator 在下次任务时派 word-read 重新生成

### Output

Present the Document Map to the user, and confirm the sidecar file path (e.g. "Document Map 已写入 {dir}/.word-agent/{name}.map.md，下游模块将直接读取该文件"). The map is delivered to downstream modules via this file — not via conversation context alone.

---

## Compare Mode: Document Diff

### Step 1: Generate Maps for Both Documents

Run Analysis Mode (Steps 1-6) for both documents. If comparing within the same conversation and one map already exists, only generate the missing one.

### Step 2: Content Comparison

```bash
python3 << 'PYEOF'
import json
from docx2python import docx2python
from deepdiff import DeepDiff

doc_a = docx2python("{file_a}")
doc_b = docx2python("{file_b}")

# Structural diff
diff = DeepDiff(
    [str(p) for p in doc_a.body],
    [str(p) for p in doc_b.body],
    ignore_order=False,
    verbose_level=2
)

print(json.dumps(json.loads(diff.to_json()), ensure_ascii=False, indent=2))
PYEOF
```

### Step 3: Text-Level Diff (For Changed Sections)

For sections identified as changed, use `redlines` for readable diff:

```bash
python3 << 'PYEOF'
from redlines import Redlines

old_text = """{old_section_text}"""
new_text = """{new_section_text}"""

diff = Redlines(old_text, new_text)
print(diff.output_markdown)
PYEOF
```

### Step 4: Assemble Diff Report

```markdown
## Diff Report: {file_A} → {file_B}

### Structural Changes
- Added: {new sections/subsections}
- Removed: {deleted sections}
- Moved: {relocated sections}

### Content Changes ({n} total)
- [P{n}] {section}: {brief description}
  - Before: "{first 80 chars}..."
  - After: "{first 80 chars}..."
...

### Formatting Changes
- {property}: {old_value} → {new_value}
...

### Asset Changes
- Tables: +{added} -{removed} ~{modified}
- Figures: +{added} -{removed} ~{modified}
- References: +{added} -{removed}
- Comments: +{added} -{removed}
```

---

## Targeted Read Mode

When called by a downstream module for specific content:

1. Receive request: "read paragraphs {start} to {end}"
2. Call `get_paragraph_text_from_document(file_path, start, end)`
3. Return raw content — no Document Map regeneration

---

## Token Efficiency Rules

1. **Never call `get_document_text` as the first step** — always start with `get_document_info` + `get_document_outline`
2. **Use `docx2python` for deep analysis** — one Python call replaces dozens of MCP calls
3. **Sample formatting, don't scan everything** — read representative paragraphs, not all paragraphs
4. **Cache the Document Map** — the map is persisted to `{doc_dir}/.word-agent/{name}.map.md` (Step 7); downstream modules Read that file, never regenerate (unless the docx mtime is newer than the map's)
5. **In Compare Mode, reuse existing maps** — if one document was already analyzed, don't re-analyze it

## Shared Resources

- `../../references/token_budget.md` — Token budget rules and anti-patterns
- `../../references/font_normalization.md` — Font inconsistency detection script
- `../../references/tool_routing.md` — Tool selection priority
