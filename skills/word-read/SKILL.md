---
name: word-read
description: >-
  Document analysis and comparison module. Reads Word documents and generates
  structured Document Maps with section structure, formatting summary, asset
  inventory, and format issue detection. Compares two document versions and
  produces structured diff reports. Triggers: 读取, 分析, 对比, 比较, read,
  analyze, compare, diff, 看看这个文档, 这两版有什么区别.
  Not for editing (word-edit) or formatting (word-format).
allowed-tools: Read Bash Glob Grep mcp__word-document-server__get_document_info mcp__word-document-server__get_document_outline mcp__word-document-server__get_document_text mcp__word-document-server__get_paragraph_text_from_document mcp__word-document-server__find_text_in_document mcp__word-document-server__get_all_comments mcp__word-document-server__get_comments_by_author mcp__word-document-server__validate_document_footnotes mcp__word-document-server__get_document_xml mcp__word-document-server__list_available_documents
metadata:
    version: "0.1.0"
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

### Step 4: Comment & Footnote Check (If Present)

```
Call: get_all_comments(file_path)        → if document has comments
Call: validate_document_footnotes(file_path) → if document has footnotes
```

### Step 5: Format Sampling

Read a few representative paragraphs to detect formatting patterns:
- First body paragraph (typical body formatting)
- First heading of each level
- First table caption / figure caption (if any)
- First reference entry

Use `get_paragraph_text_from_document` for targeted reads.

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
- Footnotes: {count}
- Endnotes: {count}
- Comments: {count} (by: {author1} ×{n}, {author2} ×{n})

### Format Issues Detected
- ⚠ P{n}: {description}
...
(empty section if no issues)
```

### Output

Present the Document Map to the user. This map is also passed to downstream modules via conversation context.

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
4. **Cache the Document Map** — downstream modules reference it from context, never regenerate
5. **In Compare Mode, reuse existing maps** — if one document was already analyzed, don't re-analyze it

## Shared Resources

- `../../references/token_budget.md` — Token budget rules and anti-patterns
- `../../references/tool_routing.md` — Tool selection priority
