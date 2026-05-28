---
name: word-check
description: >-
  Document quality checker for academic papers. Performs cross-reference
  validation (figures, tables, equations) and format compliance verification
  against user-provided requirements. Reports issues with precise locations
  and severity levels without modifying the document.
  Triggers: 检查引用, 交叉引用, 检查格式, 格式验证, 核对, 是否符合要求,
  check references, cross-reference, verify format, compliance.
allowed-tools: Read Bash Glob Grep mcp__word-document-server__get_document_text mcp__word-document-server__get_document_outline mcp__word-document-server__find_text_in_document mcp__word-document-server__get_document_info mcp__word-document-server__validate_document_footnotes mcp__word-document-server__get_all_comments mcp__word-document-server__get_paragraph_text_from_document mcp__word-document-server__get_document_xml mcp__docx-mcp__validate_paraids mcp__docx-mcp__audit_document mcp__word-mcp-live__diagnose_layout
metadata:
    version: "1.0.0"
    category: verification
    upstream-skills: [word-read, word-format, word-edit]
    downstream-skills: [word-format, word-submit]
---

# Word Checker

## Overview

Word Checker is the quality gate for academic paper Word documents. It performs two types of checks — cross-reference validation and format compliance verification — and produces structured reports. It is read-only: it finds and reports problems but never modifies the document.

## When to Use

- After formatting (word-format) to verify all changes applied correctly
- After content editing (word-edit) to verify cross-references still valid
- Before submission (word-submit) as a final quality gate
- Any time the user wants to validate their document

## When NOT to Use

- To fix issues → use word-format or word-edit
- To analyze document structure → use word-read

## Mode Detection

```
IF user mentions "交叉引用/cross-reference/检查引用/检查图表"
  → Cross-Reference Mode (Mode A)

ELIF user mentions "格式/format/compliance/是否符合要求/检查格式"
  → Format Compliance Mode (Mode B)

ELIF user mentions "结构/structural/paraId/bookmarks/文档结构"
  → Structural Validation Mode (Mode C)

ELIF user mentions "全面检查/full check/检查一下"
  → All Modes (A + B + C)

DEFAULT → Both Modes (A + B)
```

---

## Mode A: Cross-Reference Check

### Prerequisites
- Document Map from word-reader (or conversation context)
- Full document text (this mode requires full text scanning)

### Step 1: Extract All Asset Definitions

Scan document text for figure/table/equation definitions (captions):

```bash
python3 << 'PYEOF'
import re
import json

text = """${DOCUMENT_TEXT}"""

# Figure definitions (captions)
fig_defs_cn = re.findall(r'^[  ]*图\s*(\d+)[\s.．:：]', text, re.MULTILINE)
fig_defs_en = re.findall(r'^[  ]*(?:Figure|Fig\.)\s*(\d+)[\s.:]', text, re.MULTILINE)
fig_defs = sorted(set(int(n) for n in fig_defs_cn + fig_defs_en))

# Table definitions
tbl_defs_cn = re.findall(r'^[  ]*表\s*(\d+)[\s.．:：]', text, re.MULTILINE)
tbl_defs_en = re.findall(r'^[  ]*Table\s*(\d+)[\s.:]', text, re.MULTILINE)
tbl_defs = sorted(set(int(n) for n in tbl_defs_cn + tbl_defs_en))

# Equation definitions (right-aligned numbers)
eq_defs = re.findall(r'[（(]\s*(\d+)\s*[)）]\s*$', text, re.MULTILINE)
eq_defs = sorted(set(int(n) for n in eq_defs))

print(json.dumps({
    "figures_defined": fig_defs,
    "tables_defined": tbl_defs,
    "equations_defined": eq_defs
}, ensure_ascii=False))
PYEOF
```

### Step 2: Extract All References

Scan text for in-text references:

```bash
python3 << 'PYEOF'
import re
import json

text = """${DOCUMENT_TEXT}"""

# Figure references
fig_refs_cn = [(m.start(), int(m.group(1))) for m in re.finditer(r'(?:如)?图\s*(\d+)', text)]
fig_refs_en = [(m.start(), int(m.group(1))) for m in re.finditer(r'(?:Figure|Fig\.)\s*(\d+)', text)]
fig_refs = fig_refs_cn + fig_refs_en

# Table references
tbl_refs_cn = [(m.start(), int(m.group(1))) for m in re.finditer(r'(?:如|见)?表\s*(\d+)', text)]
tbl_refs_en = [(m.start(), int(m.group(1))) for m in re.finditer(r'Table\s*(\d+)', text)]
tbl_refs = tbl_refs_cn + tbl_refs_en

# Equation references
eq_refs = [(m.start(), int(m.group(1))) for m in re.finditer(r'(?:式|公式|Eq\.|Equation)\s*[（(]\s*(\d+)\s*[)）]', text)]

# Format consistency
fig_format_cn = len(re.findall(r'(?:如)?图\s*\d+', text))
fig_format_en = len(re.findall(r'(?:Figure|Fig\.)\s*\d+', text))
tbl_format_cn = len(re.findall(r'(?:如|见)?表\s*\d+', text))
tbl_format_en = len(re.findall(r'Table\s*\d+', text))

print(json.dumps({
    "figure_refs": [{"pos": p, "num": n} for p, n in fig_refs],
    "table_refs": [{"pos": p, "num": n} for p, n in tbl_refs],
    "equation_refs": [{"pos": p, "num": n} for p, n in eq_refs],
    "format_counts": {
        "fig_cn": fig_format_cn, "fig_en": fig_format_en,
        "tbl_cn": tbl_format_cn, "tbl_en": tbl_format_en
    }
}, ensure_ascii=False))
PYEOF
```

### Step 3: Apply Check Rules

See `../../references/cross_ref_rules.md` for the complete rule definitions.

Apply R1-R7 in order:

| Rule | Check | Severity |
|------|-------|----------|
| R1 | Reference target exists | ❌ ERROR |
| R2 | Numbering is sequential | ❌ ERROR |
| R3 | No duplicate numbers | ❌ ERROR |
| R4 | First-reference order | ⚠ WARNING |
| R5 | Format consistency | ⚠ WARNING |
| R6 | No orphan assets | ⚠ WARNING |
| R7 | Range reference validity | ℹ INFO |

### Step 4: Compile Report

```markdown
## Cross-Reference Check Report: {filename}

### Figures
- Defined: {list}
- Referenced: {list}
- ✅/❌ {details per rule}

### Tables
- Defined: {list}
- Referenced: {list}
- ✅/❌ {details per rule}

### Equations
- Defined: {list}
- Referenced: {list}
- ✅/❌ {details per rule}

### Summary
- ❌ Errors: {n}
- ⚠ Warnings: {n}
- ℹ Info: {n}
```

---

## Mode B: Format Compliance Check

### Prerequisites
- Document Map from word-reader
- Format Spec (from word-format session) OR user-provided format requirements

### Step 1: Obtain Format Spec

If Format Spec exists in conversation context → use it.
If not → ask user for format requirements, then parse using `../../references/format_spec_parser.md`.

### Step 2: Check Each Rule

For each non-null rule in the Format Spec, verify the current document state:

```
Page Setup:
  - Read get_document_info for page size, margins
  - Compare against Format Spec

Fonts:
  - Sample paragraphs from each category (body, headings, captions, references)
  - Use get_paragraph_text_from_document or get_document_xml for formatting details
  - Compare font family, size, bold/italic against Format Spec

Spacing:
  - Check line spacing from document XML or formatting info
  - Check paragraph before/after spacing

Numbering:
  - Check heading numbering format from outline
  - Compare against Format Spec pattern

Headers/Footers:
  - Check presence and content from document info
  - Compare against Format Spec

Special Rules:
  - Check each special rule individually
  - E.g., "首行缩进2字符" → verify indent settings
```

### Step 3: Compile Report

```markdown
## Format Compliance Report: {filename}

### ✅ Compliant ({n}/{total})
- ✅ Page size: A4 ✓
- ✅ Margins: top=2.54cm bottom=2.54cm left=3.17cm right=3.17cm ✓
- ✅ Body font: 宋体/Times New Roman, 小四 ✓
- ...

### ❌ Non-Compliant ({n}/{total})
- ❌ Line spacing: required=1.5x, actual=2.0x
  → Affected: P90, P91, P95 (3 paragraphs)
- ❌ Caption font size: required=五号(10.5pt), actual=小四(12pt)
  → Affected: Fig 1 (P92), Fig 3 (P115)
- ...

### ⚠ Unable to Verify ({n}/{total})
- ⚠ Image resolution ≥300dpi: cannot check via current tools
- ...

### Summary
- Compliant: {n}/{total} rules
- Non-compliant: {n}/{total} rules
- Recommendation: {具体建议}
```

---

## Mode C: Structural Validation

### Prerequisites
- docx-mcp server available (if unavailable, skip with warning)

### Step 1: Paragraph ID Validation

```
Call: mcp__docx-mcp__validate_paraids(file_path)
Check:
  - Zero duplicate paraIds
  - Zero missing paraIds
  - All paraIds are valid 8-digit hex format
```

### Step 2: Comprehensive Structural Audit

```
Call: mcp__docx-mcp__audit_document(file_path)
Check:
  - No broken bookmarks (bookmarkStart without matching bookmarkEnd)
  - No orphan image references (blip pointing to missing rId)
  - No invalid numbering definitions (numId referencing non-existent abstractNum)
  - No style conflicts (duplicate styleId)
  - No broken hyperlinks
  - No missing content types
  - No relationship mismatches
```

### Step 3: Compile Report

```markdown
## Structural Validation Report: {filename}

### Paragraph IDs
- Total paragraphs: {n}
- ✅/❌ Uniqueness: {details}
- ✅/❌ Format: {details}

### Document Integrity
- ✅/❌ Bookmarks: {n} checked, {n} broken
- ✅/❌ Image references: {n} checked, {n} orphaned
- ✅/❌ Numbering definitions: {n} checked, {n} invalid
- ✅/❌ Style definitions: {n} checked, {n} conflicts
- ✅/❌ Hyperlinks: {n} checked, {n} broken

### Summary
- Errors: {n}
- Warnings: {n}
```

### Fallback (docx-mcp unavailable)

```
⚠ docx-mcp 不可用，跳过结构验证（paraId、书签、图片引用完整性）。
  建议安装 docx-mcp 后重新运行检查。
```

Basic paraId duplicate scan via XML is available as minimal fallback. See `../../references/structural_validation.md`.

---

## Mode D: Layout Diagnostics (Live Mode Only)

### Prerequisites
- Word must be open with the target document (live editing mode)
- word-mcp-live must be available

### Trigger
```
"布局检查/layout check/排版诊断/页面溢出/orphan lines"
```

### Step 1: Diagnose Layout

```
Call: mcp__word-mcp-live__diagnose_layout(file_path)
Returns:
  - Actual page count vs expected
  - Orphan/widow line locations
  - Text overflow areas
  - Unexpected page breaks
  - Heading/paragraph spacing anomalies
```

### Step 2: Compile Report

```markdown
## Layout Diagnostics Report: {filename}

### Page Analysis
- Total pages: {actual} (expected: {expected})
- Unexpected page breaks: {list}

### Typography Issues
- Orphan lines: {count} (locations: P{n}, P{m}, ...)
- Widow lines: {count} (locations: P{n}, P{m}, ...)
- Text overflow: {count} areas

### Recommendations
- {specific suggestions based on findings}
```

### Fallback (Not in Live Mode)

```
⚠ 布局诊断仅在 Word 打开文档时可用（需要 Word 的排版引擎）。
  请在 Word 中打开文档后重新运行此检查。
```

---

## Combined Mode (A + B + C)

When all modes are requested (or "全面检查"):
1. Run Mode A (cross-reference) first
2. Run Mode B (format compliance)
3. Run Mode C (structural validation) — skipped with warning if docx-mcp unavailable
4. Run Mode D (layout diagnostics) — only if live mode available, otherwise skipped silently
5. Merge into a single combined report
6. Provide overall recommendation

---

## Post-Check Recommendations

Based on results, suggest next steps:

| Result | Recommendation |
|--------|---------------|
| All checks passed | "所有检查通过。可以使用 word-submit 准备投稿版本。" |
| Format issues only | "发现 {n} 处格式问题。使用 word-format 自动修复？" |
| Cross-ref issues only | "发现 {n} 处引用问题。需要手动修正后重新检查。" |
| Both types of issues | "发现 {n} 处格式问题和 {m} 处引用问题。建议先修正引用，再修复格式。" |

## Shared Resources

- `../../references/cross_ref_rules.md` — Cross-reference check rule definitions
- `../../references/format_spec_parser.md` — Format Spec structure
- `../../references/academic_formatting.md` — Academic formatting knowledge
- `../../references/structural_validation.md` — OOXML structural integrity rules (Mode C)
- `../../references/token_budget.md` — Token efficiency rules
