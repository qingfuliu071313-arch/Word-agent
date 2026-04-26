---
name: word-table-figure
description: >-
  Table and figure module for academic paper Word documents. Creates three-line
  tables, formats existing tables, inserts images with captions, and manages
  numbering consistency.
  Triggers: 表格, 三线表, 图片, 插图, 题注, table, figure, image, caption.
allowed-tools: Read Bash Glob Grep mcp__word-document-server__add_table mcp__word-document-server__format_table mcp__word-document-server__format_table_cell_text mcp__word-document-server__set_table_width mcp__word-document-server__set_table_column_widths mcp__word-document-server__auto_fit_table_columns mcp__word-document-server__merge_table_cells_horizontal mcp__word-document-server__merge_table_cells_vertical mcp__word-document-server__highlight_table_header mcp__word-document-server__set_table_cell_shading mcp__word-document-server__set_table_cell_alignment mcp__word-document-server__set_table_cell_padding mcp__word-document-server__set_table_alignment_all mcp__word-document-server__apply_table_alternating_rows mcp__word-document-server__add_picture mcp__word-document-server__add_paragraph mcp__word-document-server__format_text mcp__word-document-server__find_text_in_document mcp__word-document-server__get_paragraph_text_from_document
metadata:
    version: "0.1.0"
    category: editing
    upstream-skills: [word-read]
    downstream-skills: [word-check]
---

# Word Table & Figure

## Overview

This module creates and formats tables and figures in academic paper Word documents. Its signature feature is one-command three-line table generation following academic conventions.

## When to Use

- Creating a new table (especially three-line tables)
- Formatting an existing table (column widths, alignment, borders)
- Inserting an image with a formatted caption
- Merging table cells
- Adjusting table/figure numbering

## When NOT to Use

- Changing document-wide formatting → use word-format
- Editing text content → use word-edit
- Checking figure/table cross-references → use word-check

---

## Feature A: Three-Line Table Creation

### Input

User provides:
- Table data (header row + data rows)
- Optional: column widths, alignment, caption text
- Optional: table number (or auto-detect next number from Document Map)

### Process

```
Step 1: Add caption paragraph
  → add_paragraph: "表{n} {caption_text}"
  → format_text: apply caption font per Format Spec

Step 2: Create table
  → add_table: create table with header + data rows

Step 3: Apply three-line style
  → format_table: remove all borders first
  → format_table: set top border (1.0pt solid black)
  → format_table: set header-bottom border (0.5pt solid black)
  → format_table: set bottom border (1.0pt solid black)

Step 4: Format content
  → set_table_cell_alignment: center header, left/right-align data as needed
  → set_table_cell_padding: add consistent internal padding
  → format_table_cell_text: apply font per Format Spec

Step 5: Set dimensions
  → set_table_width: full content width
  → set_table_column_widths: user-specified or auto-fit
```

### Three-Line Table Border Specification

```
Border positions and weights:
┌─────────────────────────────────┐  ← Top line: 1.0-1.5pt solid black
│  Header 1  │  Header 2  │  ... │
├─────────────────────────────────┤  ← Header line: 0.5-0.75pt solid black
│  Data 1    │  Data 2    │  ... │
│  Data 3    │  Data 4    │  ... │
└─────────────────────────────────┘  ← Bottom line: 1.0-1.5pt solid black

No vertical lines.
No other horizontal lines between data rows.
```

## Feature B: Format Existing Table

### Input
- Table location (from Document Map or user description)
- Desired format (three-line, column widths, alignment)

### Process
1. Locate table by index or nearby text
2. Apply requested formatting changes
3. Optionally convert to three-line style

## Feature C: Figure Insertion

### Input
- Image file path
- Caption text
- Figure number (or auto-detect)
- Optional: width/height constraints

### Process

```
Step 1: Insert image
  → add_picture: insert at specified location with dimensions

Step 2: Add caption below
  → add_paragraph: "图{n} {caption_text}"
  → format_text: apply caption font per Format Spec
```

## Feature D: Cell Merging

### Input
- Table location
- Cells to merge (row/column ranges)
- Direction (horizontal or vertical)

### Process
```
→ merge_table_cells_horizontal or merge_table_cells_vertical
→ Re-apply alignment to merged cell
```

## Feature E: Numbering Check

Before creating a new table/figure:
1. Read Document Map assets section
2. Determine next sequential number
3. If user specifies a number, verify it doesn't conflict

---

## Post-Execution

After creating or modifying tables/figures:
1. Report what was done
2. Suggest running word-checker to verify cross-references still valid

## Shared Resources

- `../../references/academic_formatting.md` — Three-line table specification, caption conventions
- `../../references/tool_routing.md` — Tool selection priority
