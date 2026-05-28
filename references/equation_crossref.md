# Equations and Cross-References Reference

## Overview

This document covers inserting mathematical equations and cross-references into Word documents via word-mcp-live. Equations use OMML (Office Math Markup Language), the native Word equation format. Cross-references create live links to headings, figures, tables, and equations.

## Equations

### Backend

Equation insertion is only available via **word-mcp-live** (`insert_equation`). word-document-server and docx-mcp do not support equation insertion.

### OMML Format

Word uses OMML internally, not LaTeX or MathML. word-mcp-live accepts LaTeX input and converts it to OMML:

```
mcp__word-mcp-live__insert_equation(
  file_path,
  latex: "E = mc^2",
  position: "after",
  anchor_text: "根据爱因斯坦质能方程",
  display: true,        # true = display equation (centered, own line); false = inline
  equation_number: "(1)" # optional right-aligned numbering
)
```

### Equation Types

| Type | `display` | Use Case | Example |
|------|-----------|----------|---------|
| Display | `true` | Standalone centered equation | `$$E = mc^2 \quad (1)$$` |
| Inline | `false` | Within running text | "where $\alpha$ is the significance level" |

### Equation Numbering

Academic convention: right-aligned `(1)`, `(2)`, etc.

```
For numbered display equations:
  insert_equation(... display=true, equation_number="(1)")
  → Produces: centered equation with right-aligned "(1)"

For unnumbered display equations:
  insert_equation(... display=true)
  → Produces: centered equation, no number
```

### Common LaTeX Patterns

| Expression | LaTeX Input |
|-----------|-------------|
| Fraction | `\frac{a}{b}` |
| Subscript/superscript | `x_{i}^{2}` |
| Greek letters | `\alpha, \beta, \gamma` |
| Summation | `\sum_{i=1}^{n} x_i` |
| Integral | `\int_{0}^{\infty} f(x) dx` |
| Matrix | `\begin{pmatrix} a & b \\ c & d \end{pmatrix}` |
| Square root | `\sqrt{x^2 + y^2}` |

### Fallback

If word-mcp-live is unavailable, equations can be inserted as OMML XML directly via the XML unpack/edit/pack pipeline, but this is complex and error-prone. Recommend installing word-mcp-live for equation support.

## Cross-References

### Backend

Cross-reference insertion is available via **word-mcp-live** (`insert_cross_reference`). This creates live links that update when the document changes.

```
mcp__word-mcp-live__insert_cross_reference(
  file_path,
  ref_type: "figure",    # "heading", "figure", "table", "equation", "bookmark"
  ref_target: "图1",      # caption/heading text to reference
  ref_format: "label_number",  # what to display: "label_number", "page", "text"
  position: "after",
  anchor_text: "如"       # insert after this text → produces "如图1"
)
```

### Reference Types

| `ref_type` | Target | Example Output |
|-----------|--------|---------------|
| `heading` | Heading text | "第2章" or "2.1" |
| `figure` | Figure caption | "图1" or "Figure 1" |
| `table` | Table caption | "表2" or "Table 2" |
| `equation` | Equation number | "式(1)" or "Eq. (1)" |
| `bookmark` | Named bookmark | Custom text |

### Reference Formats

| `ref_format` | Displays |
|-------------|----------|
| `label_number` | The label and number (e.g., "图1") |
| `number_only` | Just the number (e.g., "1") |
| `page` | The page number |
| `text` | The full caption/heading text |

### Cross-Reference vs Plain Text

| Aspect | Cross-Reference (live) | Plain Text ("图1") |
|--------|----------------------|-------------------|
| Updates automatically | Yes | No |
| Ctrl+click navigation | Yes | No |
| Breaks on renumber | No | Yes |
| Requires Word | Yes (field code) | No |

**Recommendation:** Always use live cross-references for figures, tables, and equations. Use plain text only for informal references or when word-mcp-live is unavailable.

### Fallback

Without word-mcp-live, cross-references can be inserted as field codes via XML:

```xml
<w:fldSimple w:instr=" REF _Ref123456789 \h ">
  <w:r><w:t>图1</w:t></w:r>
</w:fldSimple>
```

This requires knowing the bookmark name of the target, which is complex. Prefer word-mcp-live.

## Integration Points

### word-edit: Equation Insert Mode

```
IF user says "插入公式" / "insert equation" / "add formula"
  → Equation Insert Mode
  → Collect: LaTeX expression, position, display/inline, numbering
  → Call: mcp__word-mcp-live__insert_equation(...)
```

### word-edit: Cross-Reference Insert

```
IF user says "插入引用" / "cross-reference" / "引用图/表/式"
  → Cross-Reference Insert Mode
  → Collect: reference type, target, format
  → Call: mcp__word-mcp-live__insert_cross_reference(...)
```

### word-table-figure: Equation as Asset

Equations are assets alongside tables and figures. word-table-figure can use `insert_equation` when creating equation blocks with numbering. The equation number contributes to the cross-reference check in word-checker.

## Font Normalization

Equation content uses Cambria Math font by default (Word standard). Font normalization (`normalize_fonts.py`) should NOT touch equation runs (`<m:r>` elements in OMML namespace). The script already skips these.
