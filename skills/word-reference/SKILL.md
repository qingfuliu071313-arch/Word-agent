---
name: word-reference
description: >-
  Reference and citation management module for academic paper Word documents.
  Formats reference lists using citeproc-py with CSL styles, manages
  footnotes/endnotes, integrates with Zotero for bibliography generation,
  and checks citation consistency.
  Triggers: 参考文献, 引用, 文献格式, 脚注, 尾注, Zotero, references,
  bibliography, citation, footnote, endnote.
allowed-tools: Read Write Bash Glob Grep mcp__word-document-server__add_footnote_to_document mcp__word-document-server__add_endnote_to_document mcp__word-document-server__delete_footnote_from_document mcp__word-document-server__validate_document_footnotes mcp__word-document-server__search_and_replace mcp__word-document-server__find_text_in_document mcp__word-document-server__get_document_text mcp__word-document-server__get_paragraph_text_from_document mcp__word-document-server__add_paragraph mcp__word-document-server__delete_paragraph mcp__word-document-server__replace_paragraph_block_below_header mcp__zotero__search_items mcp__zotero__get_item mcp__zotero__export_bibliography mcp__zotero__get_item_children mcp__semantic-scholar__search_papers mcp__semantic-scholar__get_paper
metadata:
    version: "0.1.0"
    category: editing
    upstream-skills: [word-read]
    downstream-skills: [word-check]
---

# Word Reference Manager

## Overview

Word Reference Manager handles all citation and reference-related operations in academic paper Word documents. It integrates with Zotero for automated bibliography generation, uses citeproc-py for precise citation formatting, and manages footnotes/endnotes.

## When to Use

- Formatting a reference list per a specific citation style
- Pulling references from Zotero and inserting into the document
- Checking that in-text citations match the reference list
- Adding, editing, or deleting footnotes/endnotes
- Converting references between citation styles

## When NOT to Use

- Checking cross-references (figures/tables) → use word-check
- Editing body text → use word-edit
- Formatting non-reference parts → use word-format

---

## Feature A: Reference List Formatting

### With Zotero Integration

```
Step 1: Identify references
  → Scan document for in-text citations: [1], [2] or (Author, Year)
  → Or user provides a Zotero collection name

Step 2: Fetch from Zotero
  → search_items / get_item to retrieve reference metadata
  → export_bibliography for batch export

Step 3: Format with citeproc-py
  → Load appropriate CSL style file
  → Generate formatted reference list

Step 4: Write to document
  → replace_paragraph_block_below_header: replace content under "参考文献" or "References"
  → Or add_paragraph for each reference entry
  → format_text: apply reference font/size per Format Spec
```

### Without Zotero (Manual Formatting)

```
Step 1: Read existing reference list from document
Step 2: Parse each entry to extract metadata (author, year, title, journal...)
Step 3: Reformat each entry per the target citation style
Step 4: Replace the reference list in the document
```

### Using citeproc-py

```bash
python3 << 'PYEOF'
import json
from citeproc import CitationStylesStyle, CitationStylesBibliography
from citeproc import Citation, CitationItem
from citeproc.source.json import CiteProcJSON

# Load CSL style
style = CitationStylesStyle('references/csl/gb-t-7714-2015-numeric.csl')

# Reference data (from Zotero or manual input)
references = [
    {
        "id": "ref1",
        "type": "article-journal",
        "title": "Biodiversity and ecosystem functioning",
        "author": [{"family": "Tilman", "given": "David"}],
        "issued": {"date-parts": [[2001]]},
        "container-title": "Nature",
        "volume": "411",
        "page": "689-695"
    }
]

source = CiteProcJSON(references)
bib = CitationStylesBibliography(style, source)

# Register citations
for ref in references:
    bib.register(Citation([CitationItem(ref["id"])]))

# Generate formatted bibliography
for item in bib.bibliography():
    print(str(item))
PYEOF
```

## Feature B: Citation Consistency Check

```
Step 1: Extract in-text citations
  → Numbered: [1], [2], [1-3], [1,3,5]
  → Author-year: (Author, Year), (Author1 & Author2, Year)

Step 2: Extract reference list entries
  → Parse numbered entries: [1] Author. Title...
  → Parse author-year entries: Author (Year). Title...

Step 3: Cross-check
  → Missing: cited in text but not in reference list
  → Orphan: in reference list but never cited
  → Order: numbered references should appear in citation order

Step 4: Report
  → List all inconsistencies with locations
```

## Feature C: Footnote/Endnote Management

### Add Footnotes
```
→ add_footnote_to_document(file_path, paragraph, text)
```

### Validate Footnotes
```
→ validate_document_footnotes(file_path)
→ Report any broken or orphaned footnotes
```

### Delete Footnotes
```
→ delete_footnote_from_document(file_path, footnote_id)
```

### Batch Operations
For adding multiple footnotes, execute sequentially to avoid position conflicts.

## Feature D: Style Conversion

Convert the entire reference list from one citation style to another:

```
1. Parse existing references to extract metadata
2. Load target CSL style
3. Re-format all entries with citeproc-py
4. Replace reference list in document
5. Update in-text citations if format changes (e.g., [1] → (Author, Year))
```

---

## Post-Execution

After any reference operation:
1. Report changes made
2. Suggest running word-checker to verify citation consistency
3. If reference format changed, remind to check in-text citation format matches

## Shared Resources

- `../../references/academic_formatting.md` — Citation style overview
- `../../references/format_spec_parser.md` — Reference formatting in Format Spec
- `../../references/tool_routing.md` — Tool selection priority
