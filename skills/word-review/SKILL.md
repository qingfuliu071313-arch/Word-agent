---
name: word-review
description: >-
  Review and revision module for academic paper Word documents. Extracts
  reviewer comments, plans revision strategies, dispatches changes to
  specialist modules in tracked-changes mode, and generates point-by-point
  response documents.
  Triggers: 审稿意见, 修订, 审阅, 批注, reviewer comments, revision,
  response to reviewers, point-by-point.
allowed-tools: Read Write Edit Bash Glob Grep mcp__word-document-server__get_all_comments mcp__word-document-server__get_comments_by_author mcp__word-document-server__get_comments_for_paragraph mcp__word-document-server__get_document_text mcp__word-document-server__get_document_outline mcp__word-document-server__get_paragraph_text_from_document mcp__word-document-server__find_text_in_document mcp__word-document-server__search_and_replace mcp__word-document-server__get_document_xml
metadata:
    version: "0.1.0"
    category: review
    upstream-skills: [word-read]
    downstream-skills: [word-edit, word-format, word-reference, word-table-figure, word-check, word-submit]
---

# Word Reviewer

## Overview

Word Reviewer processes reviewer/editor comments, plans a revision strategy for each comment, dispatches changes to appropriate modules, and generates a structured point-by-point response document.

## When to Use

- Processing reviewer comments received via email, letter, or in-document comments
- Planning a systematic revision strategy
- Executing revisions with tracked changes
- Generating a point-by-point response to reviewers

## When NOT to Use

- Self-reviewing manuscript quality → use eco-agent:ecology-review
- General content editing without reviewer context → use word-edit
- Responding to reviewers without Word document changes → handle directly

---

## Phase 1: Extract Comments

### Source A: In-Document Comments

```
1. get_all_comments(file_path)
2. Structure each comment:
   - Comment ID
   - Author
   - Text
   - Location (paragraph number)
   - Is it a reply to another comment?
```

### Source B: External Document/Email

```
1. User provides reviewer letter (paste, file, or screenshot)
2. Parse into structured list:
   - Reviewer ID (Reviewer 1, 2, 3 / Editor)
   - Comment number
   - Comment text
   - Category (major/minor/editorial)
```

### Output: Structured Comment Table

```markdown
| # | Reviewer | Category | Comment Summary | Location |
|---|----------|----------|----------------|----------|
| 1 | R1 | Major | Sample size justification needed | Methods 2.2 |
| 2 | R1 | Minor | Typo in abstract line 3 | Abstract |
| 3 | R2 | Major | Add discussion of limitations | Discussion |
| 4 | R2 | Minor | Update reference [12] | References |
| 5 | Editor | Editorial | Shorten abstract to 250 words | Abstract |
```

---

## Phase 2: Plan Revision Strategy

For each comment, propose a strategy:

```markdown
| # | Comment | Strategy | Planned Action | Module |
|---|---------|----------|---------------|--------|
| 1 | Sample size | ACCEPT | Add power analysis paragraph to 2.2 | word-edit |
| 2 | Typo | ACCEPT | Fix typo | word-edit |
| 3 | Limitations | ACCEPT | Add limitations paragraph to Discussion | word-edit |
| 4 | Update ref | ACCEPT | Update reference entry | word-reference |
| 5 | Shorten abstract | PARTIAL | Reduce to 260 words (250 too restrictive) | word-edit |
```

**Present this table to the user for confirmation.** User may adjust strategies before execution.

---

## Phase 3: Execute Revisions

After user confirms the strategy:

### Step 1: Enable Tracked Changes

All document modifications in this phase use tracked changes mode (XML manipulation):

```bash
# Unpack document for tracked changes editing
python scripts/office/unpack.py document.docx unpacked/
```

### Step 2: Execute Each Change

Process changes in document order (top to bottom) to avoid position shifts:

For each accepted/partially accepted comment:
1. Locate the target text in `unpacked/word/document.xml`
2. Apply the change with `<w:ins>` / `<w:del>` markup
3. Set author="Claude", date=current timestamp

### Step 3: Repack

```bash
python scripts/office/pack.py unpacked/ revised_document.docx --original document.docx
```

### Step 4: Verify

- Read the revised document to confirm changes applied
- Suggest opening in Word to review tracked changes

---

## Phase 4: Generate Point-by-Point Response

Create a response document using word-document-server MCP tools or the docx skill:

```markdown
Response to Reviewers

Manuscript: {title}
Date: {date}

We thank the reviewers for their constructive comments. Below we address
each comment point by point. Reviewer comments are in bold; our responses
follow. Changes in the manuscript are indicated with tracked changes.

---

## Reviewer 1

**Comment 1 (Major): Sample size justification is needed for the
experimental design described in section 2.2.**

Response: We agree with the reviewer. We have added a power analysis
paragraph to Section 2.2 (page X, lines XX-XX) demonstrating that our
sample size of N=30 provides 80% power to detect a medium effect size
(Cohen's d = 0.5) at α = 0.05.

[Change location: Section 2.2, paragraph 3]

---

**Comment 2 (Minor): Typo in abstract, line 3.**

Response: Corrected. Thank you for catching this.

[Change location: Abstract, line 3]

---

## Reviewer 2
...
```

### Response Document Format

- Create as a new Word document
- Bold for reviewer comments
- Regular text for responses
- Reference specific locations in the manuscript
- Include page/line numbers where changes were made

---

## Post-Execution

1. **Output two files:**
   - `{filename}_revised.docx` — manuscript with tracked changes
   - `response_to_reviewers.docx` — point-by-point response

2. **Suggest next steps:**
   - "修订完成。建议运行 word-checker 验证交叉引用和格式。"
   - "确认无误后，使用 word-submit 生成 clean copy 用于重新投稿。"

## Shared Resources

- `../../references/tool_routing.md` — Tool selection priority
- `../../references/token_budget.md` — Token efficiency rules
