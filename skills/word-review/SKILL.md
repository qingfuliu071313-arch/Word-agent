---
name: word-review
description: >-
  Review and revision module for academic paper Word documents. Extracts
  reviewer comments, plans revision strategies, dispatches changes to
  specialist modules in tracked-changes mode, and generates point-by-point
  response documents.
  Triggers: 审稿意见, 修订, 审阅, 批注, reviewer comments, revision,
  response to reviewers, point-by-point.
allowed-tools: Read Write Edit Bash Glob Grep mcp__word-document-server__get_all_comments mcp__word-document-server__get_comments_by_author mcp__word-document-server__get_comments_for_paragraph mcp__word-document-server__get_document_text mcp__word-document-server__get_document_outline mcp__word-document-server__get_paragraph_text_from_document mcp__word-document-server__find_text_in_document mcp__word-document-server__search_and_replace mcp__word-document-server__get_document_xml mcp__word-mcp-live__edit_text mcp__word-mcp-live__insert_text mcp__word-mcp-live__delete_text mcp__word-mcp-live__replace_text mcp__word-mcp-live__add_comment mcp__word-mcp-live__reply_to_comment mcp__word-mcp-live__resolve_comment mcp__word-mcp-live__delete_comment mcp__word-document-server__create_document mcp__word-document-server__add_paragraph mcp__word-document-server__add_heading
metadata:
    version: "1.1.0"
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

- Self-reviewing manuscript quality → 使用通用学术审稿工具或人工评审（本插件不提供稿件质量自审）
- General content editing without reviewer context → use word-edit
- Responding to reviewers without Word document changes → handle directly

---

## Phase 1: Extract Comments

### Source A: In-Document Comments

**Canonical path — `scripts/extract_comments.py` (REQUIRED; do NOT use the MCP tool to locate comments):**

```bash
python3 scripts/extract_comments.py "{file_path}"             # JSON
python3 scripts/extract_comments.py "{file_path}" --markdown  # table
# filters: --author "Reviewer 1"   --paragraph N
```

⚠️ The external `get_all_comments` tool returns the comment *text* but always
reports `paragraph_index: null` and `reference_text: ""` — it never resolves
which body text a comment is anchored to, so you cannot act on it.
`get_comments_for_paragraph` is consequently always empty. **Always read
in-document comments with `extract_comments.py`**, which resolves the anchor.

Each comment record contains:
   - `comment_id`, `author`, `initials`, `date`
   - `text` — the comment content
   - `reference_text` — **the exact body text the comment is anchored to**;
     feed this to `find_text_in_document` / `search_and_replace` to locate the edit target
   - `paragraph_index` — body paragraph index (`null` when `in_table` is true)
   - `in_table` — anchor is inside a table cell
   - `resolved` — thread already marked done
   - `parent_id` / `is_reply` — reply threading

Fallback only if the script errors (e.g. python-docx unavailable):
`get_all_comments(file_path)` for text/author, then locate anchors manually via
`get_document_xml` (parse `w:commentRangeStart`/`w:commentRangeEnd` by `w:id`).

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

For in-document comments, include the **anchored text** (from `reference_text`)
so every change has a precise edit target:

```markdown
| # | Reviewer | Category | Comment Summary | Anchored Text (¶/loc) |
|---|----------|----------|----------------|-----------------------|
| 1 | R1 | Major | Sample size justification needed | "small pilot sample of 12…" (¶24) |
| 2 | R1 | Minor | Typo in abstract line 3 | "the the effect" (¶2) |
| 3 | R2 | Major | Add discussion of limitations | "We conclude the intervention…" (¶57) |
| 4 | R2 | Minor | Update reference [12] | "[12]" (¶80) |
| 5 | Editor | Editorial | Shorten abstract to 250 words | (whole Abstract) |
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

### Step 1: Execute Changes with Native Tracked Changes (Preferred)

Uses word-mcp-live with `track_changes: true`. Process changes in document order (top to bottom) to avoid position shifts.

For each accepted/partially accepted comment:

```
1. Locate target text:
   → Use Document Map or find_text_in_document(file_path, target_text)

2. Apply change with native tracked changes:
   - For text replacement:
     mcp__word-mcp-live__replace_text(file_path, old_text, new_text, track_changes=true)
   - For new paragraph insertion:
     mcp__word-mcp-live__insert_text(file_path, anchor_text, new_text, position="after", track_changes=true)
   - For text deletion:
     mcp__word-mcp-live__delete_text(file_path, target_text, track_changes=true)
   - For paragraph rewrite:
     mcp__word-mcp-live__edit_text(file_path, old_text, new_text, track_changes=true)

3. Author is set via MCP_AUTHOR env var (default: "Claude")
```

### Step 1 (Legacy Fallback): XML Tracked Changes

If word-mcp-live is unavailable, fall back to XML manipulation:

```bash
# Unpack document for tracked changes editing
python3 scripts/office/unpack.py document.docx unpacked/
```

For each accepted/partially accepted comment:
1. Locate the target text in `unpacked/word/document.xml`
2. Apply the change with `<w:ins>` / `<w:del>` markup
3. Set author="Claude", date=current timestamp

```bash
python3 scripts/office/pack.py unpacked/ revised_document.docx --original document.docx
```

See `../../references/tracked_changes.md` for XML patterns and native mode parameter reference.

### Alternative: Bulk Edit via adeu (10+ revisions)

When Phase 2 produces 10 or more ACCEPT/PARTIAL revision items, consider using adeu for batch processing instead of individual MCP calls:

```
1. Extract document to Markdown via adeu
2. Generate CriticMarkup for all accepted changes:
   {~~old text~>new text~~} for replacements
   {++new paragraph text++} for insertions
   {--deleted text--} for deletions
3. Validate all changes atomically (adeu catches ambiguous matches)
4. Apply all changes in one pass with tracked changes
5. Font normalization gate
```

**Trigger:** Suggest to user when revision count >= 10: "有 {n} 处修改。建议使用批量模式一次性应用？"

**Fallback:** If adeu is not installed or validation fails, fall back to individual MCP calls.

See `../../references/adeu_integration.md` for pipeline details.

### Step 2: Verify

- Read the revised document to confirm changes applied
- Suggest opening in Word to review tracked changes

---

## Phase 3.5: Comment-Based Reply Workflow

After executing revisions, reply to in-document comments and mark them resolved.
This creates a threaded conversation visible in Word's comment pane.

**Canonical path — `scripts/comment_write.py` (pure XML, cross-platform, no
word-mcp-live / no Word needed).** Use the `comment_id` from `extract_comments.py`
(Phase 1). Process comments bottom-up is unnecessary here — comment ops don't
shift text positions.

### For ACCEPT/PARTIAL comments:

```bash
# 1. (Phase 3 already executed the document change)
# 2. Reply (ACCEPT: "已按建议修改。见修订标记。"  PARTIAL: "已部分采纳。[替代方案]。见修订标记。")
python3 scripts/comment_write.py "{file_path}" reply --id {comment_id} \
    --text "已按建议修改。见修订标记。" --author "Author"
# 3. Mark resolved
python3 scripts/comment_write.py "{file_path}" resolve --id {comment_id}
```

### For REBUT/DEFER comments:

```bash
# Reply with rationale, do NOT resolve (leave for user to decide):
#   REBUT: "感谢建议。我们认为原文表述更准确，理由如下：[理由]。详见 response to reviewers。"
#   DEFER: "感谢建议。此问题将在后续研究中解决。详见 response to reviewers。"
python3 scripts/comment_write.py "{file_path}" reply --id {comment_id} \
    --text "感谢建议。……详见 response to reviewers。" --author "Author"
```

The script writes a timestamped `*_backup_*.docx` before its first edit (pass
`--no-backup` to suppress). Each op edits in place by default; pass `--output PATH`
to write a copy.

### Optional alternate path

If `word-mcp-live` is installed and configured, its native `reply_to_comment` /
`resolve_comment` are equivalent. But it is NOT required and is often absent
(and capability-reduced on macOS), so **default to `comment_write.py`**. If
neither is available (e.g. python-docx missing), skip Phase 3.5 — replies will
still appear in the point-by-point response document (Phase 4).

See `../../references/comment_operations.md` for the full comment model and operations reference.

---

## Phase 4: Generate Point-by-Point Response

Create a response document using word-document-server MCP tools (`create_document` → `add_heading` / `add_paragraph`, all declared in allowed-tools) or the docx skill:

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

1. **MANDATORY: Font Normalization** — 对带修订标记的修订稿必须使用 `--skip-revisions`（跳过 `w:ins`/`w:del` 内的 run，避免静默改动污染修订记录和删除快照）：
   ```bash
   # 修订稿（含 tracked changes）：
   python3 scripts/normalize_fonts.py "{filename}_revised.docx" --unify --skip-revisions
   # response 文档（无修订，正常归一化）：
   python3 scripts/normalize_fonts.py "response_to_reviewers.docx" --unify
   ```
   See `../../references/font_normalization.md`.

2. **Output two files:**
   - `{filename}_revised.docx` — manuscript with tracked changes
   - `response_to_reviewers.docx` — point-by-point response

3. **Suggest next steps:**
   - "修订完成。建议运行 word-checker 验证交叉引用和格式。"
   - "确认无误后，使用 word-submit 生成 clean copy 用于重新投稿。"

## Shared Resources

- `../../references/tool_routing.md` — Tool selection priority
- `../../references/token_budget.md` — Token efficiency rules
- `../../references/tracked_changes.md` — Native vs Legacy tracked changes reference
- `../../references/comment_operations.md` — Threaded comment model and operations
