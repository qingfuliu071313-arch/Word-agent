# Comment Operations Reference

## Overview

This document defines how word-agent handles Word document comments (批注) across backends. Comments in OOXML are threaded: a top-level comment anchors to a text range, and replies form a chain via `w:comment/@w:parentId`.

## Backend Capabilities

| Operation | Plugin script (canonical) | word-document-server | word-mcp-live (optional) |
|-----------|---------------------------|---------------------|--------------------------|
| Read all comments (with anchors) | `extract_comments.py` | `get_all_comments` ⚠️ | — |
| Read by author | `extract_comments.py --author NAME` | `get_comments_by_author` ⚠️ | — |
| Read for paragraph | `extract_comments.py --paragraph N` | `get_comments_for_paragraph` ✗ | — |
| Reply to comment | `comment_write.py reply --id N --text …` | — | `reply_to_comment` |
| Resolve comment | `comment_write.py resolve --id N` | — | `resolve_comment` |
| Unresolve comment | `comment_write.py unresolve --id N` | — | `unresolve_comment` |
| Delete single comment | `comment_write.py delete --id N` | — | `delete_comment` |
| Delete by author | `comment_write.py delete --author NAME` | — | (loop `delete_comment`) |
| Delete all comments | `comment_write.py delete-all` | — | — |

Both scripts are pure python-docx/lxml — **cross-platform (macOS/Windows/Linux),
no word-mcp-live and no Microsoft Word required.** `word-mcp-live` remains an
optional equivalent for writes if a user has it installed, but it is not a
dependency (and is capability-reduced on macOS, where it drives Word via JXA).

### ⚠️ Why `get_all_comments` is NOT the read path

The external word-document-server tools extract the comment **text** but never
resolve **where the comment is anchored**: every record comes back with
`paragraph_index: null` and `reference_text: ""`. Because the anchor is never
filled in, `get_comments_for_paragraph` always returns an empty list (marked ✗
above). Acting on a comment ("clarify this", "soften that claim") is impossible
without knowing which body text it points at.

`scripts/extract_comments.py` (this plugin) fixes that by walking
`word/document.xml` for `w:commentRangeStart`/`w:commentRangeEnd` and joining the
spanned runs back to each comment id. It returns, per comment: `text`, `author`,
`initials`, `date`, `reference_text` (anchored body text), `paragraph_index`,
`in_table`, `resolved` (from `commentsExtended.xml`), and `parent_id`/`is_reply`
threading. **All comment reads MUST use this script.** The MCP tools are a
text-only fallback when python-docx is unavailable.

**Routing priority:**
- **Read** via `scripts/extract_comments.py` (only path that resolves anchors).
  `get_all_comments` is a last-resort text-only fallback.
- **Write** (reply/resolve/unresolve/delete) via `scripts/comment_write.py`
  (pure XML, cross-platform, no dependencies). `word-mcp-live` is an optional
  equivalent if installed.

### Comment id stability

The `comment_id` used by `comment_write.py` is the OOXML `w:id` — the same id
`extract_comments.py` returns. Always read first, then act on the ids you got.

## Comment Model

```xml
<!-- word/comments.xml -->
<w:comment w:id="1" w:author="Reviewer 1" w:date="2026-01-15T10:00:00Z">
  <w:p><w:r><w:t>Please clarify the methodology.</w:t></w:r></w:p>
</w:comment>

<!-- Reply (threaded). Each comment paragraph carries a w14:paraId; the
     parent/child link lives in commentsExtended.xml (see below), NOT here. -->
<w:comment w:id="2" w:author="Author" w:date="2026-01-16T09:00:00Z">
  <w:p w14:paraId="55555555"><w:r><w:t>Added clarification in Section 2.2.</w:t></w:r></w:p>
</w:comment>
```

```xml
<!-- word/commentsExtended.xml — threading + resolved status -->
<w15:commentEx w15:paraId="44444444" w15:done="1"/>                          <!-- parent, resolved -->
<w15:commentEx w15:paraId="55555555" w15:paraIdParent="44444444" w15:done="0"/> <!-- reply -->
```

```xml
<!-- word/document.xml — anchor range -->
<w:commentRangeStart w:id="1"/>
<w:r><w:t>target text</w:t></w:r>
<w:commentRangeEnd w:id="1"/>
<w:r><w:commentReference w:id="1"/></w:r>
```

## Resolve/Unresolve Semantics

- **Resolve** marks a comment thread as addressed (`w16:done="1"` in commentsExtended.xml)
- Resolved comments remain visible in Word but appear dimmed
- **Unresolve** reactivates a resolved comment thread
- Only the top-level comment in a thread can be resolved/unresolved; replies inherit the state

## Workflows

### word-review: Comment-Based Reply Workflow (Phase 4.5)

After executing revisions with tracked changes, reply to in-document comments:

```bash
For each comment with ACCEPT/PARTIAL strategy:
  1. Execute the document change (Phase 3)
  2. python3 scripts/comment_write.py {file} reply --id {comment_id} --author "Author" \
       --text "已按建议修改。见修订标记。"      # PARTIAL: "已部分采纳。[替代方案]。见修订标记。"
  3. python3 scripts/comment_write.py {file} resolve --id {comment_id}

For each comment with REBUT/DEFER strategy:
  1. python3 scripts/comment_write.py {file} reply --id {comment_id} --author "Author" \
       --text "感谢建议。……详见 response to reviewers。"
       # REBUT: "……我们认为原文表述更准确，理由如下：[理由]。……"
       # DEFER: "……此问题将在后续研究中解决。……"
  2. Do NOT resolve — user decides whether to resolve after reviewing
```

### word-submit: Selective Comment Management (Feature A Step 3)

```bash
Step 3: Remove comments (选择性批注管理) — all via comment_write.py (strips
        range/reference markers too, so no orphans remain)
    Option A: 删除全部批注
      → python3 scripts/comment_write.py {file} delete-all
    Option B: 仅删除已 resolved 批注
      → extract_comments.py {file}  → for each resolved==true:
        python3 scripts/comment_write.py {file} delete --id {comment_id}
    Option C: 按作者删除 (deleting a parent also removes its replies)
      → python3 scripts/comment_write.py {file} delete --author "{name}"
```

## Font Normalization

Comment write operations do NOT touch document body fonts. Font normalization is
still MANDATORY after any body-text modification but is NOT required after
comment-only operations (reply/resolve/delete).

## Implementation notes (comment_write.py)

- **Threading** is written via `commentsExtended.xml` (`w15:paraIdParent` →
  parent's `w14:paraId`), the model modern Word uses — not the legacy
  `w:comment/@w:parentId`. The script auto-creates `commentsExtended.xml` and
  registers its content-type + relationship when first needed, and generates a
  `w14:paraId` for any comment paragraph that lacks one.
- **Reply** also inserts a `<w:commentReference>` run right after the parent's in
  `document.xml`, so Word renders the reply in the same anchored thread.
- **Delete** removes the `<w:comment>` plus its `commentRangeStart/End/Reference`
  markers and its `commentsExtended` entry; deleting a parent cascades to its
  replies. No orphaned references are left (verified).
- A timestamped `*_backup_*.docx` is written before the first edit unless
  `--no-backup` is passed.
