# Comment Operations Reference

## Overview

This document defines how word-agent handles Word document comments (批注) across backends. Comments in OOXML are threaded: a top-level comment anchors to a text range, and replies form a chain via `w:comment/@w:parentId`.

## Backend Capabilities

| Operation | word-document-server | word-mcp-live | XML Manual |
|-----------|---------------------|---------------|------------|
| Read all comments | `get_all_comments` | — | Parse `word/comments.xml` |
| Read by author | `get_comments_by_author` | — | Filter XML |
| Read for paragraph | `get_comments_for_paragraph` | — | Match `commentRangeStart` |
| Add comment | — | `add_comment` | Insert XML elements |
| Reply to comment | — | `reply_to_comment` | Insert with `parentId` |
| Resolve comment | — | `resolve_comment` | Set `w16:done="1"` |
| Delete single comment | — | `delete_comment` | Remove XML elements |
| Delete all comments | — | — | Bulk XML removal (word-submit) |

**Routing priority:** Read via word-document-server (richer parsing) → Write via word-mcp-live (native operations) → XML fallback (bulk operations only).

## Comment Model

```xml
<!-- word/comments.xml -->
<w:comment w:id="1" w:author="Reviewer 1" w:date="2026-01-15T10:00:00Z">
  <w:p><w:r><w:t>Please clarify the methodology.</w:t></w:r></w:p>
</w:comment>

<!-- Reply (threaded) -->
<w:comment w:id="2" w:author="Author" w:date="2026-01-16T09:00:00Z" w:parentId="1">
  <w:p><w:r><w:t>Added clarification in Section 2.2.</w:t></w:r></w:p>
</w:comment>
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

```
For each comment with ACCEPT/PARTIAL strategy:
  1. Execute the document change (Phase 3)
  2. Reply to comment: mcp__word-mcp-live__reply_to_comment(file_path, comment_id, reply_text)
     - reply_text examples:
       ACCEPT:  "已按建议修改。见修订标记。"
       PARTIAL: "已部分采纳。[说明替代方案]。见修订标记。"
  3. Resolve comment: mcp__word-mcp-live__resolve_comment(file_path, comment_id)

For each comment with REBUT/DEFER strategy:
  1. Reply to comment with rationale (do NOT resolve — leave for user decision):
     REBUT: "感谢建议。我们认为原文表述更准确，理由如下：[理由]。详见 response to reviewers。"
     DEFER: "感谢建议。此问题将在后续研究中解决。详见 response to reviewers。"
  2. Do NOT resolve — user decides whether to resolve after reviewing
```

### word-submit: Selective Comment Management (Feature A Step 3)

```
Step 3: Remove comments (选择性批注管理)
  → User chooses:
    Option A: 删除全部批注 (legacy behavior)
      → XML bulk removal of all <w:comment>, <w:commentRangeStart/End>, <w:commentReference>
    Option B: 仅删除已 resolved 批注
      → mcp__word-mcp-live__delete_comment for each resolved comment
      → Keep unresolved comments (may need author review)
    Option C: 按作者删除
      → get_comments_by_author → filter → delete_comment for matching comments
      → E.g., delete all "Claude" comments, keep reviewer comments
```

## Font Normalization

Comment operations via word-mcp-live do NOT affect document body fonts. Font normalization is still MANDATORY after any body text modifications but is not required after comment-only operations (add/reply/resolve/delete).

## Fallback: XML Bulk Comment Removal

When word-mcp-live is unavailable, use the existing XML approach for bulk removal:

```bash
python scripts/office/unpack.py {file_path} unpacked/
# Remove all <w:comment> elements from word/comments.xml
# Remove all <w:commentRangeStart>, <w:commentRangeEnd>, <w:commentReference> from word/document.xml
python scripts/office/pack.py unpacked/ {output_path}
```

This is the only comment operation available without word-mcp-live. Individual add/reply/resolve require word-mcp-live.
