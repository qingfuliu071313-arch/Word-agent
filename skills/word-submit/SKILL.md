---
name: word-submit
description: >-
  Submission preparation module for academic paper Word documents. Generates
  clean copies, splits documents into main text and supplementary materials,
  merges multiple documents with format unification, and runs pre-submission
  quality checks.
  Triggers: 投稿, 清理, clean copy, submit, 删除批注, 接受修订, 拆分,
  补充材料, 合并, merge, split, 投稿准备.
allowed-tools: Read Write Edit Bash Glob Grep mcp__word-document-server__get_all_comments mcp__word-document-server__get_document_info mcp__word-document-server__copy_document mcp__word-document-server__protect_document mcp__word-document-server__unprotect_document mcp__word-document-server__get_document_text mcp__word-document-server__get_document_outline mcp__word-document-server__get_paragraph_text_from_document mcp__word-document-server__find_text_in_document mcp__word-document-server__get_document_xml mcp__word-document-server__delete_paragraph mcp__docx-mcp__accept_changes mcp__docx-mcp__reject_changes mcp__docx-mcp__get_tracked_changes mcp__word-mcp-live__delete_comment mcp__word-mcp-live__resolve_comment mcp__word-mcp-live__add_watermark mcp__docx-mcp__remove_watermark mcp__docx-mcp__scrub_pii mcp__docx-mcp__validate_paraids mcp__docx-mcp__audit_document
metadata:
    version: "1.1.0"
    category: submission
    upstream-skills: [word-read, word-check]
    downstream-skills: []
---

# Word Submit

## Overview

Word Submit is the final module in the word-agent pipeline. It prepares documents for journal submission by generating clean copies, splitting/merging documents, and running final quality checks.

## When to Use

- Generating a clean copy (no tracked changes, no comments, no metadata)
- Splitting a document into main text + supplementary materials
- Merging multiple author contributions into one document
- Final preparation before journal submission

## When NOT to Use

- Making content edits → use word-edit
- Formatting document → use word-format
- Handling reviewer comments → use word-review

---

## Feature A: Clean Copy Generation

### Pipeline

```
Step 1: Copy document
  → copy_document(file_path, output_path)
  → Work on the copy, never the original

Step 1.5: Review tracked changes (selective accept/reject)
  → mcp__docx-mcp__get_tracked_changes(copy_path)
  → Present revision summary to user:
    "文档包含 {n} 处修订（{ins} 处插入, {del} 处删除），作者: {authors}"
  → User chooses:
    Option A: 接受全部 → proceed to Step 2
    Option B: 逐条审查 → list each change, user marks accept/reject
      → mcp__docx-mcp__accept_changes(copy_path, ranges=[...accepted...])
      → mcp__docx-mcp__reject_changes(copy_path, ranges=[...rejected...])
      → Skip Step 2 (already handled)
    Option C: 按作者筛选 → accept all from author X, reject from author Y
  → If docx-mcp unavailable → skip to Step 2 (accept all, legacy path)

Step 2: Accept all tracked changes (if not handled by Step 1.5 Option B/C)
  → Bash: python3 scripts/accept_changes.py {copy_path} {clean_path}
  → accept_changes.py 是本插件 scripts/ 内的 XML fallback 脚本（接受全部修订）

Step 3: Remove comments (选择性批注管理)
  → User chooses:
    Option A: 删除全部批注 (default)
      → Via XML manipulation:
         a. Unpack: python3 scripts/office/unpack.py {clean_path} unpacked/
         b. Remove all <w:comment> elements from word/comments.xml
         c. Remove all <w:commentRangeStart>, <w:commentRangeEnd>,
            <w:commentReference> from word/document.xml
         d. Repack: python3 scripts/office/pack.py unpacked/ {clean_path}
    Option B: 仅删除已 resolved 批注
      → get_all_comments → filter resolved → delete each:
         mcp__word-mcp-live__delete_comment(clean_path, comment_id)
      → Keep unresolved comments for author review
    Option C: 按作者删除
      → get_comments_by_author → user selects authors to remove
      → mcp__word-mcp-live__delete_comment for matching comments
      → E.g., delete all "Claude" auto-replies, keep reviewer comments

Step 4: Strip personal metadata
  → Via XML manipulation on unpacked/docProps/core.xml:
     - Clear <dc:creator> (author)
     - Clear <cp:lastModifiedBy>
     - Clear <cp:revision> (revision count)
     - Optionally clear <dcterms:created> and <dcterms:modified>

Step 5: Final check
  → Run word-checker (both modes) on the clean copy
  → Report any remaining issues

Step 6: Output
  → {filename}_clean.docx — submission-ready document
```

### User Confirmation Points

- At Step 1.5: "文档包含 {n} 处修订。选择处理方式：A) 全部接受 B) 逐条审查 C) 按作者筛选"
- Before Step 2 (if applicable): "即将接受所有修订。确认？"（不可逆操作）
- Before Step 4: "即将清除作者信息和修改历史。确认？"
- After Step 5: Show check results, ask if ready to finalize

---

## Feature B: Document Splitting

### Use Cases

1. **Main text + Supplementary materials**
2. **Per-section split** (e.g., each chapter as a separate file)
3. **Extract figures/tables** into separate files

### Process

```
Step 1: Read Document Map
  → Identify split points (e.g., "补充材料" / "Supplementary" heading)

Step 2: Confirm split plan with user
  → "将在以下位置拆分："
  → "文件1: P1-P165 (正文)"
  → "文件2: P166-P200 (补充材料)"

Step 3: Execute split
  → Via Python/docx manipulation:
     a. Read original document
     b. Create two new documents
     c. Copy paragraphs to respective documents
     d. Preserve formatting and styles

Step 4: Verify
  → Check both output files for completeness
  → Report: "{filename}_main.docx (165段), {filename}_supplementary.docx (35段)"
```

---

## Feature C: Document Merging

### Use Cases

1. **Multiple co-author contributions** → merge into one document
2. **Separate sections** → combine into complete manuscript

### Process

```
Step 1: List input documents
  → User provides file paths for all documents to merge

Step 2: Analyze each document
  → Run word-reader on each to generate Document Maps
  → Identify potential style conflicts

Step 3: Plan merge order
  → Present to user: "合并顺序：doc1 (引言) → doc2 (方法) → doc3 (结果讨论)"

Step 4: Execute merge using docxcompose
  → Via Bash:
     python3 -c "
     from docxcompose.composer import Composer
     from docx import Document

     master = Document('{first_doc}')
     composer = Composer(master)

     for path in ['{second_doc}', '{third_doc}']:
         doc = Document(path)
         composer.append(doc)

     composer.save('{output_path}')
     "

Step 5: Post-merge formatting
  → Suggest running word-format to unify styles
  → Suggest running word-checker to verify cross-references

Step 6: Report
  → "已合并 3 个文档为 {output_path}"
  → "建议运行 word-format 统一格式"
```

---

## Feature D: Pre-Submission Checklist

Run through the checklist in `../../references/submission_checklist.md`:

```
Step 1: Automated checks
  → Run word-checker (cross-references + format compliance)
  → Verify no tracked changes remain
  → Verify no comments remain
  → Check file size (some journals have limits)

Step 1b: Structural integrity verification
  → mcp__docx-mcp__validate_paraids(file_path)
    → Must pass: zero duplicate/missing paraIds
  → mcp__docx-mcp__audit_document(file_path)
    → Must have zero errors (warnings OK with user confirmation)
  → If docx-mcp unavailable: skip with warning
  → If fails: "文档结构存在问题，建议修复后再投稿。详见检查报告。"

Step 2: Semi-automated checks
  → Word count per section (some journals have limits)
  → Figure count and resolution notes
  → Table count

Step 3: Manual check reminders
  → Display checklist items that require human verification
  → User checks off each item
```

---

## Feature E: Watermark Management

### Add Watermark

```
Step 1: User specifies watermark text (e.g., "DRAFT", "CONFIDENTIAL", "审稿版")
Step 2: mcp__word-mcp-live__add_watermark(file_path, text="{watermark_text}")
  → Options: font, color, size, rotation, transparency
Step 3: Verify watermark visible in document
```

### Remove Watermark (for clean copy)

```
Step 1: mcp__docx-mcp__remove_watermark(file_path)
Step 2: Verify watermark removed
```

Watermark removal is automatically included in Feature A (Clean Copy Generation) after Step 4 (strip metadata) as an optional step:
- "是否需要移除水印？" → If yes, remove before final check

### Fallback

If neither word-mcp-live nor docx-mcp is available, watermarks can be added/removed via XML manipulation of the header's `<v:shape>` element in `word/header*.xml`.

---

## Feature F: PII Scrubbing (Experimental)

Scans and removes Personally Identifiable Information from the document using docx-mcp's spaCy-based NER.

```
Step 1: Scan for PII
  → mcp__docx-mcp__scrub_pii(file_path, mode="scan")
  → Returns: list of detected PII (names, emails, phone numbers, addresses)

Step 2: Present findings to user
  → "检测到以下个人信息：{PII_list}"
  → "确认需要清除哪些？"

Step 3: Apply scrubbing (after user confirmation)
  → mcp__docx-mcp__scrub_pii(file_path, mode="redact", items=[...confirmed...])
  → Replaces PII with placeholders (e.g., "[AUTHOR]", "[EMAIL]")

Step 4: Verify scrubbing
  → Re-scan to confirm no PII remains
```

### Prerequisites

- docx-mcp must be available
- Optional: `python -m spacy download en_core_web_lg` (500MB+) for enhanced NER accuracy
- Without spaCy model, falls back to regex-based pattern matching (less accurate)

### Caution

- PII scrubbing is EXPERIMENTAL — always verify results manually
- Some academic content (author names in citations, acknowledgments) may be false positives
- Never auto-scrub without user confirmation
- This feature is for anonymized review submissions, not standard journal submissions

---

## Post-Execution

**MANDATORY: Font Normalization** — Run `python3 scripts/normalize_fonts.py "{file_path}" --unify` on all output documents before returning. See `../../references/font_normalization.md`.

Report final status:

```
投稿准备完成：

📄 输出文件:
  - paper_clean.docx (投稿版，无修订无批注)
  - paper_supplementary.docx (补充材料)

✅ 自动检查结果:
  - 修订痕迹: 已全部接受 ✓
  - 批注: 已全部删除 ✓
  - 个人元数据: 已清除 ✓
  - 交叉引用: 全部正确 ✓
  - 格式合规: 20/20 项通过 ✓

📋 请手动确认:
  - [ ] 作者信息和通讯作者正确
  - [ ] 致谢内容完整
  - [ ] 利益冲突声明已包含
  - [ ] 数据可用性声明已包含
  - [ ] Cover letter 已准备
```

## Shared Resources

- `../../references/submission_checklist.md` — Pre-submission checklist
- `../../references/tool_routing.md` — Tool selection priority
- `../../references/comment_operations.md` — Selective comment management operations
- `../../references/structural_validation.md` — OOXML structural integrity checks
- `../../references/equation_crossref.md` — Equation and cross-reference reference (for watermark XML context)
