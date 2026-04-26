---
name: word-submit
description: >-
  Submission preparation module for academic paper Word documents. Generates
  clean copies, splits documents into main text and supplementary materials,
  merges multiple documents with format unification, and runs pre-submission
  quality checks.
  Triggers: 投稿, 清理, clean copy, submit, 删除批注, 接受修订, 拆分,
  补充材料, 合并, merge, split, 投稿准备.
allowed-tools: Read Write Edit Bash Glob Grep mcp__word-document-server__get_all_comments mcp__word-document-server__get_document_info mcp__word-document-server__copy_document mcp__word-document-server__protect_document mcp__word-document-server__unprotect_document mcp__word-document-server__get_document_text mcp__word-document-server__get_document_outline mcp__word-document-server__get_paragraph_text_from_document mcp__word-document-server__find_text_in_document mcp__word-document-server__get_document_xml mcp__word-document-server__delete_paragraph
metadata:
    version: "0.1.0"
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

Step 2: Accept all tracked changes
  → Bash: python scripts/accept_changes.py {copy_path} {clean_path}
  → This uses the docx skill's accept_changes.py script

Step 3: Remove all comments
  → Via XML manipulation:
     a. Unpack: python scripts/office/unpack.py {clean_path} unpacked/
     b. Remove all <w:comment> elements from word/comments.xml
     c. Remove all <w:commentRangeStart>, <w:commentRangeEnd>,
        <w:commentReference> from word/document.xml
     d. Repack: python scripts/office/pack.py unpacked/ {clean_path}

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

- Before Step 2: "即将接受所有修订。确认？"（不可逆操作）
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

Step 2: Semi-automated checks
  → Word count per section (some journals have limits)
  → Figure count and resolution notes
  → Table count

Step 3: Manual check reminders
  → Display checklist items that require human verification
  → User checks off each item
```

---

## Post-Execution

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
