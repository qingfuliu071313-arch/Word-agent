---
name: word-orchestrate
description: >-
  Central coordinator for all Word document operations. Routes user requests
  to specialist modules, manages Token budget, and orchestrates multi-module
  workflows. Activates as the default entry point when a Word/docx task
  doesn't clearly map to a single specialist, or when multiple operations
  are needed in sequence.
  Triggers: any ambiguous Word request, multi-step operations, "排版并检查",
  "full preparation", "全流程".
allowed-tools: Read Bash Glob Grep mcp__word-document-server__get_document_info mcp__word-document-server__get_document_outline mcp__word-document-server__list_available_documents mcp__word-mcp-live__get_active_document
metadata:
    version: "1.1.0"
    category: coordination
    downstream-skills: [word-read, word-format, word-edit, word-reference, word-table-figure, word-review, word-check, word-submit]
---

# Word Orchestrator

## Overview

The orchestrator is the entry point and coordinator for all word-agent operations. It does not perform document operations itself — it analyzes user intent, selects the right module(s), ensures Document Map availability, and enforces Token budget rules.

## Phase 1: Intent Analysis

1. **Parse user request** — Identify what the user wants to do with which document(s)
2. **Check document availability** — Use `list_available_documents` or confirm file path
3. **Classify task type:**
   - Single-module task → route directly
   - Multi-module pipeline → plan execution order
   - Ambiguous → ask user for clarification

### Routing Rules

```
IF request mentions "对比/compare/diff"
  → word-read (compare mode)

ELIF request mentions "排版/格式/format" AND "检查/check/verify"
  → Pipeline: word-read → word-format → word-check

ELIF request mentions "排版/格式/format"
  → Pipeline: word-read → word-format

ELIF request mentions "修改内容/edit/改写"
  → Pipeline: word-read → word-edit

ELIF request mentions "投稿/submit/清理/clean"
  → Pipeline: word-read → word-check → word-submit

ELIF request mentions "审稿/review/revision"
  → Pipeline: word-read → word-review

ELIF request clearly maps to one module
  → Route to that module directly

ELSE
  → Ask user: "您具体需要对这个文档做什么操作？"
```

## Phase 2: Pre-flight Check

Before routing to any module:

1. **File format check** — Is the input file `.doc` (legacy binary format)?
   - YES → run .doc→.docx conversion pipeline:
     ```
     a. Verify LibreOffice available: `which soffice`
        - Not found → tell user: "需要 LibreOffice 来转换 .doc 文件。请安装后重试。"
     b. Convert: `soffice --headless --convert-to docx --outdir "{dir}" "{file.doc}"`
        - If same-name .docx exists → save as `{name}-converted.docx`
     c. Post-fix ALL conversion artifacts (MANDATORY):
        python3 scripts/fix_libreoffice.py "{converted.docx}"
        This single command fixes: compatibilityMode, Liberation fonts,
        semicolon fallback fonts, style chain, tables, spacing, page setup,
        AND runs normalize_fonts.py --unify automatically.
     d. Report to user: "已将 {name}.doc 转换为 .docx 并修复格式兼容性问题。后续操作基于转换后的文件。"
     e. Update file_path to point to the new .docx
     ```
     See `references/doc_conversion.md` for details.
   - NO (already .docx) → continue

2. **Document Map check** — 先查落盘的 map 文件（不依赖对话上下文）:
   ```
   a. Glob: {文档所在目录}/.word-agent/{文档名去扩展名}.map.md
   b. 文件存在 → 比较 mtime（Bash 测试 map 是否比 docx 新）:
      if [ "{dir}/.word-agent/{name}.map.md" -nt "{dir}/{name}.docx" ]; then echo VALID; else echo STALE; fi
      - VALID（map 的 mtime 晚于 docx 的 mtime）→ map 有效：
        Read 该文件，并将其内容注入目标模块的 prompt
      - STALE（docx 在 map 之后被修改过）→ map 已过期 → 派 word-read 重新生成
   c. 文件不存在 → 派 word-read 先生成（除非任务本身就是 word-read）
   ```
   word-read 生成的 map 总是写入上述路径（见 word-read SKILL Step 7），跨会话可复用。

3. **Format Spec check** — Does the task require formatting rules?
   - YES → check if user has provided format requirements
   - NO format requirements yet → ask user: "请提供格式要求文档（Word/PDF/文字描述均可）"
   - If format requirements are a Word document → remind downstream modules to extract text box content (format templates often place formatting instructions inside text boxes, which standard text extraction tools skip entirely)

4. **Token budget assessment** — Estimate operation scale:
   - Small (< 5 changes) → proceed directly
   - Medium (5-20 changes) → show plan, ask confirmation
   - Large (20+ changes) → show plan with estimated scope, ask confirmation

5. **Editing mode detection** — Is Word open with the target document?
   - Check via `mcp__word-mcp-live__get_active_document()` (if word-mcp-live available)
   - **Word open with target doc** → set mode = "live"
     - Inform downstream: use word-mcp-live for all write operations
     - Font normalization will be deferred (file locked by Word)
     - Undo support available via word-mcp-live
     - Do NOT use word-document-server write tools on the same file
   - **Word not open** → set mode = "file" (default)
     - Use standard tool routing (word-document-server primary)
     - Font normalization runs immediately after operations
   - **word-mcp-live unavailable** → set mode = "file" (no detection possible)
   
   See `references/live_editing.md` for platform support and mode rules.

## Phase 3: Execution

### Single Module
Route to the module with:
- Document Map (if available)
- Format Spec (if applicable)
- User's specific instructions

### Pipeline (Multi-Module)
Execute modules in order:
1. Run module A → collect output
2. Pass output + Document Map to module B
3. Continue chain
4. Report final results

### Error Handling
- If a module reports failure → check error type
- If recoverable → suggest alternative approach
- If not recoverable → report to user with explanation

## Phase 4: Post-Execution

### MANDATORY: Font Normalization Gate

**Before returning ANY modified document to the user, the orchestrator MUST run font normalization.** This is a non-negotiable final step that applies to ALL document-modifying pipelines (word-format, word-edit, word-table-figure, word-reference, word-review). It is NOT optional, NOT advisory, and MUST NOT be skipped.

#### File Mode (default)

```bash
python3 scripts/normalize_fonts.py "{file_path}" --unify --cn "{cn_font}" --en "{en_font}"
```

Default fonts: `--cn 宋体 --en "Times New Roman"` unless the user specified otherwise.

If the user provided custom font requirements (e.g., "正文用楷体"), use those instead:
```bash
python3 scripts/normalize_fonts.py "{file_path}" --unify --cn 楷体 --en "Times New Roman"
```

#### Live Mode (Word is open)

```bash
python3 scripts/normalize_fonts.py "{file_path}" --check-lock --unify --cn "{cn_font}" --en "{en_font}"
```

- If exit code 0: normalization succeeded (file was unlocked momentarily)
- If exit code 2: file is locked by Word → **defer normalization**
  - Warn user: "字体归一化已延迟——Word 正在使用此文件。请关闭 Word 后运行，或在下次操作时自动执行。"
  - Track as pending: `pending_normalization = true` for this file path
  - On next interaction with same file: check lock again and run if possible

This gate fixes:
- Theme font references in styles.xml (the #1 cause of font chaos)
- Bare runs with no font attributes (inherit wrong fonts from theme)
- Missing eastAsia attributes from MCP tool writes
- Mixed fonts within paragraphs from multiple write operations

**Do NOT report this step to the user unless issues were found** (file mode) **or normalization was deferred** (live mode).

### Follow-up

1. **Suggest next step** — Based on workflow table in `docs/CLAUDE.md`
2. **Report Token usage** — If the operation was large, note what was saved by caching

## What This Skill Does NOT Do

- Does not read, edit, or format documents directly
- Does not generate Document Maps (that's word-read)
- Does not make assumptions about formatting rules
