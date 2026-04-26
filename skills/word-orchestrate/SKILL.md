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
allowed-tools: Read Bash Glob Grep mcp__word-document-server__get_document_info mcp__word-document-server__get_document_outline mcp__word-document-server__list_available_documents
metadata:
    version: "0.1.0"
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

1. **Document Map check** — Does a Document Map exist in the conversation context?
   - YES → pass it to the target module
   - NO → run word-read first to generate one (unless task is word-read itself)

2. **Format Spec check** — Does the task require formatting rules?
   - YES → check if user has provided format requirements
   - NO format requirements yet → ask user: "请提供格式要求文档（Word/PDF/文字描述均可）"

3. **Token budget assessment** — Estimate operation scale:
   - Small (< 5 changes) → proceed directly
   - Medium (5-20 changes) → show plan, ask confirmation
   - Large (20+ changes) → show plan with estimated scope, ask confirmation

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

1. **Suggest next step** — Based on workflow table in `docs/CLAUDE.md`
2. **Report Token usage** — If the operation was large, note what was saved by caching

## What This Skill Does NOT Do

- Does not read, edit, or format documents directly
- Does not generate Document Maps (that's word-read)
- Does not make assumptions about formatting rules
