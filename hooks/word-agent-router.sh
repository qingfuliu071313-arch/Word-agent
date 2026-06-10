#!/bin/bash
# Word-Agent Smart Routing Hook
# Two-tier detection:
#   Tier 1 — Is this about a Word document?
#   Tier 2 — Is the operation complex enough to warrant the skill pipeline?
#
# Heavy operations → route through word-agent skills (saves tokens on complex workflows)
# Light operations → read-only MCP queries only (everything else goes through skills)

set -e

# Graceful degradation: jq is required for JSON parsing/output.
# If jq is missing, exit silently instead of breaking the prompt flow.
if ! command -v jq > /dev/null 2>&1; then
  exit 0
fi

INPUT=$(cat)
# Claude Code sends the prompt in .prompt for UserPromptSubmit events;
# .user_prompt is kept as a fallback for older harness versions.
USER_PROMPT=$(echo "$INPUT" | jq -r '.prompt // .user_prompt // empty')

if [ -z "$USER_PROMPT" ]; then
  exit 0
fi

# ── Tier 1: Word context detection ──
WORD_CONTEXT='(\.(docx|doc)\b|[Ww]ord\s*(文档|document|doc)|docx|\.doc\b)'

if ! echo "$USER_PROMPT" | grep -iE "$WORD_CONTEXT" > /dev/null 2>&1; then
  # No Word file mentioned — pass through silently
  exit 0
fi

# ── Tier 2: Complexity classification ──

# HEAVY indicators: multi-step, whole-document, or specialized workflows
HEAVY_PATTERNS='(排版|格式化|format\s*(entire|whole|document|all)|审稿|审阅|修订|revision|review|tracked\s*change|修订模式|投稿|submit|clean\s*copy|删除批注|接受修订|检查引用|交叉引用|cross.?ref|检查格式|格式验证|compliance|对比|比较|compare|diff|读取|分析|看看|\bread\b|\banalyz|参考文献|bibliography|citation\s*format|三线表|three.?line|生成目录|TOC|table\s*of\s*content|批量|bulk|batch|所有|全部|整篇|每[个一]|all\s*(paragraph|heading|table|figure)|拆分|split|合并|merge|从零写|新建文档|create\s*document|write\s*from\s*scratch|水印|watermark|PII|结构检查|structural|paraId|bookmark|布局检查|layout\s*diag|公式|equation|撤销|undo)'

if echo "$USER_PROMPT" | grep -iE "$HEAVY_PATTERNS" > /dev/null 2>&1; then
  # HEAVY: Complex operation — route through word-agent skills
  CONTEXT="[WORD-AGENT ROUTING: HEAVY OPERATION]
A complex Word document operation was detected. Invoke the appropriate word-agent skill:
  - Formatting/排版/目录 → word-agent:word-format
  - Content editing/内容编辑 → word-agent:word-edit
  - Review & revision/审稿修订 → word-agent:word-review
  - Quality check/检查 → word-agent:word-check
  - Submission prep/投稿/拆分合并 → word-agent:word-submit
  - References/参考文献/脚注 → word-agent:word-reference
  - Tables & figures/表格图片 → word-agent:word-table-figure
  - Read & analysis/读取分析对比 → word-agent:word-read
  - Multi-step or unsure → word-agent:word-orchestrate
Do NOT call MCP tools directly for these operations — the skill pipeline handles font pairing, Document Map caching, font normalization, and style-first workflow. See docs/CLAUDE.md for the full routing table."

  jq -n --arg ctx "$CONTEXT" '{"hookSpecificOutput": {"hookEventName": "UserPromptSubmit", "additionalContext": $ctx}}'
  exit 0
fi

# LIGHT: Simple query on a Word file — only read-only info tools allowed directly
CONTEXT="[WORD-AGENT ROUTING: LIGHT OPERATION]
A Word document was mentioned and the request appears to be a lightweight query.
You may call ONLY these read-only MCP tools directly (per docs/CLAUDE.md):
  - get_document_info
  - get_document_outline
  - list_available_documents
ALL other operations (including any write, replace, or formatting — even a single search_and_replace) MUST go through the appropriate word-agent:* skill (word-read, word-edit, word-format, ...). Direct write calls bypass font pairing and font normalization and cause font chaos."

jq -n --arg ctx "$CONTEXT" '{"hookSpecificOutput": {"hookEventName": "UserPromptSubmit", "additionalContext": $ctx}}'
exit 0
