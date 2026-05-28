#!/bin/bash
# Word-Agent Smart Routing Hook
# Two-tier detection:
#   Tier 1 — Is this about a Word document?
#   Tier 2 — Is the operation complex enough to warrant the Agent pipeline?
#
# Heavy operations → route through word-agent skills (saves tokens on complex workflows)
# Light operations → handle directly with MCP (avoids ~3000 token Agent overhead)

set -e

INPUT=$(cat)
USER_PROMPT=$(echo "$INPUT" | jq -r '.user_prompt // empty')

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
HEAVY_PATTERNS='(排版|格式化|format\s*(entire|whole|document|all)|审稿|审阅|修订|revision|review|tracked\s*change|修订模式|投稿|submit|clean\s*copy|删除批注|接受修订|检查引用|交叉引用|cross.?ref|检查格式|格式验证|compliance|对比|比较|compare|diff|参考文献|bibliography|citation\s*format|三线表|three.?line|生成目录|TOC|table\s*of\s*content|批量|bulk|batch|所有|全部|整篇|每[个一]|all\s*(paragraph|heading|table|figure)|拆分|split|合并|merge|从零写|新建文档|create\s*document|write\s*from\s*scratch|水印|watermark|PII|结构检查|structural|paraId|bookmark|布局检查|layout\s*diag|公式|equation|撤销|undo)'

if echo "$USER_PROMPT" | grep -iE "$HEAVY_PATTERNS" > /dev/null 2>&1; then
  # HEAVY: Complex operation — route through word-agent
  CONTEXT="[WORD-AGENT ROUTING: HEAVY OPERATION]
A complex Word document operation was detected. Use the Agent tool with the appropriate word-agent subagent:
  - Formatting/排版 → word-agent:word-formatter
  - Content editing/内容编辑 → word-agent:word-content-editor
  - Review & revision/审稿修订 → word-agent:word-reviewer
  - Quality check/检查 → word-agent:word-checker
  - Submission prep/投稿 → word-agent:word-submit
  - References/参考文献 → word-agent:word-reference
  - Tables & figures/表格图片 → word-agent:word-table-figure
  - Analysis/分析 → word-agent:word-reader
  - Multi-step or unsure → word-agent:word-orchestrator
Do NOT call MCP tools directly for these operations — the Agent pipeline handles font pairing, Document Map caching, font normalization, and style-first workflow. See docs/CLAUDE.md for the full routing table."

  jq -n --arg ctx "$CONTEXT" '{"additionalContext": $ctx, "continue": true}'
  exit 0
fi

# LIGHT: Simple operation on a Word file — no Agent needed
# Examples: read a paragraph, find text, single word replace, quick info query
CONTEXT="[WORD-AGENT ROUTING: LIGHT OPERATION]
A Word document was mentioned but the operation appears simple. You may call MCP tools directly for:
  - Single find/replace (search_and_replace)
  - Reading specific paragraphs (get_paragraph_text_from_document)
  - Quick info queries (get_document_info, get_document_outline)
  - Listing documents (list_available_documents)
REMINDER: If the task turns out to be more complex than expected (multi-step, batch changes, formatting), switch to the appropriate word-agent:* skill via the Agent tool."

jq -n --arg ctx "$CONTEXT" '{"additionalContext": $ctx, "continue": true}'
exit 0
