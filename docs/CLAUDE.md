# word-agent: Academic Paper Word Document Toolkit

You have access to the **word-agent** plugin — a suite of specialized agents for Word document operations in academic paper workflows. When Word/docx-related tasks are detected, route to the appropriate agent.

## Routing Guard

Do NOT activate word-agent modules for:
- General text editing unrelated to Word documents
- PDF, spreadsheet, or Google Docs operations
- Pure academic writing without a Word document target (use eco-agent instead)
- General coding tasks

Only activate when the request involves **Word document (.docx) operations** — reading, formatting, editing, or managing a .docx file, or explicit module invocation (`/word-agent:...`).

When ambiguous, ask: "您是需要操作 Word 文档，还是需要其他帮助？"

## Agent Routing

| User Intent | Keywords (EN) | Keywords (CN) | Route to Agent |
|-------------|---------------|---------------|----------------|
| Read/analyze document | "read", "analyze", "structure", "what's in this" | "读取", "分析", "看看", "文档结构" | `word-agent:word-read` |
| Compare documents | "compare", "diff", "what changed", "differences" | "对比", "比较", "区别", "改了什么" | `word-agent:word-read` |
| Format document | "format", "style", "apply formatting", "layout" | "排版", "格式化", "改格式", "按要求排" | `word-agent:word-format` |
| Generate TOC | "table of contents", "TOC", "outline" | "目录", "生成目录", "插入目录" | `word-agent:word-format` |
| CJK normalization | "Chinese-English spacing", "punctuation" | "中英文混排", "标点", "全角半角" | `word-agent:word-format` |
| Edit content | "change", "rewrite", "modify", "replace text" | "修改", "改写", "替换", "把XX改成YY" | `word-agent:word-edit` |
| Create document | "write from scratch", "create document", "new paper" | "从零写", "新建文档", "写一篇新的" | `word-agent:word-edit` |
| References | "references", "bibliography", "citation", "Zotero" | "参考文献", "引用", "文献格式" | `word-agent:word-reference` |
| Footnotes/endnotes | "footnote", "endnote" | "脚注", "尾注" | `word-agent:word-reference` |
| Tables | "table", "three-line table", "add table" | "表格", "三线表", "插入表格" | `word-agent:word-table-figure` |
| Figures/images | "figure", "image", "picture", "caption" | "图片", "插图", "题注" | `word-agent:word-table-figure` |
| Reviewer comments | "reviewer", "revision", "tracked changes" | "审稿意见", "修订", "审阅", "批注" | `word-agent:word-review` |
| Check cross-refs | "check references", "cross-reference", "verify" | "检查引用", "交叉引用", "核对" | `word-agent:word-check` |
| Format compliance | "check format", "verify formatting", "compliance" | "检查格式", "格式验证", "是否符合要求" | `word-agent:word-check` |
| Submission prep | "clean copy", "submit", "remove comments" | "投稿", "清理", "删除批注" | `word-agent:word-submit` |
| Split document | "split", "separate supplementary" | "拆分", "分离补充材料" | `word-agent:word-submit` |
| Merge documents | "merge", "combine", "join" | "合并", "合成一个文档" | `word-agent:word-submit` |
| Multiple operations | "format and check", "full preparation" | "排版并检查", "全流程" | `word-agent:word-orchestrate` |

## Cross-Module Workflow Suggestions

| Just Completed | Suggest Next | Reasoning |
|---------------|-------------|-----------|
| word-read | word-format | "文档结构已分析完成。需要按格式要求排版吗？" |
| word-format | word-check | "格式化完成。运行格式合规验证确认是否全部达标？" |
| word-check (issues found) | word-format | "发现 {n} 处格式问题。需要自动修复吗？" |
| word-edit | word-check | "内容修改完成。检查交叉引用是否仍然正确？" |
| word-review | word-check | "修订完成。运行最终检查确认所有修改无误？" |
| word-edit (new doc) | word-format → word-check | "文档内容已写入。运行格式化和最终检查？" |
| word-check (all pass) | word-submit | "所有检查通过。准备投稿清理版本？" |

## Interaction Principles

1. **Document Map First**: Most modules need a Document Map. If the user hasn't run word-read yet, orchestrator runs it first automatically.

2. **Language Matching**: Output language follows user input language. Chinese → Chinese. English → English.

3. **Confirm Before Batch Changes**: For operations affecting 10+ paragraphs, show a preview of planned changes and ask for confirmation before executing.

4. **Token Efficiency**: Never read the full document when an outline suffices. See `references/token_budget.md`.

5. **MCP First**: Always attempt Word MCP Server tools first. Only fall back to XML manipulation for tracked changes or operations without MCP support. See `references/tool_routing.md`.

6. **Format Spec Driven**: When formatting, always ask the user for their format requirement document first. Do not guess journal formatting rules.
