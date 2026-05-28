# Word-Agent

**A Claude Code plugin for academic paper Word document operations.**

**Claude Code 插件：学术论文 Word 文档全流程工具包。**

Bilingual support: English & 中文

---

## Features / 功能

| Module | Function | Description |
|--------|----------|-------------|
| word-reader | Document analysis | Generate Document Map, compare two document versions |
| word-formatter | Formatting | Parse format specs → batch apply, TOC generation, CJK normalization |
| word-checker | Quality check | Cross-reference validation (figures/tables/equations), format compliance |
| word-edit | Content editing | Precise replacement, paragraph rewriting, tracked changes, create from scratch |
| word-table-figure | Tables & figures | Three-line table creation, image insertion, caption management |
| word-reference | References | GB/T 7714 formatting, Zotero integration, footnote/endnote management |
| word-reviewer | Review & revision | Extract reviewer comments, plan revisions, generate point-by-point response |
| word-submit | Submission prep | Clean tracked changes/comments/metadata, split/merge documents |

## Installation / 安装

### Prerequisites / 前置条件

- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) CLI installed
- Python 3.8+
- [word-document-server](https://github.com/search?q=word-document-server+MCP) MCP server configured

### Quick Install / 一键安装

```bash
git clone https://github.com/qingfuliu071313-arch/Word-agent.git
cd Word-agent
bash scripts/setup.sh
```

### Manual Install / 手动安装

```bash
# 1. Install Python dependencies / 安装 Python 依赖
pip3 install -r scripts/requirements.txt

# 2. Register as Claude Code plugin marketplace / 注册为 Claude Code 插件市场
claude plugin marketplace add "$(pwd)"

# 3. Install the plugin / 安装插件
claude plugin install word-agent@word-agent
```

### MCP Server Configuration / MCP 服务器配置

word-agent requires the `word-document-server` MCP server. Add to your Claude Code MCP config:

word-agent 依赖 `word-document-server` MCP 服务器。请在 Claude Code 的 MCP 配置中添加：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "your-start-command",
      "args": []
    }
  }
}
```

Optional MCP servers for extended functionality / 可选 MCP 服务器（启用更多功能）：
- **Zotero MCP** — Enables Zotero library integration for word-reference / 启用 word-reference 的 Zotero 文献库集成
- **Semantic Scholar MCP** — Enables literature search and citation analysis / 启用文献搜索和引用分析

## Usage / 使用

Describe your task in Claude Code and the plugin auto-routes to the right module:

在 Claude Code 中直接描述你的需求，插件会自动路由到对应模块：

```
# Analyze document structure / 分析文档结构
Analyze this document: paper.docx
分析一下这个文档：论文.docx

# Apply formatting from a spec / 按格式要求排版
Format according to this spec: format_guide.pdf
按照这个格式要求排版：格式要求.pdf

# Check cross-references / 检查交叉引用
Check if all figure/table references in paper.docx are complete
检查论文.docx 的图表引用是否完整

# Create a document from scratch / 从零创建文档
Write a new paper based on this template
根据这个模板从零写一篇论文

# Generate a clean copy for submission / 生成投稿清理版
Generate a clean copy ready for submission
帮我生成 clean copy 准备投稿
```

You can also invoke a specific module directly / 也可以直接调用特定模块：

```
/word-agent:word-read
/word-agent:word-format
/word-agent:word-check
/word-agent:word-edit
```

## Architecture / 架构

Three-layer WHO / HOW / WHAT design:

三层 WHO / HOW / WHAT 设计：

```
agents/            → WHO  (persona, model routing, tool permissions)
                          (角色定义、模型选择、工具权限)
skills/*/SKILL.md  → HOW  (workflow phases, templates, output format)
                          (工作流程、模板、输出格式)
references/        → WHAT (domain knowledge, standards, rules)
                          (领域知识、标准、规则)
```

## Design Principles / 设计原则

- **Document Map caching** — word-reader generates a structural summary once; all downstream modules reuse it, never re-read the full document
  **Document Map 缓存** — word-reader 生成一次结构摘要，所有下游模块复用
- **MCP first, XML fallback** — Use Word MCP Server tools for direct operations; only fall back to XML for tracked changes
  **MCP 优先，XML 兜底** — 优先用 MCP 工具直接操作，仅在修订模式下使用 XML
- **Style-first** — Control formatting through paragraph styles, never apply direct formatting, avoid font chaos
  **样式先行** — 通过段落样式控制格式，禁止直接格式化，避免字体混乱
- **User-provided format specs** — No pre-built journal templates; parse the user's format requirement document into executable rules
  **用户提供格式规范** — 不预设期刊模板，解析用户的格式要求文档为可执行规则
- **Text Box awareness** — Extract text box content (`w:txbxContent`) via XML parsing, since standard text extraction tools skip them entirely. Critical for format templates with annotated instructions.
  **文本框感知** — 通过 XML 解析提取文本框内容（`w:txbxContent`），标准文本提取工具会完全跳过文本框。对包含注释说明的格式模板至关重要。

## Dependencies / 依赖

| Package | Purpose / 用途 | Phase / 阶段 |
|---------|----------------|--------------|
| docx2python | Document structure extraction / 文档结构提取 | Core / 核心 |
| deepdiff | Document comparison / 文档对比 | Core / 核心 |
| redlines | Diff visualization / 差异可视化 | Core / 核心 |
| python-docx-replace | Cross-run text replacement / 跨 run 文本替换 | Editing / 编辑 |
| docxcompose | Document merging / 文档合并 | Submission / 投稿 |
| citeproc-py | Citation formatting / 引用格式化 | References / 文献 |

## Updating the Plugin / 更新插件

After modifying source files, reinstall the cache:

修改源文件后需要重新安装缓存：

```bash
claude plugin uninstall word-agent@word-agent
claude plugin install word-agent@word-agent
```

Or use the setup script / 或使用 setup 脚本：

```bash
bash scripts/setup.sh
```

## License

MIT
