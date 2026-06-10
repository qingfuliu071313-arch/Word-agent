# Word-Agent v1.1.0

**A Claude Code plugin for academic paper Word document operations.**

**Claude Code 插件：学术论文 Word 文档全流程工具包。**

Bilingual support: English & 中文

---

## Features / 功能

### Core Modules / 核心模块

| Module | Function / 功能 | Description / 说明 |
|--------|-----------------|-------------------|
| word-reader | Document analysis / 文档分析 | Generate Document Map, compare versions, detect tracked changes / 生成文档地图、对比版本、检测修订 |
| word-formatter | Formatting / 格式化 | Parse format specs → batch apply, TOC generation, CJK normalization / 解析格式要求 → 批量应用、目录生成、中英混排 |
| word-checker | Quality check / 质量检查 | Cross-reference validation, format compliance, structural validation, layout diagnostics / 交叉引用验证、格式合规、结构验证、布局诊断 |
| word-edit | Content editing / 内容编辑 | Replacement, paragraph rewriting, tracked changes, equations, cross-references, bulk edit, undo / 替换、段落改写、修订模式、公式、交叉引用、批量编辑、撤销 |
| word-table-figure | Tables & figures / 表格与图片 | Three-line tables, image insertion, caption management, equation insertion / 三线表、插图、题注管理、公式插入 |
| word-reference | References / 参考文献 | GB/T 7714 formatting, Zotero integration, footnote/endnote management / GB/T 7714 格式化、Zotero 集成、脚注尾注管理 |
| word-reviewer | Review & revision / 审稿修订 | Extract comments, plan revisions, threaded comment replies, point-by-point response / 提取批注、规划修订、线程批注回复、逐条回复 |
| word-submit | Submission prep / 投稿准备 | Clean copy, split/merge, selective accept/reject, watermarks, PII scrubbing / 清理版、拆分合并、选择性修订处理、水印、PII 清洗 |

### v1.0.0 New Capabilities / v1.0.0 新增能力

| Capability / 能力 | Description / 说明 | Backend / 后端 |
|-------------------|-------------------|----------------|
| **Native tracked changes** / 原生修订模式 | Single MCP call per edit with Word-native revision marks / 单次调用即可产生 Word 原生修订标记 | word-mcp-live |
| **Selective accept/reject** / 选择性接受拒绝 | Review and selectively accept or reject individual tracked changes / 逐条审查并选择性接受或拒绝修订 | docx-mcp |
| **Threaded comments** / 线程批注 | Add, reply to, resolve, and delete individual comments / 添加、回复、解析、删除单条批注 | word-mcp-live |
| **Structural validation** / 结构验证 | Verify paraId uniqueness, bookmark integrity, image references / 验证段落 ID 唯一性、书签完整性、图片引用 | docx-mcp |
| **Live editing** / 实时编辑 | Edit documents while Word is open with real-time updates / Word 打开时实时编辑、即时可见 | word-mcp-live |
| **Bulk edit via Markdown** / Markdown 批量编辑 | Convert to Markdown, apply CriticMarkup, write back atomically / 转为 Markdown、CriticMarkup 批量修改、原子写回 | adeu SDK |
| **Per-action undo** / 逐步撤销 | Undo individual operations in live editing mode / 实时模式下逐步撤销单个操作 | word-mcp-live |
| **Equation insertion** / 公式插入 | Insert OMML equations from LaTeX input / 从 LaTeX 输入插入 OMML 公式 | word-mcp-live |
| **Cross-reference insertion** / 交叉引用插入 | Insert live cross-references to headings, figures, tables, equations / 插入可自动更新的交叉引用 | word-mcp-live |
| **Layout diagnostics** / 布局诊断 | Check page breaks, orphan/widow lines, text overflow / 检查页面分隔、孤行寡行、文本溢出 | word-mcp-live |
| **Watermark management** / 水印管理 | Add and remove watermarks (DRAFT, CONFIDENTIAL, etc.) / 添加和移除水印 | word-mcp-live + docx-mcp |
| **PII scrubbing** / PII 清洗 | Detect and redact personal identifiable information (experimental) / 检测并脱敏个人信息（实验性） | docx-mcp |

---

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

### Docker Install / Docker 安装

```bash
# Build and run / 构建并运行
bash scripts/docker-setup.sh --run

# Or use Docker Compose / 或使用 Docker Compose
docker compose up -d
```

Note: Docker mode only supports file-level operations. Live editing (COM/JXA) requires a native Word installation.

注意：Docker 模式仅支持文件级操作。实时编辑（COM/JXA）需要本地安装 Word。

### Windows Install / Windows 安装说明

All core features work on Windows — the Python script layer is fully cross-platform, live editing uses COM (Windows-native), and the default fonts (宋体/黑体/Times New Roman) are Windows system fonts. Three extra steps:

全部核心功能在 Windows 上可用——Python 脚本层完全跨平台，实时编辑走 Windows 原生的 COM，默认字体（宋体/黑体/Times New Roman）本就是 Windows 系统字体。需要额外做三件事：

```powershell
# 1. Install Git for Windows (Claude Code runs hooks via Git Bash)
#    安装 Git for Windows（Claude Code 通过 Git Bash 执行 hook）
winget install Git.Git

# 2. Install jq — required by the smart-routing hook.
#    Without it the hook degrades gracefully (no errors), but automatic
#    skill routing is disabled and you must invoke skills manually.
#    安装 jq——智能路由 hook 依赖它。缺失时 hook 静默降级（不报错），
#    但自动路由失效，需要手动调用 skill。
winget install jqlang.jq

# 3. (Optional) LibreOffice — only needed for legacy .doc conversion.
#    Add its program directory to PATH so `soffice` resolves.
#    （可选）LibreOffice——仅处理旧版 .doc 文件时需要，
#    安装后把程序目录加入 PATH 使 `soffice` 命令可用。
winget install TheDocumentFoundation.LibreOffice
# PATH += C:\Program Files\LibreOffice\program
```

Then follow Manual Install above in Git Bash (use `pip` instead of `pip3` if that is how Python is installed). `scripts/setup.sh` also runs in Git Bash.

之后在 Git Bash 中按上方"手动安装"步骤执行即可（若 Python 安装方式不同，`pip3` 换成 `pip`）。`scripts/setup.sh` 也可以在 Git Bash 中运行。

| Feature / 功能 | Windows 支持 |
|---|---|
| Formatting, editing, checking, references, submission prep / 排版、编辑、检查、参考文献、投稿准备 | ✅ Full / 完整 |
| Live editing (Word open) / 实时编辑（Word 打开状态） | ✅ Via COM / 走 COM |
| File-lock detection / 文件锁检测 | ✅ Native `~$` owner files / 原生 `~$` 锁文件 |
| Smart routing hook / 智能路由 hook | ✅ Needs Git Bash + jq / 需 Git Bash + jq |
| Legacy .doc conversion / 旧版 .doc 转换 | ✅ Needs LibreOffice in PATH / 需 LibreOffice 加入 PATH |

---

## MCP Server Configuration / MCP 服务器配置

word-agent uses a multi-backend architecture with three MCP servers:

word-agent 使用三后端架构，包含三台 MCP 服务器：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "your-start-command",
      "args": []
    },
    "word-mcp-live": {
      "command": "uvx",
      "args": ["word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Claude",
        "MCP_AUTHOR_INITIALS": "CL"
      }
    },
    "docx-mcp": {
      "command": "uvx",
      "args": ["docx-mcp-server"]
    }
  }
}
```

| Server | Role / 角色 | Required / 必需 |
|--------|------------|----------------|
| **word-document-server** | Core read/write, formatting, tables / 基础读写、格式化、表格 | Yes / 是 |
| **word-mcp-live** | Live editing, tracked changes, comments, equations / 实时编辑、修订、批注、公式 | Recommended / 推荐 |
| **docx-mcp** | Structural validation, selective accept/reject, PII / 结构验证、选择性修订、PII | Recommended / 推荐 |

word-agent gracefully degrades when optional servers are unavailable — features fall back to XML manipulation or skip with a warning.

当可选 server 不可用时，word-agent 自动降级到 XML 操作或跳过并给出提示。

Additional optional MCP servers / 其他可选 MCP 服务器：
- **Zotero MCP** — Zotero library integration for word-reference / 启用 word-reference 的 Zotero 文献库集成
- **Semantic Scholar MCP** — Literature search and citation analysis / 启用文献搜索和引用分析

See `references/mcp_servers.md` for detailed configuration guide / 详细配置见 `references/mcp_servers.md`。

---

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

# Edit with tracked changes / 修订模式编辑
Change "30 days" to "60 days" with tracked changes
用修订模式把"30天"改成"60天"

# Insert an equation / 插入公式
Insert equation E=mc^2 after paragraph 3
在第3段后面插入公式 E=mc²

# Review and revise / 审稿修订
Process these reviewer comments and make revisions
处理这些审稿意见并修订

# Selective accept/reject / 选择性修订处理
Show me all tracked changes and let me choose which to accept
列出所有修订让我选择接受或拒绝

# Check structural integrity / 结构检查
Run a full check including structural validation
运行全面检查（含结构验证）

# Generate a clean copy for submission / 生成投稿清理版
Generate a clean copy ready for submission
帮我生成 clean copy 准备投稿

# Bulk edit / 批量编辑
Make these 15 changes to the manuscript
批量修改这15处内容
```

You can also invoke a specific module directly / 也可以直接调用特定模块：

```
/word-agent:word-read
/word-agent:word-format
/word-agent:word-check
/word-agent:word-edit
/word-agent:word-review
/word-agent:word-submit
```

---

## Architecture / 架构

### Three-Layer Design / 三层设计

```
agents/            → WHO  (persona, model routing, tool permissions)
                          (角色定义、模型选择、工具权限)
skills/*/SKILL.md  → HOW  (workflow phases, templates, output format)
                          (工作流程、模板、输出格式)
references/        → WHAT (domain knowledge, standards, rules)
                          (领域知识、标准、规则)
```

### Five-Tier Tool Routing / 五层工具路由

```
P1: word-document-server  → Stable primary (read/write/format/tables)
                             稳定主力（读写/格式化/表格）
P2: word-mcp-live         → Live editing, tracked changes, comments, equations
                             实时编辑、修订、批注、公式
P3: docx-mcp              → Structural validation, selective accept/reject
                             结构验证、选择性接受拒绝
P4: adeu Python SDK       → Bulk editing via Markdown intermediate
                             通过 Markdown 中间表示批量编辑
P5: XML unpack/edit/pack  → Legacy fallback
                             遗留兜底方案
```

Each operation routes to the highest-priority backend available. See `references/tool_routing.md` for the complete routing table.

每个操作路由到可用的最高优先级后端。详见 `references/tool_routing.md`。

---

## Design Principles / 设计原则

- **Document Map caching** — word-reader generates a structural summary once; all downstream modules reuse it, never re-read the full document
  **Document Map 缓存** — word-reader 生成一次结构摘要，所有下游模块复用
- **Multi-backend with graceful degradation** — Three MCP servers with automatic fallback; no feature is hard-locked to a single backend
  **多后端自动降级** — 三台 MCP 服务器自动回退，无功能硬绑定到单一后端
- **Style-first** — Control formatting through paragraph styles, never apply direct formatting, avoid font chaos
  **样式先行** — 通过段落样式控制格式，禁止直接格式化，避免字体混乱
- **Mandatory font normalization** — Every write operation triggers `normalize_fonts.py` to fix theme fonts, missing eastAsia, and bare runs
  **强制字体归一化** — 每次写操作后执行字体归一化，修复主题字体、缺失 eastAsia、裸 run
- **User-provided format specs** — No pre-built journal templates; parse the user's format requirement document into executable rules
  **用户提供格式规范** — 不预设期刊模板，解析用户的格式要求文档为可执行规则
- **Text Box awareness** — Extract text box content (`w:txbxContent`) via XML parsing, since standard text extraction tools skip them entirely
  **文本框感知** — 通过 XML 解析提取文本框内容，标准文本提取工具会完全跳过文本框
- **No mixed modes** — Live editing and file editing never mix on the same file in the same session
  **模式不混用** — 同一文件同一会话中实时模式和文件模式不得混用

---

## Quality Checks / 质量检查

word-checker provides four verification modes:

word-checker 提供四种检查模式：

| Mode | Trigger / 触发词 | What It Checks / 检查内容 |
|------|-----------------|--------------------------|
| **A** Cross-Reference / 交叉引用 | "检查引用", "cross-reference" | Figure/table/equation references match assets / 图表公式引用与实际资产匹配 |
| **B** Format Compliance / 格式合规 | "检查格式", "format compliance" | Document formatting vs user-provided spec / 文档格式与用户规范对比 |
| **C** Structural Validation / 结构验证 | "结构检查", "structural" | paraId uniqueness, bookmarks, image refs / 段落 ID、书签、图片引用完整性 |
| **D** Layout Diagnostics / 布局诊断 | "布局检查", "layout" | Orphan/widow lines, page breaks, overflow (live mode only) / 孤行寡行、分页、溢出（仅实时模式） |

---

## Dependencies / 依赖

| Package | Purpose / 用途 | Priority / 优先级 |
|---------|----------------|-------------------|
| docx2python | Document structure extraction / 文档结构提取 | Required / 必需 |
| deepdiff | Document comparison / 文档对比 | Required / 必需 |
| redlines | Diff visualization / 差异可视化 | Required / 必需 |
| word-mcp-live | Live editing, tracked changes, comments, equations / 实时编辑、修订、批注、公式 | Recommended / 推荐 |
| docx-mcp-server | Structural validation, selective accept/reject / 结构验证、选择性修订 | Recommended / 推荐 |
| adeu | Markdown intermediate for bulk editing / 批量编辑 Markdown 中间表示 | Recommended / 推荐 |
| python-docx-replace | Cross-run text replacement / 跨 run 文本替换 | Optional / 可选 |
| docxcompose | Document merging / 文档合并 | Optional / 可选 |
| citeproc-py | Citation formatting / 引用格式化 | Optional / 可选 |
| pytest | Test suite / 测试套件 | Dev / 开发 |

---

## Testing / 测试

```bash
# Run all tests / 运行所有测试
bash scripts/run_tests.sh

# Run with verbose output / 详细输出
bash scripts/run_tests.sh -v

# Run specific test class / 运行特定测试类
bash scripts/run_tests.sh -k "FontNormalization"
```

Test coverage includes: font normalization, document structure parsing, text box extraction, skill routing consistency, and reference file completeness.

测试覆盖：字体归一化、文档结构解析、文本框提取、技能路由一致性、参考文档完整性。

---

## Project Structure / 项目结构

```
Word-agent/
├── .claude-plugin/          Plugin manifests / 插件清单
│   ├── plugin.json
│   └── marketplace.json
├── agents/                  9 agent definitions / 9 个 Agent 定义
├── skills/                  9 skill workflows / 9 个 Skill 工作流
│   ├── word-read/
│   ├── word-format/
│   ├── word-edit/
│   ├── word-check/
│   ├── word-review/
│   ├── word-submit/
│   ├── word-reference/
│   ├── word-table-figure/
│   └── word-orchestrate/
├── references/              Shared knowledge files / 共享知识文件
│   ├── tool_routing.md          Five-tier routing rules / 五层路由规则
│   ├── mcp_servers.md           MCP server configuration / MCP 服务器配置
│   ├── tracked_changes.md       Native vs Legacy tracked changes / 原生与遗留修订模式
│   ├── comment_operations.md    Threaded comment operations / 线程批注操作
│   ├── structural_validation.md OOXML structural checks / OOXML 结构检查
│   ├── live_editing.md          Live editing mode / 实时编辑模式
│   ├── equation_crossref.md     Equations & cross-references / 公式与交叉引用
│   ├── adeu_integration.md      adeu bulk editing pipeline / adeu 批量编辑
│   ├── font_normalization.md    Font normalization script / 字体归一化脚本
│   ├── submission_checklist.md  Pre-submission checklist / 投稿前检查清单
│   └── ...
├── scripts/                 Setup & utility scripts / 安装与工具脚本
├── test/                    Test suite / 测试套件
├── docs/                    Runtime instructions / 运行时指令
├── Dockerfile               Docker image / Docker 镜像
├── docker-compose.yml       Docker Compose config / Docker Compose 配置
└── CLAUDE.md                Developer guide / 开发者指南
```

---

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

---

## Acknowledgments / 致谢

Word-Agent's multi-backend architecture was inspired by and integrates with the following open-source projects. We thank their authors for their excellent work.

Word-Agent 的多后端架构受到以下开源项目的启发并与之集成。感谢这些项目的作者们的出色工作。

| Project / 项目 | Author / 作者 | License | Contribution / 贡献 |
|----------------|--------------|---------|---------------------|
| [word-mcp-live](https://github.com/ykarapazar/word-mcp-live) | ykarapazar | MIT | Live editing, native tracked changes, threaded comments, equation insertion / 实时编辑、原生修订、线程批注、公式插入 |
| [docx-mcp](https://github.com/SecurityRonin/docx-mcp) | SecurityRonin | MIT | Structural validation, selective accept/reject, PII scrubbing / 结构验证、选择性修订处理、PII 清洗 |
| [adeu](https://github.com/dealfluence/adeu) | dealfluence | MIT | Markdown intermediate representation for bulk editing / 批量编辑的 Markdown 中间表示 |

---

## License

[MIT](LICENSE)
