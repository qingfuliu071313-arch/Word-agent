# Word-Agent

Claude Code 插件：学术论文 Word 文档全流程工具包。

## 功能

| 模块 | 功能 | 说明 |
|------|------|------|
| word-reader | 文档分析 | 生成 Document Map，对比两版文档差异 |
| word-formatter | 格式排版 | 解析格式要求 → 批量应用，TOC 生成，中英文混排规范化 |
| word-checker | 质量检查 | 交叉引用验证（图/表/公式），格式合规性检查 |
| word-edit | 内容编辑 | 精确替换、段落改写、修订模式（tracked changes），从零创建文档 |
| word-table-figure | 表格图片 | 三线表创建、图片插入、题注管理 |
| word-reference | 参考文献 | GB/T 7714 格式化、Zotero 集成、脚注尾注管理 |
| word-reviewer | 审稿修订 | 提取审稿意见、规划修改策略、生成逐条回复 |
| word-checker | 交叉引用 | 7 条检查规则验证图表公式引用完整性 |
| word-submit | 投稿准备 | 清理修订痕迹/批注/元数据、拆分/合并文档 |

## 安装

### 前置条件

- [Claude Code](https://docs.anthropic.com/en/docs/claude-code) CLI 已安装
- Python 3.8+
- [word-document-server](https://github.com/search?q=word-document-server+MCP) MCP 服务器已配置

### 一键安装

```bash
git clone https://github.com/你的用户名/Word-agent.git
cd Word-agent
bash scripts/setup.sh
```

### 手动安装

```bash
# 1. 安装 Python 依赖
pip3 install -r scripts/requirements.txt

# 2. 注册为 Claude Code 插件市场
claude plugin marketplace add "$(pwd)"

# 3. 安装插件
claude plugin install word-agent@word-agent
```

### MCP 服务器配置

word-agent 依赖 `word-document-server` MCP 服务器。请在 Claude Code 的 MCP 配置中添加：

```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "你的启动命令",
      "args": []
    }
  }
}
```

可选 MCP 服务器（启用更多功能）：
- **Zotero MCP** — 启用 word-reference 的 Zotero 文献库集成
- **Semantic Scholar MCP** — 启用文献搜索和引用分析

## 使用

在 Claude Code 中直接描述你的需求，插件会自动路由到对应模块：

```
# 分析文档结构
分析一下这个文档：论文.docx

# 按格式要求排版
按照这个格式要求排版：格式要求.pdf

# 检查交叉引用
检查论文.docx 的图表引用是否完整

# 从零创建文档
根据这个模板从零写一篇论文

# 生成投稿清理版
帮我生成 clean copy 准备投稿
```

也可以直接调用特定模块：

```
/word-agent:word-read
/word-agent:word-format
/word-agent:word-check
/word-agent:word-edit
```

## 架构

三层 WHO/HOW/WHAT 设计：

```
agents/            → WHO  (角色定义、模型选择、工具权限)
skills/*/SKILL.md  → HOW  (工作流程、模板、输出格式)
references/        → WHAT (领域知识、标准、规则)
```

## 设计原则

- **Document Map 缓存** — word-reader 生成一次结构摘要，所有下游模块复用
- **MCP 优先，XML 兜底** — 优先用 MCP 工具直接操作，仅在修订模式下使用 XML
- **样式先行** — 通过段落样式控制格式，禁止直接格式化，避免字体混乱
- **用户提供格式规范** — 不预设期刊模板，解析用户的格式要求文档为可执行规则

## 依赖

| 包 | 用途 | 阶段 |
|---|------|------|
| docx2python | 文档结构提取 | 核心 |
| deepdiff | 文档对比 | 核心 |
| redlines | 差异可视化 | 核心 |
| python-docx-replace | 跨 run 文本替换 | 编辑 |
| docxcompose | 文档合并 | 投稿 |
| citeproc-py | 引用格式化 | 文献 |

## 更新插件

修改源文件后需要重新安装缓存：

```bash
claude plugin uninstall word-agent@word-agent
claude plugin install word-agent@word-agent
```

或使用 setup 脚本：

```bash
bash scripts/setup.sh
```

## License

MIT
