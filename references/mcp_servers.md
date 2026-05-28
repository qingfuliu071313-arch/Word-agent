# MCP Servers Reference

## 多后端架构

word-agent 使用三台 MCP server 协同工作，Claude Code 的命名空间机制 (`mcp__<server>__<tool>`) 天然隔离工具名冲突。

## Server 清单

### 1. word-document-server（主力后端）

| 属性 | 值 |
|------|---|
| 角色 | 基础读写操作、格式化、表格、图片、脚注 |
| 工具数 | 60+ |
| 命名空间 | `mcp__word-document-server__*` |
| 依赖 | 已有，无需额外安装 |
| 平台 | 全平台（Windows / macOS / Linux） |

**配置示例:**
```json
{
  "mcpServers": {
    "word-document-server": {
      "command": "node",
      "args": ["/path/to/word-document-server/dist/index.js"]
    }
  }
}
```

### 2. word-mcp-live（实时编辑 + 修订 + 批注）

| 属性 | 值 |
|------|---|
| 角色 | 实时编辑（Word 打开状态）、原生 tracked changes、threaded comments、公式、交叉引用、水印 |
| 工具数 | 124（跨平台 80 + Live 44） |
| 命名空间 | `mcp__word-mcp-live__*` |
| 安装 | `pip install word-mcp-live` |
| 平台 | 跨平台工具: 全平台；Live 工具: Windows (COM) / macOS (JXA) |

**环境变量:**
| 变量 | 用途 | 默认值 |
|------|------|-------|
| `MCP_AUTHOR` | tracked changes 和批注的作者名 | "Author" |
| `MCP_AUTHOR_INITIALS` | 批注作者缩写 | — |
| `MCP_TRANSPORT` | 传输方式 (stdio/sse/streamable-http) | stdio |

**配置示例:**
```json
{
  "mcpServers": {
    "word-mcp-live": {
      "command": "uvx",
      "args": ["word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Claude",
        "MCP_AUTHOR_INITIALS": "CL"
      }
    }
  }
}
```

**三种工作模式:**
| 模式 | 技术 | 条件 | 能力 |
|------|------|------|------|
| Windows Live | COM 自动化 | Windows + Word 已安装 | 全部 124 工具 |
| macOS Live | JXA 自动化 | macOS + Word for Mac | 120 工具（不含批注回复/resolve/水印） |
| 跨平台 | python-docx | 全平台 | 80 工具（无实时编辑） |

服务器自动检测平台并选择对应模式。

### 3. docx-mcp（结构验证 + 选择性修订）

| 属性 | 值 |
|------|---|
| 角色 | OOXML 结构验证、选择性 accept/reject 修订、文档审计、PII 清洗、水印移除 |
| 工具数 | 200+ |
| 命名空间 | `mcp__docx-mcp__*` |
| 安装 | `pip install docx-mcp-server` |
| 平台 | 全平台 |

**配置示例:**
```json
{
  "mcpServers": {
    "docx-mcp": {
      "command": "uvx",
      "args": ["docx-mcp-server"]
    }
  }
}
```

**PII 清洗依赖（可选）:**
```bash
python -m spacy download en_core_web_lg  # ~560MB
```

## 后端可用性检测

orchestrator 在 Phase 2 Pre-flight 中对每台 server 执行轻量检测：

```
word-document-server → get_document_info (已有文档) 或 list_available_documents
word-mcp-live        → 尝试调用任意 read-only 工具
docx-mcp             → 尝试调用任意 read-only 工具
```

检测失败的 server 在当前会话中标记为不可用，对应路由跳过。

## 工具命名空间对照

| 操作类别 | word-document-server | word-mcp-live | docx-mcp |
|---------|---------------------|---------------|----------|
| 基础读写 | ✅ 主力 | ✅ 跨平台模式 | ✅ |
| 格式化 | ✅ 主力 | ✅ | — |
| 表格 | ✅ 主力 | ✅ | ✅ |
| Tracked Changes | ❌ | ✅ 主力 | ✅ |
| 批注管理 | 仅读取 | ✅ 主力（完整 CRUD） | ✅ |
| 结构验证 | ❌ | ❌ | ✅ 唯一 |
| 公式/交叉引用 | ❌ | ✅ 唯一 | ❌ |
| 水印 | ❌ | ✅ 添加 | ✅ 移除 |
| PII 清洗 | ❌ | ❌ | ✅ 唯一 |
| 实时编辑 | ❌ | ✅ 唯一 | ❌ |

## 冲突解决规则

1. **读操作**: word-document-server 为权威来源（已验证稳定）
2. **写操作（基础）**: word-document-server 优先（add_paragraph, format_text 等）
3. **Tracked Changes**: word-mcp-live 优先 → XML legacy 兜底
4. **结构验证**: docx-mcp 唯一提供
5. **批注写入**: word-mcp-live 优先
6. **文档状态分歧时**: 以 word-document-server 的读取为准
