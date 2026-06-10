# Tool Routing Reference

## Core Principle

**五层路由优先级：**

```
P1: word-document-server MCP（稳定主力，基础读写）
P2: word-mcp-live MCP（实时编辑、tracked changes、批注、公式）
P3: docx-mcp MCP（结构验证、选择性 accept/reject、PII）
P4: adeu Python SDK（批量编辑 Markdown 中间路径）
P5: XML unpack/edit/pack（遗留兜底）
```

能用一次 MCP 调用完成的操作，绝不走多步路径。每个操作按优先级选择第一个可用的后端。

## 后端可用性检测

orchestrator Pre-flight 阶段检测各 server 是否可用：

```
word-document-server → list_available_documents（必须可用，否则整个 word-agent 不可用）
word-mcp-live        → 轻量 read-only 调用（不可用时跳过 P2 路由）
docx-mcp             → 轻量 read-only 调用（不可用时跳过 P3 路由）
adeu                 → python3 -c "import adeu"（不可用时跳过 P4 路由）
```

## Routing Table

### 读取操作

| 操作 | 首选路径 | 备选路径 | 说明 |
|------|---------|---------|------|
| 获取文档基本信息 | `get_document_info` | — | 页数、元数据，最便宜的调用 |
| 获取标题结构 | `get_document_outline` | `docx2python` | outline 只返回标题，极省 Token |
| 生成 Document Map | `docx2python` (Bash) | MCP 多次调用 | 一次解析获得完整层次结构 |
| 读取特定段落 | `get_paragraph_text_from_document` | `get_document_text` + 截取 | 按需读取，避免全文 |
| 读取全文 | `get_document_text` | `docx2python` | 仅在确实需要全文时使用 |
| 查找文本 | `find_text_in_document` | `get_document_text` + 搜索 | MCP 原生搜索更快 |
| 提取文本框内容 | `zipfile` + `ElementTree` 解析 `w:txbxContent` | `get_document_xml` + 手动解析 | MCP 和 docx2python 均跳过文本框 |
| 获取批注 | `get_all_comments` | `get_comments_by_author` | word-document-server 读取 |
| 获取修订列表 | `mcp__docx-mcp__get_tracked_changes` | XML 手动解析 | docx-mcp 提供结构化修订信息 |

### 内容编辑操作

| 操作 | 首选路径 | 备选路径 | 说明 |
|------|---------|---------|------|
| 替换文字 | `search_and_replace` | `python-docx-replace` → XML | 跨 run 匹配失败时用备选 |
| 添加段落 | `insert_line_or_paragraph_near_text` | XML 编辑 | 指定锚点文本插入 |
| 替换段落块 | `replace_paragraph_block_below_header` | XML 编辑 | 按标题定位替换整块 |
| 区块替换 | `replace_block_between_manual_anchors` | XML 编辑 | 两个锚点间替换 |
| 删除段落 | `delete_paragraph` | XML 编辑 | — |
| 添加标题 | `add_heading` | XML 编辑 | — |
| 添加页面分隔 | `add_page_break` | XML 编辑 | — |
| 批量内容修改（10+处） | `adeu` RedlineEngine (Bash) | 逐条 MCP 调用 | CriticMarkup → 原子写回 tracked changes |

### Tracked Changes（修订痕迹）

| 操作 | 首选路径 | 备选路径 | 说明 |
|------|---------|---------|------|
| 插入带修订 | `mcp__word-mcp-live__insert_text` (track_changes=true) | XML `<w:ins>` 拼装 | 原生修订标记，无需 unpack/pack |
| 删除带修订 | `mcp__word-mcp-live__delete_text` (track_changes=true) | XML `<w:del>` 拼装 | 同上 |
| 替换带修订 | `mcp__word-mcp-live__replace_text` (track_changes=true) | XML `<w:del>` + `<w:ins>` | 同上 |
| 查看修订列表 | `mcp__docx-mcp__get_tracked_changes` | XML 手动解析 | 结构化返回所有修订 |
| 接受所有修订 | `mcp__docx-mcp__accept_changes` | `accept_changes.py` 脚本 | docx-mcp 更可靠 |
| 选择性接受修订 | `mcp__docx-mcp__accept_changes` (range) | — | 按范围操作，仅 docx-mcp 支持 |
| 选择性拒绝修订 | `mcp__docx-mcp__reject_changes` (range) | — | 按范围操作，仅 docx-mcp 支持 |

### 批注操作

| 操作 | 首选路径 | 备选路径 | 说明 |
|------|---------|---------|------|
| 读取批注 | `get_all_comments` | `get_comments_by_author` | word-document-server |
| 添加批注 | `mcp__word-mcp-live__add_comment` | `comment.py` XML 脚本 | word-mcp-live 支持锚定到段落/文本 |
| 回复批注 | `mcp__word-mcp-live__reply_to_comment` | — | 仅 word-mcp-live 支持 |
| 解析批注 | `mcp__word-mcp-live__resolve_comment` | — | 仅 word-mcp-live 支持 |
| 取消解析 | `mcp__word-mcp-live__unresolve_comment` | — | 仅 word-mcp-live 支持 |
| 删除单条批注 | `mcp__word-mcp-live__delete_comment` | — | 仅 word-mcp-live 支持 |
| 删除所有批注 | XML 清除 comments.xml | — | word-submit 清理流程 |

### 格式化操作

| 操作 | 首选路径 | 备选路径 | 说明 |
|------|---------|---------|------|
| 修改格式 | **样式优先**：`scripts/format_document.py styles` 子命令（修改样式定义，自动传播到所有使用该样式的段落） | `format_text` 仅限零散 run 级特例（如单个词加粗），且事后必须运行 `normalize_fonts.py` | 禁止用 format_text 做段落/文档级格式化（直接格式化与样式冲突，造成字体混乱） |
| 批量格式修改 | `scripts/format_document.py styles` / `apply-spec` 一次调用 | XML 编辑 | 禁止 `format_text` 循环 —— 每次调用都可能产生 run 级字体不一致 |
| 创建自定义样式 | `create_custom_style` | XML 编辑 | — |
| 生成目录 | XML 插入 TOC 域代码 | — | 需要用户在 Word 中更新域 |

### 表格与图片

| 操作 | 首选路径 | 备选路径 | 说明 |
|------|---------|---------|------|
| 添加表格 | `add_table` | — | MCP 表格功能完整 |
| 格式化表格 | `format_table` + 系列工具 | XML 编辑 | 多个 MCP 工具配合 |
| 添加图片 | `add_picture` | XML + relationship | — |

### 结构验证（docx-mcp 专属）

| 操作 | 路径 | 说明 |
|------|------|------|
| paraId 唯一性验证 | `mcp__docx-mcp__validate_paraids` | 检测重复 ID，防止 Word 静默重写 |
| 文档综合审计 | `mcp__docx-mcp__audit_document` | 书签/图片引用/编号定义/样式完整性 |
| 水印移除 | `mcp__docx-mcp__remove_watermark` | — |
| PII 清洗 | `mcp__docx-mcp__scrub_pii` | 实验性，需 spaCy |
| 文档比较 | `mcp__docx-mcp__compare_documents` | 替代 docx2python+deepdiff |

### 实时编辑（word-mcp-live 专属）

| 操作 | 路径 | 条件 | 说明 |
|------|------|------|------|
| 实时编辑 | `mcp__word-mcp-live__*` COM/JXA | Word 打开状态 | 改一步看一步 |
| 撤销操作 | `mcp__word-mcp-live__undo_last_operation` | 仅实时模式 | 每操作 = 单次 Ctrl+Z |
| 布局诊断 | `mcp__word-mcp-live__diagnose_layout` | 仅实时模式 | 检测格式渲染问题 |
| 插入公式 | `mcp__word-mcp-live__insert_equation` | — | OMML 格式 |
| 插入交叉引用 | `mcp__word-mcp-live__insert_cross_reference` | — | 自动更新编号 |
| 添加水印 | `mcp__word-mcp-live__add_watermark` | — | — |

### 脚注/尾注与参考文献

| 操作 | 首选路径 | 备选路径 | 说明 |
|------|---------|---------|------|
| 添加脚注 | `add_footnote_to_document` | XML 编辑 | — |
| 添加尾注 | `add_endnote_to_document` | XML 编辑 | — |
| 参考文献格式化 | `citeproc-py` + CSL | 手动格式化 | 配合 Zotero MCP |

### 文档管理

| 操作 | 首选路径 | 备选路径 | 说明 |
|------|---------|---------|------|
| 文档对比 | `docx2python` + `deepdiff` + `redlines` | `mcp__docx-mcp__compare_documents` | Python 管道已验证稳定 |
| 文档合并 | `docxcompose` | 手动 XML | 自动处理样式冲突 |
| 文档复制 | `copy_document` | Bash cp | MCP 保持文档完整性 |
| .doc→.docx 转换 | `soffice --headless --convert-to docx` + 后处理 | 用户手动 Word 另存为 | 转换后必须修复，见 `doc_conversion.md` |
| 字体归一化 | `normalize_fonts.py --unify` | — | **所有后端写操作后必须执行** |
| 字体不一致检测 | `normalize_fonts.py --detect-only` | — | Document Map 生成时检测 |
| 转 PDF | `soffice --convert-to pdf` | — | LibreOffice |
| Markdown → docx | `mcp__docx-mcp__create_from_markdown` | Pandoc | docx-mcp 一步完成 |

## 降级触发条件

当首选路径失败时，按以下条件降级：

1. **MCP 工具返回错误** → 检查错误类型：
   - 文件不存在 → 提示用户确认路径
   - 文本未找到 → 尝试 `python-docx-replace`（可能是跨 run 问题）
   - 格式操作失败 → 降级到 XML 编辑

2. **MCP server 不可用** → 跳过对应优先级：
   - word-mcp-live 不可用 → tracked changes 降级到 XML unpack/edit/pack
   - docx-mcp 不可用 → 接受修订降级到 `accept_changes.py`；结构验证不可用
   - adeu 不可用 → 批量编辑降级到逐条 MCP 调用

3. **性能考虑** → 当同类操作超过 20 次时：
   - 逐次 MCP 调用可能比 adeu 批量编辑更慢
   - 评估后选择更高效的路径

4. **实时模式限制** → 部分操作在实时模式下不可用：
   - font normalization → 延迟到 Word 关闭后执行
   - XML 操作 → 文件被 Word 锁定，不可用

## 后端选择速查

```
需要 tracked changes？
  → word-mcp-live 可用？ → 用 word-mcp-live (track_changes=true)
  → 否 → XML unpack/edit/pack (Legacy)

需要批注操作？
  → 仅读取 → word-document-server (get_all_comments)
  → 添加/回复/resolve → word-mcp-live
  → 全部删除 → XML 清除

需要结构验证？
  → docx-mcp (validate_paraids / audit_document)

需要批量修改 10+ 处？
  → adeu 可用？ → adeu CriticMarkup 批量模式
  → 否 → 逐条 MCP 调用

需要实时编辑？
  → Word 打开 + word-mcp-live 可用？ → Live 模式 (COM/JXA)
  → 否 → 文件级模式（默认）

其他操作？
  → word-document-server（默认主力）
```
