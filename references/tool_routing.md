# Tool Routing Reference

## Core Principle

**MCP 单步 > MCP 多步 > Python 库 > XML unpack/edit/pack**

能用一次 MCP 调用完成的操作，绝不走多步路径。只有 MCP 不支持的操作才降级到 Python 库或 XML 操作。

## Routing Table

| 操作 | 首选路径 | 备选路径 | 说明 |
|------|---------|---------|------|
| 获取文档基本信息 | `get_document_info` | — | 页数、元数据，最便宜的调用 |
| 获取标题结构 | `get_document_outline` | `docx2python` | outline 只返回标题，极省 Token |
| 生成 Document Map | `docx2python` (Bash) | MCP 多次调用 | 一次解析获得完整层次结构 |
| 读取特定段落 | `get_paragraph_text_from_document` | `get_document_text` + 截取 | 按需读取，避免全文 |
| 读取全文 | `get_document_text` | `docx2python` | 仅在确实需要全文时使用 |
| 查找文本 | `find_text_in_document` | `get_document_text` + 搜索 | MCP 原生搜索更快 |
| 替换文字 | `search_and_replace` | `python-docx-replace` | 跨 run 匹配失败时用备选 |
| 修改格式 | `format_text` | XML 编辑 | 单段落格式用 MCP |
| 批量格式修改 | `format_text` 循环 | XML 编辑 | 同类格式变更可批量 MCP 调用 |
| 添加段落 | `insert_line_or_paragraph_near_text` | XML 编辑 | 指定锚点文本插入 |
| 替换段落块 | `replace_paragraph_block_below_header` | XML 编辑 | 按标题定位替换整块 |
| 区块替换 | `replace_block_between_manual_anchors` | XML 编辑 | 两个锚点间替换 |
| 删除段落 | `delete_paragraph` | XML 编辑 | — |
| 添加标题 | `add_heading` | XML 编辑 | — |
| 添加页面分隔 | `add_page_break` | XML 编辑 | — |
| 添加表格 | `add_table` | docx-js 创建 | MCP 表格功能完整 |
| 格式化表格 | `format_table` + 系列工具 | XML 编辑 | 多个 MCP 工具配合 |
| 添加图片 | `add_picture` | XML + relationship | — |
| 添加脚注 | `add_footnote_to_document` | XML 编辑 | — |
| 添加尾注 | `add_endnote_to_document` | XML 编辑 | — |
| 获取批注 | `get_all_comments` | XML 读取 | — |
| 创建自定义样式 | `create_custom_style` | XML 编辑 | — |
| 修订痕迹 | **必须** XML unpack/edit/pack | 无 MCP 替代 | tracked changes 只能走 XML |
| 接受所有修订 | `accept_changes.py` 脚本 | 手动 XML | docx skill 脚本 |
| 添加批注 | `comment.py` + XML 标记 | 无 MCP 替代 | docx skill 脚本 |
| 生成目录 | XML 插入 TOC 域代码 | docx-js `TableOfContents` | 需要用户在 Word 中更新域 |
| 文档对比 | `docx2python` + `deepdiff` + `redlines` | 逐段 MCP 读取 + diff | Python 管道最高效 |
| 参考文献格式化 | `citeproc-py` + CSL | 手动格式化 | 配合 Zotero MCP |
| 文档合并 | `docxcompose` | 手动 XML | 自动处理样式冲突 |
| 文档复制 | `copy_document` | Bash cp | MCP 保持文档完整性 |
| .doc→.docx 转换 | `soffice --headless --convert-to docx` + 后处理修复 | 用户手动用 Word 另存为 | 转换后必须修复兼容模式和字体，见 `doc_conversion.md` |
| 转 PDF | `soffice.py --convert-to pdf` | — | LibreOffice |

## 降级触发条件

当首选路径失败时，按以下条件降级：

1. **MCP 工具返回错误** → 检查错误类型：
   - 文件不存在 → 提示用户确认路径
   - 文本未找到 → 尝试 `python-docx-replace`（可能是跨 run 问题）
   - 格式操作失败 → 降级到 XML 编辑

2. **MCP 工具功能不足** → 直接使用备选路径：
   - 需要 tracked changes → XML
   - 需要添加批注 → `comment.py`
   - 需要生成目录 → XML 域代码

3. **性能考虑** → 当同类操作超过 20 次时：
   - 逐次 MCP 调用可能比一次 XML 批量编辑更慢
   - 评估后选择更高效的路径
