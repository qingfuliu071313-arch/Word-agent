# Tracked Changes Reference

## 双路径架构

word-agent 支持两种 tracked changes 实现方式，按优先级自动选择：

| 路径 | 后端 | 条件 | 优势 |
|------|------|------|------|
| **原生模式（首选）** | word-mcp-live | word-mcp-live MCP 可用 | 单参数启用，原生修订标记，author 可配置 |
| **Legacy XML 模式** | docx skill scripts | word-mcp-live 不可用 | 无额外依赖，但步骤多且脆弱 |

## 原生 Tracked Changes（word-mcp-live）

### 基本用法

在任何编辑操作中添加 `track_changes: true` 参数：

```json
{
  "filename": "paper.docx",
  "find_text": "旧表述",
  "replace_text": "新表述",
  "track_changes": true
}
```

### Author 配置

通过 MCP 环境变量设置修订作者：

| 变量 | 用途 | 默认值 |
|------|------|-------|
| `MCP_AUTHOR` | 修订和批注的作者名 | "Author" |
| `MCP_AUTHOR_INITIALS` | 批注缩写 | — |

建议在 MCP 配置中设为 `"Claude"` 或用户指定的名称。

### 操作对照

| 操作 | word-mcp-live 工具 | 参数 |
|------|-------------------|------|
| 插入文本（带修订） | `insert_text` | `track_changes: true` |
| 删除文本（带修订） | `delete_text` | `track_changes: true` |
| 替换文本（带修订） | `replace_text` | `track_changes: true` |
| 编辑文本（带修订） | `edit_text` | `track_changes: true` |

### 修订显示效果

- **插入**: 绿色下划线（或 Word 用户设置的颜色）
- **删除**: 红色删除线
- **替换**: 删除旧文本（红色删除线）+ 插入新文本（绿色下划线）

## 查看与管理修订

### 查看修订列表

```
mcp__docx-mcp__get_tracked_changes(file_path)
```

返回所有修订的结构化信息：作者、时间、类型（insert/delete/format）、影响的文本。

### 选择性接受/拒绝

```
mcp__docx-mcp__accept_changes(file_path, range=...)   # 接受指定范围的修订
mcp__docx-mcp__reject_changes(file_path, range=...)    # 拒绝指定范围的修订
mcp__docx-mcp__accept_changes(file_path)               # 接受所有修订
```

## Legacy XML 模式（兜底）

当 word-mcp-live 不可用时，使用 docx skill 的 XML 操作：

### 流程

```
1. unpack.py document.docx unpacked/
2. 编辑 unpacked/word/document.xml（插入 <w:ins>/<w:del> 元素）
3. pack.py unpacked/ output.docx --original document.docx
```

### XML 模式

**删除:**
```xml
<w:del w:id="1" w:author="Claude" w:date="2026-05-28T00:00:00Z">
  <w:r>
    <w:rPr>{复制原始 rPr}</w:rPr>
    <w:delText>旧文本</w:delText>
  </w:r>
</w:del>
```

**插入:**
```xml
<w:ins w:id="2" w:author="Claude" w:date="2026-05-28T00:00:00Z">
  <w:r>
    <w:rPr>{复制原始 rPr}</w:rPr>
    <w:t>新文本</w:t>
  </w:r>
</w:ins>
```

### XML 注意事项

- 替换完整 `<w:r>` 元素，不要在 run 内部注入标签
- 复制原始 `<w:rPr>` 到 tracked change run 中
- `<w:del>` 内用 `<w:delText>` 而非 `<w:t>`
- 每个 `w:id` 在文档中必须唯一

## 触发条件

用户说以下关键词时，word-edit 和 word-review 应启用 tracked changes：

| 中文 | 英文 |
|------|------|
| 修订模式 | tracked changes |
| 保留修改痕迹 | revision marks |
| 修订 | track changes |
| 显示修改 | show changes |

## Font Normalization 兼容性

`normalize_fonts.py --unify` 只修改 `w:rFonts` 属性，不影响 `<w:ins>`/`<w:del>`/`<w:delText>` 等修订标记元素。两者可安全共存。
