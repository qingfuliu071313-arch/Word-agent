# Document Creation Rules

## 新建文档防错规则

本文件定义了从零创建 Word 文档时必须遵守的规则，适用于 word-edit 和 word-format 模块。
违反这些规则会导致：字体混乱、标题不一致、目录失效。

---

## Rule 1: 样式先行，禁止直接格式化

- 所有格式通过段落样式 (paragraph style) 控制
- NEVER 对单个段落调用 `format_text` 设置字体字号
- 流程：定义样式 → 写入内容（指定 style_name）→ 内容自动继承样式格式
- 如果需要修改格式，修改样式定义，而不是修改单个段落

**原因：** `format_text` 创建 direct formatting，优先级高于样式但容易不一致。当一个段落有样式+直接格式时，Word 会混合显示，导致同一行出现多种字体。

## Rule 2: 使用 Word 内置标题样式

- 标题必须使用 Word 内置的 `Heading 1`, `Heading 2`, `Heading 3` 样式
- 可以修改内置样式的外观（字体、字号、加粗、对齐），但 NEVER 创建自定义标题样式
- 自定义样式（如 "我的标题1"、"MyHeading"）不会被 TOC 字段识别

**原因：** Word 的 TOC 字段默认扫描 `Heading 1-9` 内置样式。使用自定义样式的标题不会出现在自动目录中，即使手动更新也不行。

## Rule 3: 中英文字体必须配对设置

每个样式必须同时设置三种字体属性：

| XML 属性 | 作用 | 示例 |
|----------|------|------|
| `w:eastAsia` | 中文、日文、韩文字符 | 宋体、黑体、楷体 |
| `w:ascii` | ASCII 字符（英文、数字） | Times New Roman、Arial |
| `w:hAnsi` | 高位 ANSI 字符 | 通常与 ascii 相同 |

常见学术论文配对：

| 用途 | 中文字体 | 英文字体 |
|------|---------|---------|
| 正文 | 宋体 | Times New Roman |
| 标题 | 黑体 | Arial |
| 强调 | 楷体 | Times New Roman |

**python-docx 设置方式：**
```python
from docx.oxml.ns import qn

style.font.name = 'Times New Roman'  # 设 ascii + hAnsi
style.font.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')  # 设 eastAsia
```

**原因：** 只设 `font.name` 只影响 `w:ascii`，中文字符会回退到 Word 默认字体（通常是等线或 Calibri），导致同一行中英文字体不同。

## Rule 4: 一段一次写入

- 每个段落用一次 `add_paragraph` 调用写完全部文本
- NEVER 分多次向同一段落追加文本
- 如果段落内容很长，在调用前先拼接好完整字符串

**原因：** 每次写入操作创建一个新的 `<w:r>`（run）元素。多个 run 各自携带独立的 `<w:rPr>` 格式属性。即使它们应该相同，实践中经常出现差异，导致同一段落内字体不一致。

## Rule 5: TOC 必须用 field code 插入

- 通过 XML 插入 TOC 字段代码，不是纯文本
- 见 `skills/word-format/SKILL.md` 的 TOC 生成章节
- 插入后提醒用户在 Word 中按 Ctrl+A → F9 更新域

**TOC 字段 XML 示例：**
```xml
<w:p>
  <w:r>
    <w:fldChar w:fldCharType="begin"/>
  </w:r>
  <w:r>
    <w:instrText> TOC \o "1-3" \h \z \u </w:instrText>
  </w:r>
  <w:r>
    <w:fldChar w:fldCharType="separate"/>
  </w:r>
  <w:r>
    <w:t>（请在 Word 中更新目录）</w:t>
  </w:r>
  <w:r>
    <w:fldChar w:fldCharType="end"/>
  </w:r>
</w:p>
```

**原因：** 纯文本目录无法响应 Word 的"更新目录"功能。只有 field code 方式插入的 TOC 才能自动更新页码和标题。

---

## Anti-Patterns 速查

| ❌ 错误做法 | ✅ 正确做法 | 后果 |
|------------|-----------|------|
| `format_text(p, font="宋体")` 只设中文 | 修改样式同时设 eastAsia + ascii | 英文变 Calibri |
| `create_custom_style("MyH1")` 做标题 | 修改内置 `Heading 1` 样式 | TOC 失效 |
| 分两次 `add_paragraph` 写同一段 | 拼好字符串一次写入 | 字体不一致 |
| 先 `add_paragraph` 再 `format_text` | 写入时指定 `style_name` | 直接格式覆盖样式 |
| 用 `add_paragraph` 写目录文本 | 用 XML field code 插入 TOC | 目录无法更新 |
