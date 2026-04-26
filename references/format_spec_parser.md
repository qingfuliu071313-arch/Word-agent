# Format Spec Parser Reference

## 概述

本文档指导 word-formatter 如何从用户提供的各种格式要求文档中提取结构化格式规则（Format Spec）。

## 输入格式

用户可能以以下形式提供格式要求：

| 输入形式 | 处理方式 |
|----------|---------|
| Word 文档 (.docx) | 用 word-reader 读取，提取文字内容中的格式规则 |
| PDF 文件 | 用 Read 工具读取，提取格式规则 |
| 口头描述 | 直接从对话中提取 |
| 网页链接 | 用 WebFetch 获取内容后提取 |
| 图片/截图 | 用 Read 工具查看，识别格式规则 |

## 输出格式：Format Spec

将提取的规则统一为以下结构化格式，存储在对话上下文中供 word-formatter 使用：

```yaml
format_spec:
  # 页面设置
  page:
    size: A4 | Letter | B5 | ...
    orientation: portrait | landscape
    margins:
      top: "2.54cm"
      bottom: "2.54cm"
      left: "3.17cm"
      right: "3.17cm"

  # 字体规则
  fonts:
    body:
      chinese: "宋体"
      english: "Times New Roman"
      size: "小四"         # 中文字号 或 12pt
    title:
      chinese: "黑体"
      english: "Arial"
      size: "三号"
      bold: true
    heading_1:
      chinese: "黑体"
      english: "Arial"
      size: "四号"
      bold: true
    heading_2:
      chinese: "黑体"
      english: "Arial"
      size: "小四"
      bold: true
    heading_3:
      chinese: "楷体"
      english: "Arial"
      size: "小四"
      bold: false
    caption:               # 图题/表题
      chinese: "宋体"
      english: "Times New Roman"
      size: "五号"
    reference:             # 参考文献
      chinese: "宋体"
      english: "Times New Roman"
      size: "五号"

  # 行距与段距
  spacing:
    body:
      line_spacing: 1.5    # 倍数 或 "固定值20磅"
      before: "0pt"
      after: "0pt"
    heading_1:
      before: "24pt"
      after: "12pt"
    heading_2:
      before: "12pt"
      after: "6pt"
    reference:
      line_spacing: 1.0

  # 编号格式
  numbering:
    headings: ["1", "1.1", "1.1.1"]  # 或 ["一、", "（一）", "1."]
    figures: "图{n}"       # 或 "Figure {n}" 或 "Fig. {n}"
    tables: "表{n}"        # 或 "Table {n}"
    equations: "({n})"     # 公式编号

  # 页眉页脚
  header:
    content: "论文简短标题"  # 或 null 表示无页眉
    font: "宋体, 五号"
    position: "center"     # left | center | right
  footer:
    page_number: true
    position: "center"
    start_from: "正文"     # "第一页" | "正文" | "摘要后"
    format: "arabic"       # arabic | roman

  # 特殊要求
  special:
    - "首行缩进2字符"
    - "图片分辨率不低于300dpi"
    - "参考文献按出现顺序编号"
    - "英文摘要单独一页"
    # ... 其他自由文本规则
```

## 中文字号对照表

| 中文字号 | 磅值 (pt) | 说明 |
|----------|----------|------|
| 初号 | 42pt | — |
| 小初 | 36pt | — |
| 一号 | 26pt | — |
| 小一 | 24pt | — |
| 二号 | 22pt | — |
| 小二 | 18pt | — |
| 三号 | 16pt | 常用于论文标题 |
| 小三 | 15pt | — |
| 四号 | 14pt | 常用于一级标题 |
| 小四 | 12pt | 常用于正文 |
| 五号 | 10.5pt | 常用于图表题注、参考文献 |
| 小五 | 9pt | — |
| 六号 | 7.5pt | — |

## 提取策略

### 从 Word/PDF 格式要求文档提取

1. 读取文档全文
2. 搜索关键词定位格式规则段落：
   - 页面设置相关：`页边距|页面大小|纸张|margin|page size`
   - 字体相关：`字体|字号|font|size|宋体|黑体|Times`
   - 行距相关：`行距|行间距|段距|spacing|line`
   - 编号相关：`编号|图|表|公式|numbering|figure|table`
   - 页眉页脚：`页眉|页脚|页码|header|footer|page number`
3. 对每个匹配区域提取具体数值
4. 组装为 Format Spec 结构
5. 对未明确指定的字段标记为 `null`（不做修改）

### 从口头描述提取

1. 识别用户描述中的格式关键词
2. 对模糊描述要求确认：
   - "字体用宋体" → 确认是正文还是全部？英文用什么字体？
   - "行距 1.5" → 确认是倍数还是固定值？
3. 对未提到的字段标记为 `null`

### 缺失规则处理

当格式要求文档中某些规则未提及时：

| 策略 | 适用场景 |
|------|---------|
| 标记为 `null`，保持原样 | 默认策略 — 不改动未指定的格式 |
| 询问用户 | 关键格式（字体、字号、行距）缺失时 |
| 永不猜测 | 不基于"常见做法"填充默认值 |

## 验证步骤

Format Spec 提取完成后，向用户展示并确认：

```
我从您的格式要求中提取了以下规则：

页面：A4，页边距上下2.54cm 左右3.17cm
正文：宋体/Times New Roman，小四，1.5倍行距
一级标题：黑体/Arial，四号，加粗
...

以下未在要求中提及，将保持文档原样：
- 页眉内容
- 公式编号格式

请确认是否正确，或需要补充修改？
```

用户确认后，Format Spec 生效，word-formatter 据此执行格式修改。
