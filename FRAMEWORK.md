# Word-Agent 框架设计文档

> 最终版 — 2026-04-26

## 一、问题与目标

**痛点：** 使用 Claude Code 编辑学术论文 Word 文档时，格式修改不准确、反复试错消耗大量时间和 Token。

**根因：**
1. 现有 docx skill 是底层 XML 操作手册，缺乏学术论文领域知识
2. Word MCP Server 60+ 工具过于碎片，无高层工作流串联
3. 无文档结构缓存，每次操作重复读取全文
4. 缺少"格式要求解析 → 规则执行 → 合规验证"的闭环

**目标：** 建立一个专门处理学术论文 Word 文档的 Agent 插件，做到：
- 一次读取、结构缓存、多次复用（省 Token）
- 解析用户提供的格式要求 → 自动执行（准确）
- 工具路由自动选最优路径（高效）
- 覆盖论文写作全生命周期（完整）

---

## 二、架构：三层 WHO / HOW / WHAT

```
word-agent/
├── .claude-plugin/              # 插件清单
│   ├── plugin.json
│   └── marketplace.json
├── CLAUDE.md                    # 开发者文档
├── docs/
│   └── CLAUDE.md                # 运行时路由指令
├── agents/                      # WHO — 9 个专业 Agent
│   ├── word-orchestrator.md        # 总协调器
│   ├── word-reader.md              # 文档分析 + 文档对比
│   ├── word-formatter.md           # 格式化 + 目录 + 中英文混排
│   ├── word-content-editor.md      # 内容编辑
│   ├── word-reference.md           # 参考文献管理
│   ├── word-table-figure.md        # 表格与图片
│   ├── word-reviewer.md            # 审阅修订
│   ├── word-checker.md             # 交叉引用 + 格式合规验证
│   └── word-submit.md              # 投稿清理 + 文档拆分合并
├── skills/                      # HOW — 对应工作流
│   ├── word-orchestrate/
│   ├── word-read/
│   ├── word-format/
│   ├── word-edit/
│   ├── word-reference/
│   ├── word-table-figure/
│   ├── word-review/
│   ├── word-check/
│   └── word-submit/
└── references/                  # WHAT — 领域知识
    ├── academic_formatting.md      # 学术论文通用格式规范
    ├── chinese_standards.md        # 中文论文 GB 标准 + 中英混排规则
    ├── format_spec_parser.md       # 如何从用户格式要求文档提取规则
    ├── tool_routing.md             # MCP 工具路由决策树
    ├── token_budget.md             # Token 节约策略
    ├── common_fixes.md             # 高频格式问题修复方案
    ├── cross_ref_rules.md          # 交叉引用检查规则
    └── submission_checklist.md     # 投稿前检查清单
```

---

## 三、九大模块详细设计

---

### 模块 1: word-orchestrator（总协调器）

**职责：** 任务路由、Token 预算管理、工作流编排

**模型：** sonnet（路由不需要 opus）

**工具：** Agent 工具 + 所有模块工具的读取子集

**路由决策树：**

```
用户请求
  │
  ├─ "读取 / 分析 / 看看这个文档"
  │   └→ word-reader
  │
  ├─ "对比两个文档 / 这两版有什么区别"
  │   └→ word-reader（对比模式）
  │
  ├─ "排版 / 格式化 / 按这个要求改格式"
  │   └→ word-reader → word-formatter
  │
  ├─ "生成目录 / 插入目录"
  │   └→ word-reader → word-formatter（目录模式）
  │
  ├─ "修改内容 / 改写段落 / 把XX改成YY"
  │   └→ word-reader → word-content-editor
  │
  ├─ "参考文献 / 引用 / 脚注 / Zotero"
  │   └→ word-reference
  │
  ├─ "表格 / 图片 / 三线表 / 插入图片"
  │   └→ word-table-figure
  │
  ├─ "审稿意见 / 修改 / revision / 修订"
  │   └→ word-reader → word-reviewer
  │
  ├─ "检查交叉引用 / 检查格式是否符合要求"
  │   └→ word-reader → word-checker
  │
  ├─ "投稿准备 / 清理文档 / 拆分补充材料"
  │   └→ word-reader → word-submit
  │
  └─ 复合任务（如"排版+检查+清理"）
      └→ word-reader → 串行/并行分发多模块
```

**Token 预算策略（核心）：**

| 原则 | 做法 |
|------|------|
| 懒加载 | 用 `get_document_outline` 代替 `get_document_text`（节省 80%+ token） |
| 结构缓存 | Document Map 生成后缓存，后续模块直接引用 |
| 工具优先级 | MCP 单步 > MCP 多步 > XML unpack/edit/pack |
| 批量操作 | 同类修改合并执行，不逐段落调用 |
| 按需精读 | 只读取需要修改的段落范围 |

---

### 模块 2: word-reader（文档分析 + 文档对比）

**职责：** 读取文档生成结构化"文档地图"；对比两版文档差异

**模型：** sonnet

**MCP 工具：**
```
mcp__word-document-server__get_document_info
mcp__word-document-server__get_document_outline
mcp__word-document-server__get_document_text
mcp__word-document-server__get_paragraph_text_from_document
mcp__word-document-server__find_text_in_document
mcp__word-document-server__get_all_comments
mcp__word-document-server__get_comments_by_author
mcp__word-document-server__validate_document_footnotes
mcp__word-document-server__get_document_xml
```

**Python 工具：**
```
docx2python  — 一次性解析文档完整层次结构（标题、段落、表格、脚注、图片）
redlines     — 文档对比时生成红线差异标记
deepdiff     — 文档对比时进行深层结构对比
```

**功能 A — 文档地图（Document Map）：**

使用 `docx2python` 一次性解析文档，生成结构化摘要供所有下游模块复用：

```markdown
## Document Map: paper_v3.docx
- 总页数: 28, 总段落: 187, 语言: 中文
- 当前字体: 正文=宋体/Times New Roman, 标题=黑体
- 当前行距: 正文=1.5倍, 参考文献=单倍
- 结构:
  - [P1-P3] 标题+作者信息
  - [P4-P15] 摘要（中+英）
  - [P16-P42] 1. 引言
  - [P43-P89] 2. 材料与方法
    - [P43-P55] 2.1 研究区域
    - [P56-P72] 2.2 数据采集
    - [P73-P89] 2.3 数据分析
  - [P90-P132] 3. 结果
  - [P133-P165] 4. 讨论
  - [P166-P187] 参考文献
- 表格: Table 1 (P95), Table 2 (P108), Table 3 (P120)
- 图片: Fig 1 (P92), Fig 2 (P105), Fig 3 (P115)
- 脚注/尾注: 12 个
- 批注: 5 条（作者: 导师3条, 合作者2条）
- 格式问题检测:
  - ⚠ P23: 字体不一致（混用 Arial）
  - ⚠ P90-P132: 行距不统一
  - ⚠ 参考文献: 编号格式不统一
```

**功能 B — 文档对比：**

使用 `docx2python` 提取两个文档结构 → `deepdiff` 比较结构差异 → `redlines` 生成可视化标记：

```markdown
## Diff Report: paper_v2.docx → paper_v3.docx

### 结构变化
- 新增: 2.3.1 敏感性分析 (P80-P85)
- 删除: 无

### 内容变化 (共 23 处)
- [P18] 引言第3段: 增加了2句话关于XX的最新研究
- [P65] 方法: "采样频率从每月改为每两周"
- [P95] Table 1: 新增一列 "效应量"
- [P140] 讨论: 第2段完全重写
  - 原文: "..."（前50字）
  - 新文: "..."（前50字）
- ... (更多变化)

### 格式变化
- 全文行距: 1.5倍 → 双倍
- 参考文献格式: 未变
```

---

### 模块 3: word-formatter（格式化 + 目录 + 中英文混排）

**职责：** 解析格式要求 → 批量应用格式 → 生成目录 → 中英文混排规范化

**模型：** sonnet

**MCP 工具：**
```
mcp__word-document-server__format_text
mcp__word-document-server__create_custom_style
mcp__word-document-server__search_and_replace
mcp__word-document-server__set_table_width
mcp__word-document-server__set_table_column_widths
mcp__word-document-server__format_table
mcp__word-document-server__highlight_table_header
mcp__word-document-server__add_heading
mcp__word-document-server__add_page_break
```
**额外工具：** Read, Bash（用于 XML 操作插入 TOC 域代码）

**功能 A — 格式要求解析与执行：**

```
用户提供格式要求文档
      │
      ▼
  解析为 Format Spec（结构化规则）
      │
      ▼
  ┌─────────────────────────────────────┐
  │ Format Spec 示例                     │
  │                                     │
  │ page:                               │
  │   size: A4                          │
  │   margins: 上2.54 下2.54 左3.17 右3.17│
  │ fonts:                              │
  │   正文: 宋体+Times New Roman, 小四    │
  │   一级标题: 黑体, 三号, 加粗          │
  │   二级标题: 黑体, 四号, 加粗          │
  │   图题/表题: 宋体, 五号              │
  │ spacing:                            │
  │   正文行距: 1.5倍                    │
  │   段前段后: 0                        │
  │ numbering:                          │
  │   标题: "1", "1.1", "1.1.1"         │
  │   图: "图1", "图2"                   │
  │   表: "表1", "表2"                   │
  │ other:                              │
  │   页码: 底部居中, 从正文开始           │
  │   页眉: 论文简短标题                  │
  └─────────────────────────────────────┘
      │
      ▼
  对比 Document Map 现有格式 → 生成差异清单
      │
      ▼
  用户确认 → 批量执行修改
```

**功能 B — 目录生成：**

1. 从 Document Map 获取标题结构
2. 在指定位置插入 TOC 域代码（通过 XML 操作或 MCP 工具）
3. 提示用户在 Word 中"更新域"刷新页码

支持设置目录层级深度（如只显示到二级标题）。

**功能 C — 中英文混排规范化：**

自动检测并修正：

| 规则 | 示例 |
|------|------|
| 中英文之间加空格 | `研究area` → `研究 area` |
| 英文术语使用半角标点 | `DNA，RNA` → `DNA, RNA`（英文语境） |
| 中文语境使用全角标点 | `包括soil,water` → `包括soil、water` |
| 英文括号前后空格 | `方法(method)` → `方法 (method)` |
| 数字与单位间空格 | `30cm` → `30 cm` |
| 数字与中文间不加空格 | `第 3 组` → `第3组` |

---

### 模块 4: word-content-editor（内容编辑）

**职责：** 在不破坏格式的前提下修改文档内容

**模型：** opus（内容修改需要强推理能力）

**MCP 工具：**
```
mcp__word-document-server__search_and_replace
mcp__word-document-server__insert_line_or_paragraph_near_text
mcp__word-document-server__replace_paragraph_block_below_header
mcp__word-document-server__replace_block_between_manual_anchors
mcp__word-document-server__delete_paragraph
mcp__word-document-server__add_paragraph
mcp__word-document-server__insert_header_near_text
mcp__word-document-server__insert_numbered_list_near_text
```
**Python 工具：**
```
python-docx-replace  — 跨 XML run 文本查找替换（MCP search_and_replace 失败时的备选方案）
```
**额外工具：** Read, Write, Edit, Bash（仅 tracked changes 模式使用）

**编辑模式自动选择：**

| 模式 | 场景 | 工具路径 |
|------|------|---------|
| 精确替换 | 改几个词/句子 | `search_and_replace` — 1次调用 |
| 段落重写 | 重写某一整段 | `replace_paragraph_block_below_header` — 1次调用 |
| 区块替换 | 替换两个锚点间的内容 | `replace_block_between_manual_anchors` — 1次调用 |
| 插入 | 加新段落/小节 | `insert_line_or_paragraph_near_text` — 1次调用 |
| 修订模式 | 需保留修改痕迹 | XML unpack → tracked changes → repack（3+次调用） |

**选择原则：** 能用 MCP 一步完成的，绝不走 XML。只有"需要 tracked changes"时才启用 XML 模式。

---

### 模块 5: word-reference（参考文献管理）

**职责：** 参考文献格式化、插入、校验，Zotero 集成

**模型：** sonnet

**MCP 工具：**
```
# Zotero
mcp__zotero__search_items
mcp__zotero__get_item
mcp__zotero__export_bibliography
mcp__zotero__get_item_children

# Word
mcp__word-document-server__add_footnote_to_document
mcp__word-document-server__add_endnote_to_document
mcp__word-document-server__delete_footnote_from_document
mcp__word-document-server__validate_document_footnotes
mcp__word-document-server__search_and_replace

# 学术搜索
mcp__semantic-scholar__search_papers
mcp__semantic-scholar__get_paper
```

**Python 工具：**
```
citeproc-py  — CSL 引用格式化引擎，配合 GB/T 7714 等 CSL 样式文件自动生成规范参考文献
```

**功能：**
- 从 Zotero 检索 → 通过 `citeproc-py` + CSL 样式按格式要求生成参考文献列表并写入文档
- 检查文内引用 `[1]` / `(Author, Year)` 与参考文献列表的一致性
- 脚注/尾注的批量添加、删除、格式化
- 检测缺失引用（正文提到但参考文献没有）和孤立引用（参考文献有但正文没引用）

---

### 模块 6: word-table-figure（表格与图片）

**职责：** 按学术规范创建和格式化表格与图片

**模型：** sonnet

**MCP 工具：**
```
mcp__word-document-server__add_table
mcp__word-document-server__format_table
mcp__word-document-server__format_table_cell_text
mcp__word-document-server__set_table_width
mcp__word-document-server__set_table_column_widths
mcp__word-document-server__auto_fit_table_columns
mcp__word-document-server__merge_table_cells_horizontal
mcp__word-document-server__merge_table_cells_vertical
mcp__word-document-server__highlight_table_header
mcp__word-document-server__set_table_cell_shading
mcp__word-document-server__set_table_cell_alignment
mcp__word-document-server__set_table_cell_padding
mcp__word-document-server__set_table_alignment_all
mcp__word-document-server__apply_table_alternating_rows
mcp__word-document-server__add_picture
```

**功能：**
- **三线表**一键生成（顶线+栏目线+底线，学术论文最常用）
- 表格自适应列宽 / 手动设置列宽
- 单元格合并（横向/纵向）
- 图片插入 + 题注格式化
- 表格/图片编号自动管理

---

### 模块 7: word-reviewer（审阅修订）

**职责：** 处理审稿意见，管理修订，生成修改说明

**模型：** opus（需要理解审稿意见并制定修改策略）

**MCP 工具：**
```
mcp__word-document-server__get_all_comments
mcp__word-document-server__get_comments_by_author
mcp__word-document-server__get_comments_for_paragraph
```
**额外工具：** Read, Write, Edit, Bash（tracked changes 必须走 XML）

**工作流：**

```
审稿意见（文档/邮件/批注）
      │
      ▼
  逐条提取并分类
      │
      ├─ 格式问题 → word-formatter
      ├─ 内容修改 → word-content-editor
      ├─ 参考文献 → word-reference
      ├─ 表格/图片 → word-table-figure
      └─ 纯回复（无需改文档）→ 直接写入回复
      │
      ▼
  生成修改方案 → 用户逐条确认
      │
      ▼
  执行修改（tracked changes 模式）
      │
      ▼
  生成 Point-by-Point Response 文档
```

**输出：**
- 修改后的论文文档（带修订痕迹）
- Point-by-Point Response 文档（逐条回复审稿人）

---

### 模块 8: word-checker（交叉引用检查 + 格式合规验证）

**职责：** 文档质量检查，发现问题但不自动修复

**模型：** sonnet

**MCP 工具：**
```
mcp__word-document-server__get_document_text
mcp__word-document-server__get_document_outline
mcp__word-document-server__find_text_in_document
mcp__word-document-server__get_document_info
mcp__word-document-server__validate_document_footnotes
mcp__word-document-server__get_all_comments
```

**功能 A — 交叉引用检查：**

扫描全文，检测：

| 检查项 | 示例 |
|--------|------|
| 引用不存在的表格/图片 | 正文写"见表4"，但只有表1-3 |
| 编号不连续 | 图1, 图2, 图4（缺图3） |
| 编号重复 | 两个"表2" |
| 首次引用顺序 | 正文先提到图3再提到图1 |
| 引用格式不统一 | 混用"图1"和"Figure 1" |
| 未被引用的表格/图片 | 表3存在但正文从未提及 |

输出检查报告，列出所有问题及位置。

**功能 B — 格式合规验证：**

输入：格式要求文档（同 word-formatter 使用的）+ 当前文档
输出：逐条合规检查报告

```markdown
## 格式合规报告: paper_final.docx

### ✅ 通过 (15/20)
- ✅ 页面大小: A4
- ✅ 页边距: 上2.54 下2.54 左3.17 右3.17
- ✅ 正文字体: 宋体+Times New Roman
- ✅ 正文字号: 小四
- ✅ 一级标题: 黑体, 三号, 加粗
- ...

### ❌ 未通过 (5/20)
- ❌ 行距: 要求1.5倍, 实际部分段落为双倍 → P90, P91, P95
- ❌ 图题字号: 要求五号, 实际为小四 → Fig 1, Fig 3
- ❌ 页码: 要求从正文开始编号, 实际从第一页开始
- ❌ 参考文献行距: 要求单倍, 实际为1.5倍
- ❌ 缺少页眉
```

---

### 模块 9: word-submit（投稿清理 + 文档拆分合并）

**职责：** 投稿前最终准备，生成 clean copy，拆分/合并文档

**模型：** sonnet

**MCP 工具：**
```
mcp__word-document-server__get_all_comments
mcp__word-document-server__get_document_info
mcp__word-document-server__copy_document
mcp__word-document-server__protect_document
mcp__word-document-server__unprotect_document
```
**Python 工具：**
```
docxcompose  — 合并多个 .docx 文件，自动处理样式冲突、编号续接、页眉页脚
```
**额外工具：** Read, Write, Edit, Bash（XML 操作用于接受修订、删除批注、清理元数据）

**功能 A — 投稿前清理（Clean Copy Pipeline）：**

```
原始文档（带批注、修订、元数据）
      │
      ▼
  1. 复制文档为工作副本
  2. 接受所有修订痕迹
  3. 删除所有批注
  4. 清除个人元数据（作者、单位、修改历史）
  5. 检查图片分辨率是否满足要求（300dpi+）
  6. 运行 word-checker 做最终检查
  7. 输出 clean copy
      │
      ▼
  paper_clean.docx（投稿就绪）
```

**功能 B — 文档拆分：**

将一个文档拆分为多个文件：
- 正文 + 补充材料分离
- 按章节拆分
- 图片/表格单独提取为独立文件

**功能 C — 文档合并：**

将多个文档合并为一个：
- 多位合作者各写一部分 → 合并为完整论文
- 合并时统一格式（调用 word-formatter）
- 重新编号图表和参考文献

---

## 四、系统集成架构

```
┌─────────────────────────────────────────────────┐
│                   用户请求                        │
└──────────────────────┬──────────────────────────┘
                       │
              ┌────────▼────────┐
              │ word-orchestrator│  ← 路由 + Token 预算
              └────────┬────────┘
                       │
         ┌─────────────▼─────────────┐
         │       word-reader          │  ← 一次读取，生成 Document Map
         └─────────────┬─────────────┘
                       │ Document Map（缓存复用）
          ┌────────────┼────────────┬──────────┐
          ▼            ▼            ▼          ▼
   ┌───────────┐ ┌──────────┐ ┌────────┐ ┌────────┐
   │ formatter │ │ editor   │ │ ref    │ │ table/ │
   │           │ │          │ │ mgr    │ │ figure │
   └─────┬─────┘ └────┬─────┘ └───┬────┘ └───┬────┘
         │            │           │           │
         ▼            ▼           ▼           ▼
   ┌───────────┐ ┌──────────┐
   │ checker   │ │ reviewer │
   │ (验证)    │ │ (审阅)   │
   └─────┬─────┘ └────┬─────┘
         │            │
         ▼            ▼
   ┌─────────────────────┐
   │    word-submit       │  ← 投稿清理 + 拆分合并
   └─────────────────────┘
         │
         ▼
   ┌─────────────────────────────────────────┐
   │   底层工具自动路由                         │
   │  ┌──────────────┐  ┌──────────────────┐ │
   │  │ Word MCP     │  │ docx skill       │ │
   │  │ Server       │  │ (XML unpack/pack)│ │
   │  │ 优先使用 ✓    │  │ 仅tracked changes│ │
   │  └──────────────┘  └──────────────────┘ │
   └─────────────────────────────────────────┘
```

---

## 五、工具路由策略

**核心原则：能用 MCP 一步完成的，绝不走 XML。**

| 操作 | 首选路径 | 备选路径 |
|------|---------|---------|
| 读取文档结构 | `get_document_outline` | `docx2python`（需完整层次时） |
| 生成 Document Map | `docx2python` 一次性解析 | MCP 多次调用（逐段读取） |
| 查找文本 | `find_text_in_document` | `get_document_text` + 搜索 |
| 替换文字 | `search_and_replace` | `python-docx-replace`（跨 run 时） |
| 修改格式 | `format_text` | XML 编辑 |
| 添加段落 | `insert_line_or_paragraph_near_text` | XML 编辑 |
| 替换段落块 | `replace_paragraph_block_below_header` | XML 编辑 |
| 添加表格 | `add_table` | docx-js 创建 |
| 添加图片 | `add_picture` | XML + relationship 编辑 |
| 修订痕迹 | **必须** XML unpack/edit/pack | 无 MCP 替代 |
| 接受所有修订 | `accept_changes.py` 脚本 | 手动 XML 清理 |
| 添加批注 | `comment.py` + XML 标记 | 无 MCP 替代 |
| 生成目录 | XML 插入 TOC 域代码 | docx-js `TableOfContents` |
| 文档对比 | `docx2python` + `deepdiff` + `redlines` | 逐段 MCP 读取 + 文本 diff |
| 参考文献格式化 | `citeproc-py` + CSL 样式 | 手动格式化写入 |
| 文档合并 | `docxcompose` | 手动 XML 合并 |
| 转 PDF | `soffice.py --convert-to pdf` | — |

---

## 六、与现有系统的协作

### 与 eco-agent 的协作

```
eco-agent（写什么 — 学术内容）    word-agent（怎么呈现 — Word格式）
─────────────────────────────    ─────────────────────────────
ecology-paper-writing            word-content-editor
  → 输出 Markdown/文本草稿   ──→   → 写入 Word 文档

ecology-review                   word-reviewer
  → 审稿策略 + 回复草稿     ──→   → 执行文档修改 + 生成回复文档

ecology-polish                   word-formatter
  → 语言润色后的文本        ──→   → 应用到 Word 并保持格式

ecology-data-analysis            word-table-figure
  → 统计结果 + 表格数据     ──→   → 生成学术规范的三线表
```

### 与 docx skill 的关系

docx skill 作为**底层工具层**保留，word-agent 在需要 XML 操作时调用它的脚本（`unpack.py`, `pack.py`, `comment.py`, `accept_changes.py`）。word-agent 不替代 docx skill，而是在其之上提供高层工作流。

### 与 Zotero MCP 的集成

word-reference 模块直接调用 Zotero MCP 工具，实现从文献库到参考文献列表的自动化。

---

## 七、实施路线图

| 阶段 | 模块 | 理由 | 预计工作量 |
|------|------|------|-----------|
| **P0** | word-orchestrator + word-reader | 基础设施，所有模块依赖 | 中 |
| **P1** | word-formatter + word-checker | 解决核心痛点"格式改不对"，checker 提供验证闭环 | 大 |
| **P2** | word-content-editor | 内容编辑是高频操作 | 中 |
| **P3** | word-table-figure + word-reference | 表格和参考文献是常见需求 | 中 |
| **P4** | word-reviewer + word-submit | 审稿修改和投稿清理 | 中 |

每个阶段完成后可独立使用，不需要等全部模块完成。

---

## 八、外部依赖

### Python 库

| 库 | 版本 | 用途 | 对应模块 |
|---|------|------|---------|
| **docx2python** | `pip install docx2python` | 将 .docx 解析为嵌套 Python 列表，保留完整文档层次结构（标题、段落、表格、脚注、图片）。比逐段调用 MCP 更高效地生成 Document Map | word-reader |
| **docxcompose** | `pip install docxcompose` | 合并多个 .docx 为一个文件，自动处理样式冲突、编号续接、页眉页脚合并 | word-submit |
| **citeproc-py** | `pip install citeproc-py` | CSL 引用格式化引擎，配合 GB/T 7714 等 CSL 样式文件，自动生成符合规范的参考文献列表 | word-reference |
| **python-docx-replace** | `pip install python-docx-replace` | 跨 XML run 的文本查找替换，解决 python-docx/MCP 在跨 run 文本匹配时的已知限制 | word-content-editor（备选） |
| **redlines** | `pip install redlines` | 文本对比并生成红线标记（类似 tracked changes 的可视化差异） | word-reader（对比模式） |
| **deepdiff** | `pip install deepdiff` | 深层结构对比，用于比较两个文档的嵌套结构差异 | word-reader（对比模式） |

### CSL 样式文件

| 文件 | 来源 | 用途 |
|------|------|------|
| `gb-t-7714-2015-numeric.csl` | [citation-style-language/styles](https://github.com/citation-style-language/styles) | 中文论文 GB/T 7714-2015 数字编号格式 |
| `gb-t-7714-2015-author-date.csl` | 同上 | 中文论文 GB/T 7714-2015 作者-年份格式 |
| 其他期刊 CSL | 同上 | 按用户提供的格式要求选用 |

CSL 样式文件存放于 `references/csl/` 目录，按需下载。

### 已有依赖（保留）

| 工具 | 用途 |
|------|------|
| Word MCP Server | 60+ 工具，主要文档操作路径 |
| docx skill 脚本 | `unpack.py`, `pack.py`, `comment.py`, `accept_changes.py` — XML 操作 |
| Zotero MCP | 文献库检索与导出 |
| Semantic Scholar MCP | 学术论文搜索与元数据获取 |
| pandoc | 文本提取与格式转换 |
| LibreOffice (soffice) | PDF 转换、域更新 |

### 依赖安装

```bash
# 核心依赖（P0 阶段即需要）
pip install docx2python deepdiff redlines

# P1 阶段追加
pip install python-docx-replace

# P3 阶段追加
pip install docxcompose citeproc-py

# CSL 样式文件（按需下载）
curl -o references/csl/gb-t-7714-2015-numeric.csl \
  https://raw.githubusercontent.com/citation-style-language/styles/master/gb-t-7714-2015-numeric.csl
```

---

## 九、文件结构完整清单（更新后）

```
word-agent/
├── .claude-plugin/
│   ├── plugin.json
│   └── marketplace.json
├── CLAUDE.md
├── FRAMEWORK.md                    # 本文档
├── docs/
│   └── CLAUDE.md
├── agents/
│   ├── word-orchestrator.md
│   ├── word-reader.md
│   ├── word-formatter.md
│   ├── word-content-editor.md
│   ├── word-reference.md
│   ├── word-table-figure.md
│   ├── word-reviewer.md
│   ├── word-checker.md
│   └── word-submit.md
├── skills/
│   ├── word-orchestrate/
│   │   └── SKILL.md
│   ├── word-read/
│   │   ├── SKILL.md
│   │   └── references/
│   ├── word-format/
│   │   ├── SKILL.md
│   │   └── references/
│   ├── word-edit/
│   │   ├── SKILL.md
│   │   └── references/
│   ├── word-reference/
│   │   ├── SKILL.md
│   │   └── references/
│   ├── word-table-figure/
│   │   ├── SKILL.md
│   │   └── references/
│   ├── word-review/
│   │   ├── SKILL.md
│   │   └── references/
│   ├── word-check/
│   │   ├── SKILL.md
│   │   └── references/
│   └── word-submit/
│       ├── SKILL.md
│       └── references/
├── references/
│   ├── academic_formatting.md
│   ├── chinese_standards.md
│   ├── format_spec_parser.md
│   ├── tool_routing.md
│   ├── token_budget.md
│   ├── common_fixes.md
│   ├── cross_ref_rules.md
│   ├── submission_checklist.md
│   └── csl/
│       ├── gb-t-7714-2015-numeric.csl
│       └── gb-t-7714-2015-author-date.csl
└── scripts/
    ├── requirements.txt            # Python 依赖清单
    └── setup.sh                    # 一键安装脚本
```
