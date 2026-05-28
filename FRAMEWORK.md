# Word-Agent Framework Design / 框架设计文档

> Final version / 最终版 — 2026-04-26

## 1. Problem & Goals / 问题与目标

**Pain point / 痛点：** When using Claude Code to edit academic paper Word documents, formatting changes are inaccurate and repeated trial-and-error wastes significant time and tokens.

使用 Claude Code 编辑学术论文 Word 文档时，格式修改不准确、反复试错消耗大量时间和 Token。

**Root causes / 根因：**
1. The existing docx skill is a low-level XML operations manual, lacking academic paper domain knowledge / 现有 docx skill 是底层 XML 操作手册，缺乏学术论文领域知识
2. Word MCP Server's 60+ tools are too fragmented, with no high-level workflow orchestration / Word MCP Server 60+ 工具过于碎片，无高层工作流串联
3. No document structure caching — every operation re-reads the full document / 无文档结构缓存，每次操作重复读取全文
4. No closed loop from "parse format requirements → execute rules → verify compliance" / 缺少"格式要求解析 → 规则执行 → 合规验证"的闭环

**Goals / 目标：** Build a specialized agent plugin for academic paper Word documents:

建立一个专门处理学术论文 Word 文档的 Agent 插件，做到：

- Read once, cache structure, reuse across modules (save tokens) / 一次读取、结构缓存、多次复用（省 Token）
- Parse user-provided format requirements → auto-execute (accurate) / 解析用户提供的格式要求 → 自动执行（准确）
- Tool routing auto-selects optimal path (efficient) / 工具路由自动选最优路径（高效）
- Cover the full manuscript lifecycle (complete) / 覆盖论文写作全生命周期（完整）

---

## 2. Architecture: Three-Layer WHO / HOW / WHAT / 架构：三层设计

```
word-agent/
├── .claude-plugin/              # Plugin manifests / 插件清单
│   ├── plugin.json
│   └── marketplace.json
├── CLAUDE.md                    # Developer docs / 开发者文档
├── docs/
│   └── CLAUDE.md                # Runtime routing instructions / 运行时路由指令
├── agents/                      # WHO — 9 specialized agents / 9 个专业 Agent
│   ├── word-orchestrator.md        # Coordinator / 总协调器
│   ├── word-reader.md              # Document analysis + comparison / 文档分析 + 对比
│   ├── word-formatter.md           # Formatting + TOC + CJK / 格式化 + 目录 + 中英文混排
│   ├── word-content-editor.md      # Content editing / 内容编辑
│   ├── word-reference.md           # Reference management / 参考文献管理
│   ├── word-table-figure.md        # Tables & figures / 表格与图片
│   ├── word-reviewer.md            # Review & revision / 审阅修订
│   ├── word-checker.md             # Cross-refs + format compliance / 交叉引用 + 格式合规验证
│   └── word-submit.md              # Submission prep + split/merge / 投稿清理 + 拆分合并
├── skills/                      # HOW — Corresponding workflows / 对应工作流
│   ├── word-orchestrate/
│   ├── word-read/
│   ├── word-format/
│   ├── word-edit/
│   ├── word-reference/
│   ├── word-table-figure/
│   ├── word-review/
│   ├── word-check/
│   └── word-submit/
└── references/                  # WHAT — Domain knowledge / 领域知识
    ├── academic_formatting.md      # General academic formatting / 学术论文通用格式规范
    ├── chinese_standards.md        # Chinese paper GB standards + CJK rules / 中文论文 GB 标准 + 中英混排规则
    ├── format_spec_parser.md       # How to extract rules from format specs / 如何从格式要求文档提取规则
    ├── tool_routing.md             # MCP tool routing decision tree / MCP 工具路由决策树
    ├── token_budget.md             # Token-saving strategies / Token 节约策略
    ├── common_fixes.md             # Common formatting fixes / 高频格式问题修复方案
    ├── cross_ref_rules.md          # Cross-reference check rules / 交叉引用检查规则
    └── submission_checklist.md     # Pre-submission checklist / 投稿前检查清单
```

---

## 3. Nine Modules in Detail / 九大模块详细设计

---

### Module 1: word-orchestrator / 总协调器

**Role / 职责：** Task routing, token budget management, workflow orchestration / 任务路由、Token 预算管理、工作流编排

**Model / 模型：** sonnet (routing doesn't need opus / 路由不需要 opus)

**Tools / 工具：** Agent tool + read-only subset of all module tools

**Routing decision tree / 路由决策树：**

```
User request / 用户请求
  │
  ├─ "read / analyze / look at this doc" / "读取 / 分析 / 看看这个文档"
  │   └→ word-reader
  │
  ├─ "compare two docs / what changed" / "对比两个文档 / 这两版有什么区别"
  │   └→ word-reader (compare mode / 对比模式)
  │
  ├─ "format / apply styles" / "排版 / 格式化 / 按这个要求改格式"
  │   └→ word-reader → word-formatter
  │
  ├─ "generate TOC / insert TOC" / "生成目录 / 插入目录"
  │   └→ word-reader → word-formatter (TOC mode / 目录模式)
  │
  ├─ "edit content / rewrite / change XX to YY" / "修改内容 / 改写段落 / 把XX改成YY"
  │   └→ word-reader → word-content-editor
  │
  ├─ "references / citations / footnotes / Zotero" / "参考文献 / 引用 / 脚注 / Zotero"
  │   └→ word-reference
  │
  ├─ "tables / figures / three-line table / insert image" / "表格 / 图片 / 三线表 / 插入图片"
  │   └→ word-table-figure
  │
  ├─ "reviewer comments / revision" / "审稿意见 / 修改 / revision / 修订"
  │   └→ word-reader → word-reviewer
  │
  ├─ "check cross-refs / verify format compliance" / "检查交叉引用 / 检查格式是否符合要求"
  │   └→ word-reader → word-checker
  │
  ├─ "submission prep / clean document / split supplementary" / "投稿准备 / 清理文档 / 拆分补充材料"
  │   └→ word-reader → word-submit
  │
  └─ Compound tasks (e.g. "format + check + clean") / 复合任务（如"排版+检查+清理"）
      └→ word-reader → serial/parallel dispatch to multiple modules
```

**Token budget strategy (core) / Token 预算策略（核心）：**

| Principle / 原则 | Approach / 做法 |
|---------|----------|
| Lazy loading / 懒加载 | Use `get_document_outline` instead of `get_document_text` (saves 80%+ tokens) |
| Structure caching / 结构缓存 | Cache Document Map after generation; downstream modules reference directly |
| Tool priority / 工具优先级 | MCP single-step > MCP multi-step > XML unpack/edit/pack |
| Batch operations / 批量操作 | Merge similar changes into batch execution, not per-paragraph calls |
| On-demand reading / 按需精读 | Only read paragraph ranges that need modification |

---

### Module 2: word-reader / 文档分析 + 文档对比

**Role / 职责：** Read documents to generate structured "Document Maps"; compare two document versions / 读取文档生成结构化"文档地图"；对比两版文档差异

**Model / 模型：** sonnet

**MCP Tools:**
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

**Python Tools:**
```
docx2python  — Parse full document hierarchy in one pass (headings, paragraphs, tables, footnotes, images)
               一次性解析文档完整层次结构（标题、段落、表格、脚注、图片）
redlines     — Generate red-line diff markup for document comparison
               文档对比时生成红线差异标记
deepdiff     — Deep structural comparison for document diffing
               文档对比时进行深层结构对比
```

**Feature A — Document Map / 文档地图：**

Parse the document once with `docx2python`, generating a structured summary reused by all downstream modules:

使用 `docx2python` 一次性解析文档，生成结构化摘要供所有下游模块复用：

```markdown
## Document Map: paper_v3.docx
- Pages: 28, Paragraphs: 187, Language: Chinese
- Current fonts: body=SimSun/Times New Roman, headings=SimHei
- Current spacing: body=1.5x, references=single
- Structure:
  - [P1-P3] Title + author info
  - [P4-P15] Abstract (Chinese + English)
  - [P16-P42] 1. Introduction
  - [P43-P89] 2. Materials & Methods
    - [P43-P55] 2.1 Study area
    - [P56-P72] 2.2 Data collection
    - [P73-P89] 2.3 Data analysis
  - [P90-P132] 3. Results
  - [P133-P165] 4. Discussion
  - [P166-P187] References
- Tables: Table 1 (P95), Table 2 (P108), Table 3 (P120)
- Figures: Fig 1 (P92), Fig 2 (P105), Fig 3 (P115)
- Text Boxes: {count} (content extracted via XML parsing)
- Footnotes/Endnotes: 12
- Comments: 5 (by: supervisor ×3, collaborator ×2)
- Format issues detected:
  - ⚠ P23: Inconsistent font (mixed Arial)
  - ⚠ P90-P132: Inconsistent line spacing
  - ⚠ References: Inconsistent numbering format
```

**Feature B — Document Comparison / 文档对比：**

Extract structure from both documents with `docx2python` → compare with `deepdiff` → visualize with `redlines`:

使用 `docx2python` 提取两个文档结构 → `deepdiff` 比较结构差异 → `redlines` 生成可视化标记：

```markdown
## Diff Report: paper_v2.docx → paper_v3.docx

### Structural Changes / 结构变化
- Added: 2.3.1 Sensitivity Analysis (P80-P85)
- Removed: none

### Content Changes (23 total) / 内容变化 (共 23 处)
- [P18] Introduction para 3: Added 2 sentences about latest research on XX
- [P65] Methods: "sampling frequency changed from monthly to biweekly"
- [P95] Table 1: Added column "effect size"
- [P140] Discussion: Para 2 completely rewritten
  - Before: "..." (first 50 chars)
  - After: "..." (first 50 chars)
- ... (more changes)

### Formatting Changes / 格式变化
- Line spacing: 1.5x → double
- Reference format: unchanged
```

---

### Module 3: word-formatter / 格式化 + 目录 + 中英文混排

**Role / 职责：** Parse format requirements → batch apply formatting → generate TOC → CJK normalization / 解析格式要求 → 批量应用格式 → 生成目录 → 中英文混排规范化

**Model / 模型：** sonnet

**MCP Tools:**
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
**Additional tools:** Read, Bash (for XML operations to insert TOC field codes)

**Feature A — Format Spec Parsing & Execution / 格式要求解析与执行：**

```
User provides format requirement document
用户提供格式要求文档
      │
      ▼
  Parse into Format Spec (structured rules)
  解析为 Format Spec（结构化规则）
  ⚠ Must extract text box content — format templates
    often place instructions inside text boxes
    必须提取文本框内容 — 格式模板常将说明放在文本框中
      │
      ▼
  ┌─────────────────────────────────────┐
  │ Format Spec example / 示例           │
  │                                     │
  │ page:                               │
  │   size: A4                          │
  │   margins: T2.54 B2.54 L3.17 R3.17 │
  │ fonts:                              │
  │   body: SimSun+Times New Roman, 12pt│
  │   heading_1: SimHei, 14pt, bold     │
  │   heading_2: SimHei, 12pt, bold     │
  │   caption: SimSun, 10.5pt           │
  │ spacing:                            │
  │   body: 1.5x                        │
  │   before/after: 0                   │
  │ numbering:                          │
  │   headings: "1", "1.1", "1.1.1"    │
  │   figures: "Fig 1", "Fig 2"         │
  │   tables: "Table 1", "Table 2"     │
  │ other:                              │
  │   page numbers: bottom center       │
  │   header: short paper title         │
  └─────────────────────────────────────┘
      │
      ▼
  Compare against Document Map → generate diff list
  对比 Document Map 现有格式 → 生成差异清单
      │
      ▼
  User confirms → batch execute changes
  用户确认 → 批量执行修改
```

**Feature B — TOC Generation / 目录生成：**

1. Get heading structure from Document Map / 从 Document Map 获取标题结构
2. Insert TOC field code at specified position (via XML or MCP) / 在指定位置插入 TOC 域代码
3. Prompt user to "Update Fields" in Word to refresh page numbers / 提示用户在 Word 中"更新域"刷新页码

Supports configurable heading depth (e.g. show only down to level 2) / 支持设置目录层级深度。

**Feature C — CJK Mixed-Text Normalization / 中英文混排规范化：**

Auto-detect and fix / 自动检测并修正：

| Rule / 规则 | Example / 示例 |
|------|------|
| Add space between CJK and Latin | `研究area` → `研究 area` |
| Half-width punctuation for English terms | `DNA，RNA` → `DNA, RNA` |
| Full-width punctuation in Chinese context | `包括soil,water` → `包括soil、water` |
| Space around English parentheses | `方法(method)` → `方法 (method)` |
| Space between number and unit | `30cm` → `30 cm` |
| No space between number and CJK | `第 3 组` → `第3组` |

---

### Module 4: word-content-editor / 内容编辑

**Role / 职责：** Modify document content without breaking formatting / 在不破坏格式的前提下修改文档内容

**Model / 模型：** opus (content editing requires strong reasoning / 内容修改需要强推理能力)

**MCP Tools:**
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
**Python Tools:**
```
python-docx-replace  — Cross-XML-run text find/replace (fallback when MCP search_and_replace fails)
                       跨 XML run 文本查找替换（MCP 失败时的备选方案）
```
**Additional tools:** Read, Write, Edit, Bash (only for tracked changes mode)

**Auto-selected editing modes / 编辑模式自动选择：**

| Mode / 模式 | Scenario / 场景 | Tool path / 工具路径 |
|------|------|---------|
| Precise replace / 精确替换 | Change a few words/sentences | `search_and_replace` — 1 call |
| Paragraph rewrite / 段落重写 | Rewrite an entire paragraph | `replace_paragraph_block_below_header` — 1 call |
| Block replace / 区块替换 | Replace content between two anchors | `replace_block_between_manual_anchors` — 1 call |
| Insert / 插入 | Add new paragraph/section | `insert_line_or_paragraph_near_text` — 1 call |
| Tracked changes / 修订模式 | Need to preserve change history | XML unpack → tracked changes → repack (3+ calls) |

**Principle / 选择原则：** If MCP can do it in one step, never use XML. Only enable XML mode when tracked changes are needed.

能用 MCP 一步完成的，绝不走 XML。只有"需要 tracked changes"时才启用 XML 模式。

---

### Module 5: word-reference / 参考文献管理

**Role / 职责：** Reference formatting, insertion, validation, and Zotero integration / 参考文献格式化、插入、校验，Zotero 集成

**Model / 模型：** sonnet

**MCP Tools:**
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

# Academic search / 学术搜索
mcp__semantic-scholar__search_papers
mcp__semantic-scholar__get_paper
```

**Python Tools:**
```
citeproc-py  — CSL citation formatting engine; pairs with GB/T 7714 and other CSL styles
               CSL 引用格式化引擎，配合 GB/T 7714 等 CSL 样式文件自动生成规范参考文献
```

**Features / 功能：**
- Search Zotero → generate formatted reference list via `citeproc-py` + CSL styles → write to document / 从 Zotero 检索 → 通过 CSL 样式生成参考文献列表并写入文档
- Check consistency between in-text citations `[1]` / `(Author, Year)` and reference list / 检查文内引用与参考文献列表的一致性
- Batch add, delete, and format footnotes/endnotes / 脚注/尾注的批量添加、删除、格式化
- Detect missing citations (mentioned in text but not in references) and orphan references (in list but never cited) / 检测缺失引用和孤立引用

---

### Module 6: word-table-figure / 表格与图片

**Role / 职责：** Create and format tables and figures according to academic standards / 按学术规范创建和格式化表格与图片

**Model / 模型：** sonnet

**MCP Tools:**
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

**Features / 功能：**
- **Three-line table** one-click generation (top rule + header rule + bottom rule — most common in academic papers) / **三线表**一键生成（顶线+栏目线+底线，学术论文最常用）
- Auto-fit or manual column widths / 表格自适应列宽 / 手动设置列宽
- Cell merging (horizontal/vertical) / 单元格合并（横向/纵向）
- Image insertion + caption formatting / 图片插入 + 题注格式化
- Automatic table/figure numbering management / 表格/图片编号自动管理

---

### Module 7: word-reviewer / 审阅修订

**Role / 职责：** Process reviewer comments, manage revisions, generate response documents / 处理审稿意见，管理修订，生成修改说明

**Model / 模型：** opus (needs to understand reviewer comments and plan revision strategy / 需要理解审稿意见并制定修改策略)

**MCP Tools:**
```
mcp__word-document-server__get_all_comments
mcp__word-document-server__get_comments_by_author
mcp__word-document-server__get_comments_for_paragraph
```
**Additional tools:** Read, Write, Edit, Bash (tracked changes require XML)

**Workflow / 工作流：**

```
Reviewer comments (document/email/annotations)
审稿意见（文档/邮件/批注）
      │
      ▼
  Extract and classify item by item / 逐条提取并分类
      │
      ├─ Formatting issues → word-formatter / 格式问题
      ├─ Content changes → word-content-editor / 内容修改
      ├─ References → word-reference / 参考文献
      ├─ Tables/figures → word-table-figure / 表格/图片
      └─ Reply only (no doc changes needed) → write response directly / 纯回复
      │
      ▼
  Generate revision plan → user confirms item by item
  生成修改方案 → 用户逐条确认
      │
      ▼
  Execute changes (tracked changes mode)
  执行修改（tracked changes 模式）
      │
      ▼
  Generate Point-by-Point Response document
  生成 Point-by-Point Response 文档
```

**Output / 输出：**
- Revised paper document (with tracked changes) / 修改后的论文文档（带修订痕迹）
- Point-by-Point Response document (item-by-item replies to reviewers) / 逐条回复审稿人

---

### Module 8: word-checker / 交叉引用检查 + 格式合规验证

**Role / 职责：** Document quality checks — find problems but do not auto-fix / 文档质量检查，发现问题但不自动修复

**Model / 模型：** sonnet

**MCP Tools:**
```
mcp__word-document-server__get_document_text
mcp__word-document-server__get_document_outline
mcp__word-document-server__find_text_in_document
mcp__word-document-server__get_document_info
mcp__word-document-server__validate_document_footnotes
mcp__word-document-server__get_all_comments
```

**Feature A — Cross-Reference Check / 交叉引用检查：**

Scan full text and detect / 扫描全文，检测：

| Check item / 检查项 | Example / 示例 |
|--------|------|
| Reference to non-existent table/figure | Text says "see Table 4" but only Tables 1-3 exist |
| Non-sequential numbering | Fig 1, Fig 2, Fig 4 (missing Fig 3) |
| Duplicate numbering | Two "Table 2"s |
| First-mention order | Text mentions Fig 3 before Fig 1 |
| Inconsistent reference format | Mixing "图1" and "Figure 1" |
| Unreferenced table/figure | Table 3 exists but never mentioned in text |

Output: check report listing all issues with locations / 输出检查报告，列出所有问题及位置。

**Feature B — Format Compliance Verification / 格式合规验证：**

Input: format requirement document + current document / 输入：格式要求文档 + 当前文档
Output: item-by-item compliance report / 输出：逐条合规检查报告

```markdown
## Format Compliance Report: paper_final.docx

### ✅ Passed (15/20)
- ✅ Page size: A4
- ✅ Margins: T2.54 B2.54 L3.17 R3.17
- ✅ Body font: SimSun + Times New Roman
- ✅ Body size: 12pt
- ✅ Heading 1: SimHei, 16pt, bold
- ...

### ❌ Failed (5/20)
- ❌ Line spacing: required 1.5x, actual double in some paragraphs → P90, P91, P95
- ❌ Caption size: required 10.5pt, actual 12pt → Fig 1, Fig 3
- ❌ Page numbers: required starting from body, actual from page 1
- ❌ Reference spacing: required single, actual 1.5x
- ❌ Missing header
```

---

### Module 9: word-submit / 投稿清理 + 文档拆分合并

**Role / 职责：** Final submission preparation, generate clean copy, split/merge documents / 投稿前最终准备，生成 clean copy，拆分/合并文档

**Model / 模型：** sonnet

**MCP Tools:**
```
mcp__word-document-server__get_all_comments
mcp__word-document-server__get_document_info
mcp__word-document-server__copy_document
mcp__word-document-server__protect_document
mcp__word-document-server__unprotect_document
```
**Python Tools:**
```
docxcompose  — Merge multiple .docx files, auto-handling style conflicts, numbering continuation, headers/footers
               合并多个 .docx 文件，自动处理样式冲突、编号续接、页眉页脚
```
**Additional tools:** Read, Write, Edit, Bash (XML for accepting revisions, deleting comments, cleaning metadata)

**Feature A — Clean Copy Pipeline / 投稿前清理：**

```
Original document (with comments, tracked changes, metadata)
原始文档（带批注、修订、元数据）
      │
      ▼
  1. Copy document as working copy / 复制文档为工作副本
  2. Accept all tracked changes / 接受所有修订痕迹
  3. Delete all comments / 删除所有批注
  4. Clear personal metadata (author, org, revision history) / 清除个人元数据
  5. Check image resolution (300dpi+) / 检查图片分辨率
  6. Run word-checker for final validation / 运行 word-checker 做最终检查
  7. Output clean copy / 输出 clean copy
      │
      ▼
  paper_clean.docx (submission-ready / 投稿就绪)
```

**Feature B — Document Splitting / 文档拆分：**

Split one document into multiple files / 将一个文档拆分为多个文件：
- Separate main text + supplementary materials / 正文 + 补充材料分离
- Split by chapter / 按章节拆分
- Extract figures/tables as standalone files / 图片/表格单独提取为独立文件

**Feature C — Document Merging / 文档合并：**

Merge multiple documents into one / 将多个文档合并为一个：
- Multiple co-authors each write a section → merge into complete paper / 多位合作者各写一部分 → 合并为完整论文
- Unify formatting during merge (calls word-formatter) / 合并时统一格式（调用 word-formatter）
- Re-number figures, tables, and references / 重新编号图表和参考文献

---

## 4. System Integration Architecture / 系统集成架构

```
┌─────────────────────────────────────────────────┐
│              User Request / 用户请求               │
└──────────────────────┬──────────────────────────┘
                       │
              ┌────────▼────────┐
              │ word-orchestrator│  ← Routing + Token budget
              └────────┬────────┘
                       │
         ┌─────────────▼─────────────┐
         │       word-reader          │  ← Read once, generate Document Map
         └─────────────┬─────────────┘
                       │ Document Map (cached & reused)
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
   │ (verify)  │ │ (review) │
   └─────┬─────┘ └────┬─────┘
         │            │
         ▼            ▼
   ┌─────────────────────┐
   │    word-submit       │  ← Clean copy + split/merge
   └─────────────────────┘
         │
         ▼
   ┌─────────────────────────────────────────┐
   │   Underlying Tool Auto-Routing           │
   │   底层工具自动路由                         │
   │  ┌──────────────┐  ┌──────────────────┐ │
   │  │ Word MCP     │  │ docx skill       │ │
   │  │ Server       │  │ (XML unpack/pack)│ │
   │  │ preferred ✓  │  │ tracked changes  │ │
   │  └──────────────┘  └──────────────────┘ │
   └─────────────────────────────────────────┘
```

---

## 5. Tool Routing Strategy / 工具路由策略

**Core principle: If MCP can do it in one step, never use XML.**

**核心原则：能用 MCP 一步完成的，绝不走 XML。**

| Operation / 操作 | Primary path / 首选路径 | Fallback / 备选路径 |
|------|---------|---------|
| Read document structure | `get_document_outline` | `docx2python` (when full hierarchy needed) |
| Generate Document Map | `docx2python` one-pass | MCP multi-call (per-paragraph) |
| Extract text box content | `zipfile` + `ElementTree` parse XML `w:txbxContent` | `get_document_xml` + manual parse |
| Find text | `find_text_in_document` | `get_document_text` + search |
| Replace text | `search_and_replace` | `python-docx-replace` (cross-run) |
| Modify formatting | `format_text` | XML edit |
| Add paragraph | `insert_line_or_paragraph_near_text` | XML edit |
| Replace paragraph block | `replace_paragraph_block_below_header` | XML edit |
| Add table | `add_table` | docx-js |
| Add image | `add_picture` | XML + relationship edit |
| Tracked changes | **Must** XML unpack/edit/pack | No MCP alternative |
| Accept all revisions | `accept_changes.py` script | Manual XML |
| Add comments | `comment.py` + XML markup | No MCP alternative |
| Generate TOC | XML TOC field code insertion | docx-js `TableOfContents` |
| Document comparison | `docx2python` + `deepdiff` + `redlines` | Per-paragraph MCP + text diff |
| Reference formatting | `citeproc-py` + CSL styles | Manual formatting |
| Document merge | `docxcompose` | Manual XML merge |
| Convert to PDF | `soffice --convert-to pdf` | — |

---

## 6. Integration with Other Systems / 与现有系统的协作

### Integration with eco-agent / 与 eco-agent 的协作

```
eco-agent (what to write — academic content)     word-agent (how to present — Word formatting)
eco-agent（写什么 — 学术内容）                     word-agent（怎么呈现 — Word格式）
──────────────────────────────────               ──────────────────────────────────
ecology-paper-writing                            word-content-editor
  → Output Markdown/text draft  ──→                → Write into Word document

ecology-review                                   word-reviewer
  → Review strategy + response draft  ──→          → Execute doc changes + generate response

ecology-polish                                   word-formatter
  → Language-polished text  ──→                    → Apply to Word preserving formatting

ecology-data-analysis                            word-table-figure
  → Statistical results + table data  ──→          → Generate academic three-line tables
```

### Relationship with docx skill / 与 docx skill 的关系

The docx skill is retained as the **underlying tool layer**. word-agent calls its scripts (`unpack.py`, `pack.py`, `comment.py`, `accept_changes.py`) when XML operations are needed. word-agent does not replace docx skill — it provides high-level workflows on top of it.

docx skill 作为**底层工具层**保留，word-agent 在需要 XML 操作时调用它的脚本。word-agent 不替代 docx skill，而是在其之上提供高层工作流。

### Zotero MCP Integration / 与 Zotero MCP 的集成

The word-reference module directly calls Zotero MCP tools, automating the pipeline from library search to formatted reference list.

word-reference 模块直接调用 Zotero MCP 工具，实现从文献库到参考文献列表的自动化。

---

## 7. Implementation Roadmap / 实施路线图

| Phase / 阶段 | Modules / 模块 | Rationale / 理由 | Effort / 工作量 |
|------|------|------|-----------|
| **P0** | word-orchestrator + word-reader | Foundation — all modules depend on these / 基础设施，所有模块依赖 | Medium / 中 |
| **P1** | word-formatter + word-checker | Solves core pain point "formatting is wrong" + checker closes the loop / 解决核心痛点 + checker 提供验证闭环 | Large / 大 |
| **P2** | word-content-editor | Content editing is high-frequency / 内容编辑是高频操作 | Medium / 中 |
| **P3** | word-table-figure + word-reference | Tables and references are common needs / 表格和参考文献是常见需求 | Medium / 中 |
| **P4** | word-reviewer + word-submit | Review handling and submission prep / 审稿修改和投稿清理 | Medium / 中 |

Each phase is independently usable after completion — no need to wait for all modules.

每个阶段完成后可独立使用，不需要等全部模块完成。

---

## 8. External Dependencies / 外部依赖

### Python Libraries / Python 库

| Library / 库 | Install / 安装 | Purpose / 用途 | Module / 模块 |
|---|------|------|---------|
| **docx2python** | `pip install docx2python` | Parse .docx into nested Python lists preserving full hierarchy (headings, paragraphs, tables, footnotes, images). More efficient than per-paragraph MCP calls for Document Map generation. / 解析 .docx 为嵌套列表，保留完整层次结构 | word-reader |
| **docxcompose** | `pip install docxcompose` | Merge multiple .docx into one, auto-handling style conflicts, numbering, headers/footers / 合并多个 .docx，自动处理样式冲突 | word-submit |
| **citeproc-py** | `pip install citeproc-py` | CSL citation formatting engine; pairs with GB/T 7714 and other CSL style files / CSL 引用格式化引擎 | word-reference |
| **python-docx-replace** | `pip install python-docx-replace` | Cross-XML-run text find/replace; solves known limitation of python-docx/MCP with split runs / 跨 XML run 文本替换 | word-content-editor (fallback) |
| **redlines** | `pip install redlines` | Text comparison with red-line markup (tracked-changes-style diff visualization) / 红线差异可视化 | word-reader (compare mode) |
| **deepdiff** | `pip install deepdiff` | Deep structural comparison for nested document structures / 深层结构对比 | word-reader (compare mode) |

### CSL Style Files / CSL 样式文件

| File / 文件 | Source / 来源 | Purpose / 用途 |
|------|------|------|
| `gb-t-7714-2015-numeric.csl` | [citation-style-language/styles](https://github.com/citation-style-language/styles) | Chinese paper GB/T 7714-2015 numeric format / 中文论文数字编号格式 |
| `gb-t-7714-2015-author-date.csl` | Same / 同上 | Chinese paper GB/T 7714-2015 author-date format / 作者-年份格式 |
| Other journal CSL | Same / 同上 | Selected per user's format requirements / 按用户格式要求选用 |

CSL style files are stored in `references/csl/`, downloaded on demand.

### Existing Dependencies (retained) / 已有依赖（保留）

| Tool / 工具 | Purpose / 用途 |
|------|------|
| Word MCP Server | 60+ tools, primary document operation path / 主要文档操作路径 |
| docx skill scripts | `unpack.py`, `pack.py`, `comment.py`, `accept_changes.py` — XML operations |
| Zotero MCP | Library search and export / 文献库检索与导出 |
| Semantic Scholar MCP | Academic paper search and metadata / 学术论文搜索与元数据获取 |
| pandoc | Text extraction and format conversion / 文本提取与格式转换 |
| LibreOffice (soffice) | PDF conversion, field updates / PDF 转换、域更新 |

### Installation / 依赖安装

```bash
# Core dependencies (needed from P0) / 核心依赖（P0 阶段即需要）
pip install docx2python deepdiff redlines

# P1 additions / P1 阶段追加
pip install python-docx-replace

# P3 additions / P3 阶段追加
pip install docxcompose citeproc-py

# CSL style files (download on demand) / CSL 样式文件（按需下载）
curl -o references/csl/gb-t-7714-2015-numeric.csl \
  https://raw.githubusercontent.com/citation-style-language/styles/master/gb-t-7714-2015-numeric.csl
```

---

## 9. Complete File Structure / 文件结构完整清单

```
word-agent/
├── .claude-plugin/
│   ├── plugin.json
│   └── marketplace.json
├── CLAUDE.md
├── FRAMEWORK.md                    # This document / 本文档
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
    ├── requirements.txt            # Python dependency list / Python 依赖清单
    └── setup.sh                    # One-click install script / 一键安装脚本
```
