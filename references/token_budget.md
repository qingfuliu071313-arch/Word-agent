# Token Budget Reference

## 核心目标

减少不必要的文档读取和工具调用，将 Token 消耗控制在最低水平。

## 五大原则

### 1. 懒加载（Lazy Loading）

**永远不要一上来就读全文。** 按需逐层深入：

```
Level 0: get_document_info          ~50 tokens    基本信息
Level 1: get_document_outline       ~200 tokens   标题结构
Level 2: docx2python (summary)      ~500 tokens   完整结构（含表格/图片/脚注概要）
Level 3: get_paragraph_text (范围)  ~变动        仅读取需要的段落
Level 4: get_document_text          ~全文 tokens  最后手段
```

大多数任务只需要 Level 0-2。只有内容编辑才需要 Level 3。Level 4 几乎不应该使用。

### 2. 结构缓存（Document Map Caching）

- word-reader 生成 Document Map 后，通过对话上下文传递给下游模块
- 下游模块**禁止**重新调用 `get_document_outline` 或 `get_document_text`
- 如果需要特定段落内容，使用 `get_paragraph_text_from_document` 精确读取

### 3. 工具选择优先级

每次操作前，按优先级选择工具路径：

| 优先级 | 路径 | 典型 Token 消耗 | 适用场景 |
|--------|------|----------------|---------|
| 1 | MCP 单次调用 | 低 (~100-200) | 查找、替换、格式化单段落 |
| 2 | MCP 多次调用 | 中 (~500-1000) | 批量格式修改 |
| 3 | Python 脚本 (Bash) | 中 (~300-800) | 文档对比、结构分析 |
| 4 | XML unpack/edit/pack | 高 (~2000+) | tracked changes、批注 |

### 4. 批量操作（Batch Operations）

**错误做法：** 逐段落修改字体
```
format_text(paragraph=1, font="宋体")  # 调用 1
format_text(paragraph=2, font="宋体")  # 调用 2
format_text(paragraph=3, font="宋体")  # 调用 3
... (187 次调用)
```

**正确做法：** 按规则分类，批量处理
```
# 先分析：哪些段落需要改字体？
# Document Map 已告诉我们 P23 字体不一致
# 只修改不一致的段落
format_text(paragraph=23, font="宋体")  # 1 次调用
```

### 5. 按需精读（On-Demand Detail）

当下游模块需要某个章节的详细内容时：

1. 从 Document Map 查找目标章节的段落范围（如 `[P43-P55] 2.1 研究区域`）
2. 只读取该范围：`get_paragraph_text_from_document(start=43, end=55)`
3. 不读取其他章节

## Token 消耗估算

| 文档规模 | 全文读取 | Document Map 方式 | 节省 |
|----------|---------|------------------|------|
| 10 页论文 (~5000 字) | ~6000 tokens | ~800 tokens | ~87% |
| 30 页论文 (~15000 字) | ~18000 tokens | ~1200 tokens | ~93% |
| 50 页论文 (~25000 字) | ~30000 tokens | ~1500 tokens | ~95% |

注：Document Map 方式 = Level 0-2 读取 + 按需 Level 3 精读目标段落。

## 反模式（应避免）

| 反模式 | 问题 | 正确做法 |
|--------|------|---------|
| 先读全文再分析结构 | 浪费 Token | 先 outline → 按需精读 |
| 每次操作都重新读文档 | 重复消耗 | 复用 Document Map |
| 逐段落检查格式 | 调用次数爆炸 | 用 docx2python 一次提取 |
| 用 get_document_text 搜索关键词 | 全文加载 | 用 find_text_in_document |
| 修改一个字走 XML unpack/pack | 大材小用 | 用 search_and_replace |
