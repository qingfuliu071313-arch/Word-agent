# adeu Integration Reference

## 概述

adeu 是一个 docx ↔ Markdown 双向翻译器，作为 Python 库集成到 word-agent 中（非 MCP server）。核心价值：将 Word 文档投射为 token 高效的 Markdown，AI 直接编辑 Markdown，再原子写回为 OOXML tracked changes。

**安装:** `pip install adeu>=1.8.0`

## 三阶段流水线

### Phase 1: Read（提取）

将 .docx 文件投射为 Markdown + CriticMarkup + Semantic Appendix。

```python
from adeu import RedlineEngine
engine = RedlineEngine("document.docx")
md = engine.extract()
```

**Semantic Appendix** 附加在 Markdown 末尾，包含：
- 已定义术语和交叉引用
- 可能的拼写错误标记
- 文档结构上下文元数据

### Phase 2: Validate（安全验证）

在修改写回前，验证所有变更的合法性。

```python
result = engine.validate(redlines)
if not result.valid:
    print(result.errors)  # 歧义匹配、无效结构变更
```

**验证规则:**
- 阻止歧义文本匹配（同一文本出现多处时要求明确定位）
- 阻止无效结构变更（如删除标题但保留其下属段落）
- 阻止孤立引用（删除被引用的目标）

### Phase 3: Apply（原子写回）

将验证通过的编辑写回 .docx 文件，生成原生 Word tracked changes。

```python
engine.apply_redline(validated_redlines)
engine.save("output.docx")
```

写回特性：
- 每个修改生成原生 Word 修订标记（非 XML hack）
- 保留原始格式、字体、页面布局
- 原子事务：要么全部成功，要么全部回滚

## CriticMarkup 语法

AI 使用 CriticMarkup 语法表达修改意图：

| 操作 | 语法 | 示例 |
|------|------|------|
| 插入 | `{++text++}` | `{++新增的段落内容++}` |
| 删除 | `{--text--}` | `{--需要删除的文字--}` |
| 替换 | `{~~old~>new~~}` | `{~~旧表述~>新表述~~}` |
| 高亮 | `{==text==}` | `{==需要关注的内容==}` |
| 批注 | `{>>comment<<}` | `{>>这里需要补充数据来源<<}` |

## 触发条件

在以下场景中，word-edit 和 word-review 应优先使用 adeu 批量编辑：

| 场景 | 条件 |
|------|------|
| 批量内容修改 | 同时 10+ 处内容变更 |
| 用户显式要求 | "批量修改" / "bulk edit" |
| 审稿修订 | word-reviewer Phase 3 有 10+ 修订项 |
| 全文改写 | 多个章节需要重写 |

## Python SDK 使用模式

### 模式 A: 批量替换

```python
from adeu import RedlineEngine, ModifyText

engine = RedlineEngine("paper.docx", author="Claude")

edits = [
    ModifyText(
        target_text="State of New York",
        new_text="State of Delaware",
        comment="统一管辖法律。"
    ),
    ModifyText(
        target_text="ABC Corporation",
        new_text="XYZ Ltd",
        comment="更新公司名称。"
    ),
]

engine.apply_edits(edits)
engine.save("paper_redlined.docx")
```

### 模式 B: CLI 工具

```bash
# 提取为 Markdown
uvx adeu extract paper.docx -o paper.md

# 对比两个版本
uvx adeu diff v1.docx v2.docx

# 应用编辑（带 tracked changes）
uvx adeu apply paper.docx edits.json --author "Claude"
```

### 模式 C: 在 word-agent 中通过 Bash 调用

```bash
python3 -c "
from adeu import RedlineEngine, ModifyText
from io import BytesIO

with open('paper.docx', 'rb') as f:
    stream = BytesIO(f.read())

engine = RedlineEngine(stream, author='Claude')
edits = [
    ModifyText(target_text='旧文本', new_text='新文本', comment='原因说明')
]
engine.apply_edits(edits)

with open('paper_redlined.docx', 'wb') as f:
    f.write(engine.save_to_stream().getvalue())
"
```

## 与现有工具的关系

| 维度 | adeu | word-document-server MCP | XML unpack/edit/pack |
|------|------|--------------------------|---------------------|
| 编辑粒度 | 文档级批量 | 段落级精确 | 元素级精确 |
| Token 效率 | 高（Markdown 中间表示） | 中（JSON 参数） | 低（完整 XML） |
| Tracked Changes | 原生支持 | 不支持 | 手动 XML 拼装 |
| 安全验证 | 内置 Validate 阶段 | 无 | 无 |
| 适用场景 | 10+ 处批量修改 | 单处/少量修改 | 遗留兜底 |

## 注意事项

1. **font normalization gate 仍然适用**: adeu 写回 .docx 后，必须执行 `normalize_fonts.py --unify`
2. **Document Map 失效**: adeu 修改文档后，缓存的 Document Map 标记为 stale
3. **非 MCP 工具**: adeu 通过 Bash 调用，不在 agent 的 `tools:` 列表中（但 Bash 已在所有 agent 工具中）
4. **与 word-mcp-live 不冲突**: adeu 操作文件级 .docx，word-mcp-live 操作打开的 Word 进程，两者场景不重叠
