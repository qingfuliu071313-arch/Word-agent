# CSL 引文样式文件

word-reference 模块的 citeproc-py 流程需要本目录下的 CSL 样式文件。
（仓库当前不内置 CSL 文件 —— 构建环境无网络时无法自动下载，请按下表手动获取。）

| 样式 | 期望文件名 | 下载地址 |
|------|-----------|---------|
| GB/T 7714-2015 顺序编码制 | `china-national-standard-gb-t-7714-2015-numeric.csl` | https://raw.githubusercontent.com/citation-style-language/styles/master/china-national-standard-gb-t-7714-2015-numeric.csl |
| APA 7th | `apa.csl` | https://raw.githubusercontent.com/citation-style-language/styles/master/apa.csl |
| 其他样式 | 保持原文件名 | https://www.zotero.org/styles |

下载命令示例：

```bash
curl -sL https://raw.githubusercontent.com/citation-style-language/styles/master/china-national-standard-gb-t-7714-2015-numeric.csl \
  -o references/csl/china-national-standard-gb-t-7714-2015-numeric.csl
curl -sL https://raw.githubusercontent.com/citation-style-language/styles/master/apa.csl \
  -o references/csl/apa.csl
```

下载后请验证文件内容是 XML（首行应为 `<?xml version=...`），而非 404 页面。

**CSL 文件缺失时的行为**：word-reference 会提示上述下载地址，并降级为手工格式化（按目标样式规则逐条重排），不会调用 citeproc-py。详见 `skills/word-reference/SKILL.md` "Using citeproc-py" Step 0。
