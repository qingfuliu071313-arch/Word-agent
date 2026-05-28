# .doc → .docx 转换规范

## 触发条件

当用户提供 `.doc`（旧版 Word 二进制格式）文件时，需要先转换为 `.docx` 再进行后续操作。

## 检测方法

```python
import struct

def is_legacy_doc(file_path):
    """检测文件是否为旧版 .doc 格式（OLE2 Compound Document）"""
    if not file_path.lower().endswith('.doc'):
        return False
    try:
        with open(file_path, 'rb') as f:
            magic = f.read(8)
        # OLE2 magic bytes: D0 CF 11 E0 A1 B1 1A E1
        return magic[:4] == b'\xd0\xcf\x11\xe0'
    except Exception:
        return False
```

## 转换流程

### Step 1: 使用 LibreOffice 转换

```bash
soffice --headless --convert-to docx --outdir "{output_dir}" "{input.doc}"
```

如果 LibreOffice 不可用，提示用户：
```
需要 LibreOffice 来转换 .doc 文件。请安装 LibreOffice，或手动用 Word 打开 .doc 并另存为 .docx。
```

### Step 2: 后处理修复（MANDATORY）

**转换后必须立即运行 `scripts/fix_libreoffice.py`：**

```bash
python3 scripts/fix_libreoffice.py "{converted.docx}"
```

此脚本自动修复所有已知的 LibreOffice 转换问题：

| 修复项 | 问题描述 | 修复方式 |
|--------|---------|---------|
| compatibilityMode | LibreOffice 默认生成 Word 2003 兼容模式 (11) | 改为 15 (Word 2013+) |
| Liberation 字体 | Liberation Serif/Sans 不存在于 Windows/Mac | 映射为 Times New Roman / Arial |
| 分号后备字体 | `宋体;SimSun` 不被 Word 识别 | 去除分号及后备名 |
| 非标准中文字体 | `新宋体`、`楷体_GB2312` 等旧字体名 | 映射为标准名 (宋体、楷体) |
| 样式链断裂 | Heading1 basedOn TOC1 | 改为 basedOn Normal |
| CJK 字符间距 | 缺少 characterSpacingControl | 添加 compressPunctuation |
| 表格宽度 | 转换后表格宽度变为 0 或 auto | 设为 100%，均分列宽 |
| 表格边框 | 转换后丢失边框 | 添加默认单线边框 |
| 零宽列 | 列宽为 0 导致表格压缩 | 按比例重新分配宽度 |
| 段间距异常 | 段前段后间距过大 (>30pt) | 封顶至 24pt |
| 行距异常 | exact 模式行距过小 (<10pt) | 改为 1.5 倍行距 |
| 页面尺寸 | 尺寸为 0 或异常值 | 恢复为 A4 |
| 页边距异常 | 边距为 0 或超大值 | 恢复为合理默认值 |
| 图片引用 | 图片路径断裂 | 检测并报告 |
| **字体归一化** | **主题字体、裸 run、缺失 eastAsia** | **自动调用 normalize_fonts.py --unify** |

脚本最后会**自动调用 `normalize_fonts.py --unify`** 完成字体归一化。不需要手动单独运行。

### 可用选项

```bash
# 只检测不修复
python3 scripts/fix_libreoffice.py "{file}" --detect-only

# 输出到新文件（保留原文件）
python3 scripts/fix_libreoffice.py "{input}" "{output}"

# 跳过字体归一化（不推荐）
python3 scripts/fix_libreoffice.py "{file}" --skip-font-normalize

# JSON 输出（供程序调用）
python3 scripts/fix_libreoffice.py "{file}" --json
```

## 输出文件命名

- 转换后的文件保存在与原 `.doc` 文件同目录下
- 文件名：`{原文件名}.docx`（与 LibreOffice 默认行为一致）
- 如果同名 `.docx` 已存在，保存为 `{原文件名}-converted.docx`

## 用户提示模板

转换完成后，向用户报告：

```
已将 {filename}.doc 转换为 .docx 格式并完成格式修复。
转换后文件：{output_path}
原始文件保留不变。

后续所有操作将基于转换后的 .docx 文件进行。
```

## 注意事项

1. **永远不修改原始 .doc 文件**
2. **后处理修复是强制步骤** — 裸 LibreOffice 转换产物在 Word 中会出现严重显示问题
3. **如果 LibreOffice 不可用**，提示用户手动用 Word 打开 .doc 并另存为 .docx
4. **转换可能丢失部分格式**（如某些域代码、VBA 宏），后处理脚本会尽可能修复可修复的部分
