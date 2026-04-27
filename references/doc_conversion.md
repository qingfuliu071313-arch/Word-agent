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

### Step 2: 后处理修复（关键）

LibreOffice 转换会引入以下已知问题，必须在转换后立即修复：

| 问题 | 原因 | 修复方法 |
|------|------|---------|
| `compatibilityMode=11` | LibreOffice 默认生成 Word 2003 兼容模式 | 改为 `15`（Word 2013+） |
| `Liberation Serif/Sans` 字体 | LibreOffice 默认字体不存在于 Windows/Mac Word | 映射为 `Times New Roman` / `Arial` |
| `新宋体` 默认东亚字体 | LibreOffice 使用非标准中文字体名 | 映射为 `宋体` |
| 带分号的后备字体名 `宋体;SimSun` | LibreOffice 的字体 fallback 语法不被 Word 识别 | 去除分号及后备名 |
| `Heading1 basedOn TOC1` | 标题样式基于目录样式，样式链断裂 | 改为 `basedOn Normal` |

### Step 3: 后处理修复代码

```python
import zipfile, os, re

def fix_libreoffice_docx(docx_path):
    """修复 LibreOffice .doc→.docx 转换的已知问题"""

    def read_entry(zpath, inner):
        with zipfile.ZipFile(zpath, 'r') as z:
            return z.read(inner).decode('utf-8')

    def patch_entries(zpath, patches):
        tmp = zpath + '.tmp'
        with zipfile.ZipFile(zpath, 'r') as zin, \
             zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = patches.get(item.filename) or zin.read(item.filename)
                if isinstance(data, str):
                    data = data.encode('utf-8')
                zout.writestr(item, data)
        os.replace(tmp, zpath)

    FONT_MAP = {
        'Liberation Serif': 'Times New Roman',
        'Liberation Sans': 'Arial',
        'Liberation Mono': 'Courier New',
        '宋体;SimSun': '宋体',
        '黑体;SimHei': '黑体',
        '楷体_GB2312;KaiTi': '楷体',
        '仿宋_GB2312;FangSong': '仿宋',
    }

    patches = {}
    with zipfile.ZipFile(docx_path, 'r') as z:
        names = z.namelist()

    # Fix settings.xml: compatibilityMode 11 → 15
    if 'word/settings.xml' in names:
        settings = read_entry(docx_path, 'word/settings.xml')
        settings = re.sub(
            r'(compatibilityMode.*?w:val=")11(")',
            r'\g<1>15\2',
            settings
        )
        if 'characterSpacingControl' not in settings:
            settings = settings.replace(
                '</w:settings>',
                '<w:characterSpacingControl w:val="compressPunctuation"/></w:settings>'
            )
        patches['word/settings.xml'] = settings

    # Fix fonts in all XML files
    xml_files = [n for n in names if n.endswith('.xml') and n.startswith('word/')]
    for xml_file in xml_files:
        content = read_entry(docx_path, xml_file)
        original = content
        for old, new in FONT_MAP.items():
            content = content.replace(old, new)
        # Fix default East Asian font
        content = content.replace('w:eastAsia="新宋体"', 'w:eastAsia="宋体"')
        # Fix Heading1 basedOn TOC
        if 'Heading1' in xml_file or 'styles' in xml_file:
            content = re.sub(
                r'(<w:style[^>]*styleId="Heading1".*?<w:basedOn w:val=")TOC\d+(")',
                r'\1Normal\2',
                content,
                flags=re.DOTALL
            )
        if content != original:
            patches[xml_file] = content

    if patches:
        patch_entries(docx_path, patches)

    return len(patches)
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
2. **必须执行后处理修复** — 裸 LibreOffice 转换产物在 Word 中会出现严重显示问题
3. **如果 LibreOffice 不可用**，提示用户手动用 Word 打开 .doc 并另存为 .docx
4. **转换可能丢失部分格式**（如某些域代码、VBA宏），需告知用户
