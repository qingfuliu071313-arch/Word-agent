# OOXML Structural Validation Reference

## Overview

Word documents (.docx) are ZIP archives containing OOXML. Editing tools — including MCP servers, python-docx, and manual XML manipulation — can introduce structural corruption that Word silently repairs on open (or fails to open entirely). This reference defines the structural checks available via docx-mcp and how word-agent uses them.

## Backend

All structural validation is performed by **docx-mcp** (P3 in tool routing). word-document-server and word-mcp-live do not provide validation tools.

| Tool | Purpose | When to Use |
|------|---------|-------------|
| `mcp__docx-mcp__validate_paraids` | Check paragraph ID uniqueness and format | After any XML manipulation or merge |
| `mcp__docx-mcp__audit_document` | Comprehensive structural audit | Pre-submission, after complex edits |

## Paragraph IDs (paraId)

Every `<w:p>` element in OOXML must have a unique `w14:paraId` attribute (8-digit hex). Duplicate or missing paraIds cause:
- Track changes corruption (revisions reference wrong paragraphs)
- Comment anchoring errors
- Merge conflicts when combining documents

### Check: `validate_paraids`

```
Input:  file_path
Output: {
  "valid": true/false,
  "total_paragraphs": n,
  "duplicates": [{"paraId": "AABBCCDD", "count": 2, "locations": [...]}],
  "missing": [{"paragraph_index": n, "context": "first 50 chars..."}],
  "format_errors": [{"paraId": "XYZ", "reason": "not 8-digit hex"}]
}
```

**Common causes of paraId issues:**
- Copy-paste between documents (duplicates IDs from source)
- docxcompose merge without ID regeneration
- Manual XML editing that omits paraId attributes
- python-docx creating paragraphs without paraId

## Comprehensive Audit: `audit_document`

Checks multiple structural integrity aspects:

| Check | What It Finds | Severity |
|-------|--------------|----------|
| Broken bookmarks | `<w:bookmarkStart>` without matching `<w:bookmarkEnd>` | ERROR |
| Orphan image references | `<a:blip>` pointing to missing `rId` in relationships | ERROR |
| Invalid numbering definitions | `<w:numId>` referencing non-existent `<w:abstractNum>` | ERROR |
| Style conflicts | Multiple styles with same `w:styleId` | WARNING |
| Broken hyperlinks | `<w:hyperlink>` with invalid `r:id` | WARNING |
| Missing content types | Files in ZIP not declared in `[Content_Types].xml` | ERROR |
| Relationship mismatches | `*.rels` entries pointing to non-existent parts | ERROR |

### Output Format

```
{
  "clean": true/false,
  "issues": [
    {
      "type": "broken_bookmark",
      "severity": "error",
      "location": "word/document.xml",
      "detail": "bookmarkStart id=5 'ref_table3' has no matching bookmarkEnd",
      "suggestion": "Remove orphaned bookmarkStart or add bookmarkEnd"
    },
    ...
  ],
  "summary": {
    "errors": n,
    "warnings": n,
    "total_checks": n
  }
}
```

## Integration Points

### word-checker: Mode C — Structural Validation

Triggered by: "结构检查", "structural check", "paraId", "bookmarks", "文档结构"

```
Step 1: validate_paraids(file_path)
  → Report duplicate/missing/malformed paraIds

Step 2: audit_document(file_path)
  → Report broken bookmarks, orphan images, invalid numbering, style conflicts

Step 3: Compile into Check Report (same format as Mode A/B)
```

Mode C is automatically included when:
- word-submit runs pre-submission checks (Feature D)
- User requests "全面检查" / "full check" (combined with Mode A + B)

### word-submit: Pre-Submission Structural Gate (Feature D Step 1b)

```
Step 1b: Structural integrity verification
  → validate_paraids → must pass (zero duplicates/missing)
  → audit_document → must have zero errors (warnings OK with user confirmation)
  → If fails: "文档结构存在问题，建议修复后再投稿。详见检查报告。"
```

### Submission Checklist Items

| # | Check | Method | Pass Criteria |
|---|-------|--------|---------------|
| A5 | 段落 ID 唯一 | `validate_paraids` | 零重复、零缺失 |
| A6 | 书签完整 | `audit_document` | 零断裂书签 |
| A7 | 图片引用完整 | `audit_document` | 零孤立图片引用 |

## Fallback

If docx-mcp is unavailable, structural validation is skipped with a warning:

```
⚠ docx-mcp 不可用，跳过结构验证（paraId、书签、图片引用完整性）。
  建议安装 docx-mcp 后重新运行检查。
```

Basic paraId duplicate detection can be done via XML scan as a minimal fallback:

```bash
python3 << 'PYEOF'
import zipfile, xml.etree.ElementTree as ET, collections
with zipfile.ZipFile("{file_path}", 'r') as z:
    tree = ET.parse(z.open('word/document.xml'))
ns = {'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'}
ids = [p.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId', '')
       for p in tree.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')]
dupes = {k: v for k, v in collections.Counter(ids).items() if v > 1 and k}
print(f"Total paragraphs: {len(ids)}, Duplicates: {len(dupes)}")
if dupes:
    for pid, count in dupes.items():
        print(f"  paraId={pid} appears {count} times")
PYEOF
```

## Common Corruption Patterns

| Pattern | Cause | Fix |
|---------|-------|-----|
| Duplicate paraIds after merge | docxcompose or manual merge | Regenerate IDs with docx-mcp |
| Broken bookmarks after delete | Deleting paragraphs containing bookmarkStart/End | Remove orphaned bookmark elements |
| Orphan image refs after edit | Removing drawing elements but not updating relationships | Clean up `.rels` file |
| Style conflicts after merge | Two source documents define same styleId differently | Rename conflicting styles |
