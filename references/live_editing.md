# Live Editing Reference

## Overview

Live editing allows word-agent to modify a Word document while it is open in Microsoft Word. Changes appear in real-time in the Word UI. This mode uses word-mcp-live's COM automation (Windows) or JXA automation (macOS) to communicate directly with the running Word application.

## Platform Support

| Platform | Automation | Supported | Notes |
|----------|-----------|-----------|-------|
| macOS | JXA (JavaScript for Automation) | Yes | Requires Word for Mac |
| Windows | COM (Component Object Model) | Yes | Requires Word for Windows |
| Linux | None | No | No native Word application available |

### Detection Logic

```
Step 1: Check platform
  → macOS: proceed to JXA detection
  → Windows: proceed to COM detection
  → Linux: live editing unavailable, use file mode only

Step 2: Check Word is running with target document
  → mcp__word-mcp-live__get_active_document()
  → Returns: file_path of active document, or null if Word not open

Step 3: Report mode
  → Word open with target doc → "实时编辑模式可用。修改将即时显示在 Word 中。"
  → Word not open → "Word 未打开目标文档。使用文件模式编辑。"
```

## File Locking

When Word opens a document, it creates a lock file (`~$filename.docx`) and holds an exclusive write lock. This means:

- **File-mode MCP tools** (word-document-server) **cannot write** to the locked file
- **Live-mode MCP tools** (word-mcp-live) **can write** via automation API
- **Font normalization** (`normalize_fonts.py`) **cannot run** on a locked file

### Lock Detection

```bash
python3 scripts/normalize_fonts.py "{file_path}" --check-lock
# Exit code 0: file not locked, safe to process
# Exit code 2: file is locked, deferred processing needed
```

## Deferred Font Normalization

In live editing mode, font normalization cannot run immediately because Word holds the file lock. The orchestrator must track this as a pending operation:

```
After live editing operations:
  1. Try: python3 scripts/normalize_fonts.py "{file_path}" --check-lock --unify
  2. If exit code 2 (locked):
     → Warn user: "字体归一化已延迟——请关闭 Word 后运行，或在下次操作时自动执行。"
     → Set pending_normalization = true for this file
  3. On next interaction with same file:
     → If pending_normalization and file not locked → run normalization
     → If still locked → remind user again
```

## Per-Action Undo

Live editing through word-mcp-live supports per-action undo via the Word application's undo stack. Each MCP call corresponds to one undo step in Word.

```
mcp__word-mcp-live__undo_last_operation()
→ Equivalent to Ctrl+Z in Word
→ Undoes the last word-mcp-live operation
→ Can be called multiple times to undo multiple operations
```

**Undo is only available in live editing mode.** File-mode operations cannot be undone via this mechanism (use document backup/copy instead).

### Undo Semantics

| Operation | Undo Result |
|-----------|-------------|
| `edit_text` | Restores original text |
| `insert_text` | Removes inserted text |
| `delete_text` | Restores deleted text |
| `replace_text` | Restores original text |
| `add_comment` | Removes added comment |
| `resolve_comment` | Unresolves the comment |

## Mode Rules

### No Mixed Modes

**Critical rule:** Within the same session for the same file, do NOT mix live editing mode and file mode. Specifically:
- If Word is open and you're using word-mcp-live live operations, do NOT also use word-document-server write operations on the same file
- If you're using file-mode operations, do NOT also send live commands

**Why:** File-mode writes while Word has the file open will either fail (lock) or create conflicts if Word has unsaved changes in memory.

### Mode Selection in Orchestrator

```
Phase 2, Step 5: Editing Mode Detection

IF target document is open in Word (get_active_document returns match):
  → Set mode = "live"
  → Inform downstream modules: use word-mcp-live for all write operations
  → Defer font normalization until Word closes
ELSE:
  → Set mode = "file" (default)
  → Use standard tool routing (word-document-server primary)
  → Font normalization runs immediately after operations
```

## Layout Diagnostics (Live Mode Only)

When Word is open, word-mcp-live can inspect the rendered layout:

```
mcp__word-mcp-live__diagnose_layout(file_path)
→ Returns: page breaks, orphan/widow lines, overflow text, actual vs expected page count
→ Only available in live mode (requires Word's layout engine)
```

This is exposed as Mode D in word-checker, triggered by "布局检查" / "layout check" / "排版诊断".

## Interaction Flow

```
User: "修改第3段的内容"
  ↓
Orchestrator Phase 2 Step 5: detect Word status
  ↓
IF live mode:
  → word-edit uses mcp__word-mcp-live__replace_text (changes appear in Word immediately)
  → User sees change in real-time
  → Font normalization deferred
  → Undo available if user says "撤销"
  ↓
IF file mode:
  → word-edit uses mcp__word-document-server__search_and_replace (standard path)
  → Font normalization runs immediately
  → No undo support (use document copy as backup)
```
