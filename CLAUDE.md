# CLAUDE.md

This file provides guidance to Claude Code when working with code in this repository.

## What This Is

word-agent is a **Claude Code plugin** (v0.1.0) for academic paper Word document operations. It provides 9 specialized agents covering the full manuscript lifecycle: document analysis, formatting, content editing, reference management, table/figure formatting, review handling, quality checking, and submission preparation.

**Target users:** Researchers who write and edit academic papers in Microsoft Word, with bilingual (EN/中文) support.

## Architecture: Three-Tier WHO/HOW/WHAT

```
agents/                 → WHO  (persona, model, allowed tools)
skills/*/SKILL.md       → HOW  (workflow phases, templates, output format)
references/             → WHAT (domain knowledge, standards, rules)
```

- **Agents** define the persona and tool access. Routing agents use `model: sonnet`; content-reasoning agents use `model: opus`.
- **Skills** define multi-phase workflows. Each skill has a `references/` subdirectory for module-specific knowledge.
- **References** are lazy-loaded knowledge files. Shared references live at the plugin root `references/`; module-specific ones at `skills/<module>/references/`.

## Plugin File Layout

```
.claude-plugin/          Plugin manifest (plugin.json, marketplace.json)
agents/                  9 agent definitions (.md with YAML frontmatter)
skills/                  9 skill directories, each with SKILL.md + references/
references/              Shared reference files
docs/CLAUDE.md           Runtime instructions loaded when the plugin activates
scripts/                 Setup scripts and Python dependencies
```

`docs/CLAUDE.md` is the **plugin's operational file** — it tells Claude how to route users to agents and how modules interact at runtime. This root `CLAUDE.md` is for **developers working on the plugin itself**.

## Module Overview

| Module | Agent | Model | Status |
|--------|-------|-------|--------|
| word-orchestrator | Task router + Token budget | sonnet | ✅ |
| word-reader | Document analysis + comparison | sonnet | ✅ |
| word-formatter | Formatting + TOC + CJK normalization | sonnet | ✅ |
| word-checker | Cross-reference + format compliance | sonnet | ✅ |
| word-content-editor | Content editing | opus | ✅ |
| word-table-figure | Tables + figures | sonnet | ✅ |
| word-reference | References + Zotero | sonnet | ✅ |
| word-reviewer | Review handling + revision | opus | ✅ |
| word-submit | Submission prep + split/merge | sonnet | ✅ |

## Key Design Principles

1. **Document Map Caching** — word-reader generates a structural summary once; all downstream modules reuse it, never re-read the full document
2. **MCP First, XML Fallback** — Use Word MCP Server tools for direct operations; only fall back to XML unpack/edit/pack for tracked changes
3. **User-Provided Format Specs** — No pre-built journal templates; parse the user's format requirement document into executable rules
4. **Token Budget Enforcement** — Lazy loading, batched operations, tool priority routing

## External Dependencies

- **Python:** docx2python, deepdiff, redlines (P0); python-docx-replace (P2); docxcompose, citeproc-py (P3)
- **MCP Servers:** word-document-server (required), Zotero (optional), Semantic Scholar (optional)
- **Existing Skills:** docx skill (XML operations layer)

See `scripts/requirements.txt` for pip install commands.

## Development Workflow

### Install locally for testing
```bash
pip install -r scripts/requirements.txt
claude plugin install /path/to/Word-agent
```

### Test a module
```
# In Claude Code:
/word-agent:word-read
Analyze this document: paper.docx
```

## Cross-References to Verify After Edits

| If you edit... | Also check... |
|---------------|--------------|
| An agent's tools list | The corresponding SKILL.md tool usage |
| `docs/CLAUDE.md` routing table | Agent filenames and skill names |
| A shared reference file | All SKILL.md files that reference it |
| Plugin manifests | Version numbers match across files |
| `references/tool_routing.md` | All SKILL.md tool selection logic |
| `references/doc_conversion.md` | `skills/word-orchestrate/SKILL.md` Phase 2, `agents/word-orchestrator.md` .doc handling |
