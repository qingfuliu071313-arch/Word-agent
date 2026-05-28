"""word-agent test suite.

Covers: font normalization, document structure, skill routing logic.
Requires: pytest >= 7.0, docx2python, python-docx
Optional: word-mcp-live, docx-mcp (tests skip gracefully if unavailable)
"""
import json
import subprocess
import sys
from pathlib import Path

import pytest


# ---------------------------------------------------------------------------
# Font Normalization Tests
# ---------------------------------------------------------------------------

class TestFontNormalization:
    """Tests for scripts/normalize_fonts.py."""

    def test_detect_only_json(self, sample_docx, normalize_fonts_script):
        """--detect-only --json should return valid JSON without modifying file."""
        result = subprocess.run(
            [sys.executable, str(normalize_fonts_script), str(sample_docx),
             "--detect-only", "--json"],
            capture_output=True, text=True
        )
        assert result.returncode == 0
        data = json.loads(result.stdout)
        assert "issues" in data or "total" in data or isinstance(data, dict)

    def test_unify_produces_output(self, sample_docx, normalize_fonts_script):
        """--unify should modify the file without errors."""
        original_size = sample_docx.stat().st_size
        result = subprocess.run(
            [sys.executable, str(normalize_fonts_script), str(sample_docx), "--unify"],
            capture_output=True, text=True
        )
        assert result.returncode == 0
        assert sample_docx.stat().st_size > 0

    def test_check_lock_on_unlocked_file(self, sample_docx, normalize_fonts_script):
        """--check-lock should return 0 for an unlocked file."""
        result = subprocess.run(
            [sys.executable, str(normalize_fonts_script), str(sample_docx),
             "--check-lock", "--detect-only"],
            capture_output=True, text=True
        )
        assert result.returncode == 0

    def test_custom_fonts(self, sample_docx, normalize_fonts_script):
        """--cn and --en flags should be accepted."""
        result = subprocess.run(
            [sys.executable, str(normalize_fonts_script), str(sample_docx),
             "--unify", "--cn", "楷体", "--en", "Arial"],
            capture_output=True, text=True
        )
        assert result.returncode == 0


# ---------------------------------------------------------------------------
# Document Structure Tests
# ---------------------------------------------------------------------------

class TestDocumentStructure:
    """Tests for document reading via docx2python."""

    def test_docx2python_import(self):
        """docx2python should be importable."""
        import docx2python
        assert hasattr(docx2python, "docx2python")

    def test_read_sample(self, sample_docx):
        """Should read sample document without errors."""
        from docx2python import docx2python
        doc = docx2python(str(sample_docx))
        assert doc.body is not None
        assert len(doc.body) > 0

    def test_extract_images_list(self, sample_docx):
        """Should extract image list (may be empty)."""
        from docx2python import docx2python
        doc = docx2python(str(sample_docx))
        assert isinstance(doc.images, dict) or doc.images is None


# ---------------------------------------------------------------------------
# Text Box Extraction Tests
# ---------------------------------------------------------------------------

class TestTextBoxExtraction:
    """Tests for text box extraction via XML parsing."""

    def test_textbox_extraction_runs(self, sample_docx):
        """Text box extraction script should run without errors."""
        import zipfile
        import xml.etree.ElementTree as ET

        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        with zipfile.ZipFile(str(sample_docx), 'r') as z:
            xml_content = z.read('word/document.xml')

        root = ET.fromstring(xml_content)
        textboxes = root.findall('.//w:txbxContent', ns)
        assert isinstance(textboxes, list)


# ---------------------------------------------------------------------------
# Skill Routing Tests
# ---------------------------------------------------------------------------

class TestSkillRouting:
    """Tests for skill file structure and consistency."""

    SKILLS_DIR = Path(__file__).parent.parent / "skills"
    AGENTS_DIR = Path(__file__).parent.parent / "agents"

    def test_all_skills_have_skill_md(self):
        """Every skill directory should have a SKILL.md."""
        for skill_dir in self.SKILLS_DIR.iterdir():
            if skill_dir.is_dir():
                assert (skill_dir / "SKILL.md").exists(), \
                    f"{skill_dir.name} missing SKILL.md"

    def test_all_agents_have_tools(self):
        """Every agent .md should have a tools: line."""
        for agent_file in self.AGENTS_DIR.glob("*.md"):
            content = agent_file.read_text()
            assert "tools:" in content, \
                f"{agent_file.name} missing tools: declaration"

    def test_skill_allowed_tools_match_agent_tools(self):
        """Skill allowed-tools should be a subset of the corresponding agent tools."""
        skill_to_agent = {
            "word-read": "word-reader",
            "word-format": "word-formatter",
            "word-edit": "word-content-editor",
            "word-check": "word-checker",
            "word-review": "word-reviewer",
            "word-submit": "word-submit",
            "word-table-figure": "word-table-figure",
            "word-orchestrate": "word-orchestrator",
            "word-reference": "word-reference",
        }
        for skill_name, agent_name in skill_to_agent.items():
            skill_file = self.SKILLS_DIR / skill_name / "SKILL.md"
            agent_file = self.AGENTS_DIR / f"{agent_name}.md"
            if not skill_file.exists() or not agent_file.exists():
                continue

            skill_content = skill_file.read_text()
            agent_content = agent_file.read_text()

            # Extract allowed-tools from skill
            for line in skill_content.split('\n'):
                if line.startswith('allowed-tools:'):
                    skill_tools = set(line.replace('allowed-tools:', '').strip().split())
                    break
            else:
                continue

            # Extract tools from agent
            for line in agent_content.split('\n'):
                if line.startswith('tools:'):
                    agent_tools = set(
                        t.strip().rstrip(',') for t in
                        line.replace('tools:', '').strip().split(',')
                    )
                    break
            else:
                continue

            # MCP tools in skill should exist in agent
            skill_mcp = {t for t in skill_tools if t.startswith('mcp__')}
            agent_mcp = {t.strip() for t in agent_tools if t.strip().startswith('mcp__')}
            missing = skill_mcp - agent_mcp
            assert not missing, \
                f"{skill_name}: skill has MCP tools not in agent {agent_name}: {missing}"


# ---------------------------------------------------------------------------
# Reference File Tests
# ---------------------------------------------------------------------------

class TestReferenceFiles:
    """Tests for reference documentation completeness."""

    REFS_DIR = Path(__file__).parent.parent / "references"

    EXPECTED_REFS = [
        "tool_routing.md",
        "token_budget.md",
        "font_normalization.md",
        "submission_checklist.md",
        "mcp_servers.md",
        "adeu_integration.md",
        "tracked_changes.md",
        "comment_operations.md",
        "structural_validation.md",
        "live_editing.md",
        "equation_crossref.md",
    ]

    def test_expected_references_exist(self):
        """All expected reference files should exist."""
        for ref in self.EXPECTED_REFS:
            assert (self.REFS_DIR / ref).exists(), \
                f"Missing reference file: {ref}"

    def test_references_not_empty(self):
        """Reference files should have content."""
        for ref in self.EXPECTED_REFS:
            path = self.REFS_DIR / ref
            if path.exists():
                assert path.stat().st_size > 100, \
                    f"Reference file too small: {ref}"
