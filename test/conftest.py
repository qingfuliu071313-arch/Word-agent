"""Shared test fixtures for word-agent test suite."""
import os
import shutil
import pytest
from pathlib import Path

FIXTURES_DIR = Path(__file__).parent / "fixtures"
TEST_OUTPUT_DIR = Path(__file__).parent / "output"


@pytest.fixture(autouse=True)
def setup_output_dir():
    """Create and clean output directory for each test."""
    TEST_OUTPUT_DIR.mkdir(exist_ok=True)
    yield
    # Keep output for inspection; clean on next run


@pytest.fixture
def fixtures_dir():
    return FIXTURES_DIR


@pytest.fixture
def output_dir():
    return TEST_OUTPUT_DIR


@pytest.fixture
def sample_docx(fixtures_dir, output_dir):
    """Copy a sample docx to output dir for modification tests."""
    src = fixtures_dir / "sample_paper.docx"
    if not src.exists():
        pytest.skip("sample_paper.docx fixture not found")
    dst = output_dir / "test_copy.docx"
    shutil.copy2(src, dst)
    return dst


@pytest.fixture
def scripts_dir():
    return Path(__file__).parent.parent / "scripts"


@pytest.fixture
def normalize_fonts_script(scripts_dir):
    script = scripts_dir / "normalize_fonts.py"
    if not script.exists():
        pytest.skip("normalize_fonts.py not found")
    return script
