"""Shared test fixtures for md-word-pipeline."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest
from docx import Document

from lib.build.builder import DocumentBuilder


@pytest.fixture
def builder_factory(tmp_path):
    """Return a factory that builds a DocumentBuilder with optional config."""
    def _make(config=None):
        b = DocumentBuilder(config=config or {}, source_dir=tmp_path)
        b.setup()
        return b
    return _make


@pytest.fixture
def build_docx(tmp_path):
    """Render markdown to a temp .docx and return a loaded Document.

    The source_dir used for relative image resolution defaults to tmp_path so
    tests that reference images can drop fixture files there.
    """
    def _render(md_text, *, config=None, source_dir=None):
        src = Path(source_dir) if source_dir else tmp_path
        b = DocumentBuilder(config=config or {}, source_dir=src)
        b.setup()
        b.add_content(md_text, src)
        out = tmp_path / "out.docx"
        b.save(out)
        return Document(str(out))
    return _render
