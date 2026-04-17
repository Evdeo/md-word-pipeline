"""Tests for section-based diff used in the review workflow."""
from pathlib import Path

import pytest

from lib.section_diff import diff_documents, extract_sections


def _flatten(results, out=None):
    out = out or []
    for r in results:
        out.append(r)
        _flatten(r.children, out)
    return out


@pytest.fixture
def make_docx(tmp_path, build_docx):
    """Return a function that renders markdown to a named docx under tmp_path."""
    def _make(name, md, **kw):
        # Delegate to build_docx but rename the output
        doc = build_docx(md, **kw)
        out = tmp_path / name
        doc.save(str(out))
        return out
    return _make


def test_identical_documents_all_identical(make_docx):
    md = "# Intro\n\nHello.\n\n# Body\n\nBye.\n"
    a = make_docx("a.docx", md)
    b = make_docx("b.docx", md)
    results = diff_documents(a, b)
    flat = _flatten(results)
    assert flat, "at least cover + headings expected"
    statuses = {r.status for r in flat}
    # Cover and all headings should be identical
    assert statuses <= {"identical"}


def test_edited_section_marked_changed(make_docx):
    md_a = "# Intro\n\nHello.\n\n# Body\n\nOriginal text.\n"
    md_b = "# Intro\n\nHello.\n\n# Body\n\nEDITED text.\n"
    a = make_docx("a.docx", md_a)
    b = make_docx("b.docx", md_b)
    flat = _flatten(diff_documents(a, b))
    # "Body" section should be changed; "Intro" still identical
    body = next(r for r in flat if r.received and r.received.heading == "Body")
    intro = next(r for r in flat if r.received and r.received.heading == "Intro")
    assert body.status == "changed"
    assert intro.status == "identical"


def test_added_section(make_docx):
    md_a = "# A\n\nfoo\n"
    md_b = "# A\n\nfoo\n\n# B\n\nbar\n"
    a = make_docx("a.docx", md_a)
    b = make_docx("b.docx", md_b)
    flat = _flatten(diff_documents(a, b))
    statuses = {(r.status, r.received.heading if r.received else None) for r in flat}
    assert ("added", "B") in statuses


def test_removed_section(make_docx):
    md_a = "# A\n\nfoo\n\n# B\n\nbar\n"
    md_b = "# A\n\nfoo\n"
    a = make_docx("a.docx", md_a)
    b = make_docx("b.docx", md_b)
    flat = _flatten(diff_documents(a, b))
    statuses = {(r.status, r.baseline.heading if r.baseline else None) for r in flat}
    assert ("removed", "B") in statuses


def test_moved_section(make_docx):
    md_a = "# A\n\nfoo\n\n# B\n\nbar\n\n# C\n\nbaz\n"
    # Reorder so B moves to the end (common sections [A,B,C] vs [A,C,B])
    md_b = "# A\n\nfoo\n\n# C\n\nbaz\n\n# B\n\nbar\n"
    a = make_docx("a.docx", md_a)
    b = make_docx("b.docx", md_b)
    flat = _flatten(diff_documents(a, b))
    statuses = {(r.status, r.received.heading if r.received else None) for r in flat}
    # B moved or C moved — but something should be marked moved
    moved = {h for s, h in statuses if s in ("moved", "moved_changed")}
    assert moved, f"expected a moved section, got {statuses}"


def test_moved_and_changed(make_docx):
    md_a = "# A\n\nfoo\n\n# B\n\noriginal\n\n# C\n\nbaz\n"
    md_b = "# A\n\nfoo\n\n# C\n\nbaz\n\n# B\n\nEDITED\n"
    a = make_docx("a.docx", md_a)
    b = make_docx("b.docx", md_b)
    flat = _flatten(diff_documents(a, b))
    b_result = next(
        r for r in flat
        if r.received and r.received.heading == "B"
    )
    assert b_result.status == "moved_changed"


def test_extract_sections_preserves_heading_text(make_docx):
    md = "# First\n\n## Sub\n\n# Second\n"
    path = make_docx("a.docx", md)
    sections = extract_sections(path)
    headings = [s.heading for s in sections]
    # Cover + First + Second at top level
    assert "First" in headings
    assert "Second" in headings
