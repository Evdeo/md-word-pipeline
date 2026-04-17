"""End-to-end build of the bundled showcase project.

Instead of committing a binary .docx fixture (which would bloat the repo and
be hard to update) this test does a self-consistency check: build the
showcase twice into separate paths and assert that section_diff reports
all-identical. That single assertion covers:
  - full markdown → docx pipeline runs without raising
  - all modules (builder, tables, images, styles, extract) cooperate
  - output is deterministic (no timestamp / randomness leaks into XML)
"""
from pathlib import Path

import pytest

from lib.build_doc import (
    collect_files,
    load_config,
    load_document_info,
    load_all_yaml_files,
    substitute_properties,
)
from lib.build.builder import DocumentBuilder
from lib.section_diff import diff_documents


ROOT = Path(__file__).resolve().parent.parent
SHOWCASE = ROOT / "projects" / "md-to-docx-showcase" / "input"


def _build(output_path: Path):
    """Mirror the essential pieces of lib.build_doc.main() without argv."""
    config_path = SHOWCASE / "config.yaml"
    config = load_config(config_path)

    doc_info_path = SHOWCASE / "document-info.yaml"
    document_info, revisions = load_document_info(doc_info_path)
    if document_info:
        config["document"] = document_info

    EXCLUDED = {"config.yaml", "document-info.yaml", "revisions.yaml"}
    properties = load_all_yaml_files(SHOWCASE, exclude_files=EXCLUDED)
    for k, v in document_info.items():
        properties.setdefault(f"document.{k}", str(v))

    builder = DocumentBuilder(config=config, revisions=revisions)
    builder.setup()

    frontpage, content_files = collect_files(SHOWCASE)
    all_texts = [
        substitute_properties(f.read_text(encoding="utf-8"), properties)
        for f in content_files
    ]
    builder.prescan_labels(all_texts)

    if frontpage:
        fp_text = substitute_properties(frontpage.read_text(encoding="utf-8"), properties)
        builder.add_frontpage(fp_text, frontpage.parent)
        builder.add_toc()
    else:
        builder.add_toc()

    for cf, cf_text in zip(content_files, all_texts):
        builder.add_content(cf_text, cf.parent)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    builder.save(output_path)


@pytest.mark.skipif(not SHOWCASE.exists(), reason="showcase project not present")
def test_showcase_build_is_deterministic(tmp_path):
    a = tmp_path / "showcase-a.docx"
    b = tmp_path / "showcase-b.docx"
    _build(a)
    _build(b)
    assert a.exists() and a.stat().st_size > 0
    assert b.exists() and b.stat().st_size > 0

    results = diff_documents(a, b)

    def flatten(rs, out=None):
        out = out or []
        for r in rs:
            out.append(r)
            flatten(r.children, out)
        return out

    flat = flatten(results)
    non_identical = [r for r in flat if r.status != "identical"]
    if non_identical:
        details = "\n".join(
            f"  {r.status}: "
            f"base={r.baseline.heading if r.baseline else '-'} / "
            f"recv={r.received.heading if r.received else '-'}"
            for r in non_identical
        )
        pytest.fail(f"Showcase build is not deterministic:\n{details}")


@pytest.mark.skipif(not SHOWCASE.exists(), reason="showcase project not present")
def test_showcase_has_expected_sections(tmp_path):
    out = tmp_path / "showcase.docx"
    _build(out)
    from lib.section_diff import extract_sections
    sections = extract_sections(out)
    heading_keys = {s.key for s in sections}
    # The showcase's top-level headings should include Overview / Features / etc.
    # Rather than over-specify, just assert a healthy number.
    assert len(sections) >= 3, f"expected several top-level sections, got {len(sections)}"
