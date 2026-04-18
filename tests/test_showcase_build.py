"""End-to-end build of the bundled showcase project.

Self-consistency check: build the showcase twice into separate paths and
assert that section_diff reports all-identical. That single assertion
covers the full markdown → docx pipeline and that output is deterministic.
"""
from pathlib import Path

import pytest

from lib.config_loader import build, load_project
from lib.section_diff import diff_documents, extract_sections


ROOT = Path(__file__).resolve().parent.parent
SHOWCASE_DIR = ROOT / "projects" / "showcase"


@pytest.mark.skipif(not SHOWCASE_DIR.is_dir(),
                    reason="showcase project not present")
def test_showcase_build_is_deterministic(tmp_path):
    ctx = load_project(SHOWCASE_DIR)
    a = tmp_path / "showcase-a.docx"
    b = tmp_path / "showcase-b.docx"
    build(ctx, out_path=a)
    build(ctx, out_path=b)

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


@pytest.mark.skipif(not SHOWCASE_DIR.is_dir(),
                    reason="showcase project not present")
def test_showcase_has_expected_sections(tmp_path):
    ctx = load_project(SHOWCASE_DIR)
    out = tmp_path / "showcase.docx"
    build(ctx, out_path=out)
    sections = extract_sections(out)
    assert len(sections) >= 3, f"expected several top-level sections, got {len(sections)}"
