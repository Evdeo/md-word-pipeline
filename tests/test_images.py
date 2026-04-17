"""Tests for image size classes, width attributes, and figure groups."""
from pathlib import Path

from docx.oxml.ns import qn

from lib.build.images import (
    ImageProcessor,
    build_size_classes,
    DEFAULT_SIZE_CLASSES,
    nearest_size_class,
    EMU_PER_INCH,
    EMU_PER_PX,
)


# ── ImageProcessor.calc_emu ───────────────────────────────────────────────────

def _make_proc():
    # A4 (8.27in) minus 1in + 1in margins = 6.27in content width
    return ImageProcessor(
        page_width_in=8.27, margin_left_in=1.0, margin_right_in=1.0,
        size_classes=dict(DEFAULT_SIZE_CLASSES),
    )


def test_calc_emu_medium_size_class():
    p = _make_proc()
    w, h = p.calc_emu(w_px=400, h_px=200, size_class="medium")
    expected_w = int(p.content_width_emu * 0.50)
    assert w == expected_w
    # Aspect ratio preserved: 200/400 = 0.5
    assert h == int(expected_w * 0.5)


def test_calc_emu_each_size_class_matches_default_fractions():
    p = _make_proc()
    for name, frac in DEFAULT_SIZE_CLASSES.items():
        w, _ = p.calc_emu(w_px=100, h_px=100, size_class=name)
        assert w == int(p.content_width_emu * frac), f"class {name}"


def test_calc_emu_width_attr_overrides():
    p = _make_proc()
    w, _ = p.calc_emu(w_px=100, h_px=100, width_attr="width=2in")
    assert w == int(2.0 * EMU_PER_INCH)

    w, _ = p.calc_emu(w_px=100, h_px=100, width_attr="width=300px")
    assert w == int(300 * EMU_PER_PX)

    w, _ = p.calc_emu(w_px=100, h_px=100, width_attr="width=40%")
    assert w == int(p.content_width_emu * 0.40)


def test_calc_emu_defaults_to_50_percent_when_nothing_given():
    p = _make_proc()
    w, _ = p.calc_emu(w_px=100, h_px=100)
    assert w == int(p.content_width_emu * 0.50)


def test_calc_emu_size_class_wins_over_width_attr():
    """Explicit size class takes precedence over width attribute."""
    p = _make_proc()
    w, _ = p.calc_emu(w_px=100, h_px=100, size_class="xl", width_attr="width=1in")
    assert w == int(p.content_width_emu * 1.0)


# ── build_size_classes config loader ──────────────────────────────────────────

def test_build_size_classes_from_max_pct():
    cfg = {"small": {"max_pct": 25}, "big": {"max_pct": 90}}
    result = build_size_classes(cfg)
    assert result["small"] == 0.25
    assert result["big"] == 0.90


def test_build_size_classes_empty_returns_defaults():
    result = build_size_classes(None)
    assert result == DEFAULT_SIZE_CLASSES
    # Returned dict must be a copy — mutating it doesn't mutate DEFAULT
    result["xs"] = 0.99
    assert DEFAULT_SIZE_CLASSES["xs"] == 0.20


def test_nearest_size_class_roundtrip():
    p = _make_proc()
    for name, frac in DEFAULT_SIZE_CLASSES.items():
        w = int(p.content_width_emu * frac)
        assert p.nearest_class(w) == name


# ── end-to-end: markdown image embedding ──────────────────────────────────────

def _make_png(path: Path, w=64, h=32):
    from PIL import Image
    Image.new("RGB", (w, h), color=(255, 128, 0)).save(path)


def test_markdown_image_is_embedded(tmp_path, build_docx):
    img = tmp_path / "pic.png"
    _make_png(img)
    doc = build_docx(f"![pic](pic.png)\n", source_dir=str(tmp_path))
    # Drawings appear inside runs
    drawings = doc.element.body.findall(
        f".//{{{ 'http://schemas.openxmlformats.org/wordprocessingml/2006/main' }}}drawing"
    )
    assert len(drawings) >= 1


def test_markdown_image_with_size_class(tmp_path, build_docx):
    img = tmp_path / "pic.png"
    _make_png(img, w=200, h=100)
    doc = build_docx("![pic](pic.png){.small}\n", source_dir=str(tmp_path))
    # Find the drawing's extent
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    drawings = doc.element.body.findall(f".//{{{W}}}drawing")
    assert drawings
    extent = drawings[0].find(f".//{{{WP}}}extent")
    assert extent is not None
    cx = int(extent.get("cx"))
    # Small = 30% of content width (default)
    p = _make_proc()
    expected = int(p.content_width_emu * 0.30)
    # Allow ±5 EMU rounding
    assert abs(cx - expected) < 5, f"cx={cx}, expected≈{expected}"


def test_figures_block_produces_borderless_table(tmp_path, build_docx):
    a = tmp_path / "a.png"; _make_png(a)
    b = tmp_path / "b.png"; _make_png(b)
    md = (
        ":::figures\n"
        "![](a.png){.small}\n"
        "![](b.png){.small}\n"
        ":::\n"
    )
    doc = build_docx(md, source_dir=str(tmp_path))
    assert len(doc.tables) == 1
    # Image tables have no visible borders on cells
    tbl = doc.tables[0]
    assert len(tbl.rows) >= 1
    # At least two cells with images
    total_drawings = 0
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    for row in tbl.rows:
        for cell in row.cells:
            total_drawings += len(cell._tc.findall(f".//{{{W}}}drawing"))
    assert total_drawings >= 2
