"""Round-trip test: markdown :::overlay → docx → extract.py → markdown.

Validates that editing an overlay in Word (and re-importing the reviewed
docx) preserves the shape kinds, colors, and percent positions — the
same workflow as `section_diff` review but for annotated images.
"""
from pathlib import Path

import pytest

from lib.build.overlays import (
    parse_overlay_block,
    extract_overlay_from_group,
    NS_A,
    NS_WPG,
)


def _make_png(path: Path, w=400, h=200):
    from PIL import Image
    Image.new("RGB", (w, h), color=(220, 220, 220)).save(path)


def test_overlay_shapes_survive_docx_round_trip(tmp_path, build_docx):
    img = tmp_path / "base.png"
    _make_png(img)

    md = (
        ":::overlay {#fig-demo}\n"
        "![Annotated](base.png){.medium}\n"
        "::arrow from=10%,20% to=80%,30% color=#FF0000 stroke=2\n"
        "::rect at=5%,40% size=20%,15% color=#00FF00\n"
        '::callout at=60%,60% size=25%,12% text="Click" color=#0000FF fill=#FFFF00\n'
        ":::\n"
    )
    # First: round-trip through the markdown parser alone
    original = parse_overlay_block(md)
    assert len(original.shapes) == 3

    # Second: build a real docx and re-read the group drawing from the XML
    doc = build_docx(md, source_dir=str(tmp_path))
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    drawings = doc.element.body.findall(f".//{{{W}}}drawing")
    grouped = [
        d for d in drawings
        if d.find(f".//{{{NS_A}}}graphicData") is not None
        and d.find(f".//{{{NS_A}}}graphicData").get("uri") == NS_WPG
    ]
    assert grouped, "expected a grouped drawing"

    # Find group extent
    extent = grouped[0].find(
        f".//{{{{}}}}wp{{}}{{}}extent".format()
    ) if False else None
    # Use wp:extent directly
    WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    extent = grouped[0].find(f".//{{{WP}}}extent")
    assert extent is not None
    w_emu = int(extent.get("cx"))
    h_emu = int(extent.get("cy"))

    extracted = extract_overlay_from_group(grouped[0], w_emu, h_emu)
    assert len(extracted) == len(original.shapes)

    # Kinds match in order
    assert [s.kind for s in extracted] == ["arrow", "rect", "callout"]

    # Colors preserved
    for orig, got in zip(original.shapes, extracted):
        assert got.color.upper() == orig.color.upper()

    # Arrow positions survive within 0.5%
    arrow = extracted[0]
    assert abs(arrow.pos[0] - 10) < 0.5
    assert abs(arrow.to[0] - 80) < 0.5

    # Rect position and size survive
    rect = extracted[1]
    assert abs(rect.pos[0] - 5) < 0.5
    assert abs(rect.size[0] - 20) < 0.5

    # Callout text survives
    callout = extracted[2]
    assert callout.text == "Click"


def test_extract_module_emits_overlay_block_from_grouped_drawing(tmp_path, build_docx):
    """End-to-end with lib/extract.py: build an overlay docx and extract
    via the module used in the review workflow; the result must be a
    :::overlay block, not a plain ![]()."""
    from lib.extract import _extract_images_from_element
    from lib.build.images import DEFAULT_SIZE_CLASSES, EMU_PER_INCH

    img = tmp_path / "base.png"
    _make_png(img)
    md = (
        ":::overlay\n"
        "![](base.png){.medium}\n"
        "::rect at=10%,10% size=20%,20% color=#FF0000\n"
        ":::\n"
    )
    doc = build_docx(md, source_dir=str(tmp_path))
    # Find the paragraph element containing the drawing
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    paras = [p for p in doc.element.body.findall(f"{{{W}}}p")
             if p.find(f".//{{{W}}}drawing") is not None]
    assert paras

    images_dir = tmp_path / "extracted_images"
    results = _extract_images_from_element(
        paras[0], doc, images_dir,
        size_classes=DEFAULT_SIZE_CLASSES,
        content_width_emu=int(6.27 * EMU_PER_INCH),
        img_counter=[0],
    )
    assert results
    md_out, _, _ = results[0]
    assert md_out.startswith(":::overlay"), f"expected overlay block, got {md_out!r}"
    assert "::rect" in md_out
