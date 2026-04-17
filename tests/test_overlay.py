"""End-to-end tests: :::overlay markdown → Word grouped drawing."""
from pathlib import Path

from lib.build.overlays import NS_A, NS_WPG, NS_WPS


def _make_png(path: Path, w=200, h=100):
    from PIL import Image
    Image.new("RGB", (w, h), color=(200, 200, 255)).save(path)


def _find_drawings(doc):
    """Return every <w:drawing> element in the body."""
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    return doc.element.body.findall(f".//{{{W}}}drawing")


def test_overlay_block_produces_group_drawing(tmp_path, build_docx):
    img = tmp_path / "base.png"
    _make_png(img)
    md = (
        ":::overlay {#fig-demo}\n"
        "![Base](base.png){.small}\n"
        "::arrow from=10%,10% to=80%,20% color=#FF0000 stroke=2\n"
        ":::\n"
    )
    doc = build_docx(md, source_dir=str(tmp_path))
    drawings = _find_drawings(doc)
    assert drawings, "expected at least one drawing element"

    # At least one drawing uses the group graphicData URI
    grouped = [d for d in drawings if d.find(f".//{{{NS_A}}}graphicData") is not None
               and d.find(f".//{{{NS_A}}}graphicData").get("uri") == NS_WPG]
    assert grouped, "expected a grouped drawing with wpg:wgp graphicData"

    wgp = grouped[0].find(f".//{{{NS_WPG}}}wgp")
    assert wgp is not None

    # Group must contain the picture and exactly one shape
    NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    assert len(wgp.findall(f"{{{NS_PIC}}}pic")) == 1
    assert len(wgp.findall(f"{{{NS_WPS}}}wsp")) == 1


def test_overlay_multiple_shapes_combine(tmp_path, build_docx):
    img = tmp_path / "base.png"
    _make_png(img)
    md = (
        ":::overlay\n"
        "![](base.png){.medium}\n"
        "::rect at=10%,10% size=30%,30% color=#FF0000\n"
        "::ellipse at=55%,50% size=20%,20% color=#00FF00\n"
        '::callout at=70%,80% size=25%,10% text="Hi" color=#0000FF fill=#FFFF99\n'
        "::arrow from=10%,90% to=80%,90% color=#000000\n"
        ":::\n"
    )
    doc = build_docx(md, source_dir=str(tmp_path))
    grouped = [
        d for d in _find_drawings(doc)
        if d.find(f".//{{{NS_A}}}graphicData") is not None
        and d.find(f".//{{{NS_A}}}graphicData").get("uri") == NS_WPG
    ]
    assert grouped
    wsps = grouped[0].findall(f".//{{{NS_WPS}}}wsp")
    assert len(wsps) == 4


def test_overlay_without_shapes_still_embeds_image(tmp_path, build_docx):
    """An :::overlay block with no shape lines should still render the base image.

    (The wrap step is skipped when there are no shapes — the picture is
    embedded as a normal single image.)
    """
    img = tmp_path / "base.png"
    _make_png(img)
    md = (
        ":::overlay\n"
        "![](base.png)\n"
        ":::\n"
    )
    doc = build_docx(md, source_dir=str(tmp_path))
    drawings = _find_drawings(doc)
    assert drawings, "expected at least one drawing even without shapes"


def test_overlay_missing_image_logs_warning(tmp_path, build_docx, caplog):
    import logging
    caplog.set_level(logging.WARNING, logger="md_word_pipeline.builder")
    md = (
        ":::overlay\n"
        "![](missing.png)\n"
        "::arrow from=0%,0% to=100%,100% color=#FF0000\n"
        ":::\n"
    )
    doc = build_docx(md, source_dir=str(tmp_path))
    # Document still builds; we see an 'image not found' warning.
    msgs = " ".join(r.getMessage() for r in caplog.records)
    assert "missing.png" in msgs
    # And the fallback placeholder paragraph appears
    assert any("Image not found" in p.text for p in doc.paragraphs)
