"""Unit tests for the overlay parser and shape XML builders."""
import pytest
from lxml import etree

from lib.build.overlays import (
    parse_overlay_block,
    wrap_picture_in_group,
    extract_overlay_from_group,
    overlay_to_markdown,
    ShapeSpec,
    OverlaySpec,
    NS_A,
    NS_WPG,
    NS_WPS,
)


# ── parser ────────────────────────────────────────────────────────────────────

def test_parse_basic_overlay():
    md = (
        ":::overlay {#fig-login width=medium}\n"
        "![Login screen](screens/login.png)\n"
        "::arrow from=20%,30% to=50%,35% color=#FF0000 stroke=2\n"
        "::rect at=10%,20% size=30%,10% color=#FF0000\n"
        ":::\n"
    )
    spec = parse_overlay_block(md)
    assert spec.base_src == "screens/login.png"
    assert spec.base_alt == "Login screen"
    assert spec.attrs.get("id") == "fig-login"
    assert spec.attrs.get("width") == "medium"
    assert len(spec.shapes) == 2

    arrow = spec.shapes[0]
    assert arrow.kind == "arrow"
    assert arrow.pos == (20.0, 30.0)
    assert arrow.to == (50.0, 35.0)
    assert arrow.color == "FF0000"
    assert arrow.stroke == 2.0

    rect = spec.shapes[1]
    assert rect.kind == "rect"
    assert rect.pos == (10.0, 20.0)
    assert rect.size == (30.0, 10.0)


def test_parse_callout_with_text_and_fill():
    md = (
        ":::overlay\n"
        "![](x.png)\n"
        '::callout at=60%,40% size=22%,10% text="Click here" color=#111 fill=#FFFF00\n'
        ":::\n"
    )
    spec = parse_overlay_block(md)
    assert len(spec.shapes) == 1
    c = spec.shapes[0]
    assert c.kind == "callout"
    assert c.text == "Click here"
    assert c.fill == "FFFF00"


def test_parse_rejects_unknown_shape():
    md = ":::overlay\n![](x.png)\n::triangle at=10%,10%\n:::\n"
    # Unknown shapes are simply skipped (no crash) — verify parser tolerates them
    spec = parse_overlay_block(md)
    assert spec.shapes == []


def test_parse_requires_from_and_to_for_arrow():
    md = ":::overlay\n![](x.png)\n::arrow at=10%,10%\n:::\n"
    with pytest.raises(ValueError):
        parse_overlay_block(md)


# ── wrap + extract round trip ────────────────────────────────────────────────

def _make_fake_drawing(pic_rid="rId42"):
    """Build a minimal <w:drawing> with one <pic:pic>, mimicking what
    python-docx produces inside a run after add_picture."""
    NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    xml = f'''<w:drawing xmlns:w="{NS_W}" xmlns:wp="{NS_WP}" xmlns:a="{NS_A}" xmlns:pic="{NS_PIC}" xmlns:r="{NS_R}">
  <wp:inline>
    <wp:extent cx="3000000" cy="2000000"/>
    <wp:docPr id="1" name="Picture 1"/>
    <wp:cNvGraphicFramePr/>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic>
          <pic:nvPicPr>
            <pic:cNvPr id="1" name="base"/>
            <pic:cNvPicPr/>
          </pic:nvPicPr>
          <pic:blipFill>
            <a:blip r:embed="{pic_rid}"/>
            <a:stretch><a:fillRect/></a:stretch>
          </pic:blipFill>
          <pic:spPr>
            <a:xfrm>
              <a:off x="0" y="0"/>
              <a:ext cx="3000000" cy="2000000"/>
            </a:xfrm>
            <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          </pic:spPr>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>'''
    return etree.fromstring(xml)


def test_wrap_picture_in_group_produces_wgp_with_picture_and_shapes():
    drawing = _make_fake_drawing()
    shapes = [
        ShapeSpec(kind="arrow", pos=(20, 30), to=(50, 35), color="FF0000"),
        ShapeSpec(kind="rect",  pos=(10, 20), size=(30, 10), color="00FF00"),
    ]
    wrap_picture_in_group(drawing, shapes, w_emu=3_000_000, h_emu=2_000_000)

    # graphicData uri switches to the group namespace
    graphic_data = drawing.find(f".//{{{NS_A}}}graphicData")
    assert graphic_data.get("uri") == NS_WPG

    # Exactly one <wpg:wgp> inside
    wgp = graphic_data.find(f"{{{NS_WPG}}}wgp")
    assert wgp is not None

    # wgp should contain grpSpPr + one <pic:pic> + two <wps:wsp>
    NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    assert len(wgp.findall(f"{{{NS_PIC}}}pic")) == 1
    wsps = wgp.findall(f"{{{NS_WPS}}}wsp")
    assert len(wsps) == 2


def test_wrap_picture_adjusts_pic_xfrm_to_group_extent():
    drawing = _make_fake_drawing()
    wrap_picture_in_group(drawing, [], w_emu=1_000_000, h_emu=500_000)
    NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    pic_ext = drawing.find(f".//{{{NS_WPG}}}wgp/{{{NS_PIC}}}pic/{{{NS_PIC}}}spPr/{{{NS_A}}}xfrm/{{{NS_A}}}ext")
    assert pic_ext is not None
    assert int(pic_ext.get("cx")) == 1_000_000
    assert int(pic_ext.get("cy")) == 500_000


def test_extract_overlay_round_trip():
    """Build a group → extract → verify shape specs come back close to inputs."""
    drawing = _make_fake_drawing()
    w, h = 3_000_000, 2_000_000
    originals = [
        ShapeSpec(kind="arrow", pos=(20, 30), to=(70, 60), color="FF0000"),
        ShapeSpec(kind="rect",  pos=(10, 15), size=(25, 30), color="00AA00"),
        ShapeSpec(kind="ellipse", pos=(50, 50), size=(20, 20), color="0000FF"),
    ]
    wrap_picture_in_group(drawing, originals, w_emu=w, h_emu=h)
    extracted = extract_overlay_from_group(drawing, group_w_emu=w, group_h_emu=h)

    assert len(extracted) == len(originals)
    for orig, got in zip(originals, extracted):
        assert got.kind == orig.kind
        # Percent values should round-trip within 0.1%
        assert abs(got.pos[0] - orig.pos[0]) < 0.1
        assert abs(got.pos[1] - orig.pos[1]) < 0.1
        if orig.kind == "arrow":
            assert got.to is not None
            assert abs(got.to[0] - orig.to[0]) < 0.1
            assert abs(got.to[1] - orig.to[1]) < 0.1
        else:
            assert abs(got.size[0] - orig.size[0]) < 0.1
            assert abs(got.size[1] - orig.size[1]) < 0.1
        assert got.color.upper() == orig.color.upper()


def test_overlay_to_markdown_produces_parseable_block():
    spec = OverlaySpec(
        base_src="img.png",
        base_alt="alt",
        attrs={"id": "fig", "classes": [], "width": "medium"},
        shapes=[
            ShapeSpec(kind="arrow", pos=(20, 30), to=(50, 35), color="FF0000"),
            ShapeSpec(kind="rect", pos=(10, 20), size=(30, 10), color="00FF00"),
        ],
    )
    md = overlay_to_markdown(spec)
    assert md.startswith(":::overlay")
    assert "![alt](img.png)" in md
    assert md.endswith(":::")

    # Round-trip: parse back and compare
    reparsed = parse_overlay_block(md)
    assert reparsed.base_src == "img.png"
    assert reparsed.base_alt == "alt"
    assert len(reparsed.shapes) == 2
