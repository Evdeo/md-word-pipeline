"""Native Word grouped-shape overlays on images.

Markdown syntax::

    :::overlay {#fig-login width=medium}
    ![Login screen](screens/login.png)
    ::arrow    from=20%,30% to=50%,35% color=#FF0000 stroke=2
    ::rect     at=10%,20% size=30%,10% color=#FF0000 stroke=2
    ::ellipse  at=55%,45% size=12%,12% color=#FF0000 stroke=2
    ::callout  at=60%,40% size=22%,10% text="Click here" color=#FFFF00
    :::

The builder calls :func:`parse_overlay_block` on the raw block text and
receives an :class:`OverlaySpec`. After embedding the base image via
python-docx's ``run.add_picture``, :func:`wrap_picture_in_group` rewrites
the generated ``<w:drawing>`` so the picture becomes one child of a
``<wpg:wgp>`` group and each shape spec becomes a ``<wps:wsp>`` sibling.

Coordinates are percent-of-base-image; converted to EMU relative to the
group extent at emission time. Shapes remain editable in Word — a
reviewer can nudge an arrow and :func:`extract_overlay_from_group` reads
the adjusted coords back into markdown.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from lxml import etree

from ..log import get_logger

log = get_logger(__name__)


# ── Namespace constants ──────────────────────────────────────────────────────

NS_W   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NS_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

# Map overlay shape kinds → OOXML preset geometries and back.
KIND_TO_PRST = {
    "rect":    "rect",
    "ellipse": "ellipse",
    "callout": "wedgeRectCallout",
    "arrow":   "straightConnector1",   # handled specially (line + arrowhead)
}
PRST_TO_KIND = {v: k for k, v in KIND_TO_PRST.items()}


# ── Specs ────────────────────────────────────────────────────────────────────

@dataclass
class ShapeSpec:
    """One overlay shape. All position/size values are percents of the base
    image (0-100). `arrow` uses from_pct/to_pct; others use pos + size."""
    kind: str                                   # rect|ellipse|callout|arrow
    pos:  Tuple[float, float] = (0.0, 0.0)      # (x%, y%) — 'at' for rect/ellipse/callout, 'from' for arrow
    size: Tuple[float, float] = (10.0, 10.0)    # (w%, h%) — unused for arrow
    to:   Optional[Tuple[float, float]] = None  # (x%, y%) — arrow head for kind=arrow
    color: str = "FF0000"                       # 6-char hex, no leading #
    stroke: float = 2.0                         # stroke width in pt (converted to EMU)
    fill: Optional[str] = None                  # optional 6-char hex fill
    text: Optional[str] = None                  # callout label

    def normalized_color(self) -> str:
        return self.color.lstrip("#").upper()


@dataclass
class OverlaySpec:
    """Parsed `:::overlay` block."""
    base_src: str                           # image path from ![](path)
    base_alt: str = ""
    attrs: dict = field(default_factory=dict)  # {#id}, width=..., .small etc.
    shapes: List[ShapeSpec] = field(default_factory=list)
    caption: Optional[str] = None


# ── Parser ───────────────────────────────────────────────────────────────────

_ATTR_RE = re.compile(r"\{([^}]+)\}")
_IMG_RE = re.compile(r"^!\[([^\]]*)\]\(([^)]+)\)(\s*\{[^}]*\})?\s*$")
_SHAPE_LINE_RE = re.compile(r"^::(\w+)\s+(.*)$")
_KV_RE = re.compile(r"""(\w+)\s*=\s*(?:"([^"]*)"|(\S+))""")
_PCT_PAIR_RE = re.compile(r"^([\d.]+)%?\s*,\s*([\d.]+)%?$")


def _parse_attr_block(s: str) -> dict:
    """Parse ``{#id width=medium .small key="value"}`` into a dict.

    Anchors (``#name``) go under key ``id``; classes (``.foo``) accumulate
    under ``classes``; everything else is a key=value pair.
    """
    out: dict = {"classes": []}
    for token in re.findall(r"[#.]?[\w\-]+(?:=(?:\"[^\"]*\"|\S+))?", s):
        if token.startswith("#"):
            out["id"] = token[1:]
        elif token.startswith("."):
            out["classes"].append(token[1:])
        elif "=" in token:
            k, v = token.split("=", 1)
            out[k.strip()] = v.strip().strip('"')
    return out


def _parse_pct_pair(s: str) -> Tuple[float, float]:
    m = _PCT_PAIR_RE.match(s.strip())
    if not m:
        raise ValueError(f"expected 'X%,Y%' or 'X,Y', got {s!r}")
    return float(m.group(1)), float(m.group(2))


def _parse_shape_line(kind: str, rest: str) -> ShapeSpec:
    """Parse a single ``::kind key=value …`` line into a ShapeSpec."""
    kwargs = {}
    for m in _KV_RE.finditer(rest):
        key = m.group(1)
        val = m.group(2) if m.group(2) is not None else m.group(3)
        kwargs[key] = val

    if kind not in KIND_TO_PRST:
        raise ValueError(
            f"unknown overlay shape {kind!r}; expected one of {sorted(KIND_TO_PRST)}"
        )

    spec = ShapeSpec(kind=kind)

    if kind == "arrow":
        if "from" not in kwargs or "to" not in kwargs:
            raise ValueError("arrow requires 'from=X%,Y%' and 'to=X%,Y%'")
        spec.pos = _parse_pct_pair(kwargs["from"])
        spec.to = _parse_pct_pair(kwargs["to"])
    else:
        if "at" not in kwargs:
            raise ValueError(f"{kind} requires 'at=X%,Y%'")
        spec.pos = _parse_pct_pair(kwargs["at"])
        if "size" in kwargs:
            spec.size = _parse_pct_pair(kwargs["size"])

    if "color" in kwargs:
        spec.color = kwargs["color"].lstrip("#")
    if "fill" in kwargs:
        spec.fill = kwargs["fill"].lstrip("#")
    if "stroke" in kwargs:
        try:
            spec.stroke = float(kwargs["stroke"])
        except ValueError:
            pass
    if "text" in kwargs:
        spec.text = kwargs["text"]

    return spec


def parse_overlay_block(block_text: str) -> OverlaySpec:
    """Parse a raw `:::overlay` block (with the `:::` fences intact or stripped)."""
    # Strip opening `:::overlay` line and trailing `:::`
    lines = [l for l in block_text.splitlines() if l.strip()]
    if lines and lines[0].lstrip().startswith(":::overlay"):
        first = lines[0].lstrip()[len(":::overlay"):].strip()
        attrs = {}
        m = _ATTR_RE.search(first)
        if m:
            attrs = _parse_attr_block(m.group(1))
        lines = lines[1:]
    else:
        attrs = {}
    if lines and lines[-1].strip() == ":::":
        lines = lines[:-1]

    spec = OverlaySpec(base_src="", attrs=attrs)

    for line in lines:
        stripped = line.strip()
        img_m = _IMG_RE.match(stripped)
        if img_m and not spec.base_src:
            spec.base_alt = img_m.group(1)
            spec.base_src = img_m.group(2)
            # Fold any per-image attr block into the overlay attrs
            if img_m.group(3):
                inner = _ATTR_RE.search(img_m.group(3))
                if inner:
                    extra = _parse_attr_block(inner.group(1))
                    for k, v in extra.items():
                        if k == "classes":
                            spec.attrs.setdefault("classes", []).extend(v)
                        else:
                            spec.attrs.setdefault(k, v)
            continue
        shape_m = _SHAPE_LINE_RE.match(stripped)
        if shape_m:
            kind = shape_m.group(1)
            if kind not in KIND_TO_PRST:
                log.warning("overlay: ignoring unknown shape %r", kind)
                continue
            spec.shapes.append(_parse_shape_line(kind, shape_m.group(2)))

    return spec


# ── Shape XML emitters ───────────────────────────────────────────────────────

_SHAPE_ID_COUNTER = [100]  # simple global counter for shape IDs


def _next_id() -> int:
    _SHAPE_ID_COUNTER[0] += 1
    return _SHAPE_ID_COUNTER[0]


def _pct_to_emu(pct: float, total_emu: int) -> int:
    return int(round(total_emu * pct / 100.0))


def _pt_to_emu(pt: float) -> int:
    # EMU per pt: 914400 / 72 = 12700
    return int(pt * 12700)


def _shape_xml(spec: ShapeSpec, group_w_emu: int, group_h_emu: int) -> str:
    """Return a ``<wps:wsp>`` element (as string) for one shape."""
    sid = _next_id()
    color = spec.normalized_color()
    stroke_emu = _pt_to_emu(spec.stroke)

    if spec.kind == "arrow":
        return _arrow_xml(spec, group_w_emu, group_h_emu, sid, color, stroke_emu)

    # Rectangle-ish preset geometries.
    x_emu = _pct_to_emu(spec.pos[0], group_w_emu)
    y_emu = _pct_to_emu(spec.pos[1], group_h_emu)
    w_emu = _pct_to_emu(spec.size[0], group_w_emu)
    h_emu = _pct_to_emu(spec.size[1], group_h_emu)

    prst = KIND_TO_PRST[spec.kind]
    fill = spec.fill.lstrip("#").upper() if spec.fill else None
    fill_xml = (
        f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        if fill else '<a:noFill/>'
    )

    txbx_xml = ""
    if spec.text:
        txt_escaped = spec.text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        txbx_xml = (
            f'<wps:txbx><w:txbxContent xmlns:w="{NS_W}">'
            f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
            f'<w:r><w:t xml:space="preserve">{txt_escaped}</w:t></w:r>'
            f'</w:p></w:txbxContent></wps:txbx>'
            f'<wps:bodyPr wrap="square" anchor="ctr" anchorCtr="1"/>'
        )
    else:
        txbx_xml = '<wps:bodyPr/>'

    return (
        f'<wps:wsp xmlns:wps="{NS_WPS}" xmlns:a="{NS_A}" xmlns:w="{NS_W}">'
        f'<wps:cNvPr id="{sid}" name="{spec.kind} {sid}"/>'
        f'<wps:cNvSpPr/>'
        f'<wps:spPr>'
        f'<a:xfrm>'
        f'<a:off x="{x_emu}" y="{y_emu}"/>'
        f'<a:ext cx="{max(1, w_emu)}" cy="{max(1, h_emu)}"/>'
        f'</a:xfrm>'
        f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>'
        f'{fill_xml}'
        f'<a:ln w="{stroke_emu}">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'</a:ln>'
        f'</wps:spPr>'
        f'<wps:style>'
        f'<a:lnRef idx="0"/><a:fillRef idx="0"/>'
        f'<a:effectRef idx="0"/><a:fontRef idx="minor"/>'
        f'</wps:style>'
        f'{txbx_xml}'
        f'</wps:wsp>'
    )


def _arrow_xml(spec: ShapeSpec, group_w_emu: int, group_h_emu: int,
               sid: int, color: str, stroke_emu: int) -> str:
    assert spec.to is not None
    fx = _pct_to_emu(spec.pos[0], group_w_emu)
    fy = _pct_to_emu(spec.pos[1], group_h_emu)
    tx = _pct_to_emu(spec.to[0],  group_w_emu)
    ty = _pct_to_emu(spec.to[1],  group_h_emu)

    x_off = min(fx, tx)
    y_off = min(fy, ty)
    cx = max(1, abs(tx - fx))
    cy = max(1, abs(ty - fy))
    flipH = ' flipH="1"' if tx < fx else ''
    flipV = ' flipV="1"' if ty < fy else ''

    return (
        f'<wps:wsp xmlns:wps="{NS_WPS}" xmlns:a="{NS_A}">'
        f'<wps:cNvPr id="{sid}" name="Arrow {sid}"/>'
        f'<wps:cNvCnPr/>'
        f'<wps:spPr>'
        f'<a:xfrm{flipH}{flipV}>'
        f'<a:off x="{x_off}" y="{y_off}"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/>'
        f'</a:xfrm>'
        f'<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
        f'<a:ln w="{stroke_emu}">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:tailEnd type="triangle" w="med" len="med"/>'
        f'</a:ln>'
        f'</wps:spPr>'
        f'<wps:style>'
        f'<a:lnRef idx="0"/><a:fillRef idx="0"/>'
        f'<a:effectRef idx="0"/><a:fontRef idx="minor"/>'
        f'</wps:style>'
        f'<wps:bodyPr/>'
        f'</wps:wsp>'
    )


# ── Drawing wrapper ──────────────────────────────────────────────────────────

def wrap_picture_in_group(
    drawing_el, shapes: List[ShapeSpec], w_emu: int, h_emu: int
) -> None:
    """Mutate a ``<w:drawing>`` element so its picture becomes a wpg:wgp group
    with the picture as the first child and one wps:wsp per shape.

    Must be called on the drawing produced by ``run.add_picture`` so the
    image relationship (rId) is already registered in the docx package.
    """
    graphic_data = drawing_el.find(f".//{{{NS_A}}}graphicData")
    if graphic_data is None:
        return

    existing = list(graphic_data)
    if not existing:
        return
    original_pic = existing[0]
    for child in existing:
        graphic_data.remove(child)

    graphic_data.set("uri", NS_WPG)

    # Reset picture's inner xfrm to sit at the group origin with the group's
    # full extent — otherwise Word may render it at (0,0) with its own
    # intrinsic size which can overflow the group bbox.
    pic_xfrm = original_pic.find(f".//{{{NS_PIC}}}spPr/{{{NS_A}}}xfrm")
    if pic_xfrm is not None:
        off = pic_xfrm.find(f"{{{NS_A}}}off")
        ext = pic_xfrm.find(f"{{{NS_A}}}ext")
        if off is not None:
            off.set("x", "0"); off.set("y", "0")
        if ext is not None:
            ext.set("cx", str(w_emu)); ext.set("cy", str(h_emu))

    wgp_xml = (
        f'<wpg:wgp xmlns:wpg="{NS_WPG}" xmlns:a="{NS_A}">'
        f'<wpg:cNvGrpSpPr/>'
        f'<wpg:grpSpPr>'
        f'<a:xfrm>'
        f'<a:off x="0" y="0"/>'
        f'<a:ext cx="{w_emu}" cy="{h_emu}"/>'
        f'<a:chOff x="0" y="0"/>'
        f'<a:chExt cx="{w_emu}" cy="{h_emu}"/>'
        f'</a:xfrm>'
        f'</wpg:grpSpPr>'
        f'</wpg:wgp>'
    )
    wgp = etree.fromstring(wgp_xml)
    wgp.append(original_pic)
    for spec in shapes:
        wgp.append(etree.fromstring(_shape_xml(spec, w_emu, h_emu)))

    graphic_data.append(wgp)


# ── Reverse extraction: wpg:wgp → OverlaySpec shapes ─────────────────────────

def extract_overlay_from_group(drawing_el, group_w_emu: int, group_h_emu: int
                                ) -> List[ShapeSpec]:
    """Read back ShapeSpec objects from a wpg:wgp drawing.

    The caller is responsible for locating the image source and building
    the final markdown block; this function only recovers the shapes.
    Returns ``[]`` if the drawing isn't a group drawing.
    """
    graphic_data = drawing_el.find(f".//{{{NS_A}}}graphicData")
    if graphic_data is None or graphic_data.get("uri") != NS_WPG:
        return []

    wgp = graphic_data.find(f"{{{NS_WPG}}}wgp")
    if wgp is None:
        return []

    shapes: List[ShapeSpec] = []
    for wsp in wgp.findall(f"{{{NS_WPS}}}wsp"):
        spec = _read_shape(wsp, group_w_emu, group_h_emu)
        if spec is not None:
            shapes.append(spec)
    return shapes


def _read_shape(wsp, w_emu: int, h_emu: int) -> Optional[ShapeSpec]:
    sp_pr = wsp.find(f"{{{NS_WPS}}}spPr")
    if sp_pr is None:
        return None
    prst_el = sp_pr.find(f"{{{NS_A}}}prstGeom")
    prst = prst_el.get("prst") if prst_el is not None else None
    kind = PRST_TO_KIND.get(prst)
    if kind is None:
        return None

    xfrm = sp_pr.find(f"{{{NS_A}}}xfrm")
    if xfrm is None:
        return None
    off = xfrm.find(f"{{{NS_A}}}off")
    ext = xfrm.find(f"{{{NS_A}}}ext")
    x_emu = int(off.get("x", "0")) if off is not None else 0
    y_emu = int(off.get("y", "0")) if off is not None else 0
    cx = int(ext.get("cx", "0")) if ext is not None else 0
    cy = int(ext.get("cy", "0")) if ext is not None else 0
    flipH = xfrm.get("flipH") == "1"
    flipV = xfrm.get("flipV") == "1"

    # Color
    color = "000000"
    ln = sp_pr.find(f"{{{NS_A}}}ln")
    if ln is not None:
        sf = ln.find(f"{{{NS_A}}}solidFill/{{{NS_A}}}srgbClr")
        if sf is not None:
            color = sf.get("val", color)

    spec = ShapeSpec(kind=kind, color=color)
    if kind == "arrow":
        fx = x_emu + (cx if flipH else 0)
        fy = y_emu + (cy if flipV else 0)
        tx = x_emu + (0 if flipH else cx)
        ty = y_emu + (0 if flipV else cy)
        spec.pos = (100 * fx / w_emu, 100 * fy / h_emu)
        spec.to  = (100 * tx / w_emu, 100 * ty / h_emu)
    else:
        spec.pos = (100 * x_emu / w_emu, 100 * y_emu / h_emu)
        spec.size = (100 * cx / w_emu, 100 * cy / h_emu)
        # Fill
        solid = sp_pr.find(f"{{{NS_A}}}solidFill/{{{NS_A}}}srgbClr")
        if solid is not None:
            spec.fill = solid.get("val")
        # Text
        txbx = wsp.find(f"{{{NS_WPS}}}txbx")
        if txbx is not None:
            ts = txbx.findall(f".//{{{NS_W}}}t")
            if ts:
                spec.text = "".join((t.text or "") for t in ts)

    return spec


def overlay_to_markdown(spec: OverlaySpec) -> str:
    """Render an OverlaySpec back to its markdown block form."""
    head = ":::overlay"
    attr_bits = []
    if spec.attrs.get("id"):
        attr_bits.append(f"#{spec.attrs['id']}")
    for cls in spec.attrs.get("classes", []) or []:
        attr_bits.append(f".{cls}")
    for k, v in spec.attrs.items():
        if k in ("id", "classes"):
            continue
        attr_bits.append(f"{k}={v}")
    if attr_bits:
        head += " {" + " ".join(attr_bits) + "}"
    lines = [head, f"![{spec.base_alt}]({spec.base_src})"]
    for s in spec.shapes:
        if s.kind == "arrow":
            lines.append(
                f"::arrow from={s.pos[0]:.1f}%,{s.pos[1]:.1f}% "
                f"to={s.to[0]:.1f}%,{s.to[1]:.1f}% "
                f"color=#{s.normalized_color()} stroke={s.stroke:g}"
            )
        else:
            line = (
                f"::{s.kind} at={s.pos[0]:.1f}%,{s.pos[1]:.1f}% "
                f"size={s.size[0]:.1f}%,{s.size[1]:.1f}% "
                f"color=#{s.normalized_color()} stroke={s.stroke:g}"
            )
            if s.fill:
                line += f" fill=#{s.fill.lstrip('#').upper()}"
            if s.text:
                line += f' text="{s.text}"'
            lines.append(line)
    lines.append(":::")
    return "\n".join(lines)
