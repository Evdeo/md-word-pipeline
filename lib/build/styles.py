"""Document styles — headings, body, code, caption, blockquote.

define_styles(doc, style_cfg) applies hardcoded defaults first, then
overlays any values present in the optional style_cfg dict (drawn from
the ``styles:`` block in config.yaml).

Every key in style_cfg is optional — omitting it keeps the default.
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from typing import Optional


DARK_BLUE  = RGBColor(0x1F, 0x38, 0x64)
MID_BLUE   = RGBColor(0x2E, 0x75, 0xB6)
BLACK      = RGBColor(0x00, 0x00, 0x00)


def _get_or_add(styles, name, style_type=1):
    try:
        return styles[name]
    except KeyError:
        return styles.add_style(name, style_type)


def _hex_to_rgb(hex_str: str) -> RGBColor:
    """Convert a 6-digit hex string (with or without #) to RGBColor."""
    h = str(hex_str).lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _cfg_color(cfg: dict, key: str, default: RGBColor) -> RGBColor:
    """Return RGBColor from cfg[key] hex string, or default if absent."""
    val = cfg.get(key)
    if val is not None:
        try:
            return _hex_to_rgb(str(val))
        except Exception:
            pass
    return default


def _cfg_float(cfg: dict, key: str, default: float) -> float:
    try:
        return float(cfg.get(key, default))
    except (TypeError, ValueError):
        return default


def _cfg_bool(cfg: dict, key: str, default: bool) -> bool:
    val = cfg.get(key)
    if val is None:
        return default
    if isinstance(val, bool):
        return val
    return str(val).lower() in ("true", "1", "yes")


def _cfg_str(cfg: dict, key: str, default: str) -> str:
    val = cfg.get(key)
    return str(val) if val is not None else default


def _clear_heading_borders(s):
    """Remove any paragraph border (the H1 bottom rule) from a heading style."""
    s_pPr = s._element.get_or_add_pPr()
    existing_pBdr = s_pPr.find(qn("w:pBdr"))
    if existing_pBdr is not None:
        s_pPr.remove(existing_pBdr)
    pBdr = OxmlElement("w:pBdr")
    for side in ("top", "left", "bottom", "right"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),   "none")
        b.set(qn("w:sz"),    "0")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "auto")
        pBdr.append(b)
    s_pPr.append(pBdr)


def define_styles(doc: Document, style_cfg: Optional[dict] = None) -> None:
    """Build / update all document styles, applying overrides from style_cfg."""
    styles = doc.styles
    cfg = style_cfg or {}

    # ── Normal (body text) ──────────────────────────────────────────────────
    nc = cfg.get("normal", {}) or {}
    normal = _get_or_add(styles, "Normal")
    normal.font.name = _cfg_str(nc, "font_name", "Calibri")
    normal.font.size = Pt(_cfg_float(nc, "font_size_pt", 11))
    normal.paragraph_format.space_after = Pt(_cfg_float(nc, "space_after_pt", 6))

    # ── Headings ─────────────────────────────────────────────────────────────
    # Defaults: (style_name, cfg_key, size, color, bold, space_before, space_after, outline)
    heading_defs = [
        ("Heading 1", "heading_1", 22, DARK_BLUE, True,  12, 6,  0),
        ("Heading 2", "heading_2", 16, MID_BLUE,  True,  10, 4,  1),
        ("Heading 3", "heading_3", 13, DARK_BLUE, True,  8,  2,  2),
        ("Heading 4", "heading_4", 12, MID_BLUE,  True,  6,  2,  3),
        ("Heading 5", "heading_5", 11, DARK_BLUE, True,  4,  2,  4),
        ("Heading 6", "heading_6", 11, MID_BLUE,  False, 4,  2,  5),
    ]
    for style_name, cfg_key, def_size, def_color, def_bold, def_before, def_after, outline in heading_defs:
        hc = cfg.get(cfg_key, {}) or {}
        s = _get_or_add(styles, style_name)
        s.font.name      = _cfg_str(hc,   "font_name",       "Calibri")
        s.font.size      = Pt(_cfg_float(hc, "font_size_pt", def_size))
        s.font.bold      = _cfg_bool(hc,  "bold",            def_bold)
        s.font.color.rgb = _cfg_color(hc, "color",           def_color)
        s.paragraph_format.space_before  = Pt(_cfg_float(hc, "space_before_pt", def_before))
        s.paragraph_format.space_after   = Pt(_cfg_float(hc, "space_after_pt",  def_after))
        s.paragraph_format.outline_level = outline   # required for TOC
        _clear_heading_borders(s)

    # ── Code ─────────────────────────────────────────────────────────────────
    cc = cfg.get("code", {}) or {}
    code = _get_or_add(styles, "Code")
    code.font.name      = _cfg_str(cc,   "font_name",   "Courier New")
    code.font.size      = Pt(_cfg_float(cc, "font_size_pt", 9))
    code.font.color.rgb = BLACK
    code.paragraph_format.space_before = Pt(_cfg_float(cc, "space_before_pt", 2))
    code.paragraph_format.space_after  = Pt(_cfg_float(cc, "space_after_pt",  2))
    code.paragraph_format.left_indent  = Inches(_cfg_float(cc, "left_indent_in",  0.15))
    code.paragraph_format.right_indent = Inches(_cfg_float(cc, "right_indent_in", 0.15))

    # background shading
    bg_color = _cfg_str(cc, "background", "F0F0F0").lstrip("#")
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  bg_color)
    code._element.get_or_add_pPr().append(shd)

    # top + bottom thin border
    border_color = _cfg_str(cc, "border_color", "AAAAAA").lstrip("#")
    pPr  = code._element.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    for side in ("top", "bottom"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"),   "single")
        b.set(qn("w:sz"),    "6")
        b.set(qn("w:space"), "2")
        b.set(qn("w:color"), border_color)
        pBdr.append(b)
    pPr.append(pBdr)

    # ── Block Quote ──────────────────────────────────────────────────────────
    bqc = cfg.get("block_quote", {}) or {}
    bq  = _get_or_add(styles, "Block Quote")
    bq.font.italic      = _cfg_bool(bqc,  "font_italic",     True)
    bq.font.color.rgb   = _cfg_color(bqc, "color",           RGBColor(0x44, 0x44, 0x44))
    bq.paragraph_format.left_indent  = Inches(_cfg_float(bqc, "left_indent_in",  0.15))
    bq.paragraph_format.right_indent = Inches(_cfg_float(bqc, "right_indent_in", 0.15))
    bq.paragraph_format.space_before = Pt(_cfg_float(bqc, "space_before_pt", 4))
    bq.paragraph_format.space_after  = Pt(_cfg_float(bqc, "space_after_pt",  4))

    # left accent bar
    bar_color = _cfg_str(bqc, "bar_color", "2E75B6").lstrip("#")
    pPr2  = bq._element.get_or_add_pPr()
    pBdr2 = OxmlElement("w:pBdr")
    left  = OxmlElement("w:left")
    left.set(qn("w:val"),   "single")
    left.set(qn("w:sz"),    "12")
    left.set(qn("w:space"), "8")
    left.set(qn("w:color"), bar_color)
    pBdr2.append(left)
    pPr2.append(pBdr2)

    # ── Cover page styles (NOT in heading family → never appear in TOC) ──────
    cover_defs = [
        ("Cover Title",    "cover_title",    22, DARK_BLUE, True,  24, 8),
        ("Cover Subtitle", "cover_subtitle", 14, MID_BLUE,  False, 10, 6),
        ("Cover Body",     "cover_body",     11, BLACK,     False,  6, 4),
    ]
    for style_name, cfg_key, def_size, def_color, def_bold, def_before, def_after in cover_defs:
        cvc = cfg.get(cfg_key, {}) or {}
        s = _get_or_add(styles, style_name)
        s.font.name      = _cfg_str(cvc,   "font_name",       "Calibri")
        s.font.size      = Pt(_cfg_float(cvc, "font_size_pt", def_size))
        s.font.bold      = _cfg_bool(cvc,  "bold",            def_bold)
        s.font.color.rgb = _cfg_color(cvc, "color",           def_color)
        s.paragraph_format.space_before = Pt(_cfg_float(cvc, "space_before_pt", def_before))
        s.paragraph_format.space_after  = Pt(_cfg_float(cvc, "space_after_pt",  def_after))
        s.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.CENTER
        # outline_level intentionally NOT set → never appears in TOC

    # ── Caption ──────────────────────────────────────────────────────────────
    capc = cfg.get("caption", {}) or {}
    cap  = _get_or_add(styles, "Caption")
    cap.font.italic     = _cfg_bool(capc,  "italic",         True)
    cap.font.bold       = _cfg_bool(capc,  "bold",           True)
    cap.font.size       = Pt(_cfg_float(capc, "font_size_pt", 9))
    cap.font.color.rgb  = _cfg_color(capc, "color",          RGBColor(0x55, 0x55, 0x55))
    # Alignment is set per-paragraph in _emit_fig_caption/_emit_tbl_caption
    # so we do not set it here — that way captions follow their figure's alignment.
    cap.paragraph_format.space_before = Pt(_cfg_float(capc, "space_before_pt", 2))
    cap.paragraph_format.space_after  = Pt(_cfg_float(capc, "space_after_pt",  8))
