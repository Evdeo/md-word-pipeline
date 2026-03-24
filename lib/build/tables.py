"""Table formatting — data tables and image tables, with merged-cell support."""
import re
from typing import Dict, List, Optional, Tuple

from docx.table import Table
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# ── low-level cell helpers ─────────────────────────────────────────────────────

def _border(style="single", size=4, color="AAAAAA"):
    return {"val": style, "sz": str(size), "color": color, "space": "0"}


def _apply_tc_borders(tc, all_sides=True):
    """Apply borders directly to a raw w:tc XML element.

    Using raw tc elements is essential for merged tables — python-docx's
    row.cells de-duplicates merged cells and never visits the consumed
    tc elements, which are left without borders and render as gaps.
    """
    tcPr = tc.get_or_add_tcPr()
    # Remove any existing tcBdr to avoid duplicates
    for old in tcPr.findall(qn("w:tcBdr")):
        tcPr.remove(old)
    tcBdr = OxmlElement("w:tcBdr")
    sides = ["top", "left", "bottom", "right"]
    if all_sides:
        for s in sides:
            el = OxmlElement(f"w:{s}")
            el.set(qn("w:val"),   "single")
            el.set(qn("w:sz"),    "4")
            el.set(qn("w:color"), "AAAAAA")
            el.set(qn("w:space"), "0")
            tcBdr.append(el)
    else:
        for s in sides:
            el = OxmlElement(f"w:{s}")
            el.set(qn("w:val"), "none")
            tcBdr.append(el)
    tcPr.append(tcBdr)


def _apply_cell_borders(cell, all_sides=True):
    """Convenience wrapper for python-docx Cell objects."""
    _apply_tc_borders(cell._tc, all_sides)


def _shade_cell(cell, fill: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  fill)
    tcPr.append(shd)


def _set_cell_margins_tc(tc, top=60, bottom=60, left=100, right=100):
    """Set cell margins directly on a raw w:tc element."""
    tcPr  = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcMar")):
        tcPr.remove(old)
    tcMar = OxmlElement("w:tcMar")
    for name, val in [("top", top), ("bottom", bottom),
                      ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{name}")
        el.set(qn("w:w"),    str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)


def _set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for name, val in [("top", top), ("bottom", bottom),
                      ("left", left), ("right", right)]:
        el = OxmlElement(f"w:{name}")
        el.set(qn("w:w"),    str(val))
        el.set(qn("w:type"), "dxa")
        tcMar.append(el)
    tcPr.append(tcMar)


def apply_cell_alignment(cell, ha: Optional[str] = None, va: Optional[str] = None):
    """Set horizontal and vertical alignment on a cell.

    ha: 'l' | 'c' | 'r'   (left / centre / right)
    va: 't' | 'm' | 'b'   (top / middle / bottom)

    Horizontal alignment is applied to both the cell container and every
    paragraph inside it so text position is consistent.
    """
    h_map = {
        "l": WD_ALIGN_PARAGRAPH.LEFT,
        "c": WD_ALIGN_PARAGRAPH.CENTER,
        "r": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    v_map = {
        "t": WD_ALIGN_VERTICAL.TOP,
        "m": WD_ALIGN_VERTICAL.CENTER,
        "b": WD_ALIGN_VERTICAL.BOTTOM,
    }
    if va and va in v_map:
        cell.vertical_alignment = v_map[va]

    if ha and ha in h_map:
        para_align = h_map[ha]
        for para in cell.paragraphs:
            para.alignment = para_align


# ── merged-cell helpers ────────────────────────────────────────────────────────

# Attribute pattern: {cs=2 rs=3 ha=c va=t}  (any order, all optional)
# Must be at the END of the cell text (trailing whitespace allowed) so that
# descriptive text like "Put {rs=2} at end…" is not mistaken for a real attribute.
_ATTR_RE = re.compile(
    r'\{([^}]*)\}\s*$'
)
_CS_RE = re.compile(r'\bcs\s*=\s*(\d+)')
_RS_RE = re.compile(r'\brs\s*=\s*(\d+)')
_HA_RE = re.compile(r'\bha\s*=\s*([lcrLCR])')
_RS_MARKER = re.compile(r'^\^\^$')   # merge-up marker
_CS_MARKER = re.compile(r'^<<$')     # merge-left marker


def parse_cell_attrs(raw: str) -> Tuple[str, int, int, Optional[str], Optional[str]]:
    """Parse a raw cell string.

    Returns (clean_text, colspan, rowspan, ha, va).
    Attributes are stripped from the returned text.
    """
    text = raw.strip()
    cs, rs, ha, va = 1, 1, None, None

    m = _ATTR_RE.search(text)
    if m:
        attr_str = m.group(1)
        cs_m = _CS_RE.search(attr_str)
        rs_m = _RS_RE.search(attr_str)
        ha_m = _HA_RE.search(attr_str)
        va_m = re.search(r'\bva\s*=\s*([tmb])', attr_str, re.I)
        if cs_m: cs = max(1, int(cs_m.group(1)))
        if rs_m: rs = max(1, int(rs_m.group(1)))
        if ha_m: ha = ha_m.group(1).lower()
        if va_m: va = va_m.group(1).lower()
        # Strip the attribute block from the display text
        text = _ATTR_RE.sub("", text).strip()

    return text, cs, rs, ha, va


def build_merge_plan(
    header_cells: List[str],
    body_rows: List[List[str]],
    has_header: bool,
) -> Tuple[List[List[str]], List[List[Tuple[int,int,int,int,Optional[str],Optional[str]]]]]:
    """Resolve ^^ / << markers and {cs= rs=} attributes into a merge plan.

    Returns:
        clean_grid  — 2-D list of plain text strings (attrs stripped, markers replaced with "")
        merge_grid  — 2-D list of (anchor_row, anchor_col, rowspan, colspan, ha, va)
                      where anchor_row/col point to the top-left cell of each merge group.
                      Non-anchor cells carry the anchor coords so we know they are consumed.
    """
    # Build raw grid
    all_rows: List[List[str]] = []
    if has_header:
        all_rows.append(header_cells[:])
    for row in body_rows:
        all_rows.append(row[:])

    nrows = len(all_rows)
    ncols = max((len(r) for r in all_rows), default=0)

    # Pad short rows
    for row in all_rows:
        while len(row) < ncols:
            row.append("")

    # Parse each cell to extract clean text + span attributes
    # grid_info[r][c] = (clean_text, cs, rs, ha, va)
    grid_info = []
    for row in all_rows:
        row_info = []
        for cell in row:
            row_info.append(parse_cell_attrs(cell))
        grid_info.append(row_info)

    # Build merge_grid: each entry is (anchor_r, anchor_c, rs, cs, ha, va)
    # Initially every cell is its own anchor
    merge_grid = [
        [(r, c, grid_info[r][c][2], grid_info[r][c][1],
          grid_info[r][c][3], grid_info[r][c][4])
         for c in range(ncols)]
        for r in range(nrows)
    ]
    clean_grid = [
        [grid_info[r][c][0] for c in range(ncols)]
        for r in range(nrows)
    ]

    # First pass: resolve explicit {cs= rs=} spans from anchor cells.
    # Mark consumed cells as (-1, -1, ...) to skip them.
    for r in range(nrows):
        for c in range(ncols):
            _, cs, rs, ha, va = grid_info[r][c]
            if cs == 1 and rs == 1:
                continue  # no span
            # Validate span doesn't exceed grid
            rs = min(rs, nrows - r)
            cs = min(cs, ncols - c)
            for dr in range(rs):
                for dc in range(cs):
                    if dr == 0 and dc == 0:
                        merge_grid[r][c] = (r, c, rs, cs, ha, va)
                    else:
                        merge_grid[r + dr][c + dc] = (r, c, rs, cs, ha, va)
                        clean_grid[r + dr][c + dc] = ""

    # Second pass: resolve ^^ (merge-up) and << (merge-left) markers.
    for r in range(nrows):
        for c in range(ncols):
            raw = all_rows[r][c].strip()
            if _RS_MARKER.match(raw):
                # Find anchor above
                for ar in range(r - 1, -1, -1):
                    if merge_grid[ar][c][0] == ar:  # is itself an anchor
                        anc_r, anc_c, ars, acs, aha, ava = merge_grid[ar][c]
                        new_rs = r - anc_r + 1
                        for dr in range(new_rs):
                            merge_grid[anc_r + dr][anc_c] = (anc_r, anc_c, new_rs, acs, aha, ava)
                        merge_grid[r][c] = (anc_r, anc_c, new_rs, acs, aha, ava)
                        clean_grid[r][c] = ""
                        break

            elif _CS_MARKER.match(raw):
                # Find anchor to the left
                for ac in range(c - 1, -1, -1):
                    if merge_grid[r][ac][1] == ac:  # is itself an anchor
                        anc_r, anc_c, ars, acs, aha, ava = merge_grid[r][ac]
                        new_cs = c - anc_c + 1
                        for dc in range(new_cs):
                            merge_grid[anc_r][anc_c + dc] = (anc_r, anc_c, ars, new_cs, aha, ava)
                        merge_grid[r][c] = (anc_r, anc_c, ars, new_cs, aha, ava)
                        clean_grid[r][c] = ""
                        break

    return clean_grid, merge_grid


def apply_merges(table: Table, merge_grid, nrows: int, ncols: int):
    """Apply merges and alignments to an already-created docx table.

    merge_grid[r][c] = (anchor_r, anchor_c, rowspan, colspan, ha, va)
    Cells where (anchor_r, anchor_c) != (r, c) are consumed cells — they are
    merged into the anchor rectangle.
    """
    processed = set()

    for r in range(nrows):
        for c in range(ncols):
            anc_r, anc_c, rs, cs, ha, va = merge_grid[r][c]
            if (anc_r, anc_c) in processed:
                continue
            processed.add((anc_r, anc_c))

            if rs == 1 and cs == 1:
                # No merge needed — just alignment
                if ha or va:
                    apply_cell_alignment(table.cell(anc_r, anc_c), ha, va)
                continue

            # Merge the rectangle
            top_left     = table.cell(anc_r, anc_c)
            bottom_right = table.cell(
                min(anc_r + rs - 1, nrows - 1),
                min(anc_c + cs - 1, ncols - 1),
            )
            top_left.merge(bottom_right)

            if ha or va:
                apply_cell_alignment(top_left, ha, va)


# ── table classification ───────────────────────────────────────────────────────

def is_image_table(rows: List[List[str]]) -> bool:
    """True if every data cell contains only an image path."""
    if len(rows) < 2:
        return False
    for row in rows[1:]:
        for cell in row:
            txt = cell.strip()
            if txt and not re.match(r"!\[.*?\]\(.*?\)", txt):
                return False
    return True


# ── column width helpers ───────────────────────────────────────────────────────

def apply_col_widths(table: Table, widths_str: str, total_dxa: int = 0):
    """Apply custom column widths from '20%,50%,30%' string.

    Uses pct-based widths so Word calculates the actual pixel size at render
    time from its own page/margin settings — same as regular tables. This
    guarantees custom-width tables align precisely with all other tables.

    Each % is of the full text area. Columns summing to less than 100%
    produce a proportionally narrower table.
    """
    parts = [p.strip().rstrip("%") for p in widths_str.split(",")]
    try:
        pcts = [float(p) for p in parts]
    except ValueError:
        return

    total_pct   = sum(pcts)          # e.g. 100.0 for full-width, 70.0 for narrower
    # OOXML pct unit: 1/50th of a percent, so 100% = 5000
    col_pct     = [int(p * 50) for p in pcts]   # each column in pct units
    table_pct   = int(total_pct * 50)            # table total in pct units

    tbl = table._tbl

    # 1. Set tblW to pct — Word resolves this against the real text area
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"),    str(table_pct))
    tblW.set(qn("w:type"), "pct")
    existing = tblPr.find(qn("w:tblW"))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(tblW)

    # 2. Update tblGrid with dxa for column proportions (use total_dxa for grid only)
    if not total_dxa:
        total_dxa = 9360
    dxa_widths = [int(total_dxa * p / 100) for p in pcts]
    tblGrid = tbl.find(qn("w:tblGrid"))
    if tblGrid is not None:
        tbl.remove(tblGrid)
    new_grid = OxmlElement("w:tblGrid")
    for w in dxa_widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        new_grid.append(gc)
    tblPr_idx = list(tbl).index(tblPr) if tblPr in tbl else 0
    tbl.insert(tblPr_idx + 1, new_grid)

    # 3. Set each cell's tcW as pct so it scales with the table
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(col_pct):
                tc   = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcW  = OxmlElement("w:tcW")
                tcW.set(qn("w:w"),    str(col_pct[i]))
                tcW.set(qn("w:type"), "pct")
                existing_w = tcPr.find(qn("w:tcW"))
                if existing_w is not None:
                    tcPr.remove(existing_w)
                tcPr.append(tcW)


# ── table formatters ───────────────────────────────────────────────────────────

def _hex(c: str) -> str:
    """Normalise colour string to 6-char uppercase hex, no #."""
    return str(c).lstrip("#").upper()


def inject_table_style(doc, style_id: str,
                       hdr_bg: str, hdr_fg: str,
                       odd_bg: str, even_bg: str,
                       border_color: str = "D0D0D0") -> None:
    """Inject a named banded-row table style using lxml element construction.

    Builds the element tree programmatically to avoid XML namespace/comment
    issues that cause Word to silently reject the style.
    """
    from lxml import etree as _et

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    def w(tag):
        return f"{{{W}}}{tag}"

    def el(tag, **attrs):
        e = _et.Element(w(tag))
        for k, v in attrs.items():
            e.set(w(k), str(v))
        return e

    def sub(parent, tag, **attrs):
        e = _et.SubElement(parent, w(tag))
        for k, v in attrs.items():
            e.set(w(k), str(v))
        return e

    styles_root = doc.part.styles._element

    # Remove any existing copy
    for old in styles_root.findall(f".//{w('style')}[@{w('styleId')}='{style_id}']"):
        old.getparent().remove(old)

    h_bg  = _hex(hdr_bg)
    h_fg  = _hex(hdr_fg)
    o_bg  = _hex(odd_bg)
    e_bg  = _hex(even_bg)
    b_col = _hex(border_color)

    # ── Root style element ────────────────────────────────────────────────────
    style = el("style", type="table", styleId=style_id)
    sub(style, "name",       val=style_id)
    sub(style, "basedOn",    val="TableNormal")
    sub(style, "uiPriority", val="40")

    pPr = sub(style, "pPr")
    sub(pPr, "spacing", after="0", line="240", lineRule="auto")

    # ── Default table properties (borders + cell margins) ─────────────────────
    tblPr_s = sub(style, "tblPr")
    sub(tblPr_s, "tblStyleRowBandSize", val="1")
    sub(tblPr_s, "tblStyleColBandSize", val="1")

    tblBdr = sub(tblPr_s, "tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        sub(tblBdr, side, val="single", sz="4", space="0", color=b_col)

    tblMar = sub(tblPr_s, "tblCellMar")
    for _tag, _w in (("top","60"),("left","108"),("bottom","60"),("right","108")):
        _e = _et.SubElement(tblMar, w(_tag))
        _e.set(w("w"), _w)
        _e.set(w("type"), "dxa")

    # ── Conditional format sections ───────────────────────────────────────────
    # Order inside tblStylePr must be: pPr rPr tblPr trPr tcPr
    def band(band_type, fill, bold=False, fg_col=None):
        sp  = sub(style, "tblStylePr", type=band_type)
        pp  = sub(sp, "pPr")
        if band_type == "firstRow":
            sub(pp, "spacing", before="0", after="0")
        rp  = sub(sp, "rPr")
        if bold:
            sub(rp, "b")
        if fg_col:
            sub(rp, "color", val=fg_col)
        sub(sp, "tblPr")
        tcp = sub(sp, "tcPr")
        sub(tcp, "shd", val="clear", color="auto", fill=fill)

    band("firstRow",  h_bg, bold=True, fg_col=h_fg)
    band("band1Horz", o_bg)   # odd  body rows — correct OOXML type name
    band("band2Horz", e_bg)   # even body rows

    styles_root.append(style)



def format_data_table(table: Table, col_widths: Optional[str] = None,
                      total_dxa: int = 8748, has_header: bool = True,
                      table_cfg: Optional[dict] = None):
    """Style a data table using a named Word table style for banded rows.

    The style "MdToDocxDataTable" is injected into the document with colours
    from table_cfg (from config.yaml). Word then handles alternating row
    shading natively — no manual w:shd on every cell.

    table_cfg may contain:
      table_header.background    — header row fill hex  (default: "1F3864")
      table_header.font_color    — header text hex      (default: "FFFFFF")
      table_rows.odd_background  — odd  row fill hex    (default: "F7F7F7")
      table_rows.even_background — even row fill hex    (default: "FFFFFF")
    """
    tc_cfg   = table_cfg or {}
    hdr_cfg  = tc_cfg.get("table_header", {}) or {}
    rows_cfg = tc_cfg.get("table_rows",   {}) or {}

    hdr_bg  = str(hdr_cfg.get("background",    "1F3864")).lstrip("#")
    hdr_fg  = str(hdr_cfg.get("font_color",    "FFFFFF")).lstrip("#")
    odd_bg  = str(rows_cfg.get("odd_background",  "F7F7F7")).lstrip("#")
    even_bg = str(rows_cfg.get("even_background", "FFFFFF")).lstrip("#")

    # Inject / refresh the custom style in the document with current colours
    inject_table_style(
        table.part.document, "MdToDocxDataTable",
        hdr_bg=hdr_bg, hdr_fg=hdr_fg,
        odd_bg=odd_bg, even_bg=even_bg,
    )

    # ── Apply style reference and tblLook to this table ──────────────────────
    tbl   = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)

    # tblStyle
    tblStyle = tblPr.find(qn("w:tblStyle"))
    if tblStyle is None:
        tblStyle = OxmlElement("w:tblStyle")
        tblPr.insert(0, tblStyle)
    tblStyle.set(qn("w:val"), "MdToDocxDataTable")

    # tblW — full-width
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"),    "5000")
    tblW.set(qn("w:type"), "pct")
    existing_w = tblPr.find(qn("w:tblW"))
    if existing_w is not None:
        tblPr.remove(existing_w)
    tblPr.append(tblW)

    # tblLook — enable firstRow banding, disable column banding
    tblLook = tblPr.find(qn("w:tblLook"))
    if tblLook is None:
        tblLook = OxmlElement("w:tblLook")
        tblPr.append(tblLook)
    # tblLook val bitmask:
    #   0x0020 = firstRow highlight on
    #   0x0200 = noHBand (set = NO row banding — we want 0 = banding ON)
    #   0x0400 = noVBand (set = no column banding — we want this ON)
    # firstRow on + noVBand on = 0x0420
    look_val = "0420" if has_header else "0400"
    tblLook.set(qn("w:val"),         look_val)
    tblLook.set(qn("w:firstRow"),    "1" if has_header else "0")
    tblLook.set(qn("w:lastRow"),     "0")
    tblLook.set(qn("w:firstColumn"), "0")
    tblLook.set(qn("w:lastColumn"),  "0")
    tblLook.set(qn("w:noHBand"),     "0")   # 0 = row banding IS applied
    tblLook.set(qn("w:noVBand"),     "1")   # 1 = column banding off

    # ── Strip per-cell overrides that would suppress the style ──────────────
    # Remove tcMar (margins come from the style's tblCellMar) and any
    # leftover w:shd (style provides shading). Per-cell tcPr entries
    # override the style, so we keep each tcPr minimal: only tcW and
    # merge-related elements (vMerge, gridSpan).
    _KEEP_TCPR = {qn("w:tcW"), qn("w:vMerge"), qn("w:gridSpan"),
                  qn("w:vAlign"), qn("w:textDirection")}
    for tr in tbl.findall(qn("w:tr")):
        for tc in tr.findall(qn("w:tc")):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is not None:
                for child in list(tcPr):
                    if child.tag not in _KEEP_TCPR:
                        tcPr.remove(child)

    # ── Prevent rows from splitting across pages + repeat header ─────────────
    rows = tbl.findall(qn("w:tr"))
    for row_idx, tr in enumerate(rows):
        trPr = tr.find(qn("w:trPr"))
        if trPr is None:
            trPr = OxmlElement("w:trPr")
            tr.insert(0, trPr)
        if trPr.find(qn("w:cantSplit")) is None:
            cantSplit = OxmlElement("w:cantSplit")
            cantSplit.set(qn("w:val"), "1")
            trPr.append(cantSplit)
        # Repeat header row on each new page
        if row_idx == 0 and has_header:
            if trPr.find(qn("w:tblHeader")) is None:
                tblHeader = OxmlElement("w:tblHeader")
                trPr.append(tblHeader)

    if col_widths:
        apply_col_widths(table, col_widths, total_dxa)


def format_image_table(table: Table):
    """Remove borders from image layout tables."""
    for row in table.rows:
        for cell in row.cells:
            _apply_cell_borders(cell, all_sides=False)
            _set_cell_margins(cell, top=40, bottom=40, left=60, right=60)
            for para in cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
