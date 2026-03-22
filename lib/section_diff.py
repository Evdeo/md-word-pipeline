"""
lib/section_diff.py — Section-tree diff between two .docx files.

Philosophy
----------
Don't try to diff at character or run level. Instead:
1. Parse both docx files into a tree of Section objects (one per heading).
2. Match sections by heading text. 
3. Hash the content of each matched section.
4. If hashes match → identical. If not → changed.
5. Unmatched in baseline → removed. Unmatched in received → added.
6. Same heading but different position → moved.

Render a single self-contained HTML file with:
- Every section as a <details> card (collapsible).
- Identical sections collapsed (green).
- Changed sections expanded showing two Word-like columns.
- Removed/added/moved sections expanded with appropriate styling.
"""

from __future__ import annotations
import hashlib
import re
import html as _html
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Dict, Tuple

from docx import Document as _DocX
from docx.oxml.ns import qn

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A  = "http://schemas.openxmlformats.org/drawingml/2006/main"
R  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def _w(tag): return f"{{{W}}}{tag}"
def _e(s):   return _html.escape(str(s))


# ── Section dataclass ─────────────────────────────────────────────────────────

@dataclass
class Section:
    heading:    str               # plain text heading
    level:      int               # 1-6
    elements:   list              # raw lxml elements (w:p, w:tbl, etc.)
    children:   List[Section]     = field(default_factory=list)
    position:   int               = 0    # ordinal in original document

    @property
    def content_hash(self) -> str:
        """MD5 of normalised text + table structure (col widths, row count)."""
        parts = []
        for el in self.elements:
            tag = el.tag.split('}')[-1]
            if tag == 'p':
                text = ''.join(t.text or '' for t in el.findall(f'.//{_w("t")}'))
                # For Caption paragraphs include only text — Word strips redundant
                # explicit bold/italic on save (style already provides them), which
                # would otherwise cause false positives on every open+save cycle.
                pPr   = el.find(_w('pPr'))
                ps    = pPr.find(_w('pStyle')) if pPr is not None else None
                style = ps.get(_w('val'), '') if ps is not None else ''
                parts.append(f'{style}:{text.strip()}' if style else text.strip())
            elif tag == 'tbl':
                # Row count — catches added/removed rows
                # Column widths deliberately excluded: Word recalculates them
                # on open/save causing spurious hash mismatches.
                trs = el.findall(_w('tr'))
                parts.append(f'rows:{len(trs)}')
                for tc in el.findall(f'.//{_w("tc")}'):
                    cell_text = ''.join(t.text or '' for t in tc.findall(f'.//{_w("t")}'))
                    parts.append(cell_text.strip())
            elif tag in ('sdt', 'drawing'):
                parts.append(tag)
        # Strip trailing empty entries — Word appends empty paragraphs
        # after tables on save; these are not meaningful content changes.
        while parts and not parts[-1].strip():
            parts.pop()
        canon = '\n'.join(p for p in parts if p)
        return hashlib.md5(canon.encode('utf-8', errors='replace')).hexdigest()

    @property
    def key(self) -> str:
        """Normalised heading for matching (strip numbers, lower-case)."""
        t = self.heading.lower().strip()
        t = re.sub(r'^[a-z]?\d+(\.\d+)*\.?\s*', '', t)   # strip A.2. / 1.2.3
        t = re.sub(r'\s+', ' ', t)
        return t


# ── Parse docx into section tree ──────────────────────────────────────────────

_SKIP_STYLES = {
    'CoverTitle', 'CoverSubtitle', 'CoverBody',
    'RevisionHistory', 'toc1', 'toc2', 'toc3',
    'TOCHeading', 'TableofContents',
}

def _heading_level(el) -> Optional[int]:
    """Return heading level 1-6 if element is a heading paragraph, else None."""
    pPr = el.find(_w('pPr'))
    if pPr is None:
        return None
    ps = pPr.find(_w('pStyle'))
    if ps is None:
        return None
    style = ps.get(_w('val'), '')
    if style in _SKIP_STYLES:
        return None
    m = re.match(r'[Hh]eading\s*(\d)', style)
    if m:
        return int(m.group(1))
    # Also check outlineLvl
    ol = pPr.find(_w('outlineLvl'))
    if ol is not None:
        lvl = int(ol.get(_w('val'), 9) or 9)
        if lvl < 6:
            return lvl + 1
    return None


def _is_toc(el) -> bool:
    """True if this is inside a TOC SDT or has TOC style."""
    pPr = el.find(_w('pPr'))
    if pPr is None:
        return False
    ps = pPr.find(_w('pStyle'))
    style = ps.get(_w('val'), '') if ps is not None else ''
    return style.lower().startswith('toc') or style == 'TableofContents'


def extract_sections(docx_path: Path) -> List[Section]:
    """Parse a docx into a flat list of top-level sections, each with children.

    The cover page is captured as a special Section(heading="Cover Page", level=0)
    at position 0 so users can expand it in the review report.
    """
    doc  = _DocX(str(docx_path))
    body = doc.element.body

    # Collect all body children, separating cover from main content
    raw_elements  = []
    cover_elements = []
    in_cover = True

    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag == 'sectPr':
            continue

        if tag == 'p':
            pPr   = child.find(_w('pPr'))
            ps    = pPr.find(_w('pStyle')) if pPr is not None else None
            style = ps.get(_w('val'), '') if ps is not None else ''

            # End of cover when we hit a real Heading 1
            if style in ('Heading1', 'Heading 1') and in_cover:
                text = ''.join(t.text or '' for t in child.findall(f'.//{_w("t")}'))
                if 'Table of Contents' not in text:
                    in_cover = False

            if in_cover:
                cover_elements.append(child)
                continue
            if style in _SKIP_STYLES or _is_toc(child):
                continue

        elif tag == 'sdt':
            # Skip TOC SDT
            sdtPr = child.find(_w('sdtPr'))
            alias = sdtPr.find(_w('alias')) if sdtPr is not None else None
            if alias is not None and 'toc' in alias.get(_w('val'), '').lower():
                continue
            if in_cover:
                cover_elements.append(child)
                continue

        raw_elements.append(child)

    # Build flat section list by scanning for headings
    sections: List[Section] = []
    current_stack: List[Section] = []   # stack of open sections by level
    body_buffer: list = []              # elements before any heading
    pos = 0

    def flush_buffer(buf):
        # Attach loose elements to the innermost open section
        if current_stack and buf:
            current_stack[-1].elements.extend(buf)

    for el in raw_elements:
        tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
        if tag != 'p':
            body_buffer.append(el)
            continue

        lvl = _heading_level(el)
        if lvl is None:
            body_buffer.append(el)
            continue

        # Flush buffered content
        flush_buffer(body_buffer)
        body_buffer = []

        # Build heading text
        heading = ''.join(t.text or '' for t in el.findall(f'.//{_w("t")}'))

        sec = Section(heading=heading, level=lvl, elements=[], position=pos)
        pos += 1

        # Pop stack to find parent
        while current_stack and current_stack[-1].level >= lvl:
            current_stack.pop()

        if current_stack:
            current_stack[-1].children.append(sec)
        else:
            sections.append(sec)

        current_stack.append(sec)

    flush_buffer(body_buffer)

    # Prepend cover section if there are cover elements
    cover_elements_filtered = [
        el for el in cover_elements
        if el.tag.split('}')[-1] != 'sectPr'
    ]
    if cover_elements_filtered:
        cover_sec = Section(
            heading="Cover Page",
            level=0,
            elements=cover_elements_filtered,
            position=-1,
        )
        sections = [cover_sec] + sections

    return sections


# ── Diff section trees ────────────────────────────────────────────────────────

@dataclass
class SectionResult:
    status:   str          # "identical" | "changed" | "removed" | "added" | "moved"
    baseline: Optional[Section]
    received: Optional[Section]
    children: List[SectionResult] = field(default_factory=list)
    moved_from: int = -1   # original position if moved
    moved_to:   int = -1


def _match_sections(
    base_list: List[Section],
    recv_list: List[Section],
) -> List[SectionResult]:
    """Match sections by normalised heading key and diff recursively."""

    base_by_key: Dict[str, Section] = {}
    recv_by_key: Dict[str, Section] = {}

    for s in base_list:
        base_by_key.setdefault(s.key, s)
    for s in recv_list:
        recv_by_key.setdefault(s.key, s)

    results: List[SectionResult] = []
    used_recv = set()

    # First pass: match in received order so the report reads top-to-bottom
    # as the reviewer sees it
    for rs in recv_list:
        k = rs.key
        bs = base_by_key.get(k)
        used_recv.add(k)

        if bs is None:
            # Added in received
            results.append(SectionResult(
                status='added', baseline=None, received=rs))
        elif rs.level == 0:
            # Cover page — always identical (dates/TOC page numbers differ on rebuild)
            children = _match_sections(bs.children, rs.children)
            results.append(SectionResult(
                status='identical', baseline=bs, received=rs,
                children=children))
        else:
            # Matched by heading key — check content, ignore position.
            # Position shifts naturally when sections are added/removed before
            # this one; using position for move detection causes every section
            # after a deletion to be falsely flagged as moved.
            children = _match_sections(bs.children, rs.children)

            # Detect a genuine move: same heading appears at a different
            # ordinal within its sibling list (not just shifted by add/remove).
            base_idx = next((i for i, s in enumerate(base_list) if s.key == k), -1)
            recv_idx = next((i for i, s in enumerate(recv_list) if s.key == k), -1)
            # Only flag as moved if the relative order among COMMON sections changed.
            # Build the ordered list of keys that appear in both base and received.
            common_base = [s.key for s in base_list if s.key in recv_by_key]
            common_recv = [s.key for s in recv_list if s.key in base_by_key]
            base_common_idx = common_base.index(k) if k in common_base else -1
            recv_common_idx = common_recv.index(k) if k in common_recv else -1
            is_moved = (base_common_idx != recv_common_idx and
                        base_common_idx >= 0 and recv_common_idx >= 0)

            if is_moved:
                status = 'moved_changed' if bs.content_hash != rs.content_hash else 'moved'
                results.append(SectionResult(
                    status=status, baseline=bs, received=rs,
                    children=children,
                    moved_from=base_common_idx, moved_to=recv_common_idx))
            elif bs.content_hash == rs.content_hash and not any(
                    c.status not in ('identical',) for c in children):
                results.append(SectionResult(
                    status='identical', baseline=bs, received=rs,
                    children=children))
            elif bs.content_hash == rs.content_hash:
                # Own content unchanged — only children changed
                results.append(SectionResult(
                    status='contains_changes', baseline=bs, received=rs,
                    children=children))
            else:
                results.append(SectionResult(
                    status='changed', baseline=bs, received=rs,
                    children=children))

    # Second pass: sections in baseline not in received → removed
    for bs in base_list:
        if bs.key not in used_recv:
            results.append(SectionResult(
                status='removed', baseline=bs, received=None))

    return results


def diff_documents(baseline_path: Path, received_path: Path) -> List[SectionResult]:
    """High-level diff: extract section trees and match them."""
    base_sections = extract_sections(baseline_path)
    recv_sections = extract_sections(received_path)

    # Assign flat positions
    def assign_positions(sections, counter=[0]):
        for s in sections:
            s.position = counter[0]
            counter[0] += 1
            assign_positions(s.children, counter)

    assign_positions(base_sections)
    counter = [0]
    assign_positions(recv_sections, counter)

    return _match_sections(base_sections, recv_sections)


# ── Word-like HTML rendering ──────────────────────────────────────────────────

def _render_element(el, doc) -> str:
    """Render a single body element (p or tbl) as Word-like HTML."""
    tag = el.tag.split('}')[-1] if '}' in el.tag else el.tag
    if tag == 'p':
        return _render_para(el)
    elif tag == 'tbl':
        return _render_table(el, doc)
    elif tag == 'drawing':
        return '<div class="wd-img">[Image]</div>'
    return ''


def _render_para(el) -> str:
    """Render a w:p as Word-like HTML paragraph."""
    pPr   = el.find(_w('pPr'))
    style = ''
    jc    = ''
    indent_left = 0

    if pPr is not None:
        ps = pPr.find(_w('pStyle'))
        style = ps.get(_w('val'), '') if ps is not None else ''
        jc_el = pPr.find(_w('jc'))
        jc    = jc_el.get(_w('val'), '') if jc_el is not None else ''
        ind   = pPr.find(_w('ind'))
        if ind is not None:
            indent_left = int(ind.get(_w('left'), 0) or 0) // 120  # twips→approx px

    # Determine style-level bold/italic (e.g. Caption style has both)
    _STYLE_BOLD   = {'Caption', 'Heading1', 'Heading2', 'Heading3',
                     'Heading4', 'Heading5', 'Heading6'}
    _STYLE_ITALIC = {'Caption', 'BlockText', 'Quote', 'IntenseQuote'}
    sty_bold   = style in _STYLE_BOLD
    sty_italic = style in _STYLE_ITALIC

    # Render runs
    inner = ''
    for child in el:
        ctag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if ctag == 'r':
            inner += _render_run(child, style_bold=sty_bold, style_italic=sty_italic)
        elif ctag == 'hyperlink':
            for r in child.findall(_w('r')):
                inner += _render_run(r, style_bold=sty_bold, style_italic=sty_italic)

    if not inner.strip():
        return '<p class="wd-p wd-empty">&nbsp;</p>'

    css_class = _style_to_class(style)
    align = f' style="text-align:{jc}"' if jc in ('center','right','both') else ''
    indent = f' style="margin-left:{indent_left}px"' if indent_left > 0 else ''
    combined = ''
    if jc in ('center','right','both') and indent_left > 0:
        combined = f' style="text-align:{jc};margin-left:{indent_left}px"'
        align = indent = ''

    return f'<p class="wd-p {css_class}"{align}{indent}{combined}>{inner}</p>'


def _render_run(r, style_bold=False, style_italic=False) -> str:
    """Render a w:r as inline HTML with formatting.

    style_bold / style_italic: formatting inherited from paragraph style.
    Run-level explicit formatting is merged with style-level so that
    paragraphs whose style already defines bold/italic (e.g. Caption)
    render consistently regardless of whether Word stored explicit run props.
    """
    rPr   = r.find(_w('rPr'))
    bold  = style_bold
    italic = style_italic
    strike = False
    code  = False
    color = ''

    if rPr is not None:
        # Explicit w:b / w:i override style — but only if not explicitly turned OFF
        b_el = rPr.find(_w('b'))
        i_el = rPr.find(_w('i'))
        if b_el is not None:
            bold   = b_el.get(_w('val'), 'true').lower() not in ('0', 'false', 'off')
        if i_el is not None:
            italic = i_el.get(_w('val'), 'true').lower() not in ('0', 'false', 'off')
        strike = rPr.find(_w('strike')) is not None
        mono   = rPr.find(_w('rFonts'))
        if mono is not None:
            fn = mono.get(_w('ascii'), '') or mono.get(_w('hAnsi'), '')
            code = 'courier' in fn.lower() or 'mono' in fn.lower() or 'code' in fn.lower()
        col_el = rPr.find(_w('color'))
        if col_el is not None:
            c = col_el.get(_w('val'), '')
            if c not in ('auto', ''):
                color = f'#{c}'

    # Check for line/page break
    br = r.find(_w('br'))
    if br is not None:
        btype = br.get(_w('type'), 'line')
        if btype == 'page':
            return '<hr class="wd-pagebreak">'
        return '<br>'

    text = ''.join(t.text or '' for t in r.findall(_w('t')))
    if not text:
        return ''

    out = _e(text)
    if code:   out = f'<code>{out}</code>'
    if bold:   out = f'<strong>{out}</strong>'
    if italic: out = f'<em>{out}</em>'
    if strike: out = f'<s>{out}</s>'
    if color:  out = f'<span style="color:{color}">{out}</span>'
    return out


def _render_table(tbl, doc) -> str:
    """Render a w:tbl as an HTML table with Word-like styling.

    Handles both colspan (w:gridSpan) and rowspan (w:vMerge).
    First pass computes rowspans, second pass renders with correct attributes.
    """
    rows = tbl.findall(_w('tr'))
    n_rows = len(rows)

    # First pass: compute rowspan for each restart cell
    # rowspan_map[(ri, col_idx)] = rowspan count
    # col_idx tracks the logical column (accounting for colspan)
    rowspan_map: dict = {}
    cont_slots:  set  = set()   # (ri, logical_col) that are continuation cells

    for ri, tr in enumerate(rows):
        logical_col = 0
        for tc in tr.findall(_w('tc')):
            tcPr = tc.find(_w('tcPr'))
            cs = 1
            if tcPr is not None:
                gs = tcPr.find(_w('gridSpan'))
                if gs is not None:
                    cs = int(gs.get(_w('val'), 1) or 1)
                vm = tcPr.find(_w('vMerge'))
                if vm is not None:
                    vm_val = vm.get(_w('val'), '')
                    if vm_val == 'restart':
                        # Count how many rows this spans
                        rs = 1
                        for look in range(ri + 1, n_rows):
                            look_tcs = rows[look].findall(_w('tc'))
                            # Find tc at same logical column
                            lc = 0
                            found_cont = False
                            for ltc in look_tcs:
                                ltcPr = ltc.find(_w('tcPr'))
                                lcs = 1
                                if ltcPr is not None:
                                    lgs = ltcPr.find(_w('gridSpan'))
                                    if lgs is not None:
                                        lcs = int(lgs.get(_w('val'), 1) or 1)
                                if lc == logical_col:
                                    lvm = ltcPr.find(_w('vMerge')) if ltcPr is not None else None
                                    if lvm is not None and lvm.get(_w('val'), '') != 'restart':
                                        rs += 1
                                        found_cont = True
                                    break
                                lc += lcs
                            if not found_cont:
                                break
                        rowspan_map[(ri, logical_col)] = rs
                        # Mark continuation slots
                        for dr in range(1, rs):
                            for dc in range(cs):
                                cont_slots.add((ri + dr, logical_col + dc))
                    else:
                        # Continuation cell — skip
                        for dc in range(cs):
                            cont_slots.add((ri, logical_col + dc))
            logical_col += cs

    # Second pass: render
    rows_html = []
    for ri, tr in enumerate(rows):
        cells_html = []
        logical_col = 0
        for tc in tr.findall(_w('tc')):
            tcPr = tc.find(_w('tcPr'))
            cs = 1
            is_cont = False

            if tcPr is not None:
                gs = tcPr.find(_w('gridSpan'))
                if gs is not None:
                    cs = int(gs.get(_w('val'), 1) or 1)
                vm = tcPr.find(_w('vMerge'))
                if vm is not None and vm.get(_w('val'), '') != 'restart':
                    is_cont = True

            if is_cont or (ri, logical_col) in cont_slots:
                logical_col += cs
                continue

            rs = rowspan_map.get((ri, logical_col), 1)
            cell_inner = ''
            for p in tc.findall(_w('p')):
                cell_inner += _render_para(p)

            attrs = ''
            if cs > 1: attrs += f' colspan="{cs}"'
            if rs > 1: attrs += f' rowspan="{rs}"'
            tag = 'th' if ri == 0 else 'td'
            cells_html.append(f'<{tag}{attrs}>{cell_inner}</{tag}>')
            logical_col += cs

        rows_html.append('<tr>' + ''.join(cells_html) + '</tr>')

    return '<table class="wd-table">' + ''.join(rows_html) + '</table>'


def _style_to_class(style: str) -> str:
    """Map a Word style name to a CSS class."""
    s = style.lower()
    if re.match(r'heading\s*1', s): return 'wd-h1'
    if re.match(r'heading\s*2', s): return 'wd-h2'
    if re.match(r'heading\s*3', s): return 'wd-h3'
    if re.match(r'heading\s*4', s): return 'wd-h4'
    if re.match(r'heading\s*5', s): return 'wd-h5'
    if re.match(r'heading\s*6', s): return 'wd-h6'
    if 'listbullet' in s or 'list bullet' in s: return 'wd-li-bullet'
    if 'listnumber' in s or 'list number' in s: return 'wd-li-number'
    if 'caption' in s: return 'wd-caption'
    if 'quote' in s or 'blockquote' in s: return 'wd-quote'
    if 'code' in s or 'verbatim' in s: return 'wd-code'
    return 'wd-normal'


def _render_section_content(sec: Section, doc) -> str:
    """Render all elements in a section as Word-like HTML.

    If the section has no direct elements but has children (e.g. Images section
    is a pure container with subsections), render all children's content too
    so removed/added container sections show their full content.
    """
    parts = []
    for el in sec.elements:
        h = _render_element(el, doc)
        if h:
            parts.append(h)

    # If no direct content but has children, render children recursively
    if not parts and sec.children:
        for child in sec.children:
            # Add a mini heading for each child section
            parts.append(f'<p class="wd-h{min(child.level+1,6)}">{_e(child.heading)}</p>')
            parts.append(_render_section_content(child, doc))

    return ''.join(parts) if parts else '<p class="wd-empty wd-muted">(empty)</p>'


# ── HTML report ───────────────────────────────────────────────────────────────

_CSS = """
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px;
       background: #f5f5f5; color: #222; }
.report-header { background: #1F3864; color: #fff; padding: 18px 28px;
                 display: flex; align-items: center; gap: 16px; }
.report-header h1 { font-size: 18px; font-weight: 600; }
.report-header .sub { font-size: 12px; opacity: .7; margin-top: 2px; }
.legend { display: flex; gap: 10px; flex-wrap: wrap; padding: 10px 28px;
          background: #fff; border-bottom: 1px solid #ddd; }
.legend span { font-size: 11px; padding: 2px 8px; border-radius: 10px; }
.badge-identical { background:#e8f5e9; color:#2e7d32; }
.badge-changed   { background:#fff8e1; color:#e65100; }
.badge-removed   { background:#ffebee; color:#c62828; }
.badge-added     { background:#e3f2fd; color:#1565c0; }
.badge-moved     { background:#fff3e0; color:#e65100; }
.sections { padding: 16px 28px; max-width: 1400px; margin: 0 auto; }

/* Section cards */
details { margin-bottom: 6px; border-radius: 6px; overflow: hidden;
          border: 1px solid #ddd; }
details[open] { box-shadow: 0 2px 8px rgba(0,0,0,.08); }
summary { list-style: none; cursor: pointer; padding: 10px 16px;
          display: flex; align-items: center; gap: 10px;
          font-weight: 500; user-select: none; }
summary::-webkit-details-marker { display: none; }
summary .arrow { transition: transform .2s; font-size: 11px; color: #888; }
details[open] summary .arrow { transform: rotate(90deg); }
summary .sec-title { flex: 1; }
summary .sec-badge { font-size: 11px; padding: 2px 8px; border-radius: 10px;
                     font-weight: 600; }

.status-identical summary { background: #f1f8f2; }
.status-identical summary .sec-badge { background:#e8f5e9; color:#2e7d32; }
.status-changed   summary { background: #fffde7; }
.status-changed   summary .sec-badge { background:#fff3e0; color:#e65100; }
.status-removed   summary { background: #ffebee; }
.status-removed   summary .sec-badge { background:#ffcdd2; color:#c62828; }
.status-added     summary { background: #e3f2fd; }
.status-added     summary .sec-badge { background:#bbdefb; color:#1565c0; }
.status-moved     summary { background: #fff3e0; }
.status-moved     summary .sec-badge { background:#ffe0b2; color:#bf360c; }
.status-moved-changed summary { background: #fce4ec; }
.status-moved-changed summary .sec-badge { background:#f8bbd0; color:#880e4f; }

/* Content area */
.section-content { padding: 16px; background: #fff; }
.children { padding: 0 0 0 16px; background: #fff; }
/* Children always have neutral card backgrounds regardless of parent status */
.children details { background: #fff; }
.children .status-identical summary { background: #f9fafb; }
.children .status-changed   summary { background: #fffde7; }
.children .status-removed   summary { background: #ffebee; }
.children .status-added     summary { background: #e3f2fd; }
.children .status-moved     summary { background: #fff3e0; }

/* Two-column diff */
.diff-cols { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
.diff-col-label { font-size: 10px; font-weight: 700; text-transform: uppercase;
                  letter-spacing: .05em; color: #888; margin-bottom: 8px;
                  padding-bottom: 4px; border-bottom: 1px solid #eee; }
.diff-absent { display: flex; align-items: center; justify-content: center;
               height: 60px; color: #bbb; font-style: italic; font-size: 12px;
               border: 1px dashed #ddd; border-radius: 4px; }
.diff-single { max-width: 700px; }

.moved-banner { background: #fff3e0; border-left: 4px solid #ff8f00;
                padding: 6px 12px; font-size: 11px; color: #bf360c;
                margin-bottom: 10px; border-radius: 0 4px 4px 0; }

/* Tree overview */
.tree { padding: 12px 28px; background: #fff; border-bottom: 1px solid #eee;
        font-size: 12px; line-height: 1.8; }
.tree-title { font-size: 11px; font-weight: 700; text-transform: uppercase;
              letter-spacing: .06em; color: #888; margin-bottom: 8px; }
.tree-row { display: flex; align-items: center; gap: 6px; cursor: pointer;
            border-radius: 3px; padding: 1px 4px; }
.tree-row:hover { background: #f5f5f5; }
.tree-row .tr-icon { width: 18px; text-align: center; flex-shrink: 0; }
.tree-row .tr-name { flex: 1; color: #333; }
.tree-row .tr-badge { font-size: 10px; padding: 1px 6px; border-radius: 8px;
                      font-weight: 600; }
.tr-identical .tr-name  { color: #888; }
.tr-identical .tr-badge { background:#e8f5e9; color:#2e7d32; }
.tr-changed   .tr-badge { background:#fff3e0; color:#e65100; }
.tr-removed   .tr-badge { background:#ffcdd2; color:#c62828; }
.tr-added     .tr-badge { background:#bbdefb; color:#1565c0; }
.tr-moved     .tr-badge { background:#ffe0b2; color:#bf360c; }

/* Word-like content styles */
.wd-normal  { margin: 4px 0; line-height: 1.5; }
.wd-empty   { margin: 2px 0; }
.wd-muted   { color: #bbb; }
.wd-h1 { font-size: 18px; font-weight: 700; color: #1F3864;
          margin: 12px 0 6px; border-bottom: 1px solid #ddd; }
.wd-h2 { font-size: 15px; font-weight: 700; color: #2E4E8A; margin: 10px 0 4px; }
.wd-h3 { font-size: 13px; font-weight: 700; color: #2E4E8A; margin: 8px 0 3px; }
.wd-h4, .wd-h5, .wd-h6 { font-size: 13px; font-weight: 600; margin: 6px 0 2px; }
.wd-li-bullet { margin: 2px 0 2px 20px; line-height: 1.4; }
.wd-li-bullet::before { content: "•"; margin-right: 6px; color: #555; }
.wd-li-number { margin: 2px 0 2px 20px; line-height: 1.4; }
.wd-caption { font-size: 11px; font-style: italic; color: #555;
              text-align: center; margin: 4px 0; }
.wd-quote   { border-left: 3px solid #ccc; padding-left: 10px;
              color: #555; font-style: italic; margin: 6px 0; }
.wd-code    { font-family: 'Courier New', monospace; font-size: 11px;
              background: #f5f5f5; padding: 8px; border-radius: 4px;
              white-space: pre-wrap; margin: 4px 0; }
code { font-family: 'Courier New', monospace; font-size: 11px;
       background: #f0f0f0; padding: 0 3px; border-radius: 2px; }

/* Tables */
.wd-table { border-collapse: collapse; width: 100%; margin: 8px 0;
            font-size: 12px; }
.wd-table th { background: #1F3864; color: #fff; font-weight: 600;
               padding: 5px 8px; text-align: left;
               border: 1px solid #ddd; }
.wd-table td { padding: 4px 8px; border: 1px solid #ddd;
               vertical-align: top; }
.wd-table tr:nth-child(even) td { background: #f7f7f7; }
.wd-img { background:#f0f4f8; border:1px solid #ddd; border-radius:4px;
          padding: 20px; text-align:center; color:#888; font-size:12px;
          font-style:italic; margin: 6px 0; }
.wd-pagebreak { border: none; border-top: 1px dashed #ccc; margin: 8px 0; }
"""

_ICONS = {
    'identical':     '✅',
    'changed':       '🟡',
    'removed':       '🔴',
    'added':         '🟢',
    'moved':         '⬆️',
    'moved_changed': '⬆️🟡',
}
_LABELS = {
    'identical':     'Identical',
    'changed':       'Changed',
    'removed':       'Removed',
    'added':         'Added',
    'moved':         'Moved',
    'moved_changed': 'Moved + Changed',
}


def _render_result_html(result: SectionResult,
                         base_doc, recv_doc,
                         depth: int = 0) -> str:
    status   = result.status
    icon     = _ICONS.get(status, '?')
    label    = _LABELS.get(status, status)
    sec      = result.baseline or result.received
    heading  = sec.heading
    level    = sec.level
    is_open  = status not in ('identical', 'contains_changes')
    open_attr = ' open' if is_open else ''

    # Badge
    badge_cls = f'badge-{status.replace("_changed","").replace("moved_","moved")}'

    indent  = "&nbsp;" * depth * 4
    summary = (
        f'<summary>'
        f'<span class="arrow">▶</span>'
        f'<span class="sec-title">{indent}'
        f'{icon} {_e(heading)}</span>'
        f'<span class="sec-badge {badge_cls}">{_e(label)}</span>'
        f'</summary>'
    )

    # Content
    content = ''
    if status == 'identical':
        # Show once — content is the same on both sides
        base_html = _render_section_content(result.baseline, base_doc) if result.baseline else ''
        content = (
            f'<div class="section-content">'
            f'<div class="diff-single">'
            f'{base_html}'
            f'</div></div>'
        )
    elif status == 'removed':
        base_html = _render_section_content(result.baseline, base_doc)
        content = (
            f'<div class="section-content">'
            f'<div class="diff-single">'
            f'<div class="diff-col-label">Your version (removed in received)</div>'
            f'{base_html}'
            f'</div></div>'
        )
    elif status == 'added':
        recv_html = _render_section_content(result.received, recv_doc)
        content = (
            f'<div class="section-content">'
            f'<div class="diff-single">'
            f'<div class="diff-col-label">Added in received (not in your source)</div>'
            f'{recv_html}'
            f'</div></div>'
        )
    elif status in ('moved', 'moved_changed', 'changed'):
        # If this is a pure container section (no direct elements, only children)
        # don't render side-by-side — the children cards handle the detail.
        # Only show content if there are direct elements to compare.
        has_direct_base = result.baseline and bool(result.baseline.elements)
        has_direct_recv = result.received and bool(result.received.elements)
        banner = ''
        if 'moved' in status:
            banner = (
                f'<div class="moved-banner">'
                f'⬆️ Section moved — was position {result.moved_from + 1}, '
                f'now position {result.moved_to + 1}'
                f'</div>'
            )
        if not has_direct_base and not has_direct_recv and result.children:
            # Pure container — just show the banner if moved, children handle the rest
            content = f'<div class="section-content">{banner}</div>' if banner else ''
        else:
            base_html = _render_section_content(result.baseline, base_doc) if result.baseline else ''
            recv_html = _render_section_content(result.received, recv_doc) if result.received else ''
            content = (
                f'<div class="section-content">'
                f'{banner}'
                f'<div class="diff-cols">'
                f'<div>'
                f'<div class="diff-col-label">Your version (source)</div>'
                f'{base_html}'
                f'</div>'
                f'<div>'
                f'<div class="diff-col-label">Received</div>'
                f'{recv_html}'
                f'</div>'
                f'</div></div>'
            )

    # Children
    children_html = ''
    if result.children:
        children_html = (
            f'<div class="children">'
            + ''.join(_render_result_html(c, base_doc, recv_doc, depth + 1)
                      for c in result.children)
            + '</div>'
        )

    sec_id = f"sec-{id(result)}"
    return (
        f'<details id="{sec_id}" class="status-{status.replace("_","-")}"{open_attr}>'
        f'{summary}'
        f'{content}'
        f'{children_html}'
        f'</details>'
    )


def _render_tree(results: List[SectionResult], depth: int = 0) -> str:
    """Render a compact tree overview of all sections with status badges."""
    rows = []
    for r in results:
        sec    = r.baseline or r.received
        status = r.status.replace('moved_changed', 'moved')
        icon   = _ICONS.get(r.status, '?')
        label  = _LABELS.get(r.status, r.status)
        indent = depth * 16

        # Section ID to scroll to
        sec_id = f"sec-{id(r)}"

        badge = (f'<span class="tr-badge">{_e(label)}</span>'
                 if r.status != "identical" else "")

        rows.append(
            f'<div class="tree-row tr-{status}" '
            f'style="padding-left:{indent + 4}px" '
            f'onclick="scrollToSection(\'{sec_id}\')">' 
            f'<span class="tr-icon">{icon}</span>'
            f'<span class="tr-name">{_e(sec.heading)}</span>'
            f'{badge}'
            f'</div>'
        )
        if r.children:
            rows.append(_render_tree(r.children, depth + 1))

    return "\n".join(rows)


def build_html_report(
        results: List[SectionResult],
        baseline_path: Path,
        received_path: Path,
        baseline_label: str = "Your version (source)",
        received_label: str = "Received",
) -> str:
    """Build the full self-contained HTML report."""
    base_doc = _DocX(str(baseline_path))
    recv_doc = _DocX(str(received_path))

    # Count summary
    counts = {'identical': 0, 'changed': 0, 'removed': 0,
              'added': 0, 'moved': 0, 'moved_changed': 0}

    def _count(rs):
        for r in rs:
            # Only count sections with direct content changes — not parents
            # that are marked changed solely because a child changed.
            has_direct = (
                r.status in ('removed', 'added', 'moved', 'moved_changed') or
                (r.status == 'changed' and
                 r.baseline is not None and r.received is not None and
                 r.baseline.content_hash != r.received.content_hash)
            )
            if has_direct:
                key = r.status.replace('moved_changed', 'moved')
                counts[key] = counts.get(key, 0) + 1
            elif r.status == 'identical':
                counts['identical'] = counts.get('identical', 0) + 1
            _count(r.children)

    _count(results)

    total    = sum(counts.values())
    n_issues = total - counts.get('identical', 0)

    summary_parts = []
    for status, n in counts.items():
        if n == 0:
            continue
        icon  = _ICONS[status]
        label = _LABELS[status]
        summary_parts.append(
            f'<span class="badge-{status.replace("_changed","").replace("moved_","moved")}">'
            f'{icon} {n} {label}</span>'
        )

    tree_html     = _render_tree(results)
    sections_html = ''.join(
        _render_result_html(r, base_doc, recv_doc) for r in results
    )

    title = f"Review Report — {received_path.name}"

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>{_e(title)}</title>
<style>{_CSS}</style>
</head>
<body>
<div class="report-header">
  <div>
    <h1>📄 Document Review Report</h1>
    <div class="sub">
      Source: {_e(baseline_label)} &nbsp;·&nbsp;
      Received: {_e(received_path.name)} &nbsp;·&nbsp;
      {n_issues} section(s) with changes out of {total} total
    </div>
  </div>
</div>
<div class="legend">
  {''.join(summary_parts)}
</div>
<div class="tree">
  <div class="tree-title">Document structure</div>
  {tree_html}
</div>
<div class="sections">
{sections_html}
</div>
<script>
// Scroll to section by ID
function scrollToSection(id) {{
  const el = document.getElementById(id);
  if (el) {{ el.scrollIntoView({{behavior:'smooth', block:'start'}}); el.open = true; }}
}}
// Keyboard shortcut: press N to jump to next changed section
let cards = Array.from(document.querySelectorAll('details:not(.status-identical)'));
let idx = -1;
document.addEventListener('keydown', e => {{
  if (e.key === 'n' || e.key === 'N') {{
    idx = (idx + 1) % cards.length;
    cards[idx].scrollIntoView({{behavior:'smooth', block:'start'}});
    cards[idx].open = true;
  }}
}});
</script>
</body>
</html>"""
