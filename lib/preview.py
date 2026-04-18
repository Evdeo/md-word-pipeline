"""Live-preview server: rebuilds the project on file change, converts
the docx to PDF (Word COM on Windows) and renders each page as SVG so a
browser can show the result. Polls a version file for hot-reload.

Public API:
    handle = start_preview_server(ctx, open_browser=True)
    print(handle.url)
    handle.wait()   # blocks until Ctrl+C
    handle.stop()   # idempotent

Windows-only in practice — the docx-to-PDF step uses Word COM, with a
docx2pdf fallback. On other platforms the server still starts but PDF
conversion will fail and the browser will show the last-good pages.
"""

from __future__ import annotations

import http.server
import re
import shutil
import socket
import socketserver
import sys
import threading
import time
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from lib.config_loader import ProjectContext, build
from lib.log import get_logger

log = get_logger(__name__)


PREVIEW_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Live Preview</title>
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body { background:#525659; font-family:sans-serif; }
#toolbar {
  position:fixed; top:0; left:0; right:0; height:36px;
  background:#323639; display:flex; align-items:center;
  padding:0 16px; gap:16px; z-index:100;
  box-shadow:0 1px 4px rgba(0,0,0,.4);
}
#toolbar span { color:#ccc; font-size:12px; }
#status { color:#7cb97c; font-size:11px; }
#container { margin-top:36px; padding:24px 0;
  display:flex; flex-direction:column; align-items:center; }
.page {
  box-shadow:0 4px 24px rgba(0,0,0,.5);
  margin-bottom:24px; background:white;
  width:min(794px, 95vw);
}
.page svg { display:block; width:100%; height:auto; }
.page svg text { user-select:text; cursor:text; }
</style>
</head>
<body>
<div id="toolbar">
  <span>\U0001f4c4 Live Preview</span>
  <span id="status">Loading\u2026</span>
</div>
<div id="container"></div>
<script>
let lastVer  = null;
let loading  = false;
const container = document.getElementById('container');
const status    = document.getElementById('status');

async function loadPages() {
  if (loading) return;
  loading = true;
  try {
    let n = 0;
    while (!n) {
      try {
        const r = await fetch('preview_pages.txt?t=' + Date.now());
        if (r.ok) { n = parseInt((await r.text()).trim()); }
      } catch(e) {}
      if (!n) await new Promise(res => setTimeout(res, 800));
    }
    const t = Date.now();
    const svgs = await Promise.all(
      Array.from({length: n}, (_, i) =>
        fetch(`preview_page_${i}.svg?t=${t}`).then(r => r.text())
      )
    );
    const scrollY = window.scrollY;
    const frag = document.createDocumentFragment();
    svgs.forEach(svg => {
      const wrap = document.createElement('div');
      wrap.className = 'page';
      wrap.innerHTML = svg;
      const svgEl = wrap.querySelector('svg');
      if (svgEl) {
        svgEl.removeAttribute('width');
        svgEl.removeAttribute('height');
        svgEl.setAttribute('preserveAspectRatio', 'xMidYMid meet');
      }
      frag.appendChild(wrap);
    });
    container.innerHTML = '';
    container.appendChild(frag);
    window.scrollTo(0, scrollY);
  } finally {
    loading = false;
  }
}

async function poll() {
  if (loading) return;
  try {
    const r = await fetch('.preview_version?t=' + Date.now());
    const v = await r.text();
    if (v !== lastVer) {
      lastVer = v;
      status.textContent = 'Rebuilding\u2026';
      status.style.color = '#e0c070';
      await loadPages();
      status.textContent = '\u2713 ' + new Date().toLocaleTimeString();
      status.style.color = '#7cb97c';
    }
  } catch(e) {}
}

loadPages().then(async () => {
  try {
    const r = await fetch('.preview_version?t=' + Date.now());
    lastVer = await r.text();
  } catch(e) {}
  status.textContent = '\u2713 ' + new Date().toLocaleTimeString();
  status.style.color = '#7cb97c';
  setInterval(poll, 1000);
}).catch(e => {
  status.textContent = 'Error: ' + e.message;
  status.style.color = '#e07070';
});
</script>
</body>
</html>"""


def _docx_to_pdf(docx: Path, pdf: Path) -> None:
    """Convert docx → pdf via Word COM, with docx2pdf fallback.
    Uses Dispatch (not DispatchEx) so an already-open Word instance is
    reused — DispatchEx would conflict with the user's open doc."""
    if sys.platform == "win32":
        import pythoncom  # type: ignore
        pythoncom.CoInitialize()
    try:
        if sys.platform == "win32":
            try:
                _com_export(docx, pdf)
                return
            except Exception as e:
                log.debug("COM export failed (%s); trying docx2pdf", e)
        _docx2pdf_export(docx, pdf)
    finally:
        if sys.platform == "win32":
            import pythoncom  # type: ignore
            pythoncom.CoUninitialize()


def _com_export(docx: Path, pdf: Path) -> None:
    import win32com.client as wc  # type: ignore
    word = wc.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    doc = None
    try:
        for attempt in range(3):
            try:
                doc = word.Documents.Open(
                    str(docx),
                    ConfirmConversions=False,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    NoEncodingDialog=True,
                )
                break
            except Exception:
                if attempt == 2:
                    raise
                time.sleep(1.0)
        doc.ExportAsFixedFormat(
            str(pdf),
            17,     # wdExportFormatPDF
            False,  # OpenAfterExport
            0,      # OptimizeFor: print
        )
    finally:
        if doc is not None:
            for attempt in range(3):
                try:
                    doc.Close(SaveChanges=False)
                    break
                except Exception:
                    if attempt == 2:
                        break
                    time.sleep(0.5)
        # Don't quit Word — the user may have other docs open.


def _docx2pdf_export(docx: Path, pdf: Path) -> None:
    import io
    from docx2pdf import convert  # type: ignore

    saved_out, saved_err = sys.stdout, sys.stderr
    sys.stdout = sys.stderr = io.StringIO()
    try:
        convert(str(docx), str(pdf))
    finally:
        sys.stdout, sys.stderr = saved_out, saved_err


def _prefix_svg_ids(svg: str, page_num: int) -> str:
    """Prefix every SVG id with a per-page prefix so ids don't collide when
    multiple pages are inlined into the same HTML document."""
    prefix = f"p{page_num}-"
    ids = set(re.findall(r'\bid="([^"]+)"', svg))
    if not ids:
        return svg
    pat = "|".join(re.escape(i) for i in sorted(ids, key=len, reverse=True))
    svg = re.sub(rf'\bid="({pat})"',
                 lambda m: f'id="{prefix}{m.group(1)}"', svg)
    svg = re.sub(rf'href="#({pat})"',
                 lambda m: f'href="#{prefix}{m.group(1)}"', svg)
    svg = re.sub(rf'url\(#({pat})\)',
                 lambda m: f'url(#{prefix}{m.group(1)})', svg)
    svg = re.sub(rf'xlink:href="#({pat})"',
                 lambda m: f'xlink:href="#{prefix}{m.group(1)}"', svg)
    return svg


def _pdf_to_svgs(pdf: Path, render_dir: Path) -> int:
    import pymupdf  # type: ignore

    doc = pymupdf.open(str(pdf))
    try:
        n = doc.page_count
        for old in render_dir.glob("preview_page_*.svg"):
            try:
                idx = int(old.stem.split("_")[-1])
                if idx >= n:
                    old.unlink()
            except ValueError:
                pass

        for i, page in enumerate(doc):
            svg_visual = page.get_svg_image(text_as_path=1)
            svg_visual = _prefix_svg_ids(svg_visual, i)

            # Add an invisible text overlay so Ctrl+F works in the browser.
            svg_text = page.get_svg_image(text_as_path=0)
            text_els = re.findall(r"<text\b.*?</text>", svg_text, re.DOTALL)
            if text_els:
                overlay = (
                    '<g style="fill:transparent;pointer-events:all;" '
                    'aria-hidden="true">\n'
                    + "\n".join(text_els) + "\n</g>"
                )
                svg_visual = svg_visual.rstrip()
                if svg_visual.endswith("</svg>"):
                    svg_visual = svg_visual[:-6] + "\n" + overlay + "\n</svg>"

            (render_dir / f"preview_page_{i}.svg").write_text(
                svg_visual, encoding="utf-8")
        return n
    finally:
        doc.close()


def _rebuild(ctx: ProjectContext, render_dir: Path) -> None:
    preview_docx = render_dir / "preview.docx"
    preview_pdf  = render_dir / "preview.pdf"
    ver_file     = render_dir / ".preview_version"

    build(ctx)                          # refresh document.docx
    shutil.copy2(ctx.output_path, preview_docx)  # independent copy for COM

    _docx_to_pdf(preview_docx, preview_pdf)
    if not preview_pdf.exists() or preview_pdf.stat().st_size < 100:
        raise RuntimeError("PDF not produced by Word")

    n = _pdf_to_svgs(preview_pdf, render_dir)
    (render_dir / "preview_pages.txt").write_text(str(n), encoding="utf-8")
    ver_file.write_text(str(time.time()), encoding="utf-8")


def _clear_stale(render_dir: Path) -> None:
    for stale in list(render_dir.glob("preview_page_*.svg")) + [
            render_dir / "preview_pages.txt",
            render_dir / "preview.pdf",
            render_dir / "preview.docx",
            render_dir / ".preview_version"]:
        try:
            stale.unlink()
        except (FileNotFoundError, PermissionError):
            pass


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("", 0))
        return s.getsockname()[1]


@dataclass
class PreviewHandle:
    url: str
    server: socketserver.TCPServer
    observer: Optional[object]    # watchdog.observers.Observer
    _stop_event: threading.Event

    def wait(self) -> None:
        try:
            while not self._stop_event.wait(1.0):
                pass
        except KeyboardInterrupt:
            self.stop()

    def stop(self) -> None:
        if self._stop_event.is_set():
            return
        self._stop_event.set()
        try:
            self.server.shutdown()
        except Exception:
            pass
        if self.observer is not None:
            try:
                self.observer.stop()
                self.observer.join(timeout=2.0)
            except Exception:
                pass


def start_preview_server(ctx: ProjectContext, *,
                         open_browser: bool = True) -> PreviewHandle:
    """Build the project, start the watchdog observer and the HTTP server,
    return a handle. The caller is expected to `wait()` until Ctrl+C."""
    render_dir = ctx.render_dir
    render_dir.mkdir(parents=True, exist_ok=True)
    _clear_stale(render_dir)
    (render_dir / "preview.html").write_text(PREVIEW_HTML, encoding="utf-8")

    build_lock = threading.Lock()

    def _safe_rebuild() -> None:
        if not build_lock.acquire(blocking=False):
            return
        try:
            _rebuild(ctx, render_dir)
        except Exception as e:
            log.warning("preview rebuild failed: %s", e)
        finally:
            build_lock.release()

    # Initial build so the first browser hit shows something.
    _safe_rebuild()

    observer = None
    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler

        class _Handler(FileSystemEventHandler):
            def __init__(self) -> None:
                self._last = 0.0

            def on_modified(self, event):  # noqa: D401, ANN001
                if event.is_directory:
                    return
                p = Path(event.src_path)
                if p.suffix not in (".md", ".yaml", ".yml"):
                    return
                if p.name.startswith("."):
                    return
                now = time.time()
                if now - self._last < 0.5:
                    return
                self._last = now
                threading.Thread(target=_safe_rebuild, daemon=True).start()

        observer = Observer()
        observer.schedule(_Handler(), str(ctx.project_dir), recursive=False)
        observer.start()
    except ImportError:
        log.warning("watchdog not installed — preview will not hot-reload")

    port = _free_port()

    class _Handler(http.server.SimpleHTTPRequestHandler):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, directory=str(render_dir), **kwargs)

        def log_message(self, *args):  # silence default request log
            pass

        def end_headers(self):
            self.send_header("Cache-Control",
                             "no-store, no-cache, must-revalidate")
            self.send_header("Pragma", "no-cache")
            super().end_headers()

    server = socketserver.TCPServer(("", port), _Handler)
    server.allow_reuse_address = True
    threading.Thread(target=server.serve_forever, daemon=True).start()

    url = f"http://localhost:{port}/preview.html"
    if open_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass

    return PreviewHandle(
        url=url, server=server, observer=observer,
        _stop_event=threading.Event(),
    )
