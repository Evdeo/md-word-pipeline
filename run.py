#!/usr/bin/env python3
"""
run.py -- md-to-docx workflow tool.

Run with:  python run.py
Projects live in projects/ next to this file.
"""

import os, re, shutil, subprocess, sys, tempfile
from datetime import datetime
from pathlib import Path

# Find ROOT as the directory that contains both run.py and lib/
# Using Path(__file__).resolve() ensures this works regardless of
# where Python is invoked from (current directory, parent, etc.)
_here = Path(__file__).resolve().parent
ROOT  = _here if (_here / "lib").exists() else Path.cwd()
PROJECTS_DIR = ROOT / "projects"
sys.path.insert(0, str(ROOT))


# -- Dependency check ---------------------------------------------------------

def _ensure_requirements():
    """Install any missing packages from lib/requirements.txt."""
    import importlib
    req_path = ROOT / "lib" / "requirements.txt"
    if not req_path.exists():
        return
    import_names = {
        "python-docx": "docx", "marko": "marko",
        "pillow": "PIL", "pyyaml": "yaml", "rich": "rich",
    }
    missing = []
    for line in req_path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        pkg = re.split(r"[><=!;\[]", line)[0].strip().lower()
        mod = import_names.get(pkg, pkg.replace("-", "_"))
        try:
            importlib.import_module(mod)
        except ImportError:
            missing.append(line)
    if not missing:
        return
    print(f"\n  Installing {len(missing)} missing package(s):")
    for pkg in missing:
        print(f"    - {pkg}")
    print()
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install",
         "--break-system-packages", "--quiet", *missing],
    )
    if result.returncode != 0:
        print("\n  Installation failed. Try running manually:")
        print(f"    pip install {' '.join(missing)}")
        sys.exit(1)
    print("  All packages installed.\n")


_ensure_requirements()


# -- Third-party imports ------------------------------------------------------

try:
    from rich.console import Console
    from rich.panel   import Panel
    from rich.table   import Table
    from rich.text    import Text
    from rich.rule    import Rule
    from rich.prompt  import Prompt
    from rich         import box
    HAS_RICH = True
except ImportError:
    HAS_RICH = False

console = Console() if HAS_RICH else None

import yaml
import marko
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# -- UI helpers & terminal menu -------------------------------------


_LAST_FILE = ROOT / "lib" / ".last_project"

def _load_last() -> str:
    try:    return _LAST_FILE.read_text().strip()
    except: return ""

def _save_last(slug: str):
    try:    _LAST_FILE.write_text(slug)
    except: pass


# ── UI helpers ─────────────────────────────────────────────────────────────────

def _clear():
    os.system("cls" if os.name == "nt" else "clear")

def _pause():
    try:    input("\nPress Enter to continue…")
    except: pass

def _print(msg, style=""):
    if console: console.print(msg, style=style)
    else:       print(re.sub(r'\[/?[^\]]*\]', '', str(msg)))

def _rule(title=""):
    if console: console.print(Rule(title, style="dim"))
    else:       print(f"\n{'─'*60}  {title}")

def _inp(prompt: str, default: str = "") -> str:
    display = f"  {prompt} [{default}]: " if default else f"  {prompt}: "
    try:
        v = input(display).strip()
        return v if v else default
    except (EOFError, KeyboardInterrupt):
        return default

def _confirm(prompt: str) -> bool:
    try:    return input(f"  {prompt} [y/N] ").strip().lower() == "y"
    except: return False

def _file_age(path: Path) -> str:
    if not path.exists(): return ""
    secs = datetime.now().timestamp() - path.stat().st_mtime
    if secs < 120:   return "just now"
    if secs < 3600:  return f"{int(secs//60)}m ago"
    if secs < 86400: return f"{int(secs//3600)}h ago"
    return f"{int(secs//86400)}d ago"

def _pick_file(title: str, initial_dir: Path = None) -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.lift()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Word documents","*.docx"),("All files","*.*")],
            initialdir=str(initial_dir or Path.home()),
        )
        root.destroy()
        if path: return Path(path)
    except Exception: pass
    try:
        typed = input(f"\n{title}\nPath: ").strip().strip('"').strip("'")
        return Path(typed) if typed else None
    except: return None

def _open_path(path: Path):
    """Open a file or folder with the OS default application."""
    try:
        if sys.platform == "win32": os.startfile(str(path))
        elif sys.platform == "darwin": subprocess.run(["open", str(path)])
        else: subprocess.run(["xdg-open", str(path)])
    except Exception as e:
        _print(f"[yellow]Could not open: {e}[/yellow]")

def _open_vscode(folder: Path):
    """Open a folder in VS Code, with fallback to OS file manager."""
    # Try 1: code command (works when VS Code is in PATH)
    try:
        subprocess.Popen(["code", str(folder)])
        _print(f"[green]Opened in VS Code: {folder.name}[/green]")
        return
    except FileNotFoundError:
        pass
    except Exception:
        pass

    # Try 2: known install locations
    candidates = []
    if sys.platform == "win32":
        import os
        local = os.environ.get("LOCALAPPDATA", "")
        candidates = [
            Path(local) / "Programs" / "Microsoft VS Code" / "Code.exe",
            Path("C:/Program Files/Microsoft VS Code/Code.exe"),
            Path("C:/Program Files (x86)/Microsoft VS Code/Code.exe"),
        ]
    elif sys.platform == "darwin":
        candidates = [
            Path("/Applications/Visual Studio Code.app/Contents/Resources/app/bin/code"),
            Path(Path.home() / "Applications/Visual Studio Code.app/Contents/Resources/app/bin/code"),
        ]
    else:
        candidates = [Path("/usr/bin/code"), Path("/usr/local/bin/code"),
                      Path("/snap/bin/code")]

    for exe in candidates:
        if exe.exists():
            try:
                subprocess.Popen([str(exe), str(folder)])
                _print(f"[green]Opened in VS Code: {folder.name}[/green]")
                return
            except Exception:
                continue

    # Try 3: open as default app (file manager / Finder / Explorer)
    _print("[dim]VS Code not found — opening folder in file manager instead.[/dim]")
    try:
        _open_path(folder)
        return
    except Exception:
        pass

    _print("[yellow]Could not open folder automatically.[/yellow]")
    _print(f"  Path: {folder}")


# ── project helpers ────────────────────────────────────────────────────────────

def _slugify(text: str) -> str:
    text = text.lower().strip()
    text = re.sub(r'[^\w\s-]', '', text)
    text = re.sub(r'[\s_]+', '-', text)
    text = re.sub(r'-+', '-', text).strip('-')
    return text or "project"

def _unique_slug(name: str) -> str:
    base = _slugify(name)
    slug = base
    n    = 2
    while (PROJECTS_DIR / slug).exists():
        slug = f"{base}-{n}"; n += 1
    return slug



def _menu(items, title="", extras=None, initial=0):
    """Arrow-key menu. Returns selected index (into non-None items), or a string key.

    items:   list of (label, note) tuples, or None for a separator line
    extras:  list of (key, label) for fixed options below (e.g. Back, Quit)
    initial: which selectable item starts highlighted

    Falls back to numbered input when not running in a real terminal (e.g. Windows).
    """
    import shutil as _shutil
    extras     = extras or []
    # Separate selectable items from separators, keeping index mapping
    # sel_items: only non-None, non-disabled items are selectable
    sel_items  = [(i, it) for i, it in enumerate(items)
                  if it is not None and (len(it) < 3 or not it[2])]
    # initial refers to selectable index
    

    # ── curses path (real terminal) ───────────────────────────────────────────
    def _curses_menu(stdscr):
        import curses as _curses
        _curses.curs_set(0)
        _curses.use_default_colors()
        try:
            _curses.init_pair(1, _curses.COLOR_CYAN,  -1)  # selected
            _curses.init_pair(2, _curses.COLOR_WHITE, -1)  # normal
            _curses.init_pair(3, 8,                   -1)  # dim
        except Exception:
            _curses.init_pair(1, _curses.COLOR_CYAN,  _curses.COLOR_BLACK)
            _curses.init_pair(2, _curses.COLOR_WHITE, _curses.COLOR_BLACK)
            _curses.init_pair(3, _curses.COLOR_WHITE, _curses.COLOR_BLACK)

        sel = initial

        while True:
            stdscr.erase()
            h, w = stdscr.getmaxyx()
            max_title = max((len(it[0]) for it in items), default=20)
            col2      = min(max_title + 6, w - 30)

            row = 0
            if title:
                try:
                    stdscr.addstr(row, 2, title, _curses.A_BOLD)
                except Exception: pass
                hint = "  ↑↓ move   Enter select   Esc quit"
                try:
                    stdscr.addstr(row, 2 + len(title) + 2,
                                  hint[:w - len(title) - 6],
                                  _curses.color_pair(3))
                except Exception: pass
                row += 1
                try:
                    stdscr.addstr(row, 2, "─" * min(w - 4, 60), _curses.color_pair(3))
                except Exception: pass
                row += 2

            sel_idx = 0
            for item in items:
                if row >= h - 1: break
                if item is None:
                    try:
                        stdscr.addstr(row, 2, " ", _curses.color_pair(3))
                    except Exception: pass
                    row += 1
                    continue
                label    = item[0]
                note     = item[1] if len(item) > 1 else ""
                disabled = item[2] if len(item) > 2 else False
                is_sel   = (not disabled) and (sel_idx == sel)
                if disabled:
                    marker = " "
                    lattr  = _curses.color_pair(3)
                    nattr  = _curses.color_pair(3)
                elif is_sel:
                    marker = "▶"
                    lattr  = _curses.color_pair(1) | _curses.A_BOLD
                    nattr  = _curses.color_pair(1)
                else:
                    marker = " "
                    lattr  = _curses.color_pair(2)
                    nattr  = _curses.color_pair(3)

                label_trunc = label[:col2]
                try:
                    stdscr.addstr(row, 2, f"{marker} {label_trunc:<{col2}}", lattr)
                    if note and col2 + 6 < w:
                        stdscr.addstr(row, col2 + 5,
                                      note[:w - col2 - 7], nattr)
                except Exception: pass
                row += 1
                if not disabled:
                    sel_idx += 1

            if extras and row + 2 < h:
                row += 1
                try:
                    stdscr.addstr(row, 2,
                                  "─" * min(w - 4, 60), _curses.color_pair(3))
                except Exception: pass
                row += 1
                for key, lbl in extras:
                    if row >= h - 1: break
                    try:
                        stdscr.addstr(row, 2,
                                      f"  {key.upper()}  {lbl}",
                                      _curses.color_pair(3))
                    except Exception: pass
                    row += 1

            stdscr.refresh()
            k = stdscr.getch()

            n = len(sel_items)
            if k in (_curses.KEY_UP,   ord('k')) and sel > 0:     sel -= 1
            elif k in (_curses.KEY_DOWN, ord('j')) and sel < n-1: sel += 1
            elif k in (_curses.KEY_ENTER, 10, 13):
                return sel_items[sel][0]   # return original index
            elif k == 27:                  return "q"
            else:
                for key, _ in extras:
                    if k in (ord(key.lower()), ord(key.upper())):
                        return key.lower()

    # ── fallback path (no TTY / piped input) ─────────────────────────────────
    def _plain_menu():
        w   = _shutil.get_terminal_size((80, 24)).columns
        sep = "─" * min(w - 2, 70)
        hdr = title or "Select"
        print(f"\n{hdr}")
        print(sep)
        num = 1
        for item in items:
            if item is None:
                print()
                continue
            label    = item[0]
            note     = item[1] if len(item) > 1 else ""
            disabled = item[2] if len(item) > 2 else False
            # Truncate label to 30 chars and pad to fixed width so
            # the note column always aligns regardless of title length
            label_disp = label if len(label) <= 30 else label[:29] + "\u2026"
            padded   = f"{label_disp:<30}"
            note_str = f"  {note}" if note else ""
            if disabled:
                print(f"   {'---':>2}  {padded}{note_str}")
            else:
                print(f"  [{num:>2}] {padded}{note_str}")
                num += 1
        if extras:
            print(sep)
            for key, lbl in extras:
                print(f"  [{key.upper()}] {lbl}")
        print()
        while True:
            try:
                raw = input("Select: ").strip().lower()
            except (EOFError, KeyboardInterrupt):
                return "q"
            if raw.isdigit() and 1 <= int(raw) <= len(sel_items):
                return sel_items[int(raw) - 1][0]
            for key, _ in extras:
                if raw == key.lower():
                    return raw
            print(f"  Enter a number 1–{len(sel_items)}"
                  + (f" or a letter" if extras else ""))

    if sys.stdin.isatty():
        try:
            import curses as _curses
            return _curses.wrapper(_curses_menu)
        except Exception:
            pass
    return _plain_menu()


def _show_picker(projects: list[Path]) -> str | None:
    """Project picker with arrow-key navigation."""
    last = _load_last()

    initial = 0
    for i, p in enumerate(projects):
        if p.name == last:
            initial = i; break

    # Fixed-width columns so the note field never wraps regardless of terminal width.
    # title(32) | version(7) | author(20) | age
    items = []
    for proj in projects:
        info    = _project_display(proj)
        title   = info["title"]
        if len(title) > 28: title = title[:27] + "\u2026"
        version = (f"v{info['version']}" if info["version"] else "")[:7]
        author  = (info["author"]          if info["author"]  else "")[:20]
        age     = _file_age(proj / "input")[:10]
        note    = f"{version:<8}{author:<21}{age}"
        items.append((title, note))

    _clear()
    n_arch = len(_archived_projects())
    arch_label = f"Archive a project  ({n_arch} archived)" if n_arch else "Archive a project"

    picker_extras = [("a", "New project"), ("s", arch_label)]
    if n_arch:
        picker_extras.append(("f", "Unarchive a project"))
    picker_extras.append(("d", "Change default config"))
    picker_extras.append(("q", "Quit"))

    result = _menu(items,
                   title="Select a project",
                   extras=picker_extras,
                   initial=initial)

    if isinstance(result, int):
        return projects[result].name
    return result

# ── project dashboard ──────────────────────────────────────────────────────────

def _show_dashboard(state: dict):
    """Print the static project info panel, then show the arrow-key action menu."""
    proj_dir = state["proj_dir"]
    info     = _project_display(proj_dir)

    _clear()

    if HAS_RICH:
        t = Table(box=box.SIMPLE, show_header=False, padding=(0,1))
        t.add_column(width=12, style="dim")
        t.add_column()

        t.add_row("Project", f"[bold]{info['title']}[/bold]")
        meta_parts = filter(None, [
            f"v{info['version']}"  if info["version"]        else "",
            info["author"]          if info["author"]         else "",
            f"[dim]{info['classification']}[/dim]"
                                    if info["classification"] else "",
        ])
        t.add_row("", "  ".join(meta_parts))
        t.add_row("", "")

        n_md = len(state["md_files"])
        if n_md:
            latest = _file_age(max(state["md_files"], key=lambda f: f.stat().st_mtime))
            t.add_row("Source",
                      f"[green]✓[/green]  {n_md} file{'s' if n_md!=1 else ''}  "
                      f"[dim](last edit {latest})[/dim]")
        else:
            t.add_row("Source", "[red]✗  no markdown files found[/red]")

        if state["built_docx"].exists():
            age = _file_age(state["built_docx"])
            if state["sync_status"] == "source_newer":
                t.add_row("Built",
                          f"[yellow]⚠  source edited after last build  "
                          f"[dim]({age})[/dim][/yellow]")
            else:
                t.add_row("Built",
                          f"[green]✓[/green]  document.docx  [dim]({age})[/dim]")
        else:
            t.add_row("Built", "[dim]—  not yet built[/dim]")

        n_rx = len(state["received_files"])
        if n_rx == 0:
            t.add_row("Received", "[dim]—  nothing waiting[/dim]")
        elif n_rx == 1:
            f = state["received_files"][0]
            t.add_row("Received",
                      f"[cyan]● {f.name}  [dim]({_file_age(f)})[/dim][/cyan]")
        else:
            t.add_row("Received", f"[cyan]● {n_rx} files waiting[/cyan]")

        console.print()
        console.print(Panel(t, title="[bold]Project Dashboard[/bold]",
                            subtitle=f"[dim]projects/{proj_dir.name}[/dim]",
                            border_style="blue"))
    else:
        n_md = len(state["md_files"])
        n_rx = len(state["received_files"])
        print(f"\n  {info['title']}  v{info['version']}  {info['author']}")
        print(f"  Source: {'✓' if n_md else '✗'}  "
              f"Built: {'✓' if state['built_docx'].exists() else '—'}  "
              f"Received: {n_rx or '—'}")


def _prompt_dashboard(state: dict) -> str:
    """Arrow-key menu for the project dashboard. Returns action key."""
    no_src   = not state["md_files"]
    no_built = not state["built_docx"].exists()
    no_rx    = not state["received_files"]

    items = [
        ("Build",
         "markdown → Word" + (" — write content first" if no_src else "")),
        ("Open document",
         "open built Word file" + (" — not yet built" if no_built else "")),
        ("Export",
         "save Word file to a chosen location" + (" — not yet built" if no_built else "")),
        None,  # ── separator ──
        ("Open in VS Code", "open project folder in editor"),
        ("Edit info",       "title, author, version"),
        ("Edit properties", "{{placeholder}} values"),
        ("Inspect template","extract styles → config.yaml"),
        None,  # ── separator ──
        ("Receive file",    "import a reviewer's Word file"),
        ("Review changes",
         "compare received against source" if not no_rx
         else "receive a file first  (option 8)",
         no_rx),   # disabled when no file received
    ]

    # Build a direct map: original_index_in_items → action_key
    # This is immune to separator positions shifting the numbering.
    action_keys = ["1","2","3","4","5","6","7","8","9","10"]
    idx_to_key = {}
    key_num = 0
    for i, item in enumerate(items):
        if item is not None:
            idx_to_key[i] = action_keys[key_num]
            key_num += 1

    result = _menu(items,
                   title="",
                   extras=[("b","Back to projects"), ("q","Quit")])

    if isinstance(result, int):
        return idx_to_key.get(result, "b")
    return result  # "b" or "q"

# ── project actions ────────────────────────────────────────────────────────────



# -- Project management ---------------------------------------------

PROJECTS_DIR = ROOT / "projects"
ARCHIVE_DIR  = PROJECTS_DIR / ".archive"

def _project_display(proj_dir: Path) -> dict:
    """Read display metadata from a project's document-info.yaml."""
    di = proj_dir / "input" / "document-info.yaml"
    info = {"title": proj_dir.name, "author": "", "version": "", "classification": ""}
    if di.exists():
        try:
            import yaml
            data = yaml.safe_load(di.read_text(encoding="utf-8")) or {}
            doc  = data.get("document", {})
            info["title"]          = doc.get("title",          proj_dir.name)
            info["author"]         = doc.get("author",         "")
            info["version"]        = doc.get("version",        "")
            info["classification"] = doc.get("classification", "")
        except Exception: pass
    return info

def _project_mtime(proj_dir: Path) -> float:
    """Most recent mtime of any file inside a project."""
    try:
        return max(
            f.stat().st_mtime
            for f in (proj_dir / "input").rglob("*")
            if f.is_file()
        )
    except (ValueError, FileNotFoundError):
        return proj_dir.stat().st_mtime

def _list_projects() -> list[Path]:
    """Return project dirs sorted by most recently edited, last-used first."""
    if not PROJECTS_DIR.exists():
        return []
    dirs = [d for d in PROJECTS_DIR.iterdir()
            if d.is_dir() and (d / "input").exists()]
    last = _load_last()
    dirs.sort(key=lambda d: (d.name != last, -_project_mtime(d)))
    return dirs

def _project_state(proj_dir: Path) -> dict:
    input_dir    = proj_dir / "input"
    output_dir   = proj_dir / "output"
    built_docx   = output_dir / "document.docx"
    received_dir = output_dir / "received"

    md_files = sorted(f for f in input_dir.glob("*.md")
                      if f.name != "00-frontpage.md") if input_dir.exists() else []
    received = sorted(received_dir.glob("*.docx")) if received_dir.exists() else []

    sync_status = "unknown"
    if built_docx.exists() and md_files:
        docx_t = built_docx.stat().st_mtime
        src_t  = max(f.stat().st_mtime for f in md_files)
        yamls  = list(input_dir.glob("*.yaml")) if input_dir.exists() else []
        if yamls: src_t = max(src_t, max(f.stat().st_mtime for f in yamls))
        sync_status = "source_newer" if src_t > docx_t else "built"

    return {
        "proj_dir":       proj_dir,
        "input_dir":      input_dir,
        "output_dir":     output_dir,
        "built_docx":     built_docx,
        "received_files": received,
        "md_files":       md_files,
        "sync_status":    sync_status,
    }


# ── template content ───────────────────────────────────────────────────────────

FRONTPAGE_MD = """\
::: {toc=false align=center size=32 color=#1F3864}
# {{document.title}}
## {{document.subtitle}}
:::

::: {toc=false align=center size=14 color=#666666}
**{{document.document_type}}**
:::

::: {toc=false align=center size=12}
**Version:** {{document.version}} | **Classification:** {{document.classification}}
:::

:::space{lines=6}

{{revisions.table}}
"""

def _document_info_yaml(title: str, author: str, version: str,
                         classification: str) -> str:
    date = datetime.now().strftime("%B %Y")
    return f"""\
# Document Identity & Revision History
# Edit this file to update the document title, author, version, and revision history.

document:
  title: "{title}"
  subtitle: ""
  author: "{author}"
  date: "{date}"
  version: "{version}"
  classification: "{classification}"
  document_type: ""

revisions:
  - version: "{version}"
    date: "{datetime.now().strftime('%Y-%m-%d')}"
    author: "{author}"
    changes: "Initial version"
"""

PROPERTIES_YAML = """\
# Document Properties
# Reference these values anywhere in your markdown using {{key}} syntax.
# Example: The project version is {{project.version}}
#
# Add your own sections and keys freely.

project:
  name: "Acme Corp"
  version: "1.0"
"""

CONFIG_YAML = """
# ══════════════════════════════════════════════════════════════════════════════
#  Document Configuration
#  All formatting, layout, and style settings live here.
#  Changing any value and rebuilding will update the Word output immediately.
#  Document identity (title, author, version) lives in document-info.yaml.
# ══════════════════════════════════════════════════════════════════════════════


# ── Page layout ───────────────────────────────────────────────────────────────

page:
  size:            "A4"       # A4 or Letter
  orientation:     "portrait" # portrait or landscape

  # Margins — distance from paper edge to body text
  margin_top:      "2.54cm"
  margin_bottom:   "2.54cm"
  margin_left:     "2.54cm"
  margin_right:    "2.54cm"

  # Distance from paper edge to header/footer content.
  # Must be less than the corresponding margin or content will overlap.
  header_distance: "1.25cm"
  footer_distance: "1.25cm"


# ── Header ────────────────────────────────────────────────────────────────────
# Available placeholders: {title}  {author}  {date}  {version}  {classification}
# Each zone accepts a single string or a list of strings (one per line):
#   left:
#     - "{title}"
#     - "{author}"

header:
  left:   "{title}"
  center: ""
  right:  "{date}"

  # Logo / image in the page header.
  # Path is relative to this project's input/ folder.
  # image: images/logo.png
  # image_height_cm: 1.0      # Height in cm; width scales proportionally.
  # image_position: right     # "right" (default) or "left"


# ── Footer ────────────────────────────────────────────────────────────────────

footer:
  left:   ""
  center: "Page {page} of {total}"
  right:  "{author}"
  page_total: content   # "content" = content pages only (default) | "document" = whole document


# ── Header and footer separator lines ────────────────────────────────────────

header_line:
  show:  false
  color: "AAAAAA"   # Hex colour without #
  width: 6          # Line thickness in half-points (6 = 0.75pt)

footer_line:
  show:  false
  color: "AAAAAA"
  width: 6


# ── Front matter ─────────────────────────────────────────────────────────────
# Sets the page number shown at the start of each section.
# The header is suppressed on cover and TOC pages and only starts from content pages.
# {page} and {total} in the footer are both content-relative:
#   - {page}  counts up from content_start_page
#   - {total} counts only content pages (not cover or TOC)
# So with content_start_page: 1, a 20-page document shows "Page 1 of 20" to "Page 20 of 20".
#
# If your cover is 2 pages set cover_start_page: 1, toc_start_page: 3.

frontpage:
  cover_start_page:   1   # first page number assigned to cover pages (never shown)
  toc_start_page:     2   # first page number shown on TOC pages
  content_start_page: 1   # {page} starts here; {total} = number of content pages only

  # Use a Word file as the cover page instead of the built-in 00-frontpage.md layout.
  # Path is relative to this project's input/ folder.
  # word_cover: covers/my-cover.docx


# ── Headings ──────────────────────────────────────────────────────────────────

numbered_headings: true

styles:
  heading_1:
    font_name:       "Calibri"
    font_size_pt:    22
    bold:            true
    color:           "1F3864"
    space_before_pt: 12
    space_after_pt:  6

  heading_2:
    font_name:       "Calibri"
    font_size_pt:    16
    bold:            true
    color:           "2E75B6"
    space_before_pt: 10
    space_after_pt:  4

  heading_3:
    font_name:       "Calibri"
    font_size_pt:    13
    bold:            true
    color:           "1F3864"
    space_before_pt: 8
    space_after_pt:  2

  heading_4:
    font_name:       "Calibri"
    font_size_pt:    12
    bold:            true
    color:           "2E75B6"
    space_before_pt: 6
    space_after_pt:  2

  heading_5:
    font_name:       "Calibri"
    font_size_pt:    11
    bold:            true
    color:           "1F3864"
    space_before_pt: 4
    space_after_pt:  2

  heading_6:
    font_name:       "Calibri"
    font_size_pt:    11
    bold:            false
    color:           "2E75B6"
    space_before_pt: 4
    space_after_pt:  2


# ── Body text ─────────────────────────────────────────────────────────────────

  normal:
    font_name:      "Calibri"
    font_size_pt:   11
    space_after_pt: 6


# ── Code blocks ───────────────────────────────────────────────────────────────

  code:
    font_name:        "Courier New"
    font_size_pt:     9
    background:       "F0F0F0"
    border_color:     "AAAAAA"
    left_indent_in:   0.15
    right_indent_in:  0.15
    space_before_pt:  2
    space_after_pt:   2


# ── Block quotes ──────────────────────────────────────────────────────────────

  block_quote:
    font_italic:      true
    color:            "444444"
    bar_color:        "2E75B6"
    left_indent_in:   0.15
    right_indent_in:  0.15
    space_before_pt:  4
    space_after_pt:   4


# ── Captions ──────────────────────────────────────────────────────────────────

  caption:
    font_size_pt:    9
    color:           "555555"
    italic:          true
    space_before_pt: 2
    space_after_pt:  8


# ── Table styling ─────────────────────────────────────────────────────────────

  table_header:
    background:  "1F3864"
    font_color:  "FFFFFF"
    bold:        true

  table_rows:
    odd_background:  "F7F7F7"
    even_background: "FFFFFF"


# ── Alert boxes ───────────────────────────────────────────────────────────────

  alerts:
    note_color:    "2E75B6"
    tip_color:     "28A745"
    warning_color: "FFA500"
    caution_color: "DC3545"
    background:    "F5F5F5"


# ── Cover page styles ─────────────────────────────────────────────────────────

  cover_title:
    font_name:       "Calibri"
    font_size_pt:    22
    color:           "1F3864"
    bold:            true
    space_before_pt: 24
    space_after_pt:  8

  cover_subtitle:
    font_name:       "Calibri"
    font_size_pt:    14
    color:           "2E75B6"
    bold:            false
    space_before_pt: 10
    space_after_pt:  6

  cover_body:
    font_name:       "Calibri"
    font_size_pt:    11
    color:           "000000"
    bold:            false
    space_before_pt: 6
    space_after_pt:  4


# ── Image sizes ───────────────────────────────────────────────────────────────

image_sizes:
  xs:     { max_pct: 20  }
  small:  { max_pct: 30  }
  medium: { max_pct: 50  }
  large:  { max_pct: 75  }
  xl:     { max_pct: 100 }
"""

MINIMAL_MD = """\
# Introduction

Write your introduction here.

## Background

Add background context here.

# Main Section

Your main content goes here.

## Subsection

Detail for this subsection.

# Conclusion

Summarise your findings here.
"""

FULL_TEMPLATE_MD = (
    "# Introduction\n"
    "\n"
    "Replace this with your opening section. "
    "This template demonstrates every feature supported by the builder — "
    "work through it top to bottom, then delete what you don't need.\n"
    "\n"
    "This document was prepared by **{{document.author}}** (version {{document.version}}). "
    "Property placeholders like `{{document.author}}` are replaced at build time. "
    "The source file always keeps the placeholder text.\n"
    "\n"
    "Inline formatting: **bold**, *italic*, ***bold italic***, ~~strikethrough~~, "
    "`inline code`, and a hard line break is made with a backslash at the end of a line:\\\n"
    "This sentence starts on a new line within the same paragraph.\n"
    "\n"
    "A blank line between paragraphs creates a new paragraph — like this one.\n"
    "\n"
    "\n"
    "# Headings\n"
    "\n"
    "Headings are numbered automatically from H1 to H6. "
    "Add `{.nonumber}` to suppress the number on a specific heading. "
    "Add `{.notoc}` to exclude a heading from the Table of Contents.\n"
    "\n"
    "## Heading Level 2\n"
    "\n"
    "### Heading Level 3\n"
    "\n"
    "#### Heading Level 4\n"
    "\n"
    "##### Heading Level 5\n"
    "\n"
    "###### Heading Level 6\n"
    "\n"
    "### Unnumbered Heading {.nonumber}\n"
    "\n"
    "This H3 has no number. Useful for definitions, notes, and acknowledgements.\n"
    "\n"
    "### Excluded from Contents {.notoc}\n"
    "\n"
    "This heading does not appear in the Table of Contents.\n"
    "\n"
    "\n"
    "# Lists\n"
    "\n"
    "Unordered list:\n"
    "\n"
    "- First bullet item\n"
    "- Second bullet item\n"
    "- **Bold label:** use for definition-style lists\n"
    "\n"
    "Ordered list:\n"
    "\n"
    "1. First step\n"
    "2. Second step\n"
    "3. Third step\n"
    "\n"
    "Nested mixed list:\n"
    "\n"
    "1. **Category A**\n"
    "   - Sub-item A.1\n"
    "   - Sub-item A.2\n"
    "     - Detail A.2.a\n"
    "     - Detail A.2.b\n"
    "2. **Category B**\n"
    "   - Sub-item B.1\n"
    "\n"
    "\n"
    "# Images\n"
    "\n"
    "## Single Image\n"
    "\n"
    "Size classes: `{.xs}` (20%) `{.small}` (30%) `{.medium}` (50%) "
    "`{.large}` (75%) `{.xl}` (100% width). "
    "Alignment: `{.left}` `{.center}` `{.right}`.\n"
    "\n"
    "![Workflow diagram](images/workflow_diagram.png){.large .center}\n"
    "\n"
    "*Figure: End-to-end workflow diagram. {#fig-workflow}*\n"
    "\n"
    "Cross-references are written as `[Figure](#anchor)` and resolve to the correct "
    "number automatically — for example: see [Figure](#fig-workflow).\n"
    "\n"
    "## Side-by-Side Images\n"
    "\n"
    "Use `:::figures` to place two images next to each other:\n"
    "\n"
    ":::figures\n"
    "![System architecture](images/architecture.png){.large}\n"
    "![Comparison chart](images/comparison_chart.png){.large}\n"
    ":::\n"
    "\n"
    "*Figure: System architecture (left) and comparison chart (right). {#fig-arch}*\n"
    "\n"
    "\n"
    "# Tables\n"
    "\n"
    "## Basic Table\n"
    "\n"
    "The first row is always the header. "
    "Column alignment in the separator row: `:---` left, `:---:` centre, `---:` right.\n"
    "\n"
    "| Metric        | Baseline | Current | Change  |\n"
    "|:--------------|:--------:|--------:|:--------|\n"
    "| Response time | 240 ms   | 180 ms  | \u221225%    |\n"
    "| Error rate    | 4.2%     | 1.1%    | \u221274%    |\n"
    "| Throughput    | 1,200/s  | 2,050/s | +71%    |\n"
    "\n"
    "*Table: Performance metrics before and after optimisation. {#tbl-metrics}*\n"
    "\n"
    "Refer to [Table](#tbl-metrics) for the full breakdown.\n"
    "\n"
    "## Custom Column Widths\n"
    "\n"
    "Add `{col-widths=\"\u2026\"}` on the line immediately after a table to set widths:\n"
    "\n"
    "| Component     | Status      | Notes                              |\n"
    "|---------------|-------------|------------------------------------|\n"
    "| API gateway   | Deployed    | Running v2.3, rate limiting active |\n"
    "| Auth service  | In review   | Pending security sign-off          |\n"
    "| Data pipeline | Development | Expected end of current sprint     |\n"
    "\n"
    "{col-widths=\"20%,15%,65%\"}\n"
    "\n"
    "*Table: Component status with custom column widths. {#tbl-status}*\n"
    "\n"
    "## Merged Cells\n"
    "\n"
    "Two merge methods:\n"
    "\n"
    "- `{cs=N}` in a cell = span N columns to the right\n"
    "- `{rs=N}` in a cell = span N rows downward\n"
    "- `<<` in a cell = merge with the cell to the left\n"
    "- `^^` in a cell = merge with the cell above\n"
    "- `{ha=l/c/r}` = horizontal alignment, `{va=t/m/b}` = vertical alignment\n"
    "\n"
    "| Region {cs=2}      | <<          | Q3 Sales |\n"
    "|:-------------------|:------------|:--------:|\n"
    "| **EMEA** {rs=2}    | UK          | 142      |\n"
    "| ^^                 | Germany     | 98       |\n"
    "| **APAC** {rs=2}    | Australia   | 76       |\n"
    "| ^^                 | Japan       | 134      |\n"
    "\n"
    "*Table: Merged-cell example — region spans rows, header spans columns. {#tbl-merged}*\n"
    "\n"
    "\n"
    "# Alerts and Quotes\n"
    "\n"
    "> [!NOTE]\n"
    "> Informational note. Use for supplementary context that is helpful but not critical.\n"
    "\n"
    "> [!TIP]\n"
    "> Practical advice. Use when there is a better or faster way to do something.\n"
    "\n"
    "> [!WARNING]\n"
    "> Important caveat. Use for risks or conditions that could cause unexpected results.\n"
    "\n"
    "> [!CAUTION]\n"
    "> Hard stop. Use for actions that could cause data loss or irreversible changes.\n"
    "\n"
    "Standard blockquote (for citations or pulled quotes):\n"
    "\n"
    "> *\"Write clearly, not cleverly.\"*\n"
    "\n"
    "\n"
    "# Code Blocks\n"
    "\n"
    "Specify the language after the opening fences for syntax labelling:\n"
    "\n"
    "```python\n"
    "def process(data: list, threshold: float = 0.5) -> dict:\n"
    "    filtered = [x for x in data if x > threshold]\n"
    "    return {\"count\": len(filtered), \"mean\": sum(filtered) / len(filtered)}\n"
    "```\n"
    "\n"
    "```bash\n"
    "# Build the document\n"
    "python run.py\n"
    "```\n"
    "\n"
    "\n"
    "# Page and Spacing Controls\n"
    "\n"
    "A horizontal rule is three or more dashes on their own line:\n"
    "\n"
    "---\n"
    "\n"
    "A page break is three dashes on their own line:\n"
    "Vertical space using line units or exact points:\n"
    "\n"
    ":::space{lines=2}\n"
    "\n"
    "Text after two lines of vertical space.\n"
    "\n"
    ":::space{pt=36}\n"
    "\n"
    "Text after 36 points (0.5 inch) of vertical space.\n"
    "\n"
    "\n"
    "# Cross-References Summary\n"
    "\n"
    "All captioned figures and tables can be referenced by their anchor:\n"
    "\n"
    "- Workflow diagram: [Figure](#fig-workflow)\n"
    "- Architecture comparison: [Figure](#fig-arch)\n"
    "- Performance metrics: [Table](#tbl-metrics)\n"
    "- Component status: [Table](#tbl-status)\n"
    "- Merged cell example: [Table](#tbl-merged)\n"
    "\n"
    "Caption anchors are defined with `{#name}` at the end of the caption line. "
    "The number is assigned automatically — never write it manually.\n"
    "\n"
    "\n"
    "# Properties\n"
    "\n"
    "Properties from `properties.yaml` and `document-info.yaml` are substituted "
    "at build time using double-brace syntax:\n"
    "\n"
    "| Syntax                     | Source               | Resolves to               |\n"
    "|:---------------------------|:--------------------:|:--------------------------|\n"
    "| `{ {document.title} }`     | document-info.yaml   | {{document.title}}        |\n"
    "| `{ {document.author} }`    | document-info.yaml   | {{document.author}}       |\n"
    "| `{ {document.version} }`   | document-info.yaml   | {{document.version}}      |\n"
    "| `{ {document.date} }`      | document-info.yaml   | {{document.date}}         |\n"
    "| `{ {document.classification} }` | document-info.yaml | {{document.classification}} |\n"
    "| `{ {project.name} }`       | properties.yaml      | Acme Corp                 |\n"
    "\n"
    "Add custom properties to `properties.yaml` and reference them as `{{section.key}}`.\n"
    "\n"
    "\n"
    "# Header Image\n"
    "\n"
    "A logo or image in the top-right corner of every page header is configured "
    "in `config.yaml`, not in markdown:\n"
    "\n"
    "```yaml\n"
    "header:\n"
    "  image: images/logo.png\n"
    "  image_height_cm: 1.0\n"
    "```\n"
    "\n"
    "The height is fixed in centimetres and the width scales proportionally. "
    "Keep `image_height_cm` at or below the height of your header text rows "
    "so the image does not push the header taller than expected.\n"
    "\n"
    "\n"
    "# Appendix\n"
    "\n"
    ":::appendix\n"
    "\n"
    "## Syntax Reference\n"
    "\n"
    "Appendix sections are lettered A, B, C\u2026 automatically. "
    "Place `:::appendix` before the first appendix heading and the rest follows. "
    "Delete this entire section if your document does not need an appendix.\n"
    "\n"
    "| Feature                  | Syntax                                  |\n"
    "|:-------------------------|:----------------------------------------|\n"
    "| Bold                     | `**text**`                              |\n"
    "| Italic                   | `*text*`                                |\n"
    "| Bold italic              | `***text***`                            |\n"
    "| Strikethrough            | `~~text~~`                              |\n"
    "| Inline code              | `` `text` ``                            |\n"
    "| Line break               | Backslash `\\` at end of line           |\n"
    "| New paragraph            | Blank line between text blocks          |\n"

    "| Page break               | `---` on its own line                   |\n"
    "| Vertical space           | `:::space{lines=N}` or `:::space{pt=N}` |\n"
    "| Figure caption           | `*Figure: Description. {#anchor}*`      |\n"
    "| Table caption            | `*Table: Description. {#anchor}*`       |\n"
    "| Cross-reference          | `[Figure](#anchor)` or `[Table](#anchor)`|\n"
    "| Side-by-side images      | `:::figures` \u2026 `:::`                    |\n"
    "| Column widths            | `{col-widths=\"20%,30%,50%\"}` after table |\n"
    "| Merge right (colspan)    | Put `{cs=2}` at end of anchor cell text |\n"
    "| Merge down (rowspan)     | Put `{rs=2}` at end of anchor cell text |\n"
    "| Merge-left marker        | Put `<<` alone in consumed cell         |\n"
    "| Merge-up marker          | Put `^^` alone in consumed cell         |\n"
    "| Suppress heading number  | `{.nonumber}` after heading text        |\n"
    "| Exclude from TOC         | `{.notoc}` after heading text           |\n"
    "| Appendix sections        | `:::appendix` before first appendix H2  |\n"
    "| Property placeholder     | `{{section.key}}`                       |\n"
    "\n"
    "## Config Quick Reference\n"
    "\n"
    "| Setting                    | Where               | What it controls                 |\n"
    "|:---------------------------|:-------------------:|:---------------------------------|\n"
    "| `page.margin_top`          | config.yaml         | Space above body text            |\n"
    "| `page.header_distance`     | config.yaml         | Header distance from paper edge  |\n"
    "| `header.image`             | config.yaml         | Logo path for page header        |\n"
    "| `header.image_height_cm`   | config.yaml         | Logo height (width auto-scales)  |\n"
    "| `numbered_headings`        | config.yaml         | Toggle heading auto-numbering    |\n"
    "| `styles.normal.font_name`  | config.yaml         | Body text font family            |\n"
    "| `image_sizes.medium.max_pct` | config.yaml       | Width of `{.medium}` images      |\n"
)

# ── project creation ───────────────────────────────────────────────────────────

def _get_last_author() -> str:
    """Return the most recently used author name across all projects."""
    import yaml
    for proj in sorted(PROJECTS_DIR.glob("*/input/document-info.yaml"),
                       key=lambda f: f.stat().st_mtime, reverse=True):
        try:
            data = yaml.safe_load(proj.read_text()) or {}
            author = data.get("document", {}).get("author", "")
            if author: return author
        except Exception: pass
    try: return os.environ.get("USER") or os.environ.get("USERNAME") or ""
    except: return ""

def _get_last_classification() -> str:
    import yaml
    for proj in sorted(PROJECTS_DIR.glob("*/input/document-info.yaml"),
                       key=lambda f: f.stat().st_mtime, reverse=True):
        try:
            data = yaml.safe_load(proj.read_text()) or {}
            c = data.get("document", {}).get("classification", "")
            if c: return c
        except Exception: pass
    return "Internal Use Only"

def _create_project_files(proj_dir: Path, title: str, author: str,
                           version: str, classification: str,
                           template: str, docx_source: Path = None):
    """Scaffold a new project folder."""
    input_dir  = proj_dir / "input"
    output_dir = proj_dir / "output"
    images_dir = input_dir / "images"
    received_dir = output_dir / "received"

    for d in [input_dir, images_dir, output_dir, received_dir]:
        d.mkdir(parents=True, exist_ok=True)

    # Core yaml files
    (input_dir / "document-info.yaml").write_text(
        _document_info_yaml(title, author, version, classification), encoding="utf-8")
    (input_dir / "properties.yaml").write_text(PROPERTIES_YAML, encoding="utf-8")
    (input_dir / "config.yaml").write_text(CONFIG_YAML, encoding="utf-8")
    (input_dir / "00-frontpage.md").write_text(FRONTPAGE_MD, encoding="utf-8")
    (output_dir / "received" / "README.txt").write_text(
        "Drop reviewer's edited .docx files here, then run: python run.py\n")

    if docx_source:
        # Bootstrap from Word file — convert it
        _bootstrap_from_docx(docx_source, input_dir)
    elif template == "full":
        (input_dir / "content.md").write_text(FULL_TEMPLATE_MD, encoding="utf-8")
        # Copy showcase images so the template builds without broken image errors
        showcase_images = PROJECTS_DIR / "md-to-docx-showcase" / "input" / "images"
        if showcase_images.exists():
            for img in showcase_images.glob("*"):
                shutil.copy(img, images_dir / img.name)
    elif template == "minimal":
        (input_dir / "content.md").write_text(MINIMAL_MD, encoding="utf-8")
    else:
        (input_dir / "content.md").write_text(
            f"# {title}\n\nStart writing here.\n", encoding="utf-8")

def _bootstrap_from_docx(docx_path: Path, input_dir: Path):
    """Convert a Word file into the project's input/ folder."""
    from lib.extract import extract_body_sections, write_imported
    from lib.sync import _load_docx_size_classes
    from docx import Document

    _print("  Converting Word file to markdown…")

    sc, cw = _load_docx_size_classes(input_dir)
    doc    = Document(str(docx_path))

    with tempfile.TemporaryDirectory() as tmp:
        tmp_images = Path(tmp) / "images"
        sections, _ = extract_body_sections(doc, tmp_images, sc, cw)

        # Write directly into input_dir
        images_final = input_dir / "images"
        images_final.mkdir(exist_ok=True)
        if tmp_images.exists():
            for img in tmp_images.glob("*"):
                shutil.copy(img, images_final / img.name)

    # Write sections as content files
    file_index = 1
    for level, heading, lines in sections:
        if level == 0 and not heading: continue
        if heading.lower() == "table of contents": continue
        slug     = re.sub(r'[^\w\s-]', '', heading.lower())
        slug     = re.sub(r'[\s_]+', '-', slug).strip('-') or "section"
        fname    = f"{file_index:02d}-{slug}.md"
        body     = "\n".join(lines).strip()
        (input_dir / fname).write_text(
            f"# {heading}\n\n{body}\n", encoding="utf-8")
        file_index += 1




ARCHIVE_DIR = PROJECTS_DIR / ".archive"


def _archived_projects() -> list[Path]:
    """Return list of archived project dirs, sorted by mtime."""
    if not ARCHIVE_DIR.exists():
        return []
    dirs = [d for d in ARCHIVE_DIR.iterdir()
            if d.is_dir() and (d / "input").exists()]
    dirs.sort(key=lambda d: -_project_mtime(d))
    return dirs


def action_change_defaults(projects: list) -> None:
    """Let the user update the default config used for all new projects."""
    _clear()
    _rule("Change Default Config")
    _print("")
    _print("New projects are created with the default config.yaml baked into run.py.")
    _print("You can update it in four ways:\n")

    if HAS_RICH:
        console.print("  [cyan][1][/cyan]  Edit directly — opens the current default in your editor")
        console.print("  [cyan][2][/cyan]  Copy from project — adopt an existing project's config wholesale")
        console.print("  [cyan][3][/cyan]  Toggle fields from project — pick individual fields to copy over")
        console.print("  [cyan][4][/cyan]  Edit all fields — browse and edit every setting category by category")
        console.print("  [cyan][5][/cyan]  Cancel")
    else:
        print("  [1] Edit directly")
        print("  [2] Copy from an existing project")
        print("  [3] Toggle fields from a project")
        print("  [4] Edit all fields — browse and edit every setting")
        print("  [5] Cancel")

    _print("")
    try:    choice = input("  Choice: ").strip()
    except: return

    if choice == "1":
        _action_edit_default_directly()
    elif choice == "2":
        _action_copy_default_from_project(projects)
    elif choice == "3":
        _action_toggle_fields_from_project(projects)
    elif choice == "4":
        _action_edit_all_fields()


def _action_edit_default_directly():
    """Write the current CONFIG_YAML to a temp file, open in editor, read back."""
    import tempfile, subprocess as _sp

    _print("\n  Writing current default config to a temporary file...")

    with tempfile.NamedTemporaryFile(suffix=".yaml", mode="w",
                                     delete=False, encoding="utf-8") as f:
        f.write(CONFIG_YAML)
        tmp_path = Path(f.name)

    # Try VS Code first, then $EDITOR, then platform default
    editor = os.environ.get("EDITOR", "")
    opened = False
    for cmd in (["code", "--wait", str(tmp_path)],
                [editor, str(tmp_path)] if editor else None,
                None):
        if cmd is None:
            _open_path(tmp_path)
            _print(f"\n  Opened in default app: {tmp_path}")
            opened = True; break
        try:
            _sp.run(cmd); opened = True; break
        except (FileNotFoundError, TypeError):
            continue

    if not opened:
        _print(f"  Could not open editor. Edit manually:\n  {tmp_path}")

    _print("")
    if not _confirm("  Apply this file as the new default config?"):
        tmp_path.unlink(missing_ok=True)
        _print("[dim]  Cancelled.[/dim]" if HAS_RICH else "  Cancelled.")
        _pause(); return

    try:
        import yaml as _yaml
        new_cfg = _yaml.safe_load(tmp_path.read_text(encoding="utf-8"))
        assert isinstance(new_cfg, dict), "Not a valid YAML mapping"
    except Exception as e:
        _print(f"[red]  Invalid YAML: {e}[/red]" if HAS_RICH else f"  Invalid YAML: {e}")
        tmp_path.unlink(missing_ok=True)
        _pause(); return

    _write_default_config(tmp_path.read_text(encoding="utf-8"))
    tmp_path.unlink(missing_ok=True)
    _print("[green]\n  ✓ Default config updated.[/green]" if HAS_RICH
           else "\n  ✓ Default config updated.")
    _pause()


def _action_copy_default_from_project(projects: list):
    """Copy an existing project's config.yaml into run.py as the new default."""
    if not projects:
        _print("[yellow]  No projects available to copy from.[/yellow]")
        _pause(); return

    _clear()
    _rule("Copy Default from Project")
    _print("")
    _print("  Select a project whose config.yaml will become the new default:\n")

    items = []
    for proj in projects:
        info  = _project_display(proj)
        title = info["title"]
        if len(title) > 28: title = title[:27] + "\u2026"
        version = (f"v{info['version']}" if info["version"] else "")[:7]
        author  = (info["author"]          if info["author"]  else "")[:20]
        age     = _file_age(proj / "input")[:10]
        items.append((title, f"{version:<8}{author:<21}{age}"))

    result = _menu(items, title="Copy config from which project?",
                   extras=[("c", "Cancel")])

    if not isinstance(result, int) or result >= len(projects):
        return

    proj     = projects[result]
    cfg_path = proj / "input" / "config.yaml"
    if not cfg_path.exists():
        _print("[red]  config.yaml not found in that project.[/red]")
        _pause(); return

    info = _project_display(proj)
    _print(f"\n  Copy config from [cyan]{info['title']}[/cyan]?")
    _print(f"  [dim]{cfg_path}[/dim]")
    if not _confirm("  Confirm"):
        return

    new_yaml = cfg_path.read_text(encoding="utf-8")
    try:
        import yaml as _yaml
        _yaml.safe_load(new_yaml)
    except Exception as e:
        _print(f"[red]  Invalid YAML: {e}[/red]")
        _pause(); return

    _write_default_config(new_yaml)
    _print(f"[green]\n  ✓ Default config copied from {info['title']}.[/green]"
           if HAS_RICH else f"\n  ✓ Default config copied from {info['title']}.")
    _print("  [dim]New projects created from now on will use this config.[/dim]")
    _pause()



# ── Field descriptions ────────────────────────────────────────────────────────


# ── Config field metadata: (category, description) ────────────────────────────
# Used by the Toggle Fields feature to organise and explain every config key.

_FIELD_META = {
    # ── Page layout ───────────────────────────────────────────────────────────
    "page.size":
        ("Page layout", "Paper size for printing. A4 (210×297mm) is standard in Europe; Letter (216×279mm) in North America."),
    "page.orientation":
        ("Page layout", "Page orientation. 'portrait' is taller than wide (default for reports); 'landscape' is wider than tall."),
    "page.margin_top":
        ("Page layout", "Distance from the top of the paper to where body text begins. Increasing this adds breathing room between the header line and the first paragraph."),
    "page.margin_bottom":
        ("Page layout", "Distance from the bottom of the paper to where body text ends. Increasing this adds breathing room between the last paragraph and the footer line."),
    "page.margin_left":
        ("Page layout", "Distance from the left paper edge to body text. Also determines the usable content width — wider left margin means narrower text column."),
    "page.margin_right":
        ("Page layout", "Distance from the right paper edge to body text. Wider right margins suit documents that will be annotated or ring-bound."),
    "page.header_distance":
        ("Page layout", "Distance from the top of the paper to the header content. Must be less than margin_top, otherwise the header will print on top of the body text."),
    "page.footer_distance":
        ("Page layout", "Distance from the bottom of the paper to the footer content. Must be less than margin_bottom."),

    # ── Header / footer ───────────────────────────────────────────────────────
    "header.left":
        ("Header / footer", "Content in the left zone of the page header. Use {title}, {author}, {date}, {version}, or {classification}. A YAML list gives multiple lines stacked vertically."),
    "header.center":
        ("Header / footer", "Content in the centre zone of the page header. Leave empty for no centre content."),
    "header.right":
        ("Header / footer", "Content in the right zone of the page header. {date} is a common choice here."),
    "footer.left":
        ("Header / footer", "Content in the left zone of the page footer. Leave empty for nothing."),
    "footer.center":
        ("Header / footer", "Content in the centre zone of the page footer. 'Page {page} of {total}' is the standard pagination format."),
    "footer.right":
        ("Header / footer", "Content in the right zone of the page footer."),

    # ── Separator lines ───────────────────────────────────────────────────────
    "header_line.show":
        ("Separator lines", "Whether to draw a thin horizontal line below the header. Visually separates the header from the first paragraph of body text."),
    "header_line.color":
        ("Separator lines", "Colour of the header separator line as a 6-character hex code (without #). AAAAAA is a neutral medium grey."),
    "header_line.width":
        ("Separator lines", "Thickness of the header separator line in half-points. 6 = 0.75pt (thin/subtle). 12 = 1.5pt (moderate). 24 = 3pt (bold/prominent)."),
    "footer_line.show":
        ("Separator lines", "Whether to draw a thin horizontal line above the footer. Visually separates the footer from the last paragraph of body text."),
    "footer_line.color":
        ("Separator lines", "Colour of the footer separator line as a 6-character hex code."),
    "footer_line.width":
        ("Separator lines", "Thickness of the footer separator line in half-points."),

    # ── Headings ──────────────────────────────────────────────────────────────
    "numbered_headings":
        ("Headings", "Whether all headings are automatically numbered (1. / 1.1. / 1.1.1. etc.). Set to false for documents that don't use section numbering, such as letters, memos, or narrative reports."),
    "styles.heading_1.font_name":       ("Headings", "Font family for H1 — the top-level section heading. Must be installed on the machine opening the document."),
    "styles.heading_1.font_size_pt":    ("Headings", "H1 font size in points. 22pt is large and clearly marks major sections. 18pt is more conservative."),
    "styles.heading_1.bold":            ("Headings", "Whether H1 headings are rendered in bold."),
    "styles.heading_1.color":           ("Headings", "H1 text colour as a hex code. 1F3864 is dark navy (professional). 000000 is plain black."),
    "styles.heading_1.space_before_pt": ("Headings", "Vertical space inserted above each H1 heading in points. More space emphasises the start of a new major section."),
    "styles.heading_1.space_after_pt":  ("Headings", "Vertical space inserted below each H1 heading in points, between the heading and the first body paragraph."),
    "styles.heading_2.font_name":       ("Headings", "Font family for H2 sub-section headings."),
    "styles.heading_2.font_size_pt":    ("Headings", "H2 font size in points. Should be noticeably smaller than H1 to establish clear hierarchy."),
    "styles.heading_2.bold":            ("Headings", "Whether H2 headings are rendered in bold."),
    "styles.heading_2.color":           ("Headings", "H2 text colour as a hex code."),
    "styles.heading_2.space_before_pt": ("Headings", "Vertical space above H2 headings in points."),
    "styles.heading_2.space_after_pt":  ("Headings", "Vertical space below H2 headings in points."),
    "styles.heading_3.font_name":       ("Headings", "Font family for H3 headings."),
    "styles.heading_3.font_size_pt":    ("Headings", "H3 font size in points."),
    "styles.heading_3.bold":            ("Headings", "Whether H3 headings are bold."),
    "styles.heading_3.color":           ("Headings", "H3 colour as hex."),
    "styles.heading_3.space_before_pt": ("Headings", "Space above H3 headings in points."),
    "styles.heading_3.space_after_pt":  ("Headings", "Space below H3 headings in points."),
    "styles.heading_4.font_name":       ("Headings", "Font family for H4 headings."),
    "styles.heading_4.font_size_pt":    ("Headings", "H4 font size in points."),
    "styles.heading_4.bold":            ("Headings", "Whether H4 headings are bold."),
    "styles.heading_4.color":           ("Headings", "H4 colour as hex."),
    "styles.heading_4.space_before_pt": ("Headings", "Space above H4 headings in points."),
    "styles.heading_4.space_after_pt":  ("Headings", "Space below H4 headings in points."),
    "styles.heading_5.font_name":       ("Headings", "Font family for H5 headings."),
    "styles.heading_5.font_size_pt":    ("Headings", "H5 font size in points."),
    "styles.heading_5.bold":            ("Headings", "Whether H5 headings are bold."),
    "styles.heading_5.color":           ("Headings", "H5 colour as hex."),
    "styles.heading_5.space_before_pt": ("Headings", "Space above H5 headings in points."),
    "styles.heading_5.space_after_pt":  ("Headings", "Space below H5 headings in points."),
    "styles.heading_6.font_name":       ("Headings", "Font family for H6 — the smallest heading level."),
    "styles.heading_6.font_size_pt":    ("Headings", "H6 font size in points. Often the same size as body text since it is the least prominent heading."),
    "styles.heading_6.bold":            ("Headings", "Whether H6 headings are bold. Often false at this level to distinguish them from H5."),
    "styles.heading_6.color":           ("Headings", "H6 colour as hex."),
    "styles.heading_6.space_before_pt": ("Headings", "Space above H6 headings in points."),
    "styles.heading_6.space_after_pt":  ("Headings", "Space below H6 headings in points."),

    # ── Body text ─────────────────────────────────────────────────────────────
    "styles.normal.font_name":
        ("Body text", "Font family for all body text paragraphs. This is the most impactful font setting — it affects the appearance of the vast majority of the document."),
    "styles.normal.font_size_pt":
        ("Body text", "Body text size in points. 11pt is the professional standard. 10pt is more compact and fits more on a page. 12pt is more open and easier to read at a distance."),
    "styles.normal.space_after_pt":
        ("Body text", "Vertical space added after every body paragraph in points. This single setting controls the overall density of the document. 6pt is tight, 10pt is airy."),

    # ── Code blocks ───────────────────────────────────────────────────────────
    "styles.code.font_name":
        ("Code blocks", "Font family for code blocks. Must be a monospace font so characters align in columns. Common choices: Courier New, Consolas, Menlo, JetBrains Mono."),
    "styles.code.font_size_pt":
        ("Code blocks", "Code block font size in points. Typically set 1-2pt smaller than body text so that longer lines fit without wrapping."),
    "styles.code.background":
        ("Code blocks", "Background fill colour for code blocks as a hex code. F0F0F0 is a subtle light grey. FFFFFF removes the background entirely."),
    "styles.code.border_color":
        ("Code blocks", "Colour of the thin horizontal border lines drawn above and below each code block."),
    "styles.code.left_indent_in":
        ("Code blocks", "How far code blocks are inset from the left body margin in inches. Creates visual separation from surrounding text."),
    "styles.code.right_indent_in":
        ("Code blocks", "How far code blocks are inset from the right body margin in inches."),
    "styles.code.space_before_pt":
        ("Code blocks", "Vertical space inserted above each code block in points."),
    "styles.code.space_after_pt":
        ("Code blocks", "Vertical space inserted below each code block in points."),

    # ── Block quotes ──────────────────────────────────────────────────────────
    "styles.block_quote.font_italic":
        ("Block quotes", "Whether block quote text is rendered in italic. Standard typographic convention is to italicise quoted or pulled material."),
    "styles.block_quote.color":
        ("Block quotes", "Text colour for block quotes as a hex code. 444444 is a dark grey, slightly lighter than black body text — it subtly de-emphasises the quote."),
    "styles.block_quote.bar_color":
        ("Block quotes", "Colour of the vertical accent bar drawn on the left side of block quotes. This is the most visually defining element of the quote style."),
    "styles.block_quote.left_indent_in":
        ("Block quotes", "How far block quotes are inset from the left body margin in inches."),
    "styles.block_quote.right_indent_in":
        ("Block quotes", "How far block quotes are inset from the right body margin in inches."),
    "styles.block_quote.space_before_pt":
        ("Block quotes", "Vertical space above block quotes in points."),
    "styles.block_quote.space_after_pt":
        ("Block quotes", "Vertical space below block quotes in points."),

    # ── Captions ──────────────────────────────────────────────────────────────
    "styles.caption.font_size_pt":
        ("Captions", "Font size for figure and table captions in points. Typically 1-2pt smaller than body text to visually subordinate them to the content they describe."),
    "styles.caption.color":
        ("Captions", "Text colour for captions as a hex code. 555555 is medium grey — readable but clearly subordinate to body text."),
    "styles.caption.italic":
        ("Captions", "Whether caption text is rendered in italic. Standard convention for distinguishing captions from body paragraphs."),
    "styles.caption.space_before_pt":
        ("Captions", "Vertical space above captions in points. A small gap separates the caption from the figure or table above it."),
    "styles.caption.space_after_pt":
        ("Captions", "Vertical space below captions in points. This gap separates the caption from the text that follows the figure."),

    # ── Table styling ─────────────────────────────────────────────────────────
    "styles.table_header.background":
        ("Table styling", "Background fill colour for the header row of all data tables as a hex code. 1F3864 is dark navy. The header row is automatically bolded and set to white text."),
    "styles.table_header.font_color":
        ("Table styling", "Text colour for table header cells as a hex code. FFFFFF is white, which contrasts well against dark header backgrounds."),
    "styles.table_header.bold":
        ("Table styling", "Whether table header row text is rendered in bold."),
    "styles.table_rows.odd_background":
        ("Table styling", "Background fill for odd-numbered body rows (rows 1, 3, 5…) as a hex code. F7F7F7 is a very subtle light grey — creates zebra striping without being distracting."),
    "styles.table_rows.even_background":
        ("Table styling", "Background fill for even-numbered body rows (rows 2, 4, 6…) as a hex code. FFFFFF is white."),

    # ── Alert boxes ───────────────────────────────────────────────────────────
    "styles.alerts.note_color":
        ("Alert boxes", "Border and label colour for [!NOTE] alert boxes as a hex code. Blue conveys a neutral informational tone."),
    "styles.alerts.tip_color":
        ("Alert boxes", "Border and label colour for [!TIP] alert boxes as a hex code. Green conveys a positive or helpful tone."),
    "styles.alerts.warning_color":
        ("Alert boxes", "Border and label colour for [!WARNING] alert boxes as a hex code. Orange conveys a cautionary tone."),
    "styles.alerts.caution_color":
        ("Alert boxes", "Border and label colour for [!CAUTION] alert boxes as a hex code. Red conveys danger or a hard stop."),
    "styles.alerts.background":
        ("Alert boxes", "Background fill for all alert box types as a hex code. F5F5F5 is light grey. FFFFFF removes the fill entirely."),

    # ── Cover page ────────────────────────────────────────────────────────────
    "styles.cover_title.font_name":
        ("Cover page", "Font family for the document title on the cover page."),
    "styles.cover_title.font_size_pt":
        ("Cover page", "Font size for the cover page title in points. This is the most prominent text in the document — typically 20-28pt."),
    "styles.cover_title.color":
        ("Cover page", "Colour of the cover page title as a hex code."),
    "styles.cover_title.bold":
        ("Cover page", "Whether the cover page title is bold."),
    "styles.cover_title.space_before_pt":
        ("Cover page", "Vertical space above the cover title in points. Pushes the title down from the top of the page."),
    "styles.cover_title.space_after_pt":
        ("Cover page", "Vertical space below the cover title in points."),
    "styles.cover_subtitle.font_name":
        ("Cover page", "Font family for the subtitle on the cover page."),
    "styles.cover_subtitle.font_size_pt":
        ("Cover page", "Font size for the cover page subtitle in points."),
    "styles.cover_subtitle.color":
        ("Cover page", "Colour of the cover page subtitle as a hex code."),
    "styles.cover_subtitle.bold":
        ("Cover page", "Whether the cover subtitle is bold."),
    "styles.cover_subtitle.space_before_pt":
        ("Cover page", "Vertical space above the cover subtitle in points."),
    "styles.cover_subtitle.space_after_pt":
        ("Cover page", "Vertical space below the cover subtitle in points."),
    "styles.cover_body.font_name":
        ("Cover page", "Font family for supporting text on the cover page (document type, version, classification)."),
    "styles.cover_body.font_size_pt":
        ("Cover page", "Font size for cover page supporting text in points."),
    "styles.cover_body.color":
        ("Cover page", "Colour of cover page supporting text as a hex code."),
    "styles.cover_body.bold":
        ("Cover page", "Whether cover page supporting text is bold."),
    "styles.cover_body.space_before_pt":
        ("Cover page", "Vertical space above cover page supporting text in points."),
    "styles.cover_body.space_after_pt":
        ("Cover page", "Vertical space below cover page supporting text in points."),

    # ── Image sizes ───────────────────────────────────────────────────────────
    "image_sizes.xs.max_pct":
        ("Image sizes", "{.xs} image width as a percentage of usable content width. Default 20% — suitable for small icons or thumbnails shown inline."),
    "image_sizes.small.max_pct":
        ("Image sizes", "{.small} image width as a percentage of content width. Default 30% — a small image that sits alongside text."),
    "image_sizes.medium.max_pct":
        ("Image sizes", "{.medium} image width as a percentage of content width. Default 50% — half page width, the most common choice for diagrams."),
    "image_sizes.large.max_pct":
        ("Image sizes", "{.large} image width as a percentage of content width. Default 75% — a prominent figure that dominates its section."),
    "image_sizes.xl.max_pct":
        ("Image sizes", "{.xl} image width as a percentage of content width. Default 100% — full content width, edge to edge."),
}

# Ordered list of category names for display
_CATEGORY_ORDER = [
    "Page layout",
    "Header / footer",
    "Separator lines",
    "Headings",
    "Body text",
    "Code blocks",
    "Block quotes",
    "Captions",
    "Table styling",
    "Alert boxes",
    "Cover page",
    "Image sizes",
]


def _flatten_yaml(d: dict, prefix: str = "") -> list:
    """Flatten a nested dict into a list of (dotted_key, value) pairs."""
    items = []
    for k, v in d.items():
        full = f"{prefix}.{k}" if prefix else k
        if isinstance(v, dict):
            items.extend(_flatten_yaml(v, full))
        else:
            items.append((full, v))
    return items


def _set_nested(d: dict, dotted_key: str, value) -> dict:
    """Return a deep copy of d with the value at dotted_key replaced."""
    import copy
    result = copy.deepcopy(d)
    keys   = dotted_key.split(".")
    node   = result
    for k in keys[:-1]:
        node = node.setdefault(k, {})
    node[keys[-1]] = value
    return result


def _action_toggle_fields_from_project(projects: list):
    """Pick a project then walk through differing config fields category by category."""
    if not projects:
        _print("[yellow]  No projects available.[/yellow]")
        _pause(); return

    # ── Step 1: pick a project ────────────────────────────────────────────────
    _clear()
    _rule("Toggle Fields — Choose Source Project")
    _print("")
    _print("  Select a project to compare against the current default:\n")

    items = []
    for proj in projects:
        info  = _project_display(proj)
        title = info["title"]
        if len(title) > 28: title = title[:27] + "\u2026"
        version = (f"v{info['version']}" if info["version"] else "")[:7]
        author  = (info["author"]         if info["author"]  else "")[:20]
        age     = _file_age(proj / "input")[:10]
        items.append((title, f"{version:<8}{author:<21}{age}"))

    result = _menu(items, title="Compare against which project?",
                   extras=[("c", "Cancel")])
    if not isinstance(result, int) or result >= len(projects):
        return

    proj     = projects[result]
    cfg_path = proj / "input" / "config.yaml"
    if not cfg_path.exists():
        _print("[red]  config.yaml not found in that project.[/red]")
        _pause(); return

    import yaml as _yaml, re as _re

    try:
        proj_cfg = _yaml.safe_load(cfg_path.read_text(encoding="utf-8"))
    except Exception as e:
        _print(f"[red]  Could not parse config.yaml: {e}[/red]")
        _pause(); return

    # Load current default
    src      = open(Path(__file__).resolve(), encoding="utf-8").read()
    m        = _re.search(r'CONFIG_YAML\s*=\s*"""(.*?)"""', src, _re.DOTALL)
    curr_cfg = _yaml.safe_load(m.group(1)) if m else {}

    curr_flat = dict(_flatten_yaml(curr_cfg))
    proj_flat = dict(_flatten_yaml(proj_cfg))

    # Build diff: only fields in both configs that have different values
    diff = {
        k: {"curr": curr_flat[k], "proj": proj_flat[k], "chosen": None}
        for k in curr_flat
        if k in proj_flat and str(curr_flat[k]) != str(proj_flat[k])
    }

    if not diff:
        _clear()
        _rule("Toggle Fields")
        _print("\n  [green]The project config is identical to the current default. Nothing to toggle.[/green]\n"
               if HAS_RICH else "\n  Configs are identical — nothing to toggle.\n")
        _pause(); return

    # Group diffs by category
    categorised = {}
    for k in _CATEGORY_ORDER:
        fields = [f for f in diff if _FIELD_META.get(f, ("",))[0] == k]
        if fields:
            categorised[k] = fields

    info = _project_display(proj)
    _toggle_categories(diff, categorised, curr_cfg, info["title"])


def _toggle_categories(diff: dict, categorised: dict, curr_cfg: dict, proj_name: str):
    """Main loop: show categories, let user pick one, walk its fields."""
    import yaml as _yaml

    while True:
        _clear()
        _rule("Toggle Fields")
        _print(f"\n  Source project: [cyan]{proj_name}[/cyan]\n" if HAS_RICH
               else f"\n  Source project: {proj_name}\n")

        # Build category summary
        cat_names = list(categorised.keys())
        items     = []
        for cat in cat_names:
            fields    = categorised[cat]
            done      = sum(1 for f in fields if diff[f]["chosen"] is not None)
            total     = len(fields)
            status    = f"{done}/{total} reviewed"
            items.append((cat, status))

        # Count overall progress
        total_fields  = sum(len(v) for v in categorised.values())
        decided       = sum(1 for d in diff.values() if d["chosen"] is not None)
        adopted       = sum(1 for d in diff.values() if d["chosen"] == "adopt")
        edited        = sum(1 for d in diff.values() if d["chosen"] == "edit")

        _print(f"  {decided}/{total_fields} fields reviewed  |  "
               f"{adopted + edited} change(s) queued\n")

        result = _menu(items,
                       title="Select a category to review its fields",
                       extras=[("a", "Apply all queued changes"),
                                ("r", "Reset all decisions"),
                                ("q", "Cancel — discard everything")])

        if result == "q":
            _print("\n[dim]  Cancelled — default config unchanged.[/dim]\n"
                   if HAS_RICH else "\n  Cancelled.\n")
            _pause(); return

        if result == "r":
            for d in diff.values():
                d["chosen"] = None
                d.pop("new_val", None)
            continue

        if result == "a":
            changes = [(k, d) for k, d in diff.items() if d["chosen"] in ("adopt", "edit")]
            if not changes:
                _print("\n[yellow]  No changes queued yet.[/yellow]\n"
                       if HAS_RICH else "\n  No changes queued.\n")
                _pause(); continue
            _apply_changes(changes, curr_cfg)
            return

        if isinstance(result, int) and result < len(cat_names):
            cat    = cat_names[result]
            fields = categorised[cat]
            _review_category(cat, fields, diff)


def _review_category(cat: str, fields: list, diff: dict):
    """Walk through fields in one category, one at a time."""
    import shutil as _sh

    i = 0
    while i < len(fields):
        key  = fields[i]
        d    = diff[key]
        meta = _FIELD_META.get(key, ("", "No description available."))
        desc = meta[1]

        curr_val = str(d["curr"])
        proj_val = str(d["proj"])

        # Current decision label
        if d["chosen"] == "adopt":
            decided = f"[adopt]  will use: {proj_val}"
        elif d["chosen"] == "edit":
            decided = f"[edit]   will use: {d.get('new_val', proj_val)}"
        elif d["chosen"] == "keep":
            decided = f"[keep]   keeping:  {curr_val}"
        else:
            decided = "not yet decided"

        w   = _sh.get_terminal_size((80, 24)).columns
        sep = "\u2500" * min(w - 2, 70)

        _clear()
        _rule(f"{cat}  \u2014  field {i + 1} of {len(fields)}")
        print()
        # Field path
        print(f"  {key}")
        print(f"  {sep}")
        # Description — word-wrap to terminal width
        words    = desc.split()
        line     = "  "
        for word in words:
            if len(line) + len(word) + 1 > w - 2:
                print(line)
                line = "  " + word
            else:
                line = line + (" " if line.strip() else "") + word
        if line.strip():
            print(line)
        print()
        print(f"  Current default : {curr_val}")
        print(f"  Project value   : {proj_val}")
        if d["chosen"] is not None:
            print(f"  Decision        : {decided}")
        print(f"  {sep}")
        print()
        print("  [Enter] Keep current default")
        print("  [A]     Adopt project value")
        print("  [F]     Edit — type your own value")
        print("  [S]     Back to categories")
        print()

        try:
            raw = input("  Choice: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            return

        if raw == "":                 # Enter = keep
            d["chosen"] = "keep"
            d.pop("new_val", None)
            i += 1
        elif raw == "a":
            d["chosen"] = "adopt"
            d.pop("new_val", None)
            i += 1
        elif raw == "f":
            print(f"\n  Current: {curr_val}")
            print(f"  Project: {proj_val}")
            try:
                new_val = input("  New value: ").strip()
            except (EOFError, KeyboardInterrupt):
                continue
            if new_val:
                d["chosen"]  = "edit"
                d["new_val"] = new_val
                i += 1
            else:
                print("  (empty input — unchanged)")
        elif raw == "s":
            return
        else:
            print("  Please press Enter, A, F, or S.")


def _apply_changes(changes: list, curr_cfg: dict):
    """Apply queued changes to curr_cfg and write to run.py."""
    import yaml as _yaml

    cfg = curr_cfg
    for key, d in changes:
        if d["chosen"] == "adopt":
            raw = d["proj"]
        else:  # edit
            raw = d.get("new_val", d["proj"])

        # Attempt to parse the value into the right Python type
        try:
            parsed = _yaml.safe_load(str(raw))
        except Exception:
            parsed = str(raw)

        cfg = _set_nested(cfg, key, parsed)

    new_yaml = _yaml.dump(cfg, default_flow_style=False,
                          allow_unicode=True, sort_keys=False)
    _write_default_config(new_yaml)

    # Summary
    _clear()
    _rule("Changes Applied")
    print()
    adopted = [(k, d) for k, d in changes if d["chosen"] == "adopt"]
    edited  = [(k, d) for k, d in changes if d["chosen"] == "edit"]
    kept    = sum(1 for _, d in changes if d["chosen"] == "keep")

    if adopted:
        print(f"  Adopted from project ({len(adopted)} field(s)):")
        for key, d in adopted:
            print(f"    {key:<45}  {str(d['curr']):<14}  \u2192  {str(d['proj'])}")
    if edited:
        print(f"\n  Custom values set ({len(edited)} field(s)):")
        for key, d in edited:
            print(f"    {key:<45}  {str(d['curr']):<14}  \u2192  {d.get('new_val','')}")
    print()
    print(f"  Default config updated. New projects will use these values.")
    print()
    input("  Press Enter to continue\u2026")



def _action_edit_all_fields():
    """Browse and edit every config field, category by category."""
    import yaml as _yaml, re as _re

    src      = open(Path(__file__).resolve(), encoding="utf-8").read()
    m        = _re.search(r'CONFIG_YAML\s*=\s*"""(.*?)"""', src, _re.DOTALL)
    if not m:
        _print("[red]  Could not read CONFIG_YAML from run.py.[/red]")
        _pause(); return

    curr_cfg  = _yaml.safe_load(m.group(1))
    curr_flat = dict(_flatten_yaml(curr_cfg))

    # pending: key -> new_value (only keys the user explicitly changed)
    pending = {}

    while True:
        _clear()
        _rule("Edit All Config Fields")
        _print("")

        changed = len(pending)
        _print(f"  Browse every setting category by category.\n"
               f"  {changed} change(s) pending.\n")

        items = []
        for cat in _CATEGORY_ORDER:
            fields  = [k for k in curr_flat if _FIELD_META.get(k, ("",))[0] == cat]
            if not fields:
                continue
            edited  = sum(1 for f in fields if f in pending)
            note    = f"{edited} edited" if edited else f"{len(fields)} fields"
            items.append((cat, note))

        result = _menu(items,
                       title="Select a category",
                       extras=[("a", "Apply all changes"),
                                ("r", "Reset all changes"),
                                ("q", "Cancel — discard everything")])

        if result == "q":
            _print("\n[dim]  Cancelled — default config unchanged.[/dim]\n"
                   if HAS_RICH else "\n  Cancelled.\n")
            _pause(); return

        if result == "r":
            pending.clear()
            continue

        if result == "a":
            if not pending:
                _print("\n[yellow]  No changes to apply.[/yellow]\n"
                       if HAS_RICH else "\n  No changes.\n")
                _pause(); continue
            _apply_all_edits(pending, curr_cfg)
            return

        if isinstance(result, int):
            cat_name = _CATEGORY_ORDER[[
                c for c in _CATEGORY_ORDER
                if any(_FIELD_META.get(k, ("",))[0] == c for k in curr_flat)
            ].index(list(items[result][0:1])[0]) if False else result]
            # simpler: extract from items list directly
            cat_name = items[result][0]
            fields   = [k for k in curr_flat if _FIELD_META.get(k, ("",))[0] == cat_name]
            _edit_category_fields(cat_name, fields, curr_flat, pending)


def _edit_category_fields(cat: str, fields: list,
                           curr_flat: dict, pending: dict):
    """Walk through every field in a category, show value, allow editing."""
    import shutil as _sh

    i = 0
    while i < len(fields):
        key  = fields[i]
        meta = _FIELD_META.get(key, ("", "No description available."))
        desc = meta[1]

        # Show the pending value if the user has already edited it, else current
        current_val = str(pending[key]) if key in pending else str(curr_flat[key])
        original    = str(curr_flat[key])
        is_edited   = key in pending

        w   = _sh.get_terminal_size((80, 24)).columns
        sep = "\u2500" * min(w - 2, 70)

        _clear()
        _rule(f"{cat}  \u2014  field {i + 1} of {len(fields)}")
        print()
        print(f"  {key}")
        print(f"  {sep}")

        # Word-wrap description
        words = desc.split()
        line  = "  "
        for word in words:
            if len(line) + len(word) + 1 > w - 2:
                print(line)
                line = "  " + word
            else:
                line = (line + " " + word) if line.strip() else "  " + word
        if line.strip():
            print(line)

        print()
        if is_edited:
            print(f"  Original value  :  {original}")
            print(f"  Current edit    :  {current_val}  ◄")
        else:
            print(f"  Current value   :  {current_val}")
        print()
        print(f"  {sep}")
        print()
        print("  [Enter] Keep / continue to next field")
        print("  [F]     Edit — type a new value")
        if is_edited:
            print("  [D]     Undo — revert to original")
        print("  [S]     Back to category list")
        print()

        try:
            raw = input("  Choice: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            return

        if raw == "":                 # Enter = keep/advance
            i += 1
        elif raw == "f":
            print(f"\n  Current: {current_val}")
            try:
                new_val = input("  New value: ").strip()
            except (EOFError, KeyboardInterrupt):
                continue
            if new_val == "":
                print("  (empty input — unchanged)")
            elif new_val == original:
                pending.pop(key, None)
                print("  (same as original — edit cleared)")
                i += 1
            else:
                pending[key] = new_val
                i += 1
        elif raw == "d" and is_edited:
            pending.pop(key, None)
            # don't advance — let user see the reverted value
        elif raw == "s":
            return
        else:
            print("  Please press Enter, F, or S.")


def _apply_all_edits(pending: dict, curr_cfg: dict):
    """Write pending edits into curr_cfg and save to run.py."""
    import yaml as _yaml

    cfg = curr_cfg
    for key, new_val in pending.items():
        try:
            parsed = _yaml.safe_load(str(new_val))
        except Exception:
            parsed = str(new_val)
        cfg = _set_nested(cfg, key, parsed)

    new_yaml = _yaml.dump(cfg, default_flow_style=False,
                          allow_unicode=True, sort_keys=False)
    _write_default_config(new_yaml)

    _clear()
    _rule("Changes Applied")
    print()
    print(f"  {len(pending)} field(s) updated:\n")
    for key, new_val in pending.items():
        orig = dict(_flatten_yaml(curr_cfg)).get(key, "")
        print(f"    {key:<45}  {str(orig):<16}  \u2192  {new_val}")
    print()
    print("  Default config updated. New projects will use these values.")
    print()
    input("  Press Enter to continue\u2026")


def _write_default_config(new_yaml_text: str):
    """Overwrite the CONFIG_YAML constant in run.py with new_yaml_text."""
    import re as _re
    run_py = Path(__file__).resolve()
    src    = run_py.read_text(encoding="utf-8")

    # Locate CONFIG_YAML by finding its marker and matching close-fence
    tq = chr(34) * 3          # three double-quotes, avoids parser confusion
    marker = "CONFIG_YAML = " + tq
    start  = src.find(marker)
    if start < 0:
        raise RuntimeError("Could not locate CONFIG_YAML in run.py")
    end = src.find(tq, start + len(marker)) + 3

    clean = new_yaml_text.strip("\n")
    new_assignment = marker + "\n" + clean + "\n" + tq
    run_py.write_text(src[:start] + new_assignment + src[end:], encoding="utf-8")


def action_archive_project():
    """Move a project to the archive."""
    projects = _list_projects()
    if not projects:
        _print("[yellow]No active projects to archive.[/yellow]")
        _pause(); return

    _clear()
    _rule("Archive a Project")
    _print("")
    _print("Select a project to archive (it will be hidden from the main list):\n")

    import io as _io
    items = []
    for proj in projects:
        info  = _project_display(proj)
        title = info["title"]
        if len(title) > 28: title = title[:27] + "\u2026"
        version = (f"v{info['version']}" if info["version"] else "")[:7]
        author  = (info["author"]          if info["author"]  else "")[:20]
        age     = _file_age(proj / "input")[:10]
        items.append((title, f"{version:<8}{author:<21}{age}"))

    result = _menu(items, title="Archive which project?",
                   extras=[("c","Cancel")])

    if not isinstance(result, int) or result >= len(projects):
        return

    proj  = projects[result]
    info  = _project_display(proj)

    _print(f"\n  Archive [cyan]{info['title']}[/cyan]?")
    if not _confirm("  Confirm"):
        return

    ARCHIVE_DIR.mkdir(exist_ok=True)
    dst = ARCHIVE_DIR / proj.name
    if dst.exists():
        ts  = datetime.now().strftime("%Y%m%d-%H%M")
        dst = ARCHIVE_DIR / f"{proj.name}-{ts}"

    shutil.move(str(proj), dst)
    _print(f"\n[green]✓ Archived: {info['title']}[/green]")
    _print(f"[dim]  Stored in projects/.archive/{dst.name}[/dim]")
    _pause()


def action_unarchive_project():
    """Restore an archived project back to the active list."""
    archived = _archived_projects()
    if not archived:
        _print("[yellow]No archived projects.[/yellow]")
        _pause(); return

    _clear()
    _rule("Unarchive a Project")
    _print("")
    _print("Select a project to restore:\n")

    items = []
    for proj in archived:
        info  = _project_display(proj)
        title = info["title"]
        if len(title) > 28: title = title[:27] + "\u2026"
        version = (f"v{info['version']}" if info["version"] else "")[:7]
        author  = (info["author"]          if info["author"]  else "")[:20]
        age     = _file_age(proj / "input")[:10]
        items.append((title, f"{version:<8}{author:<21}{age}"))

    result = _menu(items, title="Restore which project?",
                   extras=[("c","Cancel")])

    if not isinstance(result, int) or result >= len(archived):
        return

    proj = archived[result]
    info = _project_display(proj)
    dst  = PROJECTS_DIR / proj.name
    if dst.exists():
        ts  = datetime.now().strftime("%Y%m%d-%H%M")
        dst = PROJECTS_DIR / f"{proj.name}-{ts}"

    shutil.move(str(proj), dst)
    _print(f"\n[green]✓ Restored: {info['title']}[/green]")
    _pause()


# ── main ──────────────────────────────────────────────────────────────────────



def action_new_project():
    """Interactively create a new project."""
    _rule("New Project")
    _print("")

    title = _inp("Project title", "Untitled")
    if not title.strip():
        _pause(); return

    author         = _inp("Author", _get_last_author())
    version        = _inp("Version", "1.0")
    classification = _inp("Classification", _get_last_classification())

    templates = ["minimal", "full"]
    _print("\nTemplate:")
    for i, t in enumerate(templates, 1):
        _print(f"  [{i}] {t}")
    raw = _inp("\nSelect template", "1").strip()
    try:    tmpl_idx = int(raw) - 1
    except: tmpl_idx = 0
    template = templates[tmpl_idx] if 0 <= tmpl_idx < len(templates) else "minimal"

    slug     = _unique_slug(title)
    proj_dir = PROJECTS_DIR / slug
    _create_project_files(proj_dir, title, author, version, classification, template)

    _print(f"\n[green]Created: {proj_dir}[/green]\n" if HAS_RICH
           else f"\n  Created: {proj_dir}\n")
    _pause()

# -- Document actions -----------------------------------------------

def action_build(state: dict):
    _rule("Build — Markdown → Word")
    _print("")

    if not state["md_files"]:
        _print("[red]No markdown files found in input/[/red]")
        _pause(); return

    from lib.build_doc import (load_config, load_document_info, load_all_yaml_files,
                               substitute_properties, collect_files)
    from lib.build.builder import DocumentBuilder

    input_dir  = state["input_dir"]
    output_dir = state["output_dir"]
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "document.docx"

    lock = output_dir / f"~${output_path.name}"
    if lock.exists():
        _print("[red]document.docx is open in Word — close it and try again.[/red]")
        _pause(); return

    config = load_config(input_dir / "config.yaml")
    doc_info, revisions = load_document_info(input_dir / "document-info.yaml")
    if doc_info: config["document"] = doc_info
    config.setdefault("document", {})

    EXCL = {"config.yaml", "document-info.yaml", "revisions.yaml"}
    try:
        props = load_all_yaml_files(input_dir, exclude_files=EXCL)
        for k, v in doc_info.items(): props.setdefault(f"document.{k}", str(v))
    except ValueError as e:
        _print(f"[red]Error: {e}[/red]"); _pause(); return

    builder = DocumentBuilder(config=config, revisions=revisions, source_dir=input_dir)
    builder.setup()
    frontpage, content_files = collect_files(input_dir)

    all_texts = []
    for cf in content_files:
        try:    all_texts.append(substitute_properties(cf.read_text(encoding="utf-8"), props))
        except: all_texts.append("")
    builder.prescan_labels(all_texts)

    word_cover_rel = config.get("frontpage", {}).get("word_cover", "")
    if word_cover_rel:
        word_cover_path = input_dir / word_cover_rel
        if word_cover_path.exists():
            builder.add_word_cover(word_cover_path)
        else:
            _print(f"[yellow]Warning: word_cover not found: {word_cover_path}[/yellow]"
                   if HAS_RICH else f"Warning: word_cover not found: {word_cover_path}")
            if frontpage:
                builder.add_frontpage(
                    substitute_properties(frontpage.read_text(encoding="utf-8"), props),
                    frontpage.parent)
    elif frontpage:
        builder.add_frontpage(
            substitute_properties(frontpage.read_text(encoding="utf-8"), props),
            frontpage.parent)
    builder.add_toc()

    for cf, ct in zip(content_files, all_texts):
        try:    builder.add_content(ct, cf.parent)
        except Exception as e: _print(f"[yellow]Warning: {cf.name}: {e}[/yellow]")

    try:
        builder.save(output_path)
    except PermissionError:
        _print("\n[red]Cannot write document.docx — the file is open in Word.\n"
               "  Close it in Word and try again.[/red]"
               if HAS_RICH else
               "\nCannot save — document.docx is open. Close it in Word first.")
        _pause(); return
    _print(f"\n[green]✓ Built: {output_path}[/green]")
    _pause()



def _add_revision(input_dir: Path, new_version: str,
                   author: str, changes: str) -> bool:
    """Add a revision entry to document-info.yaml and update the version field.

    Returns True if the file was updated.
    """
    import yaml
    di_path = input_dir / "document-info.yaml"
    if not di_path.exists():
        return False
    try:
        data = yaml.safe_load(di_path.read_text(encoding="utf-8")) or {}
    except Exception:
        return False

    # Update version in document block
    doc = data.get("document", {})
    doc["version"] = new_version
    doc["date"]    = datetime.now().strftime("%B %Y")
    data["document"] = doc

    # Prepend new revision entry (newest first)
    new_entry = {
        "version": new_version,
        "date":    datetime.now().strftime("%Y-%m-%d"),
        "author":  author,
        "changes": changes,
    }
    revisions = data.get("revisions", [])
    revisions.insert(0, new_entry)
    data["revisions"] = revisions

    _backup_file(di_path)
    di_path.write_text(
        yaml.dump(data, default_flow_style=False, allow_unicode=True, sort_keys=False),
        encoding="utf-8")
    return True


def action_export(state: dict):
    """Prompt for revision info, rebuild, then export to a chosen location."""
    _rule("Export Word Document")
    _print("")

    if not state["md_files"]:
        _print("[red]No source files found. Nothing to export.[/red]"
               if HAS_RICH else "No source files found.")
        _pause(); return

    from lib.build_doc import load_document_info
    di_path  = state["input_dir"] / "document-info.yaml"
    doc_info, _ = load_document_info(di_path)

    cur_version = doc_info.get("version", "1.0")
    cur_author  = doc_info.get("author",  "")

    # ── Step 1: Revision info ─────────────────────────────────────────────────
    _clear()
    _rule("Export — Revision History")
    _print("")
    _print("  Before exporting, record what changed in this version.\n")
    _print(f"  [dim]Current version: {cur_version}[/dim]\n")

    new_version = _inp("New version number", cur_version)
    author      = _inp("Author",             cur_author)
    _print("")
    _print("  Describe what changed in this version:")
    _print("  [dim](Press Enter to skip)[/dim]")
    try:
        changes = input("  Changes: ").strip()
    except (EOFError, KeyboardInterrupt):
        changes = ""

    # ── Step 2: Update document-info.yaml if version changed ─────────────────
    revision_added = False
    if new_version != cur_version or changes:
        _add_revision(state["input_dir"], new_version, author, changes or "—")
        revision_added = True
        _print(f"\n  [green]✓ Revision {new_version} recorded.[/green]"
               if HAS_RICH else f"\n  ✓ Revision {new_version} recorded.")

    # ── Step 3: Rebuild so the cover page shows updated revision table ────────
    _print("\n  Rebuilding document…")
    from lib.build_doc import (load_config, load_document_info as _ldi,
                               load_all_yaml_files, substitute_properties,
                               collect_files)
    from lib.build.builder import DocumentBuilder

    input_dir  = state["input_dir"]
    output_dir = state["output_dir"]
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / "document.docx"

    lock = output_dir / f"~${output_path.name}"
    if lock.exists():
        _print("[red]document.docx is open in Word — close it and try again.[/red]"
               if HAS_RICH else "document.docx is open. Close it first.")
        _pause(); return

    config = load_config(input_dir / "config.yaml")
    doc_info2, revisions = _ldi(input_dir / "document-info.yaml")
    if doc_info2: config["document"] = doc_info2
    config.setdefault("document", {})

    EXCL = {"config.yaml", "document-info.yaml", "revisions.yaml"}
    try:
        props = load_all_yaml_files(input_dir, exclude_files=EXCL)
        for k, v in doc_info2.items(): props.setdefault(f"document.{k}", str(v))
    except ValueError as e:
        _print(f"[red]Error: {e}[/red]"); _pause(); return

    builder = DocumentBuilder(config=config, revisions=revisions, source_dir=input_dir)
    builder.setup()
    frontpage, content_files = collect_files(input_dir)
    all_texts = []
    for cf in content_files:
        try:    all_texts.append(substitute_properties(cf.read_text(encoding="utf-8"), props))
        except: all_texts.append("")
    builder.prescan_labels(all_texts)
    word_cover_rel = config.get("frontpage", {}).get("word_cover", "")
    if word_cover_rel:
        word_cover_path = input_dir / word_cover_rel
        if word_cover_path.exists():
            builder.add_word_cover(word_cover_path)
        else:
            _print(f"[yellow]Warning: word_cover not found: {word_cover_path}[/yellow]"
                   if HAS_RICH else f"Warning: word_cover not found: {word_cover_path}")
            if frontpage:
                builder.add_frontpage(
                    substitute_properties(frontpage.read_text(encoding="utf-8"), props),
                    frontpage.parent)
    elif frontpage:
        builder.add_frontpage(
            substitute_properties(frontpage.read_text(encoding="utf-8"), props),
            frontpage.parent)
    builder.add_toc()
    for cf, ct in zip(content_files, all_texts):
        try:    builder.add_content(ct, cf.parent)
        except Exception as e: _print(f"[yellow]Warning: {cf.name}: {e}[/yellow]")
    try:
        builder.save(output_path)
    except PermissionError:
        _print("\n[red]Cannot write document.docx — the file is open in Word.\n"
               "  Close it in Word and try again.[/red]"
               if HAS_RICH else
               "\nCannot save — document.docx is open. Close it in Word first.")
        _pause(); return
    _print(f"  [green]✓ Document rebuilt.[/green]" if HAS_RICH else "  ✓ Rebuilt.")

    # ── Step 4: Pick export destination and copy ──────────────────────────────
    _print("\n  Choose where to save the exported file.\n")

    # Build suggested filename from title + version
    slug = re.sub(r"[^\w\s-]", "", doc_info2.get("title", "document").lower())
    slug = re.sub(r"[\s_]+", "-", slug).strip("-") or "document"
    ver  = doc_info2.get("version", new_version)
    suggested = f"{slug}-v{ver}.docx"

    dst = None
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk(); root.withdraw(); root.lift()
        root.attributes("-topmost", True)
        chosen = filedialog.asksaveasfilename(
            title="Export Word document",
            defaultextension=".docx",
            filetypes=[("Word document", "*.docx")],
            initialfile=suggested,
            initialdir=str(Path.home() / "Documents"),
        )
        root.destroy()
        if chosen:
            dst = Path(chosen)
    except Exception:
        pass

    if dst is None:
        # Fallback to typed path
        try:
            typed = input(f"  Save to [{suggested}]: ").strip().strip('"').strip("'")
            dst = Path(typed) if typed else Path.home() / "Documents" / suggested
        except (EOFError, KeyboardInterrupt):
            _print("[dim]Export cancelled.[/dim]")
            _pause(); return

    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(output_path, dst)
    _print(f"\n[green]✓ Exported to: {dst}[/green]"
           if HAS_RICH else f"\n✓ Exported to: {dst}")
    _pause()

def action_sync(state: dict):
    """Build fresh from source, diff against received, show terminal summary.
    Optionally open the full HTML section report.
    """
    _rule("Review Changes")
    _print("")

    if not state["received_files"]:
        _print("[yellow]No file in received/ — nothing to compare against.[/yellow]\n"
               "Use 'Receive file' first to import a reviewer's Word file."
               if HAS_RICH else "No file in received/. Use 'Receive file' first.")
        _pause(); return

    from lib.sync import _find_received_docx
    from lib.section_diff import diff_documents as _section_diff, build_html_report

    input_dir  = state["input_dir"]
    output_dir = state["output_dir"]

    received = _find_received_docx(output_dir)
    if not received:
        _print("[yellow]No file in received/.[/yellow]")
        _pause(); return

    _print(f"  Building from source…\n" if not HAS_RICH
           else "  [dim]Building from source…[/dim]\n")

    import tempfile, shutil as _sh
    tmp_dir  = Path(tempfile.mkdtemp())
    baseline = tmp_dir / "baseline.docx"

    try:
        from lib.build_doc import (load_config, load_document_info,
                                    load_all_yaml_files, substitute_properties,
                                    collect_files)
        from lib.build.builder import DocumentBuilder

        config = load_config(input_dir / "config.yaml")
        doc_info, revisions = load_document_info(input_dir / "document-info.yaml")
        if doc_info: config["document"] = doc_info
        config.setdefault("document", {})
        EXCL = {"config.yaml", "document-info.yaml", "revisions.yaml"}
        props = load_all_yaml_files(input_dir, exclude_files=EXCL)
        for k, v in doc_info.items():
            props.setdefault(f"document.{k}", str(v))

        builder = DocumentBuilder(config=config, revisions=revisions,
                                   source_dir=input_dir)
        builder._verbose = False   # suppress informational prints during silent build
        builder.setup()
        frontpage, content_files = collect_files(input_dir)
        all_texts = []
        for cf in content_files:
            try:    all_texts.append(substitute_properties(
                        cf.read_text(encoding="utf-8"), props))
            except: all_texts.append("")
        builder.prescan_labels(all_texts)
        word_cover_rel = config.get("frontpage", {}).get("word_cover", "")
        if word_cover_rel:
            wcp = input_dir / word_cover_rel
            if wcp.exists():
                builder.add_word_cover(wcp)
            elif frontpage:
                builder.add_frontpage(
                    substitute_properties(frontpage.read_text(encoding="utf-8"), props),
                    frontpage.parent)
        elif frontpage:
            builder.add_frontpage(
                substitute_properties(frontpage.read_text(encoding="utf-8"), props),
                frontpage.parent)
        builder.add_toc()
        for cf, ct in zip(content_files, all_texts):
            try:    builder.add_content(ct, cf.parent)
            except Exception as e:
                _print(f"[yellow]Warning: {cf.name}: {e}[/yellow]")
        builder.save(baseline)
    except Exception as e:
        _print(f"[red]  Build failed: {e}[/red]")
        _sh.rmtree(str(tmp_dir), ignore_errors=True)
        _pause(); return

    try:
        results = _section_diff(baseline, received)
    except Exception as e:
        _print(f"[red]  Diff failed: {e}[/red]")
        _sh.rmtree(str(tmp_dir), ignore_errors=True)
        _pause(); return
    # Note: tmp_dir is NOT cleaned up here — baseline is still needed
    # by build_html_report to render section content. Cleanup happens below.

    # ── Terminal summary ──────────────────────────────────────────────────────
    def _count(rs, counts=None):
        if counts is None:
            counts = {"changed": 0, "added": 0, "removed": 0, "moved": 0}
        for r in rs:
            # Only count a section if it has direct content changes — not just
            # because a child changed. A parent is "changed" only when its own
            # content_hash differs; the child is counted separately.
            has_direct_change = (
                r.status in ("removed", "added", "moved", "moved_changed") or
                (r.status == "changed" and
                 r.baseline is not None and r.received is not None and
                 r.baseline.content_hash != r.received.content_hash)
            )
            if has_direct_change:
                key = r.status.replace("moved_changed", "moved")
                counts[key] = counts.get(key, 0) + 1
            _count(r.children, counts)
        return counts

    counts = _count(results)
    total  = sum(counts.values())

    if total == 0:
        _print("[green]\n  ✓ Documents are identical — no differences found.[/green]\n"
               if HAS_RICH else "\n  Documents are identical.")
        _pause(); return

    _print("")
    if HAS_RICH:
        from rich.table import Table as _T
        from rich import box as _box
        t = _T(box=_box.SIMPLE, show_header=False, padding=(0, 2))
        t.add_column(style="dim")
        t.add_column()
        if counts.get("changed"): t.add_row("🟡 Changed",  str(counts["changed"]))
        if counts.get("added"):   t.add_row("🟢 Added",    str(counts["added"]))
        if counts.get("removed"): t.add_row("🔴 Removed",  str(counts["removed"]))
        if counts.get("moved"):   t.add_row("⬆️  Moved",    str(counts["moved"]))
        console.print(t)
    else:
        for key, label in [("changed","Changed"),("added","Added"),
                           ("removed","Removed"),("moved","Moved")]:
            if counts.get(key):
                print(f"  {label}: {counts[key]}")

    _print("")
    if _confirm("  Open HTML report?"):
        try:
            html        = build_html_report(results, baseline, received,
                                             baseline_label="Your current source",
                                             received_label=received.name)
            report_path = output_dir / "review_report.html"
            report_path.write_text(html, encoding="utf-8")
            _print(f"[green]  ✓ Report written.[/green]\n"
                   f"  [dim]{report_path}[/dim]" if HAS_RICH
                   else f"  ✓ Report: {report_path}")
            import webbrowser
            try:
                webbrowser.open(report_path.as_uri())
            except Exception:
                pass
        except Exception as e:
            _print(f"[red]  Could not write report: {e}[/red]")

    # Clean up temp baseline now that we are done with it
    _sh.rmtree(str(tmp_dir), ignore_errors=True)

    _pause()


def action_receive_file(state: dict):
    _rule("Receive File")
    _print("")
    _print("Select the reviewer's Word file to import.\n")

    received_dir = state["output_dir"] / "received"
    received_dir.mkdir(parents=True, exist_ok=True)

    src = _pick_file("Select reviewer's Word file", initial_dir=Path.home())
    if not src or not src.exists():
        _print("[yellow]No file selected.[/yellow]")
        _pause(); return

    dst = received_dir / src.name
    # If a file with the same name exists, add a timestamp suffix
    if dst.exists():
        ts  = datetime.now().strftime("%Y%m%d-%H%M")
        dst = received_dir / f"{src.stem}_{ts}{src.suffix}"

    shutil.copy(src, dst)
    _print(f"\n[green]✓ Received: {dst.name}[/green]")
    _print(f"  Saved to: output/received/{dst.name}")
    _print("\n[dim]Run Sync Check or Merge to review the changes.[/dim]")
    _pause()


def action_open_document(state: dict):
    if not state["built_docx"].exists():
        _print("[yellow]Document hasn't been built yet. Run Build first.[/yellow]")
        _pause(); return
    _print(f"\nOpening [cyan]{state['built_docx'].name}[/cyan]…")
    _open_path(state["built_docx"])
    _pause()


def action_open_vscode(state: dict):
    _print(f"\nOpening [cyan]projects/{state['proj_dir'].name}/[/cyan] in VS Code…")
    _open_vscode(state["proj_dir"])
    _pause()


def action_edit_info(state: dict):
    """Edit document-info.yaml fields inline in the terminal."""
    _rule("Edit Document Info")
    _print("")

    import yaml
    di_path = state["input_dir"] / "document-info.yaml"
    try:
        data = yaml.safe_load(di_path.read_text(encoding="utf-8")) or {}
    except Exception as e:
        _print(f"[red]Could not read document-info.yaml: {e}[/red]")
        _pause(); return

    doc = data.get("document", {})

    _print("  Edit values (press Enter to keep current):\n")
    fields = [
        ("title",          "Document title"),
        ("subtitle",       "Subtitle"),
        ("author",         "Author"),
        ("date",           "Date"),
        ("version",        "Version"),
        ("classification", "Classification"),
        ("document_type",  "Document type"),
    ]

    changed = False
    for key, label in fields:
        current  = doc.get(key, "")
        new_val  = _inp(label, current)
        if new_val != current:
            doc[key]  = new_val
            changed   = True

    if not changed:
        _print("\n[dim]No changes made.[/dim]")
        _pause(); return

    data["document"] = doc

    # Backup before writing
    _backup_file(di_path)

    di_path.write_text(
        yaml.dump(data, default_flow_style=False, allow_unicode=True, sort_keys=False),
        encoding="utf-8")
    _print(f"\n[green]✓ document-info.yaml updated.[/green]")
    _pause()


def action_edit_properties(state: dict):
    """Edit properties.yaml key/value pairs inline."""
    _rule("Edit Properties")
    _print("")

    import yaml
    prop_path = state["input_dir"] / "properties.yaml"
    try:
        raw  = prop_path.read_text(encoding="utf-8")
        data = yaml.safe_load(raw) or {}
    except Exception as e:
        _print(f"[red]Could not read properties.yaml: {e}[/red]")
        _pause(); return

    # Flatten for display
    flat: dict[str, str] = {}
    def _flatten(d, prefix=""):
        for k, v in d.items():
            full = f"{prefix}{k}" if prefix else k
            if isinstance(v, dict): _flatten(v, f"{full}.")
            elif not isinstance(v, list): flat[full] = str(v)
    _flatten(data)

    if flat:
        _print("  Current properties:\n")
        for key, val in flat.items():
            _print(f"  [dim]{key}[/dim] = [cyan]{val}[/cyan]")
        _print("")

    _print("  [dim]Options:[/dim]")
    _print("  [1] Edit an existing value")
    _print("  [2] Add a new property")
    _print("  [3] Done\n")

    try:    choice = input("  Choice: ").strip()
    except: choice = "3"

    changed = False

    if choice == "1" and flat:
        keys = list(flat.keys())
        for i, k in enumerate(keys, 1):
            _print(f"  [{i}] {k} = {flat[k]}")
        _print("")
        try:    idx = int(input("  Which property? ").strip()) - 1
        except: idx = -1
        if 0 <= idx < len(keys):
            key     = keys[idx]
            new_val = _inp(f"New value for '{key}'", flat[key])
            if new_val != flat[key]:
                flat[key] = new_val
                changed   = True

    elif choice == "2":
        _print("\n  New property key (use dots for nesting, e.g. client.name):")
        try:    key = input("  Key: ").strip()
        except: key = ""
        if key:
            try:    val = input(f"  Value for '{key}': ").strip()
            except: val = ""
            flat[key] = val
            changed   = True

    if not changed:
        _print("\n[dim]No changes made.[/dim]")
        _pause(); return

    # Rebuild nested dict from flat keys
    def _unflatten(flat_dict):
        result = {}
        for dotted_key, val in flat_dict.items():
            parts = dotted_key.split(".")
            d = result
            for part in parts[:-1]:
                d = d.setdefault(part, {})
            d[parts[-1]] = val
        return result

    _backup_file(prop_path)
    # Preserve header comments
    comments = "\n".join(l for l in raw.splitlines()
                         if l.startswith("#") or l.strip() == "")
    prop_path.write_text(
        comments + "\n\n" +
        yaml.dump(_unflatten(flat), default_flow_style=False,
                  allow_unicode=True, sort_keys=False),
        encoding="utf-8")
    _print(f"\n[green]✓ properties.yaml updated.[/green]")
    _pause()


def action_inspect(state: dict):
    _rule("Inspect Template — Extract Styles")
    _print("")
    _print("Select a Word file to extract its styles.\n")

    docx_path = _pick_file("Select Word file to inspect",
                           initial_dir=state["output_dir"])
    if not docx_path or not docx_path.exists():
        _print("[yellow]No file selected.[/yellow]")
        _pause(); return

    from lib.inspect_template import inspect as do_inspect
    do_inspect(docx_path)

    _print("")
    if _confirm("Update config.yaml with values from this template?"):
        _apply_template_to_config(docx_path, state["input_dir"])

    _pause()


def _backup_file(path: Path):
    """Copy path to input/backup/ with a timestamp suffix."""
    if not path.exists(): return
    backup_dir = path.parent / "backup"
    backup_dir.mkdir(exist_ok=True)
    ts  = datetime.now().strftime("%Y-%m-%d_%H-%M")
    dst = backup_dir / f"{path.stem}_{ts}{path.suffix}"
    shutil.copy(path, dst)
    _print(f"[dim]  Backup: input/backup/{dst.name}[/dim]")


def _apply_template_to_config(docx_path: Path, input_dir: Path):
    import re, yaml
    from docx import Document

    config_path = input_dir / "config.yaml"
    if not config_path.exists():
        _print("[red]config.yaml not found.[/red]"); return
    try:
        doc = Document(str(docx_path))
    except Exception as e:
        _print(f"[red]Could not open file: {e}[/red]"); return

    updates: dict = {}

    def _rgb(color):
        try:
            if color and color.rgb:
                r, g, b = color.rgb; return f"{r:02X}{g:02X}{b:02X}"
        except: pass
        return None

    def _pt(size):
        try:
            if size: return round(size.pt, 1)
        except: pass
        return None

    sec = doc.sections[0]
    if sec.page_width:
        updates.setdefault("page", {})["size"] = \
            "A4" if abs(sec.page_width.cm - 21.0) < 0.5 else "Letter"
    for attr, key in [("top_margin","margin_top"),("bottom_margin","margin_bottom"),
                      ("left_margin","margin_left"),("right_margin","margin_right")]:
        val = getattr(sec, attr, None)
        if val: updates.setdefault("page", {})[key] = f"{val.cm:.2f}cm"

    for sname, ckey in [("Heading 1","heading_1"),("Heading 2","heading_2"),
                        ("Heading 3","heading_3"),("Heading 4","heading_4"),
                        ("Heading 5","heading_5"),("Heading 6","heading_6"),
                        ("Normal","normal")]:
        try:
            s   = doc.styles[sname]
            cfg = {}
            if s.font.name:              cfg["font_name"]       = s.font.name
            pt = _pt(s.font.size)
            if pt:                       cfg["font_size_pt"]    = pt
            if s.font.bold is not None:  cfg["bold"]            = bool(s.font.bold)
            col = _rgb(s.font.color)
            if col:                      cfg["color"]           = col
            pb = _pt(s.paragraph_format.space_before)
            if pb is not None:           cfg["space_before_pt"] = pb
            pa = _pt(s.paragraph_format.space_after)
            if pa is not None:           cfg["space_after_pt"]  = pa
            if cfg: updates.setdefault("styles", {})[ckey] = cfg
        except KeyError: pass

    raw = config_path.read_text(encoding="utf-8")
    try:    existing = yaml.safe_load(raw) or {}
    except: existing = {}

    def _merge(base, new):
        for k, v in new.items():
            if isinstance(v, dict) and isinstance(base.get(k), dict): _merge(base[k], v)
            else: base[k] = v
    _merge(existing, updates)

    comments = "\n".join(l for l in raw.splitlines()
                         if l.startswith("#") or l.strip() == "")

    _backup_file(config_path)
    config_path.write_text(
        comments + "\n\n" +
        yaml.dump(existing, default_flow_style=False, allow_unicode=True, sort_keys=False),
        encoding="utf-8")

    n = sum(len(v) if isinstance(v, dict) else 1 for v in updates.values())
    _print(f"\n[green]✓ config.yaml updated with {n} values from {docx_path.name}[/green]")


# ── project runner ─────────────────────────────────────────────────────────────

def _run_project(proj_dir: Path):
    """Inner loop: project dashboard for a single project."""
    _save_last(proj_dir.name)

    actions = {
        "1":  action_build,
        "2":  action_open_document,
        "3":  action_export,
        "4":  action_open_vscode,
        "5":  action_edit_info,
        "6":  action_edit_properties,
        "7":  action_inspect,
        "8":  action_receive_file,
        "9":  action_sync,
    }

    while True:
        _clear()
        state = _project_state(proj_dir)
        _show_dashboard(state)

        choice = _prompt_dashboard(state)

        if choice == "q":
            _print("\n[dim]Goodbye.[/dim]\n")
            sys.exit(0)

        if choice == "b":
            return  # back to project picker

        if choice in actions:
            _clear()
            try:
                actions[choice](state)
            except KeyboardInterrupt:
                _print("\n[yellow]Interrupted.[/yellow]")
                _pause()
            except PermissionError as e:
                # Most common cause: the output .docx is open in Word
                fname = getattr(e, "filename", None)
                if fname and "docx" in str(fname).lower():
                    _print("\n[red]Cannot write the document — it is open in Word.\n"
                           "  Close the file in Word and try again.[/red]"
                           if HAS_RICH else
                           "\nCannot save — the .docx file is open. Close it in Word first.")
                else:
                    _print(f"\n[red]Permission denied: {e}[/red]"
                           if HAS_RICH else f"\nPermission denied: {e}")
                _pause()
            except Exception as e:
                _print(f"\n[red]Unexpected error: {e}[/red]"
                       if HAS_RICH else f"\nError: {e}")
                import traceback; traceback.print_exc()
                _pause()


# ── archive ───────────────────────────────────────────────────────────────────



# -- Entry point --------------------------------------------------------------

def main():
    PROJECTS_DIR.mkdir(exist_ok=True)

    while True:
        _clear()
        projects = _list_projects()

        if not projects:
            if HAS_RICH:
                console.print()
                console.print(Panel(
                    "[bold]Welcome![/bold]\n\n"
                    "No projects yet. Create your first one to get started.\n\n"
                    f"[dim]Projects folder: {PROJECTS_DIR}[/dim]",
                    border_style="blue"))
                console.print()
                console.print("  [cyan][A][/cyan]  New project")
                console.print("  [dim][Q]  Quit[/dim]")
                console.print()
            else:
                print(f"\nWelcome! No projects yet.\n  Projects folder: {PROJECTS_DIR}\n  [A] New project  [Q] Quit\n")
            try:    raw = input("Select: ").strip().lower()
            except: raw = "q"
            if raw == "q": break
            if raw == "a": action_new_project()
            continue

        choice = _show_picker(projects)
        if choice is None:  continue
        if choice == "q":   break
        if choice == "a":   action_new_project(); continue
        if choice == "s":   action_archive_project(); continue
        if choice == "f":   action_unarchive_project(); continue
        if choice == "d":   action_change_defaults(projects); continue

        proj_dir = PROJECTS_DIR / choice
        if proj_dir.exists():
            _run_project(proj_dir)

    _print("\n[dim]Goodbye.[/dim]\n" if HAS_RICH else "\nGoodbye.\n")


if __name__ == "__main__":
    main()
