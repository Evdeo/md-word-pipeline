#!/usr/bin/env python3
"""md.py — one entry point for the markdown → Word pipeline.

Usage:
    python md.py                          pick a project and an action
    python md.py <project>                live preview (default action)
    python md.py <project> -build         build once
    python md.py <project> -diff          section-diff vs output/received/
    python md.py <project> -open          open the built docx in Word
    python md.py <project> -import FILE   extract docx into this project
    python md.py -new <name>              scaffold a new empty project

Single-dash short flags. `-b`, `-d`, `-o`, `-i` are aliases. Unknown
flags print usage and exit.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path
from typing import NamedTuple, Optional


REPO_ROOT   = Path(__file__).resolve().parent
PROJECTS    = REPO_ROOT / "projects"
TEMPLATE    = PROJECTS / "_template"

ACTIONS    = ("live", "build", "diff", "open", "import", "new")
ALIASES    = {"b": "build", "d": "diff", "o": "open", "i": "import"}


class Args(NamedTuple):
    project: Optional[str]
    action: str                    # one of ACTIONS
    extra: Optional[str] = None    # e.g. docx path for -import, new name for -new


USAGE = (
    "usage: python md.py [<project>] [-build | -diff | -open "
    "| -import FILE | -new <name>]"
)


def parse_args(argv: list[str]) -> Args:
    """Hand-rolled parser. argv is sys.argv[1:] (no program name)."""
    if not argv:
        raise ValueError("no arguments")

    # `-new <name>` is the only flag that precedes the project name.
    if argv[0] in ("-new",):
        if len(argv) != 2:
            raise ValueError("-new needs exactly one name argument")
        return Args(project=None, action="new", extra=argv[1])

    # Otherwise the first positional is the project name.
    project = argv[0]
    rest = argv[1:]
    if not rest:
        return Args(project=project, action="live")

    flag = rest[0]
    if not flag.startswith("-"):
        raise ValueError(f"expected a flag, got {flag!r}")
    name = flag[1:]
    action = ALIASES.get(name, name)
    if action not in ACTIONS or action == "new":
        raise ValueError(f"unknown flag {flag!r}")

    if action == "import":
        if len(rest) != 2:
            raise ValueError("-import needs a file argument")
        return Args(project=project, action="import", extra=rest[1])

    if len(rest) != 1:
        raise ValueError(f"{flag} takes no extra arguments")
    return Args(project=project, action=action)


# ── action implementations ────────────────────────────────────────────────────

def run_live(ctx) -> None:
    from lib.preview import start_preview_server

    handle = start_preview_server(ctx, open_browser=True)
    print(f"preview: {handle.url}   (Ctrl+C to stop)")
    handle.wait()


def run_build(ctx) -> None:
    from lib.config_loader import build

    out = build(ctx)
    print(f"built: {out}")


def run_diff(ctx) -> None:
    from lib.section_diff import run as _run
    _run(ctx)


def run_open(ctx) -> None:
    out = ctx.output_path
    if not out.exists():
        from lib.config_loader import build
        print("no docx yet — building first")
        build(ctx)
    if sys.platform == "win32":
        os.startfile(str(out))          # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        os.system(f'open "{out}"')
    else:
        os.system(f'xdg-open "{out}" >/dev/null 2>&1 &')


def run_import(ctx, docx: str) -> None:
    docx_path = Path(docx).expanduser().resolve()
    if not docx_path.is_file():
        sys.exit(f"no such file: {docx_path}")

    md_file = ctx.project_dir / "content.md"
    if md_file.exists() and md_file.read_text(encoding="utf-8").strip():
        sys.exit(
            f"refusing to import: {md_file} is not empty.\n"
            "Clear it first, or use `-new` for a fresh project.")

    from lib.extract import extract_to_project

    extract_to_project(docx_path, ctx.project_dir, config=ctx.config)
    print(f"imported {docx_path.name} → {ctx.project_dir}")


def scaffold_new(name: str) -> None:
    if not TEMPLATE.is_dir():
        sys.exit(f"template folder missing at {TEMPLATE}")
    if "/" in name or "\\" in name or name.startswith("_"):
        sys.exit(f"invalid project name: {name!r}")

    dest = PROJECTS / name
    if dest.exists():
        sys.exit(f"refusing to overwrite: {dest} already exists")

    shutil.copytree(str(TEMPLATE), str(dest))
    # Drop the .gitkeep from the images folder — not needed outside the template.
    gitkeep = dest / "images" / ".gitkeep"
    if gitkeep.exists():
        gitkeep.unlink()
    print(f"created {dest}")


# ── interactive menu ──────────────────────────────────────────────────────────

def pick_from_list(title: str, items: list[str],
                   default_index: Optional[int] = None) -> Optional[int]:
    """Prompt with a numbered list. Enter picks `default_index`. Empty list →
    None. Returns the 0-based index picked, or None on quit."""
    if not items:
        print(f"{title}: nothing to pick.")
        return None

    print(f"\n{title}")
    for i, it in enumerate(items, 1):
        marker = "  (default)" if default_index is not None and i - 1 == default_index else ""
        print(f"  [{i}] {it}{marker}")
    print("  [Q] quit")

    while True:
        try:
            raw = input("> ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            return None
        if raw in ("q", "quit", "exit"):
            return None
        if raw == "" and default_index is not None:
            return default_index
        if raw.isdigit():
            n = int(raw)
            if 1 <= n <= len(items):
                return n - 1
        # match by unique prefix (case-insensitive)
        matches = [i for i, it in enumerate(items)
                   if it.lower().startswith(raw)]
        if len(matches) == 1:
            return matches[0]
        print(f"  pick 1-{len(items)} (or Q to quit)")


def prompt_project_and_action() -> Args:
    from lib.config_loader import list_projects

    projects = [p.name for p in list_projects()]
    if not projects:
        print("No projects yet. Create one with:  python md.py -new <name>")
        sys.exit(0)

    idx = pick_from_list("Pick a project:", projects)
    if idx is None:
        sys.exit(0)
    project = projects[idx]

    labels = [
        "Live preview",
        "Build once",
        "Diff received docx",
        "Open in Word",
        "Import docx",
    ]
    actions = ["live", "build", "diff", "open", "import"]
    idx = pick_from_list(f"Project: {project}", labels, default_index=0)
    if idx is None:
        sys.exit(0)

    action = actions[idx]
    extra: Optional[str] = None
    if action == "import":
        try:
            extra = input("Path to .docx: ").strip()
        except (EOFError, KeyboardInterrupt):
            sys.exit(0)
        if not extra:
            sys.exit("cancelled")
    return Args(project=project, action=action, extra=extra)


# ── main ──────────────────────────────────────────────────────────────────────

def main(argv: list[str]) -> None:
    if len(argv) == 1:
        args = prompt_project_and_action()
    else:
        try:
            args = parse_args(argv[1:])
        except ValueError as e:
            print(f"error: {e}\n{USAGE}", file=sys.stderr)
            sys.exit(2)

    if args.action == "new":
        scaffold_new(args.extra)  # type: ignore[arg-type]
        return

    from lib.config_loader import load_project, resolve_project_dir

    project_dir = resolve_project_dir(args.project)  # type: ignore[arg-type]
    if not (project_dir / "project.yaml").is_file():
        sys.exit(f"not a project: {project_dir} (no project.yaml)")

    ctx = load_project(project_dir)

    if   args.action == "live":   run_live(ctx)
    elif args.action == "build":  run_build(ctx)
    elif args.action == "diff":   run_diff(ctx)
    elif args.action == "open":   run_open(ctx)
    elif args.action == "import": run_import(ctx, args.extra)  # type: ignore[arg-type]


if __name__ == "__main__":
    main(sys.argv)
