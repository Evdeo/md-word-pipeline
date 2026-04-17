#!/usr/bin/env python3
"""
md_to_docx — Convert Markdown files to professional Word documents.

Usage:
  python md_to_docx.py                              # Uses default paths (input/ → output/document.docx)
  python md_to_docx.py input.md -o output.docx      # Custom paths
  python md_to_docx.py docs/ -o output.docx         # Directory source
  python md_to_docx.py --template my_template.docx  # Use a Word template as base

Config files (all in input/ by default):
  config.yaml        — page layout, header/footer, image sizes  (edit rarely)
  document-info.yaml — title, author, date, revisions           (edit per project)
  properties.yaml    — placeholder values for {{key}} syntax    (project data)
"""

import argparse
import re
import sys
from pathlib import Path
from typing import Optional

import yaml

from lib.build.builder import DocumentBuilder
from lib.log import configure as _configure_logging, get_logger

log = get_logger(__name__)


# ── path helpers ──────────────────────────────────────────────────────────────

def get_project_root() -> Path:
    return Path(__file__).parent.resolve()


def resolve_user_path(path_str: str) -> Path:
    p = Path(path_str)
    return p.resolve() if p.is_absolute() else (Path.cwd() / p).resolve()


def get_default_path(relative: str) -> Path:
    return (get_project_root() / relative).resolve()


# ── YAML loaders ──────────────────────────────────────────────────────────────

def load_config(path: Path) -> dict:
    if path and path.exists():
        try:
            with open(path, encoding="utf-8") as f:
                return yaml.safe_load(f) or {}
        except Exception as e:
            log.warning("could not parse config.yaml (%s) — using defaults.", e)
            return {}
    return {}


def load_document_info(path: Path) -> tuple:
    """Load document-info.yaml.

    Returns (document_dict, revisions_list).
    Falls back gracefully if the file is missing.
    """
    if path and path.exists():
        with open(path, encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        return data.get("document", {}), data.get("revisions", [])
    return {}, []


def load_all_yaml_files(src_dir: Path, exclude_files: Optional[set] = None) -> dict:
    """Load all YAML files in directory, flatten to dotted keys, detect duplicates.

    Excluded files are skipped entirely (not flattened into properties).
    """
    exclude_files = exclude_files or set()
    all_properties: dict = {}
    seen_keys: dict = {}

    yaml_files = sorted(
        list(src_dir.glob("*.yaml")) + list(src_dir.glob("*.yml"))
    )

    for yaml_path in yaml_files:
        if yaml_path.name in exclude_files:
            continue

        with open(yaml_path, encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}

        def flatten(d, prefix=""):
            for key, value in d.items():
                full_key = f"{prefix}{key}" if prefix else key
                if isinstance(value, dict):
                    flatten(value, f"{full_key}.")
                elif isinstance(value, list):
                    pass  # lists (e.g. revisions) are not flattened into properties
                else:
                    if full_key in seen_keys:
                        raise ValueError(
                            f"Duplicate property '{full_key}' found in both "
                            f"'{seen_keys[full_key]}' and '{yaml_path.name}'. "
                            f"Property names must be unique across all YAML files."
                        )
                    seen_keys[full_key] = yaml_path.name
                    all_properties[full_key] = str(value)

        flatten(data)

    return all_properties


def substitute_properties(text: str, properties: dict) -> str:
    """Replace {{property.name}} placeholders with actual values."""
    if not properties:
        return text

    def replace_match(match):
        key = match.group(1).strip()
        return properties.get(key, match.group(0))

    return re.sub(r'\{\{(\s*[\w.]+\s*)\}\}', replace_match, text)




# ── file collection ───────────────────────────────────────────────────────────

def collect_files(source: Path):
    """Return (frontpage_path_or_None, [content_paths…])."""
    if source.is_file():
        return None, [source]

    all_md = sorted(source.glob("*.md"))
    frontpage = None
    content   = []
    for p in all_md:
        if p.name == "00-frontpage.md" or "::: {toc=false}" in p.read_text(encoding="utf-8"):
            frontpage = p
        else:
            content.append(p)
    return frontpage, content


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    project_root = get_project_root()

    ap = argparse.ArgumentParser(
        description="Convert Markdown → Word (.docx)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Examples:
  python md_to_docx.py                              # Use defaults
  python md_to_docx.py input/ -o output.docx        # Custom source/output
  python md_to_docx.py --template brand.docx        # Apply Word template styles

Default paths (relative to Project folder):
  Source:        input/
  Output:        output/document.docx
  Config:        input/config.yaml
  Document info: input/document-info.yaml

Project location: {project_root}
        """
    )
    ap.add_argument("source", nargs="?", default=None,
                    help="Markdown file or directory (default: input/)")
    ap.add_argument("-o", "--output", default=None,
                    help="Output .docx path (default: output/document.docx)")
    ap.add_argument("-c", "--config", default=None,
                    help="config.yaml path (default: input/config.yaml)")
    ap.add_argument("--template", default=None,
                    help="Word .docx template to use as base document")
    ap.add_argument("-v", "--verbose", action="store_true",
                    help="Enable debug-level logging")
    ap.add_argument("-q", "--quiet", action="store_true",
                    help="Suppress warnings (errors only)")
    args = ap.parse_args()

    _configure_logging(verbose=args.verbose, quiet=args.quiet)

    # Resolve paths
    source      = resolve_user_path(args.source) if args.source else get_default_path("input")
    output      = resolve_user_path(args.output) if args.output else get_default_path("output/document.docx")
    config_path = resolve_user_path(args.config) if args.config else get_default_path("input/config.yaml")
    template    = resolve_user_path(args.template) if args.template else None

    if not source.exists():
        print(f"Error: Source not found: {source}", file=sys.stderr)
        sys.exit(1)

    if template and not template.exists():
        print(f"Error: Template not found: {template}", file=sys.stderr)
        sys.exit(1)

    # Load config.yaml
    config = load_config(config_path)

    src_dir = source.parent if source.is_file() else source

    # Load document-info.yaml — document identity + revisions
    doc_info_path = src_dir / "document-info.yaml"
    document_info, revisions = load_document_info(doc_info_path)

    if document_info:
        # Inject into config so header/footer substitution and builder work
        config["document"] = document_info
        print(f"Loaded document info: {doc_info_path.name}")
    elif "document" not in config:
        # Backwards compatibility: old config.yaml with document: block
        config.setdefault("document", {})

    if revisions:
        print(f"Loaded revisions: {len(revisions)} entries")

    # Load all other YAML files for {{placeholder}} substitution.
    # Exclude config files that are handled separately above.
    EXCLUDED = {"config.yaml", "document-info.yaml", "revisions.yaml"}
    try:
        properties = load_all_yaml_files(src_dir, exclude_files=EXCLUDED)

        # Also expose document.* fields as properties for use in markdown
        for key, value in document_info.items():
            prop_key = f"document.{key}"
            if prop_key not in properties:
                properties[prop_key] = str(value)

        if properties:
            print(f"Loaded properties: {len(properties)} values from YAML files")

    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

    # Build document
    builder = DocumentBuilder(
        config=config,
        revisions=revisions,
        template_path=str(template) if template else None,
    )
    builder.setup()

    frontpage, content_files = collect_files(source)

    # Pre-scan to build figure/table label map
    all_texts = []
    for cf in content_files:
        try:
            all_texts.append(substitute_properties(cf.read_text(encoding="utf-8"), properties))
        except Exception:
            all_texts.append("")
    builder.prescan_labels(all_texts)

    if frontpage:
        fp_text = substitute_properties(frontpage.read_text(encoding="utf-8"), properties)
        builder.add_frontpage(fp_text, frontpage.parent)
        builder.add_toc()
    else:
        builder.add_toc()

    for cf, cf_text in zip(content_files, all_texts):
        try:
            builder.add_content(cf_text, cf.parent)
        except Exception as e:
            log.warning("error processing %s: %s", cf.name, e)

    output.parent.mkdir(parents=True, exist_ok=True)

    lock_file = output.parent / f"~${output.name}"
    if lock_file.exists():
        print(f"Error: {output.name} is currently open in Microsoft Word.", file=sys.stderr)
        print(f"Please close the document and try again.", file=sys.stderr)
        sys.exit(1)

    try:
        builder.save(output)
        print(f"Created: {output}")
    except PermissionError:
        print(f"Error: Cannot write to {output}", file=sys.stderr)
        print(f"The file may be open in another program.", file=sys.stderr)
        sys.exit(1)



if __name__ == "__main__":
    main()
