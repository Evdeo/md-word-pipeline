"""Reusable markdown → docx helpers.

The user-facing CLI lives in ``md.py`` at the repo root. This module only
exposes the building blocks (``load_config``, ``load_document_info``,
``load_all_yaml_files``, ``substitute_properties``, ``collect_files``) used
by ``lib/config_loader.py`` and friends.
"""

import re
from pathlib import Path
from typing import Optional

import yaml

from lib.build.config_schema import validate_config
from lib.log import get_logger

log = get_logger(__name__)


# ── YAML loaders ──────────────────────────────────────────────────────────────

def load_config(path: Path) -> dict:
    if path and path.exists():
        try:
            with open(path, encoding="utf-8") as f:
                raw = yaml.safe_load(f) or {}
            return validate_config(raw)
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


_PLACEHOLDER_RE = re.compile(r'(?<!\\)\{\{(\s*[\w.]+\s*)\}\}')
_ESCAPED_RE = re.compile(r'\\(\{\{[^}]*\}\})')


def substitute_properties(text: str, properties: dict) -> str:
    r"""Replace ``{{property.name}}`` placeholders with their values.

    Behavior:
    - Unknown keys are left literal and a warning is logged naming the key.
    - ``\{{literal}}`` is an escape — the backslash is stripped, the braces
      survive. Nothing inside is substituted.
    - Braces that don't match the placeholder grammar (e.g. ``{"x": 1}``) are
      untouched.
    """
    if not text:
        return text

    unknown: set = set()

    def _replace(m: "re.Match") -> str:
        key = m.group(1).strip()
        if not properties or key not in properties:
            unknown.add(key)
            return m.group(0)
        return properties[key]

    result = _PLACEHOLDER_RE.sub(_replace, text)
    # Strip the single-backslash escape on \{{...}} so the braces render literally.
    result = _ESCAPED_RE.sub(r'\1', result)

    for key in sorted(unknown):
        log.warning("undefined property {{%s}} — left as literal", key)

    return result




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
