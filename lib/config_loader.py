"""Load a project's merged configuration into a ProjectContext."""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import yaml

from lib.build.config_schema import validate_config
from lib.build_doc import (collect_files, load_all_yaml_files,
                           load_document_info, substitute_properties)
from lib.log import get_logger

log = get_logger(__name__)

REPO_ROOT   = Path(__file__).resolve().parent.parent
CONFIGS_DIR = REPO_ROOT / "configs"
PROJECTS_DIR = REPO_ROOT / "projects"


@dataclass
class ProjectContext:
    project_dir: Path
    name: str
    config: dict
    document_info: dict
    revisions: list
    properties: dict
    frontpage: Optional[Path]
    content_files: list
    output_path: Path

    @property
    def output_dir(self) -> Path:
        return self.output_path.parent

    @property
    def render_dir(self) -> Path:
        return self.output_dir / "render"


def _deep_merge(base: dict, overlay: dict) -> dict:
    """Return a new dict: overlay's keys win, nested dicts merged recursively."""
    out = copy.deepcopy(base)
    for key, val in overlay.items():
        if isinstance(val, dict) and isinstance(out.get(key), dict):
            out[key] = _deep_merge(out[key], val)
        else:
            out[key] = copy.deepcopy(val)
    return out


def _load_yaml(path: Path) -> dict:
    if not path.is_file():
        return {}
    with path.open(encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def resolve_project_dir(project: str | Path) -> Path:
    """Accept a bare name (looked up under projects/), a relative path, or
    an absolute path. Return the resolved directory."""
    p = Path(project)
    if not p.is_absolute():
        direct = Path.cwd() / p
        if direct.is_dir():
            p = direct
        else:
            p = PROJECTS_DIR / project
    return p.resolve()


def load_project(project_dir: Path) -> ProjectContext:
    """Parse `project_dir/project.yaml`, deep-merge with the named config
    under `configs/`, and return a fully-populated `ProjectContext`."""
    project_dir = project_dir.resolve()
    proj_yaml = _load_yaml(project_dir / "project.yaml")

    config_name = proj_yaml.get("config", "default")
    default_path = CONFIGS_DIR / "default.yaml"
    named_path   = CONFIGS_DIR / f"{config_name}.yaml"
    if not named_path.is_file():
        raise FileNotFoundError(
            f"config '{config_name}' not found at {named_path}")

    merged: dict = {}
    if default_path.is_file() and default_path != named_path:
        merged = _deep_merge(merged, _load_yaml(default_path))
    merged = _deep_merge(merged, _load_yaml(named_path))

    doc_info, revisions = load_document_info(
        project_dir / "document-info.yaml")
    if "title_override" in proj_yaml:
        doc_info = dict(doc_info)
        doc_info["title"] = proj_yaml["title_override"]
    if doc_info:
        merged["document"] = doc_info
    merged.setdefault("document", {})

    merged = validate_config(merged)

    excluded = {"project.yaml", "document-info.yaml", "revisions.yaml"}
    try:
        properties = load_all_yaml_files(project_dir, exclude_files=excluded)
    except ValueError as e:
        log.error("property collision: %s", e)
        raise
    for key, value in doc_info.items():
        properties.setdefault(f"document.{key}", str(value))

    frontpage, content_files = collect_files(project_dir)

    output_name = proj_yaml.get("output", "document.docx")
    output_path = project_dir / "output" / output_name

    return ProjectContext(
        project_dir=project_dir,
        name=project_dir.name,
        config=merged,
        document_info=doc_info,
        revisions=revisions,
        properties=properties,
        frontpage=frontpage,
        content_files=content_files,
        output_path=output_path,
    )


def build(ctx: ProjectContext, out_path: Optional[Path] = None) -> Path:
    """Build the docx from a ProjectContext. Returns the output path."""
    from lib.build.builder import DocumentBuilder

    out = Path(out_path) if out_path else ctx.output_path
    out.parent.mkdir(parents=True, exist_ok=True)

    builder = DocumentBuilder(
        config=ctx.config,
        revisions=ctx.revisions,
        source_dir=ctx.project_dir,
    )
    builder._verbose = False
    builder.setup()

    texts = []
    for cf in ctx.content_files:
        try:
            texts.append(substitute_properties(
                cf.read_text(encoding="utf-8"), ctx.properties))
        except Exception:
            texts.append("")
    builder.prescan_labels(texts)

    word_cover = ctx.config.get("frontpage", {}).get("word_cover", "")
    if word_cover and (ctx.project_dir / word_cover).exists():
        builder.add_word_cover(ctx.project_dir / word_cover)
    elif ctx.frontpage:
        builder.add_frontpage(
            substitute_properties(
                ctx.frontpage.read_text(encoding="utf-8"), ctx.properties),
            ctx.frontpage.parent)
    builder.add_toc()
    for cf, text in zip(ctx.content_files, texts):
        try:
            builder.add_content(text, cf.parent)
        except Exception as e:
            log.warning("error processing %s: %s", cf.name, e)

    builder.save(out)
    return out


def list_projects() -> list[Path]:
    """Return sorted list of project folders that have a project.yaml."""
    if not PROJECTS_DIR.is_dir():
        return []
    return sorted(
        p for p in PROJECTS_DIR.iterdir()
        if p.is_dir() and not p.name.startswith("_")
        and (p / "project.yaml").is_file()
    )
