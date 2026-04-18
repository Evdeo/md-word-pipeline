"""load_project() deep-merges default + named + project-local overrides."""
from pathlib import Path

import pytest
import yaml

from lib.config_loader import _deep_merge, load_project


def test_deep_merge_combines_nested_dicts():
    base    = {"page": {"size": "A4", "margin_top": "2cm"}, "styles": {"h1": {"bold": True}}}
    overlay = {"page": {"margin_top": "3cm"},             "styles": {"h1": {"size": 20}}}
    merged  = _deep_merge(base, overlay)

    assert merged["page"]["size"]       == "A4"       # from base
    assert merged["page"]["margin_top"] == "3cm"      # overlay wins
    assert merged["styles"]["h1"]       == {"bold": True, "size": 20}


def test_deep_merge_overlay_replaces_non_dict():
    base    = {"images": [1, 2, 3], "count": 1}
    overlay = {"images": [9],       "count": 2}
    merged  = _deep_merge(base, overlay)
    assert merged == {"images": [9], "count": 2}


def _write_yaml(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data), encoding="utf-8")


def test_load_project_deep_merges_default_and_named(monkeypatch, tmp_path):
    import lib.config_loader as loader

    configs = tmp_path / "configs"
    projects = tmp_path / "projects"
    _write_yaml(configs / "default.yaml", {
        "page":   {"size": "A4", "margin_left": "2cm"},
        "styles": {"heading_1": {"bold": True}},
    })
    _write_yaml(configs / "tech.yaml", {
        "page":   {"margin_left": "3cm"},     # overrides default
        "styles": {"heading_1": {"color": "1F3864"}},
    })
    _write_yaml(projects / "demo" / "project.yaml", {
        "config": "tech",
        "title_override": "Draft",
    })
    _write_yaml(projects / "demo" / "document-info.yaml", {
        "document": {"title": "Original"},
    })
    (projects / "demo" / "content.md").write_text("# Hello\n", encoding="utf-8")
    (projects / "demo" / "images").mkdir()

    monkeypatch.setattr(loader, "CONFIGS_DIR", configs)
    monkeypatch.setattr(loader, "PROJECTS_DIR", projects)

    ctx = load_project(projects / "demo")

    assert ctx.config["page"]["size"]         == "A4"      # from default
    assert ctx.config["page"]["margin_left"]  == "3cm"     # overridden by tech
    assert ctx.config["styles"]["heading_1"]["bold"]  is True
    assert ctx.config["styles"]["heading_1"]["color"] == "1F3864"
    assert ctx.document_info["title"] == "Draft"           # title_override wins
    assert ctx.output_path == projects / "demo" / "output" / "document.docx"


def test_load_project_rejects_missing_config(tmp_path, monkeypatch):
    import lib.config_loader as loader

    configs = tmp_path / "configs"
    _write_yaml(configs / "default.yaml", {"page": {"size": "A4"}})
    projects = tmp_path / "projects"
    _write_yaml(projects / "demo" / "project.yaml", {"config": "doesnotexist"})

    monkeypatch.setattr(loader, "CONFIGS_DIR", configs)
    monkeypatch.setattr(loader, "PROJECTS_DIR", projects)

    with pytest.raises(FileNotFoundError):
        load_project(projects / "demo")
