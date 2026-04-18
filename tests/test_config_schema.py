"""Tests for config validation via pydantic schema."""
import logging

from lib.build.config_schema import validate_config


def test_valid_default_config_passes_silently(caplog):
    """The bundled default.yaml must validate without warnings."""
    import yaml
    from pathlib import Path
    default_path = Path(__file__).resolve().parent.parent / "configs" / "default.yaml"
    raw = yaml.safe_load(default_path.read_text(encoding="utf-8"))
    caplog.set_level(logging.WARNING, logger="md_word_pipeline.config_schema")
    validate_config(raw)
    assert not [r for r in caplog.records if r.levelno >= logging.WARNING], \
        f"unexpected warnings: {[r.getMessage() for r in caplog.records]}"


def test_unknown_top_level_key_warns(caplog):
    caplog.set_level(logging.WARNING, logger="md_word_pipeline.config_schema")
    validate_config({"pager": {"size": "A4"}})  # typo: pager
    messages = [r.getMessage() for r in caplog.records]
    assert any("pager" in m for m in messages)


def test_invalid_hex_color_warns(caplog):
    caplog.set_level(logging.WARNING, logger="md_word_pipeline.config_schema")
    validate_config({
        "styles": {"heading_1": {"color": "not-a-hex"}}
    })
    messages = [r.getMessage() for r in caplog.records]
    assert any("hex color" in m for m in messages)


def test_invalid_page_size_warns(caplog):
    caplog.set_level(logging.WARNING, logger="md_word_pipeline.config_schema")
    validate_config({"page": {"size": "quatre-cent"}})
    messages = [r.getMessage() for r in caplog.records]
    assert any("unknown page size" in m for m in messages)


def test_invalid_margin_format_warns(caplog):
    caplog.set_level(logging.WARNING, logger="md_word_pipeline.config_schema")
    validate_config({"page": {"margin_top": "2 bananas"}})
    messages = [r.getMessage() for r in caplog.records]
    assert any("length" in m for m in messages)


def test_invalid_image_pct_warns(caplog):
    caplog.set_level(logging.WARNING, logger="md_word_pipeline.config_schema")
    validate_config({"image_sizes": {"huge": {"max_pct": 250}}})
    messages = [r.getMessage() for r in caplog.records]
    assert any("100" in m or "percentage" in m for m in messages)


def test_validation_is_non_breaking():
    """Even on a horribly wrong config, validate_config returns the dict."""
    bad = {
        "page": {"size": 42},
        "styles": "this should be a dict",
        "image_sizes": {"x": {"max_pct": -5}},
    }
    result = validate_config(bad)
    assert result is bad


def test_empty_config_is_ok():
    assert validate_config({}) == {}
    assert validate_config(None) is None
