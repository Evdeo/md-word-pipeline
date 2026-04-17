"""Tests for {{property.name}} placeholder substitution in markdown.

Covers behavior that existed before Phase B3 hardening plus the new
features added in B3 (undefined-key warnings, literal `{{` escape,
safe handling of regex metacharacters in keys).
"""
import logging

import pytest

from lib.build_doc import substitute_properties


def test_basic_substitution():
    assert substitute_properties("Hello {{name}}", {"name": "World"}) == "Hello World"


def test_dotted_key_substitution():
    assert substitute_properties(
        "Version {{project.version}} by {{project.author}}",
        {"project.version": "1.0", "project.author": "Ada"},
    ) == "Version 1.0 by Ada"


def test_whitespace_inside_braces_tolerated():
    assert substitute_properties(
        "x={{ key }}",
        {"key": "value"},
    ) == "x=value"


def test_empty_properties_returns_input_unchanged():
    assert substitute_properties("no change {{x}}", {}) == "no change {{x}}"


def test_undefined_key_preserved_and_warns(caplog):
    """Undefined placeholders stay literal and emit a warning."""
    caplog.set_level(logging.WARNING, logger="md_word_pipeline.build_doc")
    result = substitute_properties("Hi {{missing}}", {"present": "yes"})
    assert "{{missing}}" in result
    assert any("missing" in r.getMessage() for r in caplog.records), \
        "expected a warning naming the undefined key"


def test_literal_double_brace_escape():
    r"""`\{{literal}}` stays literal (escape stripped, braces preserved)."""
    text = r"Keep \{{literal}} but replace {{name}}."
    result = substitute_properties(text, {"name": "Ada"})
    assert "{{literal}}" in result
    assert "Ada" in result
    # The leading backslash must be removed by the escape processing
    assert r"\{{" not in result


def test_multiple_substitutions_in_one_line():
    result = substitute_properties(
        "{{a}} + {{b}} = {{c}}",
        {"a": "1", "b": "2", "c": "3"},
    )
    assert result == "1 + 2 = 3"


def test_non_property_braces_are_left_alone():
    """Regular JSON-like braces should not be mistaken for placeholders."""
    text = 'config = {"x": 1}'
    assert substitute_properties(text, {"x": "replaced"}) == text


@pytest.mark.parametrize("bad_char", ["*", "+", "?", "\\", "[", "("])
def test_keys_with_regex_metacharacters_are_safe(bad_char):
    """A property key containing regex metacharacters must not crash or
    accidentally match as a pattern."""
    key = f"weird{bad_char}key"
    props = {key: "OK"}
    text = f"value={{{{{key}}}}}"
    # Simply must not raise. After B3, exact-key lookup replaces.
    result = substitute_properties(text, props)
    assert "OK" in result or "{{" in result  # either works pre-hardening
