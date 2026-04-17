"""Lightweight pydantic schema for config.yaml.

Goals:
- Catch common user typos (wrong value types, invalid hex colors, unknown
  page sizes, malformed margins) with a clear error pointing at the key.
- Stay non-breaking: unknown keys are *warned* about, not rejected, because
  the existing style system tolerates unknown entries and we don't want to
  block a user whose config predates this schema.
- Run after yaml.safe_load in build_doc.load_config; if validation fails
  the raw dict is returned as before, so no behaviour regression.

Usage:

    from lib.build.config_schema import validate_config
    config = validate_config(raw_dict)  # returns the dict unchanged on success
"""
from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Union

from pydantic import BaseModel, ConfigDict, Field, field_validator

from ..log import get_logger

log = get_logger(__name__)


_HEX_RE = re.compile(r"^#?[0-9A-Fa-f]{6}$")
_LEN_RE = re.compile(r"^\s*\d+(\.\d+)?\s*(cm|in|mm|pt)\s*$")
_KNOWN_PAGE_SIZES = {"A3", "A4", "A5", "LETTER", "LEGAL"}
_KNOWN_ORIENTATIONS = {"portrait", "landscape"}


def _validate_hex_optional(v: Optional[str]) -> Optional[str]:
    if v is None or v == "":
        return v
    if not isinstance(v, str) or not _HEX_RE.match(v):
        raise ValueError(f"expected a 6-char hex color (e.g. '1F3864'), got {v!r}")
    return v


def _validate_length(v: Optional[str]) -> Optional[str]:
    if v is None:
        return v
    if not isinstance(v, str) or not _LEN_RE.match(v):
        raise ValueError(
            f"expected a length like '2.54cm' / '1in' / '18pt', got {v!r}"
        )
    return v


class PageConfig(BaseModel):
    model_config = ConfigDict(extra="allow")

    size: str = "A4"
    orientation: str = "portrait"
    margin_top: Optional[str] = None
    margin_bottom: Optional[str] = None
    margin_left: Optional[str] = None
    margin_right: Optional[str] = None
    header_distance: Optional[str] = None
    footer_distance: Optional[str] = None

    @field_validator("size")
    @classmethod
    def _check_size(cls, v: str) -> str:
        if v.upper() not in _KNOWN_PAGE_SIZES:
            raise ValueError(
                f"unknown page size {v!r}; expected one of {sorted(_KNOWN_PAGE_SIZES)}"
            )
        return v

    @field_validator("orientation")
    @classmethod
    def _check_orient(cls, v: str) -> str:
        if v.lower() not in _KNOWN_ORIENTATIONS:
            raise ValueError(
                f"unknown orientation {v!r}; expected 'portrait' or 'landscape'"
            )
        return v

    _check_margins = field_validator(
        "margin_top", "margin_bottom", "margin_left", "margin_right",
        "header_distance", "footer_distance",
    )(staticmethod(_validate_length))


class StyleBlock(BaseModel):
    """Any named style block — we validate the commonly-mistyped fields only."""
    model_config = ConfigDict(extra="allow")

    font_name: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color: Optional[str] = None
    background: Optional[str] = None
    border_color: Optional[str] = None
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None

    _check_colors = field_validator("color", "background", "border_color")(
        staticmethod(_validate_hex_optional)
    )


class FrontpageConfig(BaseModel):
    model_config = ConfigDict(extra="allow")

    cover_start_page: int = 1
    toc_start_page: int = 2
    content_start_page: int = 1


class ImageSizeEntry(BaseModel):
    model_config = ConfigDict(extra="allow")

    max_pct: Optional[float] = None
    pct: Optional[float] = None

    @field_validator("max_pct", "pct")
    @classmethod
    def _positive(cls, v: Optional[float]) -> Optional[float]:
        if v is not None and (v <= 0 or v > 100):
            raise ValueError(f"image size percentage must be in (0, 100], got {v!r}")
        return v


# Recognized top-level keys; anything else will trigger a warning.
_KNOWN_TOP_LEVEL = {
    "page", "header", "footer", "header_line", "footer_line",
    "frontpage", "numbered_headings", "styles", "image_sizes",
    "document",   # injected from document-info.yaml
}


def validate_config(raw: Dict[str, Any]) -> Dict[str, Any]:
    """Validate a loaded config dict in-place.

    On validation error, logs a warning and returns the dict unchanged so
    builds proceed with whatever defaults the builder fills in. Strict
    behavior can be added later behind a flag.
    """
    if not isinstance(raw, dict) or not raw:
        return raw

    # Warn on unknown top-level keys — highly useful for spotting typos like
    # "fontpage:" (missing 'r') or "pager:" (extra 'r').
    for key in raw:
        if key not in _KNOWN_TOP_LEVEL:
            log.warning("config: unknown top-level key %r (typo?)", key)

    try:
        if "page" in raw and isinstance(raw["page"], dict):
            PageConfig(**raw["page"])
        if "frontpage" in raw and isinstance(raw["frontpage"], dict):
            FrontpageConfig(**raw["frontpage"])
        if "styles" in raw and isinstance(raw["styles"], dict):
            for style_name, entry in raw["styles"].items():
                if isinstance(entry, dict):
                    StyleBlock(**entry)
        if "image_sizes" in raw and isinstance(raw["image_sizes"], dict):
            for name, entry in raw["image_sizes"].items():
                if isinstance(entry, dict):
                    ImageSizeEntry(**entry)
    except Exception as e:
        log.warning("config validation: %s", e)

    return raw
