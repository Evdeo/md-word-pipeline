"""Lightweight logging setup for md-word-pipeline.

One-line usage in every module:

    from lib.log import get_logger
    log = get_logger(__name__)
    log.info("...")
    log.warning("image not found: %s", path)

At process start (e.g. in run.py or build_doc.py) call configure(verbose=..., quiet=...).
If nothing calls configure(), the handler defaults to WARNING-level on stderr
so test runs stay quiet.
"""
from __future__ import annotations

import logging
import os
from typing import Optional

_LOGGER_NAME = "md_word_pipeline"
_configured = False


def _make_handler() -> logging.Handler:
    try:
        from rich.logging import RichHandler  # type: ignore
        return RichHandler(show_time=False, show_path=False, markup=False)
    except Exception:
        h = logging.StreamHandler()
        h.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        return h


def configure(verbose: bool = False, quiet: bool = False,
              level: Optional[int] = None) -> None:
    """Configure the package logger. Safe to call multiple times.

    Precedence: explicit level > verbose/quiet > env MD_WORD_PIPELINE_LOG > WARNING.
    """
    global _configured
    root = logging.getLogger(_LOGGER_NAME)

    # Drop any previously attached handlers so re-configuration is idempotent
    for h in list(root.handlers):
        root.removeHandler(h)

    if level is None:
        if verbose:
            level = logging.DEBUG
        elif quiet:
            level = logging.ERROR
        else:
            env = os.environ.get("MD_WORD_PIPELINE_LOG", "").upper()
            level = getattr(logging, env, logging.WARNING) if env else logging.WARNING

    root.setLevel(level)
    handler = _make_handler()
    handler.setLevel(level)
    root.addHandler(handler)
    # Propagate to root so pytest's caplog fixture and any user-configured
    # root logger can observe our records. Duplicate output is avoided because
    # root logging handlers default to WARNING only when unconfigured.
    root.propagate = True
    _configured = True


def get_logger(name: str) -> logging.Logger:
    """Return a child logger under the package namespace.

    `name` is typically `__name__`; only the trailing component is kept so
    log lines stay readable.
    """
    if not _configured:
        configure()
    short = name.rsplit(".", 1)[-1]
    return logging.getLogger(f"{_LOGGER_NAME}.{short}")
