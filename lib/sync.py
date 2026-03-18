"""lib/sync.py — Utility helpers for the workflow tool."""
from pathlib import Path
from typing import Optional, Tuple


def _load_docx_size_classes(input_dir: Path) -> Tuple[dict, int]:
    from lib.build.images import build_size_classes, DEFAULT_SIZE_CLASSES
    config_path = input_dir / "config.yaml"
    size_classes = dict(DEFAULT_SIZE_CLASSES)
    content_width_emu = int(17 / 2.54 * 914_400)
    if config_path.exists():
        try:
            with open(config_path, encoding="utf-8") as f:
                cfg = yaml.safe_load(f) or {}
            size_classes = build_size_classes(cfg.get("image_sizes"))
            page_cfg = cfg.get("page", {})
            def cm_to_emu(s):
                try:
                    val = float(str(s).replace("cm","").replace("in","").strip())
                    return int(val * 914_400) if "in" in str(s) else int(val / 2.54 * 914_400)
                except Exception:
                    return int(2.54 / 2.54 * 914_400)
            size_str  = page_cfg.get("size","A4").upper()
            page_w    = int(21.0/2.54*914_400) if "A4" in size_str else int(21.59/2.54*914_400)
            ml = cm_to_emu(page_cfg.get("margin_left",  "2.54cm"))
            mr = cm_to_emu(page_cfg.get("margin_right", "2.54cm"))
            content_width_emu = page_w - ml - mr
        except Exception:
            pass
    return size_classes, content_width_emu




def _find_received_docx(output_dir: Path) -> Optional[Path]:
    """Look in output/received/ for reviewer-submitted Word files.

    Returns the single .docx found, or None if folder is missing/empty.
    Prompts the user if multiple files are present.
    """
    received_dir = output_dir / "received"
    if not received_dir.exists():
        return None

    docx_files = sorted(received_dir.glob("*.docx"))
    if not docx_files:
        return None

    if len(docx_files) == 1:
        return docx_files[0]

    # Multiple files — ask which one to use
    print("Multiple files found in output/received/:")
    for i, f in enumerate(docx_files, 1):
        from datetime import datetime
        size_kb = f.stat().st_size // 1024
        ts = datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        print(f"  [{i}] {f.name}  ({size_kb} KB, modified {ts})")
    print()
    while True:
        try:
            raw = input("Which file to compare? ").strip()
        except (EOFError, KeyboardInterrupt):
            return None
        if raw.isdigit() and 1 <= int(raw) <= len(docx_files):
            return docx_files[int(raw) - 1]
        print(f"Please enter a number between 1 and {len(docx_files)}")


