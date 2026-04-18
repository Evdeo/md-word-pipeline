"""lib/sync.py — Helpers for the review workflow."""
from pathlib import Path
from typing import Optional


def _find_received_docx(output_dir: Path) -> Optional[Path]:
    """Look in ``output/received/`` for reviewer-submitted Word files.

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
