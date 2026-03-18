"""Image loading and dimension calculation."""
import re
from pathlib import Path
from typing import Dict, Optional, Tuple

# 1 inch = 914 400 EMU;  1 px (96 dpi) = 9 525 EMU
EMU_PER_INCH = 914_400
EMU_PER_PX   = 9_525

# Default size classes used when no config is provided.
# Keys are class names; values are fraction of content width (0.0–1.0).
DEFAULT_SIZE_CLASSES: Dict[str, float] = {
    "xs":     0.20,
    "small":  0.30,
    "medium": 0.50,
    "large":  0.75,
    "xl":     1.00,
}

# Keep this alias so any code that imported it directly still works.
SIZE_CLASS_PCT = DEFAULT_SIZE_CLASSES


def build_size_classes(image_sizes_cfg: Optional[dict]) -> Dict[str, float]:
    """Convert the image_sizes config block into a {name: fraction} dict.

    config block example:
        image_sizes:
          xs:     { max_pct: 15 }
          small:  { max_pct: 35 }
          medium: { max_pct: 55 }

    Falls back to DEFAULT_SIZE_CLASSES if cfg is None or empty.
    """
    if not image_sizes_cfg:
        return dict(DEFAULT_SIZE_CLASSES)
    result = {}
    for name, entry in image_sizes_cfg.items():
        if isinstance(entry, dict):
            pct = entry.get("max_pct", entry.get("pct", 50))
        else:
            pct = float(entry)
        result[str(name)] = float(pct) / 100.0
    return result if result else dict(DEFAULT_SIZE_CLASSES)


def nearest_size_class(
    width_emu: int,
    content_width_emu: int,
    size_classes: Dict[str, float],
) -> str:
    """Return the size class name whose threshold is closest to the actual width fraction.

    Used by docx_to_md.py when reverse-converting image widths.
    size_classes maps name → fraction (0.0–1.0).
    """
    if content_width_emu <= 0:
        return "medium"
    actual_frac = width_emu / content_width_emu
    # Find the class whose max_pct is closest to the actual fraction
    best_name = "medium"
    best_diff = float("inf")
    for name, frac in size_classes.items():
        diff = abs(frac - actual_frac)
        if diff < best_diff:
            best_diff = diff
            best_name = name
    return best_name


class ImageProcessor:
    def __init__(
        self,
        page_width_in: float = 8.27,
        margin_left_in: float = 1.0,
        margin_right_in: float = 1.0,
        size_classes: Optional[Dict[str, float]] = None,
    ):
        self.content_width_emu = int(
            (page_width_in - margin_left_in - margin_right_in) * EMU_PER_INCH
        )
        self.size_classes: Dict[str, float] = (
            size_classes if size_classes is not None else dict(DEFAULT_SIZE_CLASSES)
        )

    def load(self, src: str, base_dir: Path) -> Optional[Tuple[Path, int, int]]:
        """Return (path, width_px, height_px) or None if not found."""
        p = Path(src)
        if not p.is_absolute():
            p = base_dir / p
        if not p.exists():
            return None
        try:
            from PIL import Image
            with Image.open(p) as im:
                return p, im.width, im.height
        except Exception:
            return None

    def calc_emu(
        self,
        w_px: int,
        h_px: int,
        size_class: Optional[str] = None,
        width_attr: Optional[str] = None,
    ) -> Tuple[int, int]:
        """Return (width_emu, height_emu) preserving aspect ratio."""
        if size_class and size_class in self.size_classes:
            w_emu = int(self.content_width_emu * self.size_classes[size_class])
        elif width_attr:
            m = re.match(r"width\s*=\s*([0-9.]+)(in|px|%)", width_attr)
            if m:
                val, unit = float(m.group(1)), m.group(2)
                if unit == "in":
                    w_emu = int(val * EMU_PER_INCH)
                elif unit == "px":
                    w_emu = int(val * EMU_PER_PX)
                else:  # %
                    w_emu = int(self.content_width_emu * val / 100)
            else:
                w_emu = int(self.content_width_emu * 0.5)
        else:
            w_emu = int(self.content_width_emu * 0.5)  # default: 50%

        aspect = h_px / w_px if w_px else 1.0
        h_emu  = int(w_emu * aspect)
        return w_emu, h_emu

    def inches(self, emu: int) -> float:
        return emu / EMU_PER_INCH

    def nearest_class(self, width_emu: int) -> str:
        """Return size class name nearest to the given EMU width."""
        return nearest_size_class(width_emu, self.content_width_emu, self.size_classes)
