"""Measure the real rendered width of label text using the actual installed font.

The previous approach counted characters and compared against a per-font-size table
(CHARS_PER_LINE_30PP). That ignored two facts that make labels overflow: fonts are
proportional (a 'W' is far wider than an 'i') and bold text is wider than regular.
A 38-character line could therefore be reported as "fits" while Word wrapped it.

This module measures the true width of a string in its real font family, point size,
and weight using Pillow (already a project dependency), so the fit check matches what
Word actually renders. Widths are compared against the label cell's usable width
(cell width minus its left/right margins), read straight from the docx.
"""

import math
import os
from functools import lru_cache

from PIL import ImageFont

# We load the font once at a large reference size and scale linearly (width is exactly
# proportional to point size). Measuring at a large size avoids integer-pixel rounding
# error that shows up when measuring directly at small sizes like 9 or 11 pt.
_REF_SIZE = 1000

# Fallback family used when a label's named font is not installed on this machine.
DEFAULT_FAMILY = "Arial"

# Directories searched for installed font files (macOS first, then common Linux/Windows).
_FONT_DIRS = [
    "/System/Library/Fonts/Supplemental",
    "/System/Library/Fonts",
    "/Library/Fonts",
    os.path.expanduser("~/Library/Fonts"),
    "/usr/share/fonts",
    "/usr/local/share/fonts",
    "C:/Windows/Fonts",
]


def _candidate_filenames(family: str, bold: bool):
    """Yields likely font filenames for a family and weight, most specific first.

    Covers the common macOS/Windows naming conventions ('Arial Bold.ttf',
    'Arial-Bold.ttf', 'arialbd.ttf', ...) so a family name from the docx can be
    matched to an installed file without a full font-database lookup.
    """
    fam = family.strip()
    fam_nospace = fam.replace(" ", "")
    exts = (".ttf", ".ttc", ".otf")
    if bold:
        stems = [f"{fam} Bold", f"{fam}-Bold", f"{fam_nospace}-Bold",
                 f"{fam_nospace}bd", f"{fam_nospace}-BoldMT", f"{fam}bd"]
    else:
        stems = [fam, f"{fam}-Regular", f"{fam_nospace}-Regular",
                 f"{fam_nospace}MT", fam_nospace]
    for stem in stems:
        for ext in exts:
            yield stem + ext


@lru_cache(maxsize=None)
def _find_font_path(family: str, bold: bool):
    """Locate an installed font file for a family and weight, or None if not found."""
    for fname in _candidate_filenames(family, bold):
        for directory in _FONT_DIRS:
            candidate = os.path.join(directory, fname)
            if os.path.exists(candidate):
                return candidate
    return None


@lru_cache(maxsize=None)
def resolve_font(family: str, bold: bool):
    """Resolve a family and weight to an installed font file.

    Args:
        family (str): Font family name from the label (e.g. 'Arial', 'Calibri').
            May be None/empty, in which case the default family is used.
        bold (bool): Whether the bold weight is wanted.

    Returns:
        tuple: (font_path, family_used, fallback). font_path is the located file or
            None if nothing usable was found. family_used is the family actually
            resolved. fallback is True when the requested family was unavailable and
            the default family was substituted (so the width is an approximation).
    """
    requested = (family or DEFAULT_FAMILY).strip() or DEFAULT_FAMILY

    # Exact family, requested weight.
    path = _find_font_path(requested, bold)
    if path is not None:
        return path, requested, False
    # Same family, regular weight (bold file missing but family present).
    if bold:
        path = _find_font_path(requested, False)
        if path is not None:
            return path, requested, False
    # Fall back to the default family.
    path = _find_font_path(DEFAULT_FAMILY, bold) or _find_font_path(DEFAULT_FAMILY, False)
    return path, DEFAULT_FAMILY, (path is not None)


@lru_cache(maxsize=None)
def _ref_font(font_path: str):
    """Return a Pillow font for font_path loaded once at the reference size."""
    return ImageFont.truetype(font_path, size=_REF_SIZE)


def width_pt(font_path: str, text: str, size_pt: float) -> float:
    """Rendered advance width of text in points, for a font file and point size."""
    return _ref_font(font_path).getlength(text) * size_pt / _REF_SIZE


def text_width_pt(text: str, family: str, size_pt: float, bold: bool):
    """Rendered width of text in points, or None if no usable font was found."""
    if not size_pt:
        return None
    path, _, _ = resolve_font(family, bold)
    if path is None:
        return None
    return width_pt(path, text, size_pt)


def max_fitting_size_pt(font_path: str, text: str, current_size_pt: float,
                        usable_pt: float) -> float:
    """Largest point size (in 0.5 pt steps) at which text fits within usable_pt.

    Width is proportional to size, so the exact largest fitting size is
    ``usable_pt * current_size / width_at_current``. It is rounded down to the nearest
    half point so the suggested size is guaranteed to fit.
    """
    w = width_pt(font_path, text, current_size_pt)
    if not w:
        return None
    exact = usable_pt * current_size_pt / w
    return math.floor(exact * 2) / 2
