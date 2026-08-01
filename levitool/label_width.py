"""Measure the real rendered width of label text using the actual installed font.

The previous approach counted characters and compared against a per-font-size table
(CHARS_PER_LINE_30PP). That ignored two facts that make labels overflow: fonts are
proportional (a 'W' is far wider than an 'i') and bold text is wider than regular.
A 38-character line could therefore be reported as "fits" while Word wrapped it.

This module measures the true width of a string in its real font family, point size,
and weight using Pillow (already a project dependency), so the fit check matches what
Word actually renders. Widths are compared against the label cell's usable width
(cell width minus its left/right margins), read straight from the docx.

A label line is measured as a list of *segments* rather than as one string, because Word
mixes formatting inside a single line: the CFCG logo run is a different family and size
from the sentence beside it. Measuring a whole line in one family collapsed that logo to
an Arial '.notdef' advance and under-reported the line by about 5 pt, which was enough to
call a wrapping line "fits". Each segment is a dict with the keys 'text', 'family',
'size', 'bold' and 'is_symbol'; see get_label_line_segments() in
find_longest_value_in_label_rows.py for how they are built.
"""

import math
import os
from functools import lru_cache

from PIL import ImageFont

from .constants import MAX_SUGGESTED_SIZE_PT

# We load the font once at a large reference size and scale linearly (width is exactly
# proportional to point size). Measuring at a large size avoids integer-pixel rounding
# error that shows up when measuring directly at small sizes like 9 or 11 pt.
_REF_SIZE = 1000

# Fallback family used when a label's named font is not installed on this machine.
DEFAULT_FAMILY = "Arial"

# Fraction of the usable width a line may occupy before it is called TOO WIDE. Pillow and
# Word do not agree to the last fraction of a point (kerning and hinting differ), and a
# real overflow can be as small as 0.3 pt, so a line that measures within 1% of the limit
# is treated as overflowing. This cannot be loosened much: a correct line in a real label
# was measured at 98.5% of the usable width.
FIT_TOLERANCE = 0.99

# Fraction of the usable width a *suggested* font size is sized to fit. Looser than
# FIT_TOLERANCE on purpose: once a line has to be resized anyway, landing a little small
# is better than landing exactly on the edge and wrapping.
SUGGEST_TOLERANCE = 0.97

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


def segment_width_pt(segment: dict) -> float:
    """Rendered width of one segment in points, or None if it cannot be measured."""
    text = segment.get('text') or ''
    size_pt = segment.get('size')
    if not text or not size_pt:
        return 0.0 if not text else None
    path, _, _ = resolve_font(segment.get('family'), bool(segment.get('bold')))
    if path is None:
        return None
    return width_pt(path, text, size_pt)


def segments_width_pt(segments) -> float:
    """Total rendered width of a line's segments in points.

    Each segment is measured in its own family, size, and weight, then the advances are
    summed - that is how Word lays a line out, and it is the whole point of keeping the
    line split into segments rather than measuring it as one string.

    Args:
        segments (list): Segment dicts with 'text', 'family', 'size', and 'bold' keys.

    Returns:
        float: Summed width in points, or None if any non-empty segment could not be
            measured (no font size, or no usable font file on this machine).
    """
    total = 0.0
    for seg in segments:
        w = segment_width_pt(seg)
        if w is None:
            return None
        total += w
    return total


def dominant_text_size(text_segments):
    """Point size of the text segment covering the most characters, or None."""
    by_size = {}
    for seg in text_segments:
        size_pt = seg.get('size')
        if size_pt:
            by_size[size_pt] = by_size.get(size_pt, 0) + len(seg.get('text') or '')
    if not by_size:
        return None
    return max(by_size, key=by_size.get)


def max_fitting_size_pt(segments, usable_pt: float) -> float:
    """Largest point size (in 0.5 pt steps) for a line's text so the line fits.

    Answers the question in both directions: it is what to shrink an overflowing line to,
    and equally what a line with room to spare could be grown to. The result is capped at
    MAX_SUGGESTED_SIZE_PT - label text is never set larger than that, and without the cap a
    short line measures as fitting at a size that would push the label past its fixed row
    height, where Word clips the text rather than growing the row.

    Symbol segments (the CFCG logo and the like) are held at their current size and only
    the text segments are resized, which is what happens in practice: the sentence gets
    made smaller and the logo is left alone. Width is proportional to size, so the text
    segments can be scaled by a single factor:

        factor = (usable_pt * SUGGEST_TOLERANCE - symbol_width) / text_width

    The returned size is that factor applied to the dominant text size and rounded down to
    the nearest half point. Rounding down is not enough on its own when a line carries more
    than one text size - the others are scaled by the same ratio and may still land over -
    so the result is re-measured and stepped down by half a point until it really fits.

    Args:
        segments (list): Segment dicts for the line, as measured by segments_width_pt().
        usable_pt (float): The line's usable width in points.

    Returns:
        float: Suggested point size, never above MAX_SUGGESTED_SIZE_PT, or None when no
            text can be resized (no text segments, unmeasurable widths, or the symbols
            alone already overflow).
    """
    text_segments = [s for s in segments if not s.get('is_symbol') and (s.get('text') or '')]
    symbol_segments = [s for s in segments if s.get('is_symbol')]

    text_w = segments_width_pt(text_segments)
    symbol_w = segments_width_pt(symbol_segments)
    if not text_w or symbol_w is None:
        return None

    dominant_size = dominant_text_size(text_segments)
    if not dominant_size:
        return None

    budget = usable_pt * SUGGEST_TOLERANCE
    factor = (budget - symbol_w) / text_w
    if factor <= 0:
        return None

    suggested = min(math.floor(factor * dominant_size * 2) / 2, MAX_SUGGESTED_SIZE_PT)
    while suggested >= 5:
        scaled = [dict(s, size=s['size'] * suggested / dominant_size) for s in text_segments]
        width = segments_width_pt(scaled)
        if width is None:
            return None
        if symbol_w + width <= budget:
            return suggested
        suggested -= 0.5
    return suggested
