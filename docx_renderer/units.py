"""OpenXML unit conversions and color math for MS Word layout fidelity."""

from __future__ import annotations
import colorsys
from typing import Optional, Tuple

# OpenXML Namespace Registry
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "v": "urn:schemas-microsoft-com:vml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
}


def qn(tag: str) -> str:
    """Convert a prefixed tag like 'w:pPr' into Clark notation '{uri}pPr'."""
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def twips_to_pt(twips: Optional[int | float | str]) -> float:
    """Convert twips (1/20th of a point, 1/1440th of an inch) to points."""
    if twips is None:
        return 0.0
    try:
        return float(twips) / 20.0
    except (ValueError, TypeError):
        return 0.0


def half_pts_to_pt(hps: Optional[int | float | str]) -> float:
    """Convert half-points (used for w:sz font size) to points."""
    if hps is None:
        return 0.0
    try:
        return float(hps) / 2.0
    except (ValueError, TypeError):
        return 0.0


def eighths_pt_to_pt(eighths: Optional[int | float | str]) -> float:
    """Convert eighths of a point (used for border widths w:sz) to points."""
    if eighths is None:
        return 0.0
    try:
        return float(eighths) / 8.0
    except (ValueError, TypeError):
        return 0.0


def emu_to_pt(emu: Optional[int | float | str]) -> float:
    """Convert English Metric Units (914,400 per inch = 12,700 per pt) to points."""
    if emu is None:
        return 0.0
    try:
        return float(emu) / 12700.0
    except (ValueError, TypeError):
        return 0.0


def parse_bool_attr(elem, attr_name: Optional[str] = None) -> bool:
    """Parse an OpenXML boolean property element (e.g., <w:b/> or <w:b w:val="0"/>)."""
    if elem is None:
        return False
    val = elem.get(qn(attr_name or "w:val"))
    if val is None:
        return True
    return val.lower() not in ("0", "false", "off", "none")


def hex_to_rgb_norm(hex_str: Optional[str]) -> Optional[Tuple[float, float, float]]:
    """Convert a 6-character hex string like '2F5496' to normalized (r, g, b) in [0, 1]."""
    if not hex_str:
        return None
    cleaned = hex_str.strip().lstrip("#")
    if cleaned.lower() in ("auto", "none", "transparent"):
        return None
    if len(cleaned) != 6:
        return None
    try:
        r = int(cleaned[0:2], 16) / 255.0
        g = int(cleaned[2:4], 16) / 255.0
        b = int(cleaned[4:6], 16) / 255.0
        return (r, g, b)
    except ValueError:
        return None


def apply_tint_or_shade(
    hex_color: str,
    theme_tint: Optional[str] = None,
    theme_shade: Optional[str] = None,
    lum_mod: Optional[int] = None,
    lum_off: Optional[int] = None,
) -> str:
    """
    Apply Word's themeTint, themeShade, or DrawingML lumMod/lumOff to a 6-char hex color.
    Matches ECMA-376 Part 1 section 17.3.2.6 HSL luminance calculation.
    """
    rgb = hex_to_rgb_norm(hex_color)
    if rgb is None:
        return hex_color

    r, g, b = rgb
    h, l, s = colorsys.rgb_to_hls(r, g, b)

    if lum_mod is not None or lum_off is not None:
        mod = (lum_mod / 100000.0) if lum_mod is not None else 1.0
        off = (lum_off / 100000.0) if lum_off is not None else 0.0
        l = max(0.0, min(1.0, l * mod + off))
    elif theme_tint:
        try:
            tint = int(theme_tint, 16) / 255.0
            l = l * tint + (1.0 - tint)
        except ValueError:
            pass
    elif theme_shade:
        try:
            shade = int(theme_shade, 16) / 255.0
            l = l * shade
        except ValueError:
            pass

    r2, g2, b2 = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
    return f"{int(round(r2 * 255)):02X}{int(round(g2 * 255)):02X}{int(round(b2 * 255)):02X}"


HIGHLIGHT_COLORS = {
    "yellow": "FFFF00",
    "green": "00FF00",
    "cyan": "00FFFF",
    "magenta": "FF00FF",
    "blue": "0000FF",
    "red": "FF0000",
    "darkBlue": "000080",
    "darkCyan": "008080",
    "darkGreen": "008000",
    "darkMagenta": "800080",
    "darkRed": "800000",
    "darkYellow": "808000",
    "darkGray": "808080",
    "lightGray": "C0C0C0",
    "black": "000000",
    "white": "FFFFFF",
}
