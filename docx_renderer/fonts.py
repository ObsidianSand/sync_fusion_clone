"""Cross-platform Font Resolver and Metric Calculator using PyMuPDF (fitz)."""

from __future__ import annotations
import os
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import pymupdf as fitz


# Word font family -> candidate filenames (in priority order, including Linux metric-compatible clones)
FONT_CANDIDATES: Dict[str, Dict[Tuple[bool, bool], List[str]]] = {
    "calibri": {
        (False, False): ["calibri.ttf", "Carlito-Regular.ttf", "LiberationSans-Regular.ttf", "arial.ttf", "DejaVuSans.ttf"],
        (True, False): ["calibrib.ttf", "Carlito-Bold.ttf", "LiberationSans-Bold.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf"],
        (False, True): ["calibrii.ttf", "Carlito-Italic.ttf", "LiberationSans-Italic.ttf", "ariali.ttf", "DejaVuSans-Oblique.ttf"],
        (True, True): ["calibriz.ttf", "Carlito-BoldItalic.ttf", "LiberationSans-BoldItalic.ttf", "arialbi.ttf", "DejaVuSans-BoldOblique.ttf"],
    },
    "calibri light": {
        (False, False): ["calibril.ttf", "calibri.ttf", "Carlito-Regular.ttf", "LiberationSans-Regular.ttf"],
        (True, False): ["calibrib.ttf", "Carlito-Bold.ttf", "LiberationSans-Bold.ttf"],
        (False, True): ["calibrili.ttf", "calibrii.ttf", "Carlito-Italic.ttf", "LiberationSans-Italic.ttf"],
        (True, True): ["calibriz.ttf", "Carlito-BoldItalic.ttf", "LiberationSans-BoldItalic.ttf"],
    },
    "cambria": {
        (False, False): ["cambria.ttc", "cambria.ttf", "Caladea-Regular.ttf", "LiberationSerif-Regular.ttf", "times.ttf", "DejaVuSerif.ttf"],
        (True, False): ["cambriab.ttf", "Caladea-Bold.ttf", "LiberationSerif-Bold.ttf", "timesbd.ttf", "DejaVuSerif-Bold.ttf"],
        (False, True): ["cambriai.ttf", "Caladea-Italic.ttf", "LiberationSerif-Italic.ttf", "timesi.ttf", "DejaVuSerif-Italic.ttf"],
        (True, True): ["cambriaz.ttf", "Caladea-BoldItalic.ttf", "LiberationSerif-BoldItalic.ttf", "timesbi.ttf", "DejaVuSerif-BoldItalic.ttf"],
    },
    "arial": {
        (False, False): ["arial.ttf", "LiberationSans-Regular.ttf", "Arimo-Regular.ttf", "DejaVuSans.ttf"],
        (True, False): ["arialbd.ttf", "LiberationSans-Bold.ttf", "Arimo-Bold.ttf", "DejaVuSans-Bold.ttf"],
        (False, True): ["ariali.ttf", "LiberationSans-Italic.ttf", "Arimo-Italic.ttf", "DejaVuSans-Oblique.ttf"],
        (True, True): ["arialbi.ttf", "LiberationSans-BoldItalic.ttf", "Arimo-BoldItalic.ttf", "DejaVuSans-BoldOblique.ttf"],
    },
    "times new roman": {
        (False, False): ["times.ttf", "LiberationSerif-Regular.ttf", "Tinos-Regular.ttf", "DejaVuSerif.ttf"],
        (True, False): ["timesbd.ttf", "LiberationSerif-Bold.ttf", "Tinos-Bold.ttf", "DejaVuSerif-Bold.ttf"],
        (False, True): ["timesi.ttf", "LiberationSerif-Italic.ttf", "Tinos-Italic.ttf", "DejaVuSerif-Italic.ttf"],
        (True, True): ["timesbi.ttf", "LiberationSerif-BoldItalic.ttf", "Tinos-BoldItalic.ttf", "DejaVuSerif-BoldItalic.ttf"],
    },
    "courier new": {
        (False, False): ["cour.ttf", "LiberationMono-Regular.ttf", "Cousine-Regular.ttf", "DejaVuSansMono.ttf"],
        (True, False): ["courbd.ttf", "LiberationMono-Bold.ttf", "Cousine-Bold.ttf", "DejaVuSansMono-Bold.ttf"],
        (False, True): ["couri.ttf", "LiberationMono-Italic.ttf", "Cousine-Italic.ttf", "DejaVuSansMono-Oblique.ttf"],
        (True, True): ["courbi.ttf", "LiberationMono-BoldItalic.ttf", "Cousine-BoldItalic.ttf", "DejaVuSansMono-BoldOblique.ttf"],
    },
    "consolas": {
        (False, False): ["consola.ttf", "LiberationMono-Regular.ttf", "DejaVuSansMono.ttf", "cour.ttf"],
        (True, False): ["consolab.ttf", "LiberationMono-Bold.ttf", "DejaVuSansMono-Bold.ttf", "courbd.ttf"],
        (False, True): ["consolai.ttf", "LiberationMono-Italic.ttf", "DejaVuSansMono-Oblique.ttf", "couri.ttf"],
        (True, True): ["consolaz.ttf", "LiberationMono-BoldItalic.ttf", "DejaVuSansMono-BoldOblique.ttf", "courbi.ttf"],
    },
    "segoe ui": {
        (False, False): ["segoeui.ttf", "calibri.ttf", "Carlito-Regular.ttf", "LiberationSans-Regular.ttf"],
        (True, False): ["segoeuib.ttf", "calibrib.ttf", "Carlito-Bold.ttf", "LiberationSans-Bold.ttf"],
        (False, True): ["segoeuii.ttf", "calibrii.ttf", "Carlito-Italic.ttf", "LiberationSans-Italic.ttf"],
        (True, True): ["segoeuiz.ttf", "calibriz.ttf", "Carlito-BoldItalic.ttf", "LiberationSans-BoldItalic.ttf"],
    },
    "verdana": {
        (False, False): ["verdana.ttf", "DejaVuSans.ttf", "LiberationSans-Regular.ttf"],
        (True, False): ["verdanab.ttf", "DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf"],
        (False, True): ["verdanai.ttf", "DejaVuSans-Oblique.ttf", "LiberationSans-Italic.ttf"],
        (True, True): ["verdanaz.ttf", "DejaVuSans-BoldOblique.ttf", "LiberationSans-BoldItalic.ttf"],
    },
    "georgia": {
        (False, False): ["georgia.ttf", "LiberationSerif-Regular.ttf", "DejaVuSerif.ttf"],
        (True, False): ["georgiab.ttf", "LiberationSerif-Bold.ttf", "DejaVuSerif-Bold.ttf"],
        (False, True): ["georgiai.ttf", "LiberationSerif-Italic.ttf", "DejaVuSerif-Italic.ttf"],
        (True, True): ["georgiaz.ttf", "LiberationSerif-BoldItalic.ttf", "DejaVuSerif-BoldItalic.ttf"],
    },
}

SERIF_FAMILIES = {"cambria", "times new roman", "times", "georgia", "garamond", "palatino", "book antiqua", "caladea"}
MONO_FAMILIES = {"courier new", "courier", "consolas", "lucida console", "monaco", "dejavu sans mono", "liberation mono"}


class FontManager:
    """Discovers system/bundled fonts and provides fast glyph measurement via fitz.Font."""

    def __init__(self, extra_font_dirs: Optional[List[str | Path]] = None) -> None:
        self._font_file_index: Dict[str, str] = {}
        self._font_cache: Dict[Tuple[str, bool, bool], Tuple[fitz.Font, Optional[str], str]] = {}
        self._discover_fonts(extra_font_dirs or [])

    def _discover_fonts(self, extra_dirs: List[str | Path]) -> None:
        search_dirs: List[Path] = []
        for d in extra_dirs:
            p = Path(d)
            if p.is_dir():
                search_dirs.append(p)

        # Package local fonts folder (if bundled for air-gapped Linux)
        pkg_fonts = Path(__file__).parent / "fonts"
        if pkg_fonts.is_dir():
            search_dirs.append(pkg_fonts)

        # Standard OS font locations (Windows, Linux, macOS)
        win_dir = os.environ.get("WINDIR", r"C:\Windows")
        candidates = [
            Path(win_dir) / "Fonts",
            Path("/usr/share/fonts"),
            Path("/usr/local/share/fonts"),
            Path.home() / ".fonts",
            Path.home() / ".local/share/fonts",
            Path("/System/Library/Fonts"),
            Path("/Library/Fonts"),
        ]
        for c in candidates:
            if c.is_dir():
                search_dirs.append(c)

        for root_dir in search_dirs:
            try:
                for dirpath, _, filenames in os.walk(root_dir):
                    for fn in filenames:
                        lower = fn.lower()
                        if lower.endswith((".ttf", ".otf", ".ttc")) and lower not in self._font_file_index:
                            self._font_file_index[lower] = os.path.join(dirpath, fn)
            except OSError:
                continue

    def get_font(
        self, family: Optional[str], bold: bool = False, italic: bool = False
    ) -> Tuple[fitz.Font, Optional[str], str]:
        """
        Return (fitz_font, font_file_path_or_None, builtin_base_name).
        Caches loaded fitz.Font instances for speed.
        """
        fam_key = (family or "Calibri").strip().lower()
        cache_key = (fam_key, bool(bold), bool(italic))
        if cache_key in self._font_cache:
            return self._font_cache[cache_key]

        # 1. Check explicit candidate table
        candidates: List[str] = []
        if fam_key in FONT_CANDIDATES:
            candidates.extend(FONT_CANDIDATES[fam_key].get((bool(bold), bool(italic)), []))
        else:
            # Try direct filename heuristic
            clean = fam_key.replace(" ", "")
            if bold and italic:
                candidates.extend([f"{clean}bi.ttf", f"{clean}z.ttf", f"{clean}-bolditalic.ttf"])
            elif bold:
                candidates.extend([f"{clean}bd.ttf", f"{clean}b.ttf", f"{clean}-bold.ttf"])
            elif italic:
                candidates.extend([f"{clean}i.ttf", f"{clean}-italic.ttf"])
            candidates.append(f"{clean}.ttf")

            # Add category fallback candidates
            if fam_key in SERIF_FAMILIES:
                candidates.extend(FONT_CANDIDATES["cambria"][(bool(bold), bool(italic))])
            elif fam_key in MONO_FAMILIES:
                candidates.extend(FONT_CANDIDATES["consolas"][(bool(bold), bool(italic))])
            else:
                candidates.extend(FONT_CANDIDATES["calibri"][(bool(bold), bool(italic))])

        for cand in candidates:
            path = self._font_file_index.get(cand.lower())
            if path and os.path.exists(path):
                try:
                    fobj = fitz.Font(fontfile=path)
                    res = (fobj, path, "F0")
                    self._font_cache[cache_key] = res
                    return res
                except Exception:
                    continue

        # 2. Builtin Base-14 PDF/MuPDF fallback (works 100% offline even with 0 OS fonts)
        if fam_key in SERIF_FAMILIES:
            if bold and italic:
                bname = "tibi"
            elif bold:
                bname = "tibo"
            elif italic:
                bname = "tiit"
            else:
                bname = "tiro"
        elif fam_key in MONO_FAMILIES:
            if bold and italic:
                bname = "cobi"
            elif bold:
                bname = "cobo"
            elif italic:
                bname = "coit"
            else:
                bname = "cour"
        else:
            if bold and italic:
                bname = "hebi"
            elif bold:
                bname = "hebo"
            elif italic:
                bname = "heit"
            else:
                bname = "helv"

        fobj = fitz.Font(bname)
        res = (fobj, None, bname)
        self._font_cache[cache_key] = res
        return res

    def measure_text(
        self,
        text: str,
        family: Optional[str] = "Calibri",
        size_pt: float = 11.0,
        bold: bool = False,
        italic: bool = False,
        char_spacing_pt: float = 0.0,
    ) -> float:
        """Measure exact horizontal advance width in points for a string."""
        if not text:
            return 0.0
        fobj, _, _ = self.get_font(family, bold, italic)
        try:
            width = fobj.text_length(text, fontsize=size_pt)
        except Exception:
            width = len(text) * size_pt * 0.52
        if char_spacing_pt and len(text) > 0:
            width += char_spacing_pt * len(text)
        return width

    def get_metrics(
        self,
        family: Optional[str] = "Calibri",
        size_pt: float = 11.0,
        bold: bool = False,
        italic: bool = False,
    ) -> Tuple[float, float, float]:
        """
        Return (ascender_pt, descender_pt, default_single_line_height_pt).
        Note: descender_pt is returned as a positive distance below baseline.
        In MS Word, ' single' line spacing (240/240) uses the font's winAscent + winDescent + externalLeading,
        which for Calibri 11pt is ~1.15 * size_pt (~12.65pt) and for Times New Roman is ~1.15 * size_pt.
        """
        fobj, _, _ = self.get_font(family, bold, italic)
        asc = max(0.75, float(fobj.ascender)) * size_pt
        desc = max(0.20, abs(float(fobj.descender))) * size_pt
        # Word single line height (1.0 line = 240 twips-of-line) is typically ~1.15-1.20x point size
        single_line_pt = max(asc + desc, size_pt * 1.15)
        return asc, desc, single_line_pt
