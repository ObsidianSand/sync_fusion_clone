"""Record the Microsoft fonts' metrics into `docx_renderer/data/font_metrics.json.gz`.

    .venv\\Scripts\\python.exe tools_record_font_metrics.py [output path]

Run on a machine that has the Microsoft fonts (Windows). The renderer then lays text out with
these numbers wherever it runs, so line breaks and page breaks do not depend on the fonts
installed; see `fonts.py` for how they are used. Only measurements are stored -- no font data.

Per family and style it records:

* the ascender and descender, as fractions of the em, exactly as PyMuPDF reports them -- the
  numbers line heights are computed from;
* the advance width of every character the font has in `RANGES`, in 1/2048 em. Every width in
  these fonts is a whole number of those units, so the table reproduces `Font.text_length`
  exactly (this script checks that before writing).

Characters outside `RANGES` are measured with whatever font is found, for layout and drawing
alike, so they still cannot spill out of their space -- they just may wrap differently on a
machine without the font.
"""

from __future__ import annotations

import gzip
import json
import sys
from pathlib import Path

PROJECT = Path(__file__).resolve().parent
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))

from docx_renderer import fonts as F  # noqa: E402

OUT = PROJECT / "docx_renderer" / "data" / "font_metrics.json.gz"

#: Unicode blocks documents realistically use: Latin, Greek, Cyrillic, punctuation, currency,
#: letterlike symbols, number forms, arrows, math, technical, enclosed numbers, box drawing,
#: shapes, symbols, dingbats, and the Latin ligatures.
RANGES = [
    (0x0020, 0x024F), (0x0370, 0x03FF), (0x0400, 0x04FF), (0x2000, 0x206F), (0x20A0, 0x20CF),
    (0x2100, 0x214F), (0x2150, 0x218F), (0x2190, 0x21FF), (0x2200, 0x22FF), (0x2300, 0x23FF),
    (0x2460, 0x24FF), (0x2500, 0x257F), (0x25A0, 0x25FF), (0x2600, 0x26FF), (0x2700, 0x27BF),
    (0xFB00, 0xFB06),
]
UNITS = 2048
STYLES = [(False, False), (True, False), (False, True), (True, True)]


def _runs(widths: dict) -> list:
    """{codepoint: width} -> [[first_codepoint, [width, width, ...]], ...] over consecutive runs."""
    out: list = []
    prev = None
    for cp in sorted(widths):
        if prev is not None and cp == prev + 1:
            out[-1][1].append(widths[cp])
        else:
            out.append([cp, [widths[cp]]])
        prev = cp
    return out


def main(argv=None) -> int:
    args = sys.argv[1:] if argv is None else list(argv)
    out = Path(args[0]) if args else OUT
    fm = F.FontManager()
    tables: list = []
    seen: dict = {}
    families: dict = {}
    for fam, by_style in F.FONT_CANDIDATES.items():
        styles = {}
        for bold, italic in STYLES:
            fobj, path, _ = fm.get_font(fam, bold, italic)
            expected = by_style[(bold, italic)][0].lower()
            if path is None or Path(path).name.lower() != expected:
                print(f"error: {fam} {bold=} {italic=} resolved to {path or 'a built-in font'}, "
                      f"not {expected}. Run this where the Microsoft fonts are installed.", file=sys.stderr)
                return 1
            have = set(fobj.valid_codepoints())
            widths = {}
            for lo, hi in RANGES:
                for cp in range(lo, hi + 1):
                    if cp in have:
                        units = fobj.glyph_advance(cp) * UNITS
                        if abs(units - round(units)) > 1e-6:
                            print(f"error: {Path(path).name} U+{cp:04X} width is not a whole number of "
                                  f"1/{UNITS} em", file=sys.stderr)
                            return 1
                        widths[cp] = round(units)
            sample = "".join(chr(cp) for cp in sorted(widths))
            if abs(sum(widths.values()) / UNITS - fobj.text_length(sample, fontsize=1.0)) > 1e-6:
                print(f"error: summed widths do not reproduce text_length for {Path(path).name}", file=sys.stderr)
                return 1
            key = json.dumps(_runs(widths), separators=(",", ":"))
            if key not in seen:
                seen[key] = len(tables)
                tables.append(_runs(widths))
            styles["%d%d" % (bold, italic)] = {
                "ascender": float(fobj.ascender),
                "descender": abs(float(fobj.descender)),
                "widths": seen[key],
                "source": Path(path).name,
            }
        families[fam] = styles

    payload = {"units_per_em": UNITS, "families": families, "width_tables": tables}
    out.parent.mkdir(parents=True, exist_ok=True)
    raw = json.dumps(payload, separators=(",", ":"), sort_keys=True).encode("utf-8")
    out.write_bytes(gzip.compress(raw, compresslevel=9, mtime=0))
    print(f"wrote {out}: {len(families)} families, {len(tables)} distinct width tables, "
          f"{out.stat().st_size // 1024} KB")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
