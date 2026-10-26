"""Render a .docx with the engine and sanity-check the Table of Contents layout.

A no-Jupyter companion to `inspect_render.ipynb` - same renders, same two checks, but it needs
nothing beyond the project venv. Run it from the repo root:

    .venv\\Scripts\\python.exe inspect_render.py                      # every page + checks
    .venv\\Scripts\\python.exe inspect_render.py --pages 1-4          # just the front matter
    .venv\\Scripts\\python.exe inspect_render.py --pages 1,4,9-10 --open 1 4
    .venv\\Scripts\\python.exe inspect_render.py --pages 1-4 --ruler --bboxes --json
    .venv\\Scripts\\python.exe inspect_render.py --docx other.docx --dpi 150

The document is always parsed and laid out in full, so --pages only skips rasterization: page
numbers, headers/footers and the TOC checks are identical to a full render.
"""

from __future__ import annotations

import argparse
import sys
import time
from collections import Counter
from pathlib import Path
from typing import Iterator, List, Sequence

PROJECT = Path(__file__).resolve().parent
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))

from docx_renderer import CanvasRenderer, DocxParser, FontManager, LayoutEngine, parse_page_spec  # noqa: E402
from docx_renderer.model import ParagraphBlock, ParsedDocument, TableBlock  # noqa: E402

DEFAULT_DOCX = PROJECT / "NDA BLA XXXXXX - Safety Tables and Figures - 2 Columns RD.docx"


def iter_paragraphs(blocks) -> Iterator[ParagraphBlock]:
    """Yield every ParagraphBlock, descending into table rows and cells."""
    for b in blocks:
        if isinstance(b, ParagraphBlock):
            yield b
        elif isinstance(b, TableBlock):
            for row in b.rows:
                for cell in row.cells:
                    yield from iter_paragraphs(cell.blocks)


def check_toc_not_hyperlinked(parser: DocxParser, doc: ParsedDocument) -> bool:
    """TOC entries carry <w:rStyle w:val="Hyperlink"/>; Word does not render them blue."""
    paras = [p for p in iter_paragraphs(doc.blocks) if parser.is_toc_entry_style(p.style_id)]
    if not paras:
        print("  SKIP - this document has no TOC entry paragraphs")
        return True

    looks = {(r.style.color, r.style.underline) for p in paras for r in p.runs if r.kind == "text"}
    styles = ", ".join(sorted({p.style_id for p in paras}))

    print(f"  TOC entry paragraphs : {len(paras)}  ({styles})")
    print(f"  (color, underline)   : {looks or '{}'}")
    bad = {lk for lk in looks if lk != ("000000", None)}
    print("  FAIL - styled as hyperlinks: " + str(bad) if bad else "  PASS - rendered black, no underline")
    return not bad


def check_toc_dot_leaders(parser: DocxParser, doc: ParsedDocument, pages) -> bool:
    """Each TOC entry needs a dot leader running out to its right tab stop."""
    n_entries = sum(1 for p in iter_paragraphs(doc.blocks) if parser.is_toc_entry_style(p.style_id))
    if not n_entries:
        print("  SKIP - this document has no TOC entry paragraphs")
        return True

    leaders = [(bb.style_id, bb.x1) for pg in pages for bb in pg.bboxes if bb.kind == "tab_leader"]

    print(f"  dot leaders drawn    : {len(leaders)}  (TOC entry paragraphs: {n_entries})")
    print(f"  by style             : {dict(Counter(s for s, _ in leaders))}")
    print(f"  leader ends at x     : {dict(Counter(round(x) for _, x in leaders))} pt")

    ok = bool(leaders) and len(leaders) == n_entries
    print("  PASS - one leader per entry" if ok else "  FAIL - leaders missing or wrapped away")
    return ok


def iter_cells(blocks):
    """Yield every table cell, descending into nested tables."""
    for b in blocks:
        if isinstance(b, TableBlock):
            for row in b.rows:
                for cell in row.cells:
                    yield cell
                    yield from iter_cells(cell.blocks)


def check_table_shading(doc: ParsedDocument) -> bool:
    """`w:fill="auto"` is an absent fill, not a black one -- resolving it to black blacks out tables."""
    cells = list(iter_cells(doc.blocks))
    if not cells:
        print("  SKIP - this document has no tables")
        return True

    fills = Counter(c.shading_fill for c in cells)
    print(f"  table cells          : {len(cells)}")
    print(f"  resolved cell fills  : {dict(fills)}")
    black = fills.get("000000", 0)
    print("  PASS - no cell is filled black" if not black else f"  FAIL - {black} cells filled black")
    return not black


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--docx", type=Path, default=DEFAULT_DOCX, help="input .docx")
    ap.add_argument("-o", "--output-dir", type=Path, default=PROJECT / "inspect_out")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--pages", metavar="SPEC", default=None,
                    help="which pages to write: '3', '1-4', '1,3,5-7', '8-' (default: all)")
    ap.add_argument("--ruler", action="store_true", help="overlay the 0.5\" tab ruler")
    ap.add_argument("--bboxes", action="store_true", help="overlay element bounding boxes")
    ap.add_argument("--json", action="store_true", help="export the XPath -> bbox map")
    ap.add_argument("--open", nargs="*", type=int, metavar="PAGE",
                    help="open these rendered pages in the default image viewer")
    ap.add_argument("--no-checks", action="store_true", help="render only, skip the TOC checks")
    args = ap.parse_args(argv)

    if not args.docx.exists():
        print(f"Error: file not found: {args.docx}", file=sys.stderr)
        return 1

    t0 = time.time()
    fm = FontManager()
    parser = DocxParser(args.docx)
    doc = parser.parse()
    pages = LayoutEngine(doc, font_manager=fm).layout_document()
    t_layout = time.time() - t0

    try:
        wanted = set(parse_page_spec(args.pages, len(pages)))
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    written: List[Path] = CanvasRenderer(font_manager=fm).render_to_png_files(
        [pg for pg in pages if pg.page_number in wanted],
        output_dir=args.output_dir, prefix="page", dpi=args.dpi,
        show_ruler=args.ruler, show_bboxes=args.bboxes, export_layout_json=args.json,
    )
    print(f"{args.docx.name}")
    print(f"  parse + layout : {t_layout:.1f}s -> {len(pages)} pages")
    print(f"  rendered       : {len(written)} of {len(pages)} pages at {args.dpi} dpi -> {args.output_dir}")

    ok = True
    if not args.no_checks:
        print("\nCheck 1 - TOC entries must not be styled as hyperlinks")
        ok &= check_toc_not_hyperlinked(parser, doc)
        print("\nCheck 2 - TOC entries must keep their dot leader and right-aligned page number")
        ok &= check_toc_dot_leaders(parser, doc, pages)
        print("\nCheck 3 - table cells must not be filled black")
        ok &= check_table_shading(doc)

    for n in args.open or []:
        target = args.output_dir / f"page_{n}.png"
        if target.exists():
            print(f"\nopening {target}")
            if sys.platform == "win32":
                import os
                os.startfile(target)  # noqa: S606
            else:
                import subprocess
                subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", str(target)])
        else:
            print(f"\nno such page: {target}", file=sys.stderr)

    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
