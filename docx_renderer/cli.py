"""Command-line interface for rendering .docx files to high-fidelity PNGs."""

from __future__ import annotations
import argparse
import sys
from pathlib import Path
from typing import List, Optional

from .renderer import render_docx_to_png


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="docx_renderer",
        description="Render a Word (.docx) document to high-fidelity page PNGs (offline Syncfusion-style engine).",
    )
    parser.add_argument("docx_path", type=Path, help="Path to the input .docx file")
    parser.add_argument(
        "-o", "--output-dir", type=Path, default=Path("rendered_pages"), help="Directory to write PNG pages"
    )
    parser.add_argument("--prefix", type=str, default="page", help="Filename prefix for rendered PNGs")
    parser.add_argument("--dpi", type=int, default=150, help="Render resolution in DPI (default: 150)")
    parser.add_argument(
        "--pages",
        type=str,
        default=None,
        metavar="SPEC",
        help="Which pages to write: '3', '1-4', '1,3,5-7', '8-' (default: all). The whole document "
        "is still laid out, so page numbers and headers/footers match a full render.",
    )
    parser.add_argument(
        "--ruler",
        action="store_true",
        help="Draw top 0.5-inch tab-stop ruler and margin guide lines for LLM visual inspection",
    )
    parser.add_argument(
        "--bboxes",
        action="store_true",
        help="Draw color-coded bounding boxes around paragraphs, tables, and list markers",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Export a companion JSON file mapping every XML element XPath to its page bounding box and tab indent",
    )
    parser.add_argument(
        "--font-dir",
        action="append",
        default=[],
        help="Additional directory containing .ttf/.otf font files (can be specified multiple times)",
    )

    args = parser.parse_args(argv)
    if not args.docx_path.exists():
        print(f"Error: File not found: {args.docx_path}", file=sys.stderr)
        return 1

    try:
        pngs = render_docx_to_png(
            docx_path=args.docx_path,
            output_dir=args.output_dir,
            prefix=args.prefix,
            dpi=args.dpi,
            show_ruler=args.ruler,
            show_bboxes=args.bboxes,
            export_layout_json=args.json,
            extra_font_dirs=args.font_dir,
            pages=args.pages,
        )
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    for p in pngs:
        print(p)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
