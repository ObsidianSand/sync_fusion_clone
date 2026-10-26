"""Score our pagination against Word's own, using the breaks Word recorded in the .docx.

Word writes a <w:lastRenderedPageBreak/> into the run where each page began, the last time it laid
the document out. That is ground truth for pagination without needing Word installed. This script
lines those markers up against the first element on each of our rendered pages and reports how many
agree.

    .venv\\Scripts\\python.exe tools_compare_word_pagination.py
    .venv\\Scripts\\python.exe tools_compare_word_pagination.py --docx other.docx --from "w:p[1]"

Only compare from a point after the table of contents. Word's markers reflect its last save, so if
the TOC has been regenerated since, the front matter will not line up and tells you nothing --
hence the --from default below. See NOTES.md.
"""

from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path
from typing import Dict, List, Sequence

from lxml import etree

PROJECT = Path(__file__).resolve().parent
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))

from docx_renderer import DocxParser, FontManager, LayoutEngine  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DEFAULT_DOCX = PROJECT / "NDA BLA XXXXXX - Safety Tables and Figures - 2 Columns RD.docx"


def word_page_starts(docx: Path) -> List[str]:
    """Labels of the body elements Word began a page in, in document order."""
    root = etree.fromstring(zipfile.ZipFile(docx).read("word/document.xml"))
    body = root.find(W + "body")
    labels: Dict[etree._Element, str] = {}
    p_i = t_i = 0
    for child in body:
        if child.tag == W + "p":
            p_i += 1
            labels[child] = f"w:p[{p_i}]"
        elif child.tag == W + "tbl":
            t_i += 1
            labels[child] = f"w:tbl[{t_i}]"
    return [
        labels[child]
        for child in body
        if child in labels
        for _ in child.iter(W + "lastRenderedPageBreak")
    ]


def our_page_starts(docx: Path) -> tuple[list, List[str]]:
    """Our laid-out pages, and the label of the first element on each page after the first."""
    fm = FontManager()
    pages = LayoutEngine(DocxParser(docx).parse(), font_manager=fm).layout_document()
    starts = [
        next((bb.xml_path.split("/")[-1] for bb in pg.bboxes if bb.xml_path), None)
        for pg in pages[1:]
    ]
    return pages, starts


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--docx", type=Path, default=DEFAULT_DOCX)
    ap.add_argument("--from", dest="anchor", default="w:p[146]",
                    help="align both sequences on this element before scoring (default: w:p[146], "
                         "the first body element after the sample's front matter)")
    ap.add_argument("--show", type=int, default=14, help="how many boundaries to print")
    args = ap.parse_args(argv)

    if not args.docx.exists():
        print(f"Error: file not found: {args.docx}", file=sys.stderr)
        return 1

    word = word_page_starts(args.docx)
    pages, ours = our_page_starts(args.docx)
    blank = [pg.page_number for pg in pages if not (pg.texts or pg.images or pg.rects)]

    print(f"{args.docx.name}")
    print(f"  our pages            : {len(pages)}")
    print(f"  Word's cached layout : {len(word) + 1} pages")
    print(f"  blank pages          : {blank or 'none'}")

    if args.anchor not in word or args.anchor not in ours:
        print(f"\n  {args.anchor!r} is not a page start in both sequences; nothing to score.")
        print("  Pick an element that starts a page in each -- see the lists above/below.")
        return 1

    w2 = word[word.index(args.anchor):]
    o2 = ours[ours.index(args.anchor):]
    n = min(len(w2), len(o2))
    agree = sum(1 for a, b in zip(w2, o2) if a == b)
    print(f"  boundaries agreeing  : {agree}/{n} from {args.anchor} onward")

    print(f"\n  {'word':<12}{'ours':<12}")
    for i in range(min(args.show, max(len(w2), len(o2)))):
        a = w2[i] if i < len(w2) else ""
        b = o2[i] if i < len(o2) else ""
        print(f"  {a:<12}{b:<12}{'' if a == b else '  <-- differs'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
