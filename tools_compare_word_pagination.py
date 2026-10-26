"""Score our pagination against Word's own record of it, read out of the .docx.

    .venv\\Scripts\\python.exe tools_compare_word_pagination.py
    .venv\\Scripts\\python.exe tools_compare_word_pagination.py --docx other.docx --show 40

Three independent facts come out of the file (see `word_truth.py` for the details and the traps):

* `docProps/app.xml` `<Pages>` -- Word's own total page count. Compared first, because it decides
  whether the rest of Word's cached layout data can be trusted at all.
* every TOC / Table of Tables / Table of Figures entry's `PAGEREF` result -- the page Word resolved
  that heading, table caption or figure caption to. This is the score to optimize: one assertable
  anchor per entry, spread over the whole document.
* the `<w:lastRenderedPageBreak/>` markers -- reported last and only as a hint, because Word
  refreshes them lazily and they are stale in this repo's sample (83 pages against `<Pages>` 88).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Sequence

PROJECT = Path(__file__).resolve().parent
if str(PROJECT) not in sys.path:
    sys.path.insert(0, str(PROJECT))

import word_truth as wt  # noqa: E402
from docx_renderer import DocxParser, FontManager, LayoutEngine  # noqa: E402

DEFAULT_DOCX = PROJECT / "NDA BLA XXXXXX - Safety Tables and Figures - 2 Columns RD.docx"


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--docx", type=Path, default=DEFAULT_DOCX)
    ap.add_argument("--show", type=int, default=20,
                    help="how many mismatching anchors / markers to list (default: 20)")
    args = ap.parse_args(argv)

    if not args.docx.exists():
        print(f"Error: file not found: {args.docx}", file=sys.stderr)
        return 1

    pages = LayoutEngine(DocxParser(args.docx).parse(), font_manager=FontManager()).layout_document()
    blank = [pg.page_number for pg in pages if not (pg.texts or pg.images or pg.rects)]
    recorded = wt.recorded_page_count(args.docx)

    print(f"{args.docx.name}")
    print(f"  our pages            : {len(pages)}")
    print(f"  Word's <Pages>       : {recorded if recorded is not None else 'not recorded'}")
    print(f"  blank pages          : {blank or 'none'}")
    ok = not blank
    if recorded is not None:
        agree = recorded == len(pages)
        print(f"  page count           : {'AGREE' if agree else f'DIFFER by {len(pages) - recorded:+d}'}")
        ok &= agree

    # ---- the score: every PAGEREF anchor on the page Word recorded for it --------------------
    targets = wt.pageref_targets(args.docx)
    page_of = wt.our_page_of_label(pages)
    print("\n  TOC PAGEREF anchors (Word's page number for each heading / caption)")
    if not targets:
        print("    SKIP - this document has no TOC entries with PAGEREF fields")
    else:
        hits, misses, unresolved = 0, [], []
        for t in targets:
            our = page_of.get(t.target_label) if t.target_label else None
            if our is None:
                unresolved.append(t)
            elif our == t.word_page:
                hits += 1
            else:
                misses.append((t, our))
        print(f"    anchors            : {len(targets)}"
              + (f"  ({len(unresolved)} unresolved)" if unresolved else ""))
        print(f"    on Word's page     : {hits}/{len(targets)}")
        for t, our in misses[:args.show]:
            print(f"      {t.entry_label:<10}{t.style_id:<16}{t.target_label:<11}"
                  f"Word={t.word_page:<4}ours={our:<4}{t.text[:34]}")
        if len(misses) > args.show:
            print(f"      ... and {len(misses) - args.show} more")
        for t in unresolved[:args.show]:
            print(f"      {t.entry_label:<10}{t.style_id:<16}{'?':<11}"
                  f"Word={t.word_page:<4}bookmark {t.bookmark} not found in the body")
        print("    " + ("PASS - every anchor on the page Word recorded"
                        if hits == len(targets) else f"FAIL - {len(misses) + len(unresolved)} off"))
        ok &= hits == len(targets)

    # ---- the hint: Word's cached page-start markers ------------------------------------------
    markers = wt.rendered_page_starts(args.docx)
    print(f"\n  <w:lastRenderedPageBreak/> markers: {len(markers) + 1} pages  (a hint, not a score)")
    if recorded is not None and len(markers) + 1 != recorded:
        print(f"    STALE - Word's own <Pages> says {recorded}; these markers predate its last")
        print("            field update, so their positions cannot be scored. See word_truth.py.")

    # A page-break-only paragraph holds no run for Word to put the marker in, so Word's marker
    # lands one paragraph later than our page start. Map our starts forward before comparing.
    seq = wt.body_label_sequence(args.docx)
    nxt = {a: b for a, b in zip(seq, seq[1:])}
    pb_only = wt.page_break_only_labels(args.docx)
    ours = wt.our_page_starts(pages)
    equiv = {nxt.get(lb, lb) if lb in pb_only else lb for lb in ours if lb}

    missing = [m for m in markers if m not in equiv]
    print(f"    markers that are also one of our page starts: {len(markers) - len(missing)}/{len(markers)}")
    print(f"      (after mapping {sum(1 for lb in ours if lb in pb_only)} page-break-only "
          "paragraphs forward)")
    for m in missing[:args.show]:
        print(f"      no page of ours starts at {m}")
    if len(missing) > args.show:
        print(f"      ... and {len(missing) - args.show} more")

    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
