"""Our pagination of the sample document must match Word's own record of it.

A `.docx` carries Word's layout of itself, so this needs no Word install. The trustworthy records
are `docProps/app.xml` `<Pages>` and the cached `PAGEREF` result of every TOC / Table of Tables /
Table of Figures entry -- one page-number anchor per heading and caption. `word_truth.py` documents
why the `<w:lastRenderedPageBreak/>` markers are deliberately *not* asserted on here.

The anchor assertion is brittle on purpose: one extra page early in the document shifts every later
anchor, so a pagination regression fails loudly instead of drifting by a point.
"""

from __future__ import annotations

from pathlib import Path

import pytest

import word_truth as wt
from docx_renderer import DocxParser, FontManager, LayoutEngine

PROJECT = Path(__file__).resolve().parents[1]
SAMPLE = PROJECT / "NDA BLA XXXXXX - Safety Tables and Figures - 2 Columns RD.docx"

pytestmark = pytest.mark.skipif(not SAMPLE.exists(), reason=f"sample not present: {SAMPLE.name}")


@pytest.fixture(scope="module")
def pages():
    """Lay the sample out once -- it is a 88-page document and takes a couple of seconds."""
    return LayoutEngine(DocxParser(SAMPLE).parse(), font_manager=FontManager()).layout_document()


def test_page_count_matches_words_own(pages) -> None:
    recorded = wt.recorded_page_count(SAMPLE)
    assert recorded is not None, "sample has no docProps/app.xml <Pages>"
    assert len(pages) == recorded, f"we lay out {len(pages)} pages, Word recorded {recorded}"


def test_every_pageref_anchor_lands_on_words_page(pages) -> None:
    targets = wt.pageref_targets(SAMPLE)
    assert len(targets) > 100, "expected the sample's TOC / ToT / ToF anchors to be found"

    page_of = wt.our_page_of_label(pages)
    misses = [
        f"{t.entry_label} {t.style_id} -> {t.target_label}: "
        f"Word={t.word_page} ours={page_of.get(t.target_label)} ({t.text[:40]})"
        for t in targets
        if page_of.get(t.target_label) != t.word_page
    ]
    assert not misses, "anchors not on the page Word recorded:\n  " + "\n  ".join(misses[:12])


def test_no_blank_pages(pages) -> None:
    """A page with nothing drawn means content was pushed off the page before it."""
    blank = [pg.page_number for pg in pages if not (pg.texts or pg.images or pg.rects)]
    assert not blank, f"nothing drawn on page(s) {blank}"


def test_body_label_ignores_table_cell_paragraphs() -> None:
    """A cell's first paragraph is also `w:p[1]`; keying on the last path segment confuses them."""
    body = wt.BODY_PREFIX
    assert wt._body_label(body + "w:p[133]") == "w:p[133]"
    assert wt._body_label(body + "w:tbl[22]") == "w:tbl[22]"
    assert wt._body_label(body + "w:tbl[22]/w:tr[1]/w:tc[1]/w:p[1]") is None
    assert wt._body_label("word/header1.xml:/w:hdr/w:p[1]") is None
    assert wt._body_label(None) is None


def test_page_break_only_paragraphs_are_recognized() -> None:
    """These are why Word's page-start markers land one paragraph later than our page starts."""
    labels = wt.page_break_only_labels(SAMPLE)
    # w:p[170] holds a single <w:br w:type="page"/> and no text; w:p[171] is the heading after it.
    assert "w:p[170]" in labels
    assert "w:p[171]" not in labels
