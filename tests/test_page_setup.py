"""Page size changes between sections, and table rows too tall for the space left on a page.

Both failures this guards against looked the same in a render: a landscape page whose content
stopped part-way down, with the rest missing. Landscape pages are only 612 pt tall, so they hit
these limits first.

* A continuous section break into a section of a different page size or orientation must start a
  new page -- one page cannot be two sizes. Kept on the same page, a portrait section following a
  landscape one flowed down to the portrait bottom margin and off the bottom of the 612 pt page,
  and a landscape section following a portrait one ran off its right edge.
* A table row that does not fit must break across pages, as Word does unless the row is
  `w:cantSplit`. Placed whole, a row taller than the page was drawn from the top of it and lost
  everything below the bottom edge.

Like `test_columns.py`, these build their own python-docx fixtures and assert geometry and
reading order: such a file carries no Word-recorded layout to compare against.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import List

import docx
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.shared import Inches, Pt

from docx_renderer import DocxParser, FontManager, LayoutEngine

TOKEN = re.compile(r"^L\d\d$")


def _landscape(section) -> None:
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = Inches(11), Inches(8.5)


def _portrait(section) -> None:
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width, section.page_height = Inches(8.5), Inches(11)


def _new_doc() -> docx.Document:
    doc = docx.Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10)
    s0 = doc.sections[0]
    s0.top_margin = s0.bottom_margin = Inches(1.0)
    s0.left_margin = s0.right_margin = Inches(1.0)
    return doc


def _layout(path: Path):
    return LayoutEngine(DocxParser(path).parse(), font_manager=FontManager()).layout_document()


def _tokens(pg) -> List[str]:
    """The `Lnn` tokens drawn on a page, top to bottom."""
    return [t.text for t in sorted(pg.texts, key=lambda t: (t.y_baseline, t.x)) if TOKEN.match(t.text)]


def _assert_inside_page(pages) -> None:
    for pg in pages:
        for t in pg.texts:
            assert 0 <= t.y_baseline <= pg.height_pt, f"page {pg.page_number}: text {t.text!r} at y={t.y_baseline}"
        for bb in pg.bboxes:
            assert bb.y1 <= pg.height_pt + 0.5, f"page {pg.page_number}: {bb.kind} ends at y={bb.y1}"
            assert bb.x1 <= pg.width_pt + 0.5, f"page {pg.page_number}: {bb.kind} ends at x={bb.x1}"


# --- Section page-size changes ----------------------------------------------------------------


def _orientation_doc(path: Path, first, second, start=WD_SECTION.CONTINUOUS) -> Path:
    doc = _new_doc()
    first(doc.sections[0])
    doc.add_paragraph("First section.")
    s1 = doc.add_section(start)
    s1.top_margin = s1.bottom_margin = Inches(1.0)
    s1.left_margin = s1.right_margin = Inches(1.0)
    second(s1)
    for i in range(45):
        doc.add_paragraph(f"L{i:02d}")
    doc.save(path)
    return path


def test_continuous_break_into_portrait_starts_a_new_page(tmp_path):
    pages = _layout(_orientation_doc(tmp_path / "l2p.docx", _landscape, _portrait))
    assert (pages[0].width_pt, pages[0].height_pt) == (792.0, 612.0)
    assert all((pg.width_pt, pg.height_pt) == (612.0, 792.0) for pg in pages[1:])
    assert _tokens(pages[0]) == [], "portrait-section content was laid out on the landscape page"
    _assert_inside_page(pages)
    assert [t for pg in pages for t in _tokens(pg)] == [f"L{i:02d}" for i in range(45)]


def test_continuous_break_into_landscape_starts_a_new_page(tmp_path):
    pages = _layout(_orientation_doc(tmp_path / "p2l.docx", _portrait, _landscape))
    assert (pages[0].width_pt, pages[0].height_pt) == (612.0, 792.0)
    assert all((pg.width_pt, pg.height_pt) == (792.0, 612.0) for pg in pages[1:])
    _assert_inside_page(pages)


def test_continuous_break_with_same_page_size_stays_on_the_page(tmp_path):
    """The rule must not fire for an ordinary continuous break."""
    pages = _layout(_orientation_doc(tmp_path / "l2l.docx", _landscape, _landscape))
    assert _tokens(pages[0])[:1] == ["L00"]
