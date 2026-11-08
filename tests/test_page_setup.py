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
from docx.oxml import OxmlElement
from docx.oxml.ns import qn as docx_qn
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


# --- Table rows split across pages -------------------------------------------------------------


def _set_row_flag(row, tag: str) -> None:
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement(tag))


def _tall_row_doc(path: Path, lead_paras: int = 0, n_lines: int = 60, cant_split: bool = False,
                  header: bool = False) -> Path:
    """Landscape page; optional lead-in paragraphs, then a table whose body row holds `n_lines`."""
    doc = _new_doc()
    _landscape(doc.sections[0])
    for i in range(lead_paras):
        doc.add_paragraph(f"lead {i}")
    table = doc.add_table(rows=2 if header else 1, cols=2)
    table.style = "Table Grid"
    if header:
        table.cell(0, 0).text = "HEAD"
        _set_row_flag(table.rows[0], "w:tblHeader")
    body = table.rows[-1]
    cell = body.cells[0]
    cell.text = "L00"
    for i in range(1, n_lines):
        cell.add_paragraph(f"L{i:02d}")
    body.cells[1].text = "side"
    if cant_split:
        _set_row_flag(body, "w:cantSplit")
    doc.save(path)
    return path


def test_row_taller_than_page_splits_without_losing_lines(tmp_path):
    pages = _layout(_tall_row_doc(tmp_path / "tall.docx"))
    assert len(pages) == 2
    _assert_inside_page(pages)
    # Every line once, in order, and each page's share ends inside its bottom margin.
    assert [t for pg in pages for t in _tokens(pg)] == [f"L{i:02d}" for i in range(60)]
    for pg in pages:
        assert max(t.y_baseline for t in pg.texts) <= pg.height_pt - pg.margin_bottom_pt
    # The first page is filled before the break, not left with a token amount.
    assert len(_tokens(pages[0])) > 30


def test_row_that_does_not_fit_splits_rather_than_moving(tmp_path):
    """Word's default: a row starting part-way down keeps the lines that fit on this page."""
    pages = _layout(_tall_row_doc(tmp_path / "mid.docx", lead_paras=15, n_lines=30))
    first, second = _tokens(pages[0]), _tokens(pages[1])
    assert first and first[0] == "L00"
    assert second and second[-1] == "L29"
    assert first + second == [f"L{i:02d}" for i in range(30)]


def test_cant_split_row_moves_whole_to_the_next_page(tmp_path):
    pages = _layout(_tall_row_doc(tmp_path / "cant.docx", lead_paras=15, n_lines=30, cant_split=True))
    assert _tokens(pages[0]) == []
    assert _tokens(pages[1]) == [f"L{i:02d}" for i in range(30)]


def test_unsplittable_row_taller_than_the_page_terminates_under_a_repeating_header(tmp_path):
    """A repeated header row means a fresh page is no longer "top of page"; the row loop once
    took that as a reason to move on again and started new pages forever."""
    import threading

    path = _tall_row_doc(tmp_path / "loop.docx", lead_paras=5, n_lines=60, cant_split=True, header=True)
    result: dict = {}
    worker = threading.Thread(target=lambda: result.setdefault("pages", _layout(path)), daemon=True)
    worker.start()
    worker.join(timeout=30)
    assert not worker.is_alive(), "layout did not terminate"
    # Header and lead-in on page 1; the unsplittable row moves once, then is placed regardless.
    assert len(result["pages"]) == 2


def test_header_row_repeats_above_the_continuation_of_a_split_row(tmp_path):
    pages = _layout(_tall_row_doc(tmp_path / "head.docx", n_lines=60, header=True))
    assert len(pages) == 2
    _assert_inside_page(pages)
    for pg in pages:
        ordered = [t.text for t in sorted(pg.texts, key=lambda t: (t.y_baseline, t.x))]
        assert ordered[0] == "HEAD", f"page {pg.page_number} starts with {ordered[:3]}"
    assert [t for pg in pages for t in _tokens(pg)] == [f"L{i:02d}" for i in range(60)]
