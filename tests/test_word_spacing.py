"""Vertical spacing rules measured in Word, and the figure report whose page 2 they fixed.

Page 2 of `Figure_Comparison_Report_Redacted (1).docx` lost the Figure 3 heading and header row to
page 3: each heading + table block laid out ~20 pt taller than Word's, and two blocks plus the next
heading no longer fit. Measured against a screenshot of Word's page 2, four rules were off:

* Space between paragraphs is the *larger* of the first's space-after and the second's
  space-before, not their sum (Word: a 30pt-after / 30pt-before pair is 30 pt apart).
* A line is as tall as its content. The paragraph mark's font only sizes an empty line -- a 13pt
  Heading 2 style holding 10pt text is a 10pt line in Word.
* No space is added below a table.
* In Word 2010 compatibility mode and earlier, the paragraph mark after a closing page break stays
  on the old page; Word 2013+ moves it to the top of the new page as an empty line.
"""

from __future__ import annotations

from pathlib import Path

import docx
import pytest
from docx.oxml import OxmlElement
from docx.oxml.ns import qn as docx_qn
from docx.enum.text import WD_BREAK
from docx.shared import Inches, Pt

import word_truth as wt
from docx_renderer import DocxParser, FontManager, LayoutEngine

PROJECT = Path(__file__).resolve().parents[1]
REPORT = PROJECT / "Figure_Comparison_Report_Redacted (1).docx"


def _layout(path: Path):
    return LayoutEngine(DocxParser(path).parse(), font_manager=FontManager()).layout_document()


def _new_doc(compat_mode: int = 15) -> docx.Document:
    doc = docx.Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    pf = normal.paragraph_format
    pf.space_before = pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    s0 = doc.sections[0]
    s0.top_margin = s0.bottom_margin = s0.left_margin = s0.right_margin = Inches(1.0)
    settings = doc.settings.element
    compat = settings.find(docx_qn("w:compat"))
    if compat is None:
        compat = OxmlElement("w:compat")
        settings.append(compat)
    for cs in compat.findall(docx_qn("w:compatSetting")):
        if cs.get(docx_qn("w:name")) == "compatibilityMode":
            compat.remove(cs)
    cs = OxmlElement("w:compatSetting")
    cs.set(docx_qn("w:name"), "compatibilityMode")
    cs.set(docx_qn("w:uri"), "http://schemas.microsoft.com/office/word")
    cs.set(docx_qn("w:val"), str(compat_mode))
    compat.append(cs)
    return doc


def _baseline(pg, text: str) -> float:
    return next(t.y_baseline for t in pg.texts if t.text == text)


def _body_bbox(pg, child: str):
    """The bbox of a body-level block, e.g. `w:p[2]` -- never a cell's paragraph of the same name."""
    return next(b for b in pg.bboxes if b.xml_path.endswith(":/w:body/" + child))


def test_paragraph_gap_is_the_larger_spacing_not_the_sum(tmp_path):
    doc = _new_doc(14)
    doc.add_paragraph("ref1")
    doc.add_paragraph("ref2")
    a = doc.add_paragraph("upper")
    a.paragraph_format.space_after = Pt(30)
    b = doc.add_paragraph("lower")
    b.paragraph_format.space_before = Pt(30)
    doc.save(tmp_path / "gap.docx")
    pg = _layout(tmp_path / "gap.docx")[0]
    line = _baseline(pg, "ref2") - _baseline(pg, "ref1")
    assert _baseline(pg, "lower") - _baseline(pg, "upper") == pytest.approx(line + 30.0, abs=0.5)


def test_line_height_comes_from_the_text_not_the_paragraph_mark(tmp_path):
    doc = _new_doc()
    big = doc.styles.add_style("BigMark", 1)
    big.base_style = doc.styles["Normal"]
    big.font.size = Pt(26)
    big.paragraph_format.space_before = big.paragraph_format.space_after = Pt(0)
    small = doc.add_paragraph(style="BigMark")
    small.add_run("small").font.size = Pt(11)
    plain = doc.add_paragraph("plain")
    empty = doc.add_paragraph(style="BigMark")
    doc.add_paragraph("after")
    doc.save(tmp_path / "mark.docx")
    pg = _layout(tmp_path / "mark.docx")[0]
    h_small = _body_bbox(pg, "w:p[1]").y1 - _body_bbox(pg, "w:p[1]").y0
    h_plain = _body_bbox(pg, "w:p[2]").y1 - _body_bbox(pg, "w:p[2]").y0
    h_empty = _body_bbox(pg, "w:p[3]").y1 - _body_bbox(pg, "w:p[3]").y0
    assert h_small == pytest.approx(h_plain, abs=0.1)
    assert h_empty > 2 * h_plain, "an empty paragraph is sized by its 26pt mark"


def test_no_space_is_added_below_a_table(tmp_path):
    doc = _new_doc()
    doc.add_table(rows=1, cols=1).cell(0, 0).text = "cell"
    doc.add_paragraph("next")
    doc.save(tmp_path / "tbl.docx")
    pg = _layout(tmp_path / "tbl.docx")[0]
    assert _body_bbox(pg, "w:p[1]").y0 == pytest.approx(_body_bbox(pg, "w:tbl[1]").y1, abs=0.01)


def _page_break_doc(path: Path, compat_mode: int) -> Path:
    doc = _new_doc(compat_mode)
    doc.add_paragraph("first page")
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    nxt = doc.add_paragraph("second page")
    nxt.paragraph_format.space_before = Pt(10)
    doc.save(path)
    return path


def test_word_2010_keeps_the_mark_after_a_closing_page_break_on_the_old_page(tmp_path):
    pages = _layout(_page_break_doc(tmp_path / "pb14.docx", 14))
    assert len(pages) == 2
    first = min(pages[1].bboxes, key=lambda b: b.y0)
    assert first.xml_path.endswith(":/w:body/w:p[3]"), "page 2 should open with the paragraph after the break"
    # ... at the top margin: space-before is dropped at the top of the page.
    assert first.y0 == pytest.approx(pages[1].margin_top_pt, abs=0.01)


def test_word_2013_moves_the_mark_after_a_closing_page_break_to_the_new_page(tmp_path):
    pages = _layout(_page_break_doc(tmp_path / "pb15.docx", 15))
    assert len(pages) == 2
    first = min(pages[1].bboxes, key=lambda b: b.y0)
    assert first.xml_path.endswith(":/w:body/w:p[2]"), "page 2 should open with the break paragraph's mark"


@pytest.mark.skipif(not REPORT.exists(), reason=f"sample not present: {REPORT.name}")
def test_figure_report_paginates_like_word():
    """Word's <Pages> is 11, and its 10 page-start markers agree with that count, so they are
    current: every page must start where Word's does. Each page holds two figures and the next
    figure's heading and header row, with that figure's image row opening the following page."""
    pages = _layout(REPORT)
    assert len(pages) == wt.recorded_page_count(REPORT) == 11
    page2 = [b.xml_path for b in pages[1].bboxes]
    assert any(p.endswith("/w:body/w:p[9]") for p in page2), "Figure 3 heading not on page 2"
    assert any("/w:body/w:tbl[3]/w:tr[1]/" in p for p in page2), "Figure 3 header row not on page 2"
    page3_rows = [b.xml_path for b in pages[2].bboxes if "/w:body/w:tbl[3]/w:tr[" in b.xml_path]
    assert page3_rows and all("/w:tr[2]/" in p for p in page3_rows), "page 3 should open with the image row"
    page3_top = min(pages[2].bboxes, key=lambda b: b.y0)
    assert page3_top.xml_path.split(":", 1)[1].startswith("/w:body/w:tbl[3]")
