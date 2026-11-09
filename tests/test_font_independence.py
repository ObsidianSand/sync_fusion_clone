"""Layout must not depend on which fonts the machine has installed.

The renderer runs in MCP containers with no Microsoft fonts. Line heights used to come from
whatever font file was found, so there the built-in Helvetica/Times stood in for Calibri/Cambria
and every page shifted: the figure report pushed its Figure 3 heading off page 2, and the 88-page
sample came out at 91 pages with 7/118 anchors. Line heights now come from
`FAMILY_VERTICAL_METRICS`, recorded from the Microsoft fonts, whatever file is found.

`_NoFontsManager` simulates such a container by indexing no font files at all, so every family
falls through to PyMuPDF's built-in faces. On a machine that has no Microsoft fonts either, the
"with fonts" side of these comparisons is the same thing, and the tests still hold.

Also guards the failure that a workaround for this introduced: a table row cut through the middle
of an image, which then drew past the bottom of the page.
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import pytest

import word_truth as wt
from docx_renderer import DocxParser, FontManager, LayoutEngine
from docx_renderer.fonts import FAMILY_VERTICAL_METRICS

PROJECT = Path(__file__).resolve().parents[1]
SAMPLE = PROJECT / "NDA BLA XXXXXX - Safety Tables and Figures - 2 Columns RD.docx"
REPORT = PROJECT / "Figure_Comparison_Report_Redacted (1).docx"


class _NoFontsManager(FontManager):
    """A FontManager that finds no font files -- a bare Linux container."""

    def _discover_fonts(self, extra_dirs) -> None:
        pass


_layouts: Dict[tuple, list] = {}


def _layout(path: Path, no_fonts: bool) -> list:
    key = (path, no_fonts)
    if key not in _layouts:
        fm = _NoFontsManager() if no_fonts else FontManager()
        _layouts[key] = LayoutEngine(DocxParser(path).parse(), font_manager=fm).layout_document()
    return _layouts[key]


def _anchor_hits(path: Path, pages) -> int:
    page_of = wt.our_page_of_label(pages)
    return sum(1 for t in wt.pageref_targets(path) if t.target_label and page_of.get(t.target_label) == t.word_page)


def _body_overflow(pages) -> List[tuple]:
    """Body content (not headers/footers) that ends below the bottom margin.

    Checks the drawn images as well as the bboxes: splitting a table row clips the bboxes at the
    cut, so an image drawn whole across the cut shows up only in its draw command.
    """
    out = []
    for pg in pages:
        limit = pg.height_pt - pg.margin_bottom_pt + 0.5
        out += [(pg.page_number, b.kind, b.xml_path, round(b.y1, 1))
                for b in pg.bboxes if b.xml_path.startswith("word/document.xml:") and b.y1 > limit]
        out += [(pg.page_number, "drawn image", im.xml_path, round(im.y + im.height, 1))
                for im in pg.images
                if (im.xml_path or "").startswith("word/document.xml:") and im.y + im.height > limit]
    return out


@pytest.mark.parametrize("family", sorted(FAMILY_VERTICAL_METRICS))
@pytest.mark.parametrize("bold,italic", [(False, False), (True, False), (False, True), (True, True)])
def test_line_metrics_do_not_depend_on_installed_fonts(family, bold, italic):
    assert _NoFontsManager().get_metrics(family, 11.0, bold, italic) == FontManager().get_metrics(
        family, 11.0, bold, italic
    )


def test_unknown_family_without_a_file_uses_its_fallback_familys_metrics():
    fm = _NoFontsManager()
    assert fm.get_metrics("Some Missing Sans", 11.0) == fm.get_metrics("Calibri", 11.0)
    assert fm.get_metrics("Garamond", 11.0) == fm.get_metrics("Cambria", 11.0)
    assert fm.get_metrics("Lucida Console", 11.0) == fm.get_metrics("Consolas", 11.0)


@pytest.mark.skipif(not SAMPLE.exists(), reason=f"sample not present: {SAMPLE.name}")
def test_sample_paginates_like_word_without_any_fonts():
    pages = _layout(SAMPLE, no_fonts=True)
    assert len(pages) == wt.recorded_page_count(SAMPLE) == 88
    targets = wt.pageref_targets(SAMPLE)
    assert _anchor_hits(SAMPLE, pages) == len(targets) == 118


@pytest.mark.skipif(not REPORT.exists(), reason=f"sample not present: {REPORT.name}")
def test_figure_report_page_2_without_any_fonts():
    pages = _layout(REPORT, no_fonts=True)
    assert len(pages) == wt.recorded_page_count(REPORT) == 11
    page2 = [b.xml_path for b in pages[1].bboxes]
    assert any(p.endswith("/w:body/w:p[9]") for p in page2), "Figure 3 heading not on page 2"
    assert any("/w:body/w:tbl[3]/w:tr[1]/" in p for p in page2), "Figure 3 header row not on page 2"


@pytest.mark.parametrize("no_fonts", [False, True], ids=["installed-fonts", "no-fonts"])
@pytest.mark.parametrize("path", [SAMPLE, REPORT], ids=["sample", "figure-report"])
def test_nothing_in_the_body_is_drawn_past_the_bottom_margin(path, no_fonts):
    if not path.exists():
        pytest.skip(f"sample not present: {path.name}")
    assert _body_overflow(_layout(path, no_fonts)) == []
