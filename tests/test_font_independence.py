"""Layout must not depend on which fonts the machine has installed.

The renderer runs in MCP containers with no Microsoft fonts. Text used to be measured with
whatever font file was found, so there the built-in Helvetica/Times stood in for Calibri/Cambria
and every page shifted: the figure report pushed its Figure 3 heading off page 2, and the 88-page
sample came out at 91 pages with 7/118 anchors.

Text is now measured with metrics recorded from the Microsoft fonts
(`docx_renderer/data/font_metrics.json.gz`, written by `tools_record_font_metrics.py`): line
heights and per-character widths. So the guarantee here is exact, not approximate -- every word
on every page has the same position and width with and without fonts. When the real font is
missing, the substitute is drawn scaled to the recorded width, so no word runs past its space.

`_NoFontsManager` simulates such a container by indexing no font files at all, so every family
falls through to PyMuPDF's built-in faces. On a machine without the Microsoft fonts the "installed
fonts" side is the same thing, and these tests still hold.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Dict, List

import pymupdf as fitz
import pytest

import word_truth as wt
from docx_renderer import DocxParser, FontManager, LayoutEngine
from docx_renderer.fonts import FAMILY_VERTICAL_METRICS, METRICS_FILE
from docx_renderer.renderer import CanvasRenderer

PROJECT = Path(__file__).resolve().parents[1]
SAMPLE = PROJECT / "NDA BLA XXXXXX - Safety Tables and Figures - 2 Columns RD.docx"
REPORT = PROJECT / "Figure_Comparison_Report_Redacted (1).docx"
TOC = PROJECT / "Table of Contents template example.docx"

STYLES = [(False, False), (True, False), (False, True), (True, True)]
TEXT = "Section 3 — ADAVCG | Figure 1: L1: Avg Daily % Time <70 mg/dL | CGM-Stratified, “quoted” ½ ±"


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


def _words(pages) -> List[tuple]:
    return [(pg.page_number, t.text, t.x, t.y_baseline, t.width_pt) for pg in pages for t in pg.texts]


def _require(path: Path) -> None:
    if not path.exists():
        pytest.skip(f"sample not present: {path.name}")


# --- metrics ----------------------------------------------------------------------------------


def test_metrics_cover_the_known_families():
    assert METRICS_FILE.exists()
    assert {"calibri", "cambria", "arial", "times new roman"} <= set(FAMILY_VERTICAL_METRICS)


@pytest.mark.parametrize("family", sorted(FAMILY_VERTICAL_METRICS))
@pytest.mark.parametrize("bold,italic", STYLES)
def test_metrics_and_widths_do_not_depend_on_installed_fonts(family, bold, italic):
    have, bare = FontManager(), _NoFontsManager()
    assert bare.get_metrics(family, 11.0, bold, italic) == have.get_metrics(family, 11.0, bold, italic)
    assert bare.measure_text(TEXT, family, 11.0, bold, italic) == have.measure_text(TEXT, family, 11.0, bold, italic)


def test_unknown_family_without_a_file_uses_its_fallback_familys_metrics():
    fm = _NoFontsManager()
    for unknown, fallback in (("Some Missing Sans", "Calibri"), ("Garamond", "Cambria"), ("Lucida Console", "Consolas")):
        assert fm.get_metrics(unknown, 11.0) == fm.get_metrics(fallback, 11.0)
        assert fm.measure_text(TEXT, unknown, 11.0) == fm.measure_text(TEXT, fallback, 11.0)


def test_recorded_metrics_match_a_fresh_recording(tmp_path):
    """The data file must be what the generator writes from the real fonts -- it is checked in, so
    a hand edit or a stale copy would otherwise go unnoticed. Needs the Microsoft fonts."""
    sys.path.insert(0, str(PROJECT))
    import tools_record_font_metrics as rec

    out = tmp_path / "font_metrics.json.gz"
    if rec.main([str(out)]) != 0:
        pytest.skip("the Microsoft fonts are not installed here")
    assert out.read_bytes() == METRICS_FILE.read_bytes()


# --- layout -----------------------------------------------------------------------------------


@pytest.mark.parametrize("path", [SAMPLE, REPORT, TOC], ids=["sample", "figure-report", "toc"])
def test_every_word_lands_in_the_same_place_without_fonts(path):
    _require(path)
    have, bare = _words(_layout(path, False)), _words(_layout(path, True))
    assert len(have) == len(bare)
    first_diff = next((pair for pair in zip(have, bare) if pair[0] != pair[1]), None)
    assert first_diff is None, f"with fonts {first_diff[0]} != without {first_diff[1]}"


def test_sample_paginates_like_word_without_any_fonts():
    _require(SAMPLE)
    pages = _layout(SAMPLE, no_fonts=True)
    assert len(pages) == wt.recorded_page_count(SAMPLE) == 88
    page_of = wt.our_page_of_label(pages)
    targets = wt.pageref_targets(SAMPLE)
    hits = sum(1 for t in targets if t.target_label and page_of.get(t.target_label) == t.word_page)
    assert hits == len(targets) == 118


def test_figure_report_page_2_without_any_fonts():
    _require(REPORT)
    pages = _layout(REPORT, no_fonts=True)
    assert len(pages) == wt.recorded_page_count(REPORT) == 11
    page2 = [b.xml_path for b in pages[1].bboxes]
    assert any(p.endswith("/w:body/w:p[9]") for p in page2), "Figure 3 heading not on page 2"
    assert any("/w:body/w:tbl[3]/w:tr[1]/" in p for p in page2), "Figure 3 header row not on page 2"


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


@pytest.mark.parametrize("no_fonts", [False, True], ids=["installed-fonts", "no-fonts"])
@pytest.mark.parametrize("path", [SAMPLE, REPORT], ids=["sample", "figure-report"])
def test_nothing_in_the_body_is_drawn_past_the_bottom_margin(path, no_fonts):
    _require(path)
    assert _body_overflow(_layout(path, no_fonts)) == []


# --- drawing ----------------------------------------------------------------------------------


def _drawn_width(renderer: CanvasRenderer, tcmd) -> float:
    """Width of one text command as actually drawn: first glyph origin to last glyph's advance."""
    doc = fitz.open()
    page = doc.new_page(width=2000, height=200)
    renderer._draw_text_cmd(page, tcmd.__class__(**{**tcmd.__dict__, "x": 10.0, "y_baseline": 100.0}), {})
    chars = [c for b in page.get_text("rawdict")["blocks"] for ln in b.get("lines", [])
             for sp in ln["spans"] for c in sp["chars"]]
    return chars[-1]["bbox"][2] - chars[0]["origin"][0] if chars else 0.0


@pytest.mark.parametrize("path", [SAMPLE, REPORT], ids=["sample", "figure-report"])
def test_no_word_is_drawn_wider_than_layout_gave_it_without_fonts(path):
    """The substitute font is fitted to each word's recorded width. Unfitted, Helvetica standing in
    for Calibri overran by up to 10.6 pt on these pages, into the next word or a cell border."""
    _require(path)
    pages = _layout(path, no_fonts=True)[:4]
    renderer = CanvasRenderer(font_manager=_NoFontsManager())
    overruns = [(pg.page_number, t.text, round(_drawn_width(renderer, t) - t.width_pt, 3))
                for pg in pages for t in pg.texts if t.text.strip()]
    assert overruns and max(o[2] for o in overruns) <= 0.01, max(overruns, key=lambda o: o[2])
