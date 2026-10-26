"""2D Canvas Renderer using PyMuPDF (fitz) to render PageLayout objects to PNGs."""

from __future__ import annotations
import json
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple, Union
import pymupdf as fitz

from .fonts import FontManager
from .layout import LayoutEngine
from .model import (
    DrawImageCmd,
    DrawLineCmd,
    DrawRectCmd,
    DrawTextCmd,
    PageLayout,
)
from .parser import DocxParser
from .units import hex_to_rgb_norm

PageSpec = Union[str, int, Iterable[int], None]


def parse_page_spec(spec: PageSpec, total_pages: int) -> List[int]:
    """Resolve a page selection to a sorted list of 1-based page numbers.

    `spec` may be None (every page), a single int, an iterable of ints, or a string of
    comma-separated numbers and ranges: ``"3"``, ``"1-4"``, ``"1,3,5-7"``, ``"8-"`` (page 8 to the
    end), ``"-3"`` (start through page 3). Pages outside the document are dropped.
    """
    if spec is None:
        return list(range(1, total_pages + 1))

    wanted: set = set()
    if isinstance(spec, int):
        wanted.add(spec)
    elif isinstance(spec, str):
        try:
            for part in spec.replace(" ", "").split(","):
                if not part:
                    continue
                if "-" in part:
                    lo_s, _, hi_s = part.partition("-")
                    lo = int(lo_s) if lo_s else 1
                    hi = int(hi_s) if hi_s else total_pages
                    if lo < 1 or hi < 1:
                        raise ValueError(part)
                    wanted.update(range(min(lo, hi), max(lo, hi) + 1))
                else:
                    wanted.add(int(part))
        except ValueError:
            raise ValueError(
                f"invalid page selection {spec!r}; expected e.g. '3', '1-4', '1,3,5-7' or '8-'"
            ) from None
    else:
        wanted.update(int(n) for n in spec)

    selected = sorted(n for n in wanted if 1 <= n <= total_pages)
    if not selected:
        raise ValueError(f"page selection {spec!r} matches no pages (document has {total_pages})")
    return selected


class CanvasRenderer:
    """Renders laid-out PageLayout objects to PNG files or bytes using PyMuPDF."""

    def __init__(self, font_manager: Optional[FontManager] = None) -> None:
        self.fm = font_manager or FontManager()

    def render_pages_to_pdf_doc(
        self,
        pages: List[PageLayout],
        show_ruler: bool = False,
        show_bboxes: bool = False,
    ) -> fitz.Document:
        """Render all PageLayout objects into an in-memory fitz.Document."""
        pdf_doc = fitz.open()
        for pg in pages:
            page = pdf_doc.new_page(width=pg.width_pt, height=pg.height_pt)
            self._draw_page(page, pg, show_ruler=show_ruler, show_bboxes=show_bboxes)
        return pdf_doc

    def render_to_png_files(
        self,
        pages: List[PageLayout],
        output_dir: str | Path,
        prefix: str = "page",
        dpi: int = 150,
        show_ruler: bool = False,
        show_bboxes: bool = False,
        export_layout_json: bool = False,
    ) -> List[Path]:
        """Render each page to a PNG file in `output_dir` and return the list of file paths.

        Files are named after each page's own `page_number`, so a subset of a document keeps the
        page numbers you asked for (`page_9.png` is page 9, not the first file written).
        """
        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        pdf_doc = self.render_pages_to_pdf_doc(pages, show_ruler=show_ruler, show_bboxes=show_bboxes)
        png_paths: List[Path] = []
        try:
            for layout_pg, page in zip(pages, pdf_doc):
                pix = page.get_pixmap(dpi=dpi, alpha=False)
                out_path = out_dir / f"{prefix}_{layout_pg.page_number}.png"
                pix.save(str(out_path))
                png_paths.append(out_path)
        finally:
            pdf_doc.close()

        if export_layout_json:
            json_path = out_dir / f"{prefix}_layout.json"
            self.save_layout_json(pages, json_path)

        return png_paths

    def render_to_png_bytes(
        self,
        pages: List[PageLayout],
        dpi: int = 150,
        show_ruler: bool = False,
        show_bboxes: bool = False,
    ) -> List[bytes]:
        """Render each page to PNG bytes in memory."""
        pdf_doc = self.render_pages_to_pdf_doc(pages, show_ruler=show_ruler, show_bboxes=show_bboxes)
        out_bytes: List[bytes] = []
        try:
            for page in pdf_doc:
                pix = page.get_pixmap(dpi=dpi, alpha=False)
                out_bytes.append(pix.tobytes("png"))
        finally:
            pdf_doc.close()
        return out_bytes

    def save_layout_json(self, pages: List[PageLayout], json_path: str | Path) -> None:
        """Export structured coordinate map of all elements per page for LLM grounding."""
        data = []
        for pg in pages:
            data.append(
                {
                    "page_number": pg.page_number,
                    "width_pt": round(pg.width_pt, 2),
                    "height_pt": round(pg.height_pt, 2),
                    "margins_pt": {
                        "top": round(pg.margin_top_pt, 2),
                        "bottom": round(pg.margin_bottom_pt, 2),
                        "left": round(pg.margin_left_pt, 2),
                        "right": round(pg.margin_right_pt, 2),
                    },
                    "elements": [
                        {
                            "kind": bb.kind,
                            "xml_path": bb.xml_path,
                            "style_id": bb.style_id,
                            "bbox_pt": [
                                round(bb.x0, 2),
                                round(bb.y0, 2),
                                round(bb.x1, 2),
                                round(bb.y1, 2),
                            ],
                            "indent_from_left_margin_pt": round(bb.x0 - pg.margin_left_pt, 2),
                            "indent_in_tabs_0_5in": round((bb.x0 - pg.margin_left_pt) / 36.0, 2),
                            "summary": bb.summary,
                        }
                        for bb in pg.bboxes
                    ],
                }
            )
        Path(json_path).write_text(json.dumps(data, indent=2), encoding="utf-8")

    # -------------------------------------------------------------------------
    # Internal Drawing Operations
    # -------------------------------------------------------------------------
    def _draw_page(
        self, page: fitz.Page, pg: PageLayout, show_ruler: bool = False, show_bboxes: bool = False
    ) -> None:
        # White page background
        page.draw_rect(fitz.Rect(0, 0, pg.width_pt, pg.height_pt), color=None, fill=(1.0, 1.0, 1.0))

        # 1. Background rectangles (z_order <= 0)
        for rcmd in sorted([r for r in pg.rects if r.z_order <= 0], key=lambda x: x.z_order):
            self._draw_rect_cmd(page, rcmd)

        # 2. Background images (z_order < 0)
        for imcmd in [im for im in pg.images if im.z_order < 0]:
            self._draw_image_cmd(page, imcmd)

        # 3. Foreground images (z_order >= 0)
        for imcmd in [im for im in pg.images if im.z_order >= 0]:
            self._draw_image_cmd(page, imcmd)

        # 4. Text runs (including run shading, highlight, glyphs, underlines, strikethroughs)
        registered_fonts: Dict[Tuple[str, bool, bool], str] = {}
        for tcmd in pg.texts:
            self._draw_text_cmd(page, tcmd, registered_fonts)

        # 5. Foreground rectangles (z_order > 0) and lines/borders
        for rcmd in [r for r in pg.rects if r.z_order > 0]:
            self._draw_rect_cmd(page, rcmd)

        for lncmd in pg.lines:
            self._draw_line_cmd(page, lncmd)

        # 6. Optional LLM Visual Inspection Guides (Ruler & Tab Grid / Bounding Boxes)
        if show_ruler:
            self._draw_llm_ruler(page, pg)
        if show_bboxes:
            self._draw_llm_bboxes(page, pg)

    def _draw_rect_cmd(self, page: fitz.Page, rcmd: DrawRectCmd) -> None:
        rect = fitz.Rect(rcmd.x, rcmd.y, rcmd.x + rcmd.width, rcmd.y + rcmd.height)
        fill_rgb = hex_to_rgb_norm(rcmd.fill_color)
        stroke_rgb = hex_to_rgb_norm(rcmd.stroke_color)
        page.draw_rect(
            rect,
            color=stroke_rgb,
            fill=fill_rgb,
            width=rcmd.stroke_width if stroke_rgb else 0.0,
        )

    def _draw_image_cmd(self, page: fitz.Page, imcmd: DrawImageCmd) -> None:
        if not imcmd.image_bytes:
            return
        rect = fitz.Rect(imcmd.x, imcmd.y, imcmd.x + imcmd.width, imcmd.y + imcmd.height)
        try:
            page.insert_image(rect, stream=imcmd.image_bytes, keep_proportion=False)
        except Exception:
            # Draw placeholder box if image stream is unsupported
            page.draw_rect(rect, color=(0.7, 0.7, 0.7), fill=(0.95, 0.95, 0.95), width=0.5)

    def _draw_line_cmd(self, page: fitz.Page, lncmd: DrawLineCmd) -> None:
        rgb = hex_to_rgb_norm(lncmd.color) or (0.0, 0.0, 0.0)
        p0 = fitz.Point(lncmd.x0, lncmd.y0)
        p1 = fitz.Point(lncmd.x1, lncmd.y1)
        w = max(0.25, lncmd.width)

        if lncmd.style == "double":
            offset = max(1.0, w * 1.1)
            if abs(lncmd.y1 - lncmd.y0) < 0.1:
                page.draw_line(
                    fitz.Point(lncmd.x0, lncmd.y0 - offset / 2),
                    fitz.Point(lncmd.x1, lncmd.y1 - offset / 2),
                    color=rgb,
                    width=w * 0.7,
                )
                page.draw_line(
                    fitz.Point(lncmd.x0, lncmd.y0 + offset / 2),
                    fitz.Point(lncmd.x1, lncmd.y1 + offset / 2),
                    color=rgb,
                    width=w * 0.7,
                )
            else:
                page.draw_line(
                    fitz.Point(lncmd.x0 - offset / 2, lncmd.y0),
                    fitz.Point(lncmd.x1 - offset / 2, lncmd.y1),
                    color=rgb,
                    width=w * 0.7,
                )
                page.draw_line(
                    fitz.Point(lncmd.x0 + offset / 2, lncmd.y0),
                    fitz.Point(lncmd.x1 + offset / 2, lncmd.y1),
                    color=rgb,
                    width=w * 0.7,
                )
            return

        dashes = None
        if lncmd.style in ("dashed", "dashSmallGap", "dotDash"):
            dashes = f"[{max(2, int(w * 3))} {max(1, int(w * 2))}] 0"
        elif lncmd.style == "dotted":
            dashes = f"[{max(1, int(w))} {max(1, int(w * 1.5))}] 0"

        page.draw_line(p0, p1, color=rgb, width=w, dashes=dashes)

    def _draw_text_cmd(
        self,
        page: fitz.Page,
        tcmd: DrawTextCmd,
        registered_fonts: Dict[Tuple[str, bool, bool], str],
    ) -> None:
        if not tcmd.text:
            return

        # 1. Run shading or highlight background
        bg_hex = tcmd.highlight or tcmd.shading_fill
        if bg_hex and tcmd.width_pt > 0:
            bg_rgb = hex_to_rgb_norm(bg_hex)
            if bg_rgb:
                rect = fitz.Rect(
                    tcmd.x,
                    tcmd.y_baseline - tcmd.ascender_pt,
                    tcmd.x + tcmd.width_pt,
                    tcmd.y_baseline + tcmd.descender_pt,
                )
                page.draw_rect(rect, color=None, fill=bg_rgb)

        # 2. Resolve and register font on page
        f_key = (tcmd.font_family.lower(), tcmd.bold, tcmd.italic)
        if f_key not in registered_fonts:
            _, fpath, bname = self.fm.get_font(tcmd.font_family, tcmd.bold, tcmd.italic)
            if fpath:
                alias = f"F_{len(registered_fonts)}"
                try:
                    page.insert_font(fontname=alias, fontfile=fpath)
                    registered_fonts[f_key] = alias
                except Exception:
                    registered_fonts[f_key] = bname
            else:
                registered_fonts[f_key] = bname

        font_alias = registered_fonts[f_key]
        txt_rgb = hex_to_rgb_norm(tcmd.color) or (0.0, 0.0, 0.0)

        if tcmd.text.strip():
            page.insert_text(
                fitz.Point(tcmd.x, tcmd.y_baseline),
                tcmd.text,
                fontsize=tcmd.font_size_pt,
                fontname=font_alias,
                color=txt_rgb,
            )

        # 3. Underline & Strikethrough decorations
        if tcmd.underline and tcmd.width_pt > 0:
            uy = tcmd.y_baseline + max(1.0, tcmd.font_size_pt * 0.10)
            uw = max(0.5, tcmd.font_size_pt * 0.05)
            ustyle = "double" if tcmd.underline == "double" else ("dotted" if "dot" in tcmd.underline else "single")
            self._draw_line_cmd(
                page,
                DrawLineCmd(
                    x0=tcmd.x,
                    y0=uy,
                    x1=tcmd.x + tcmd.width_pt,
                    y1=uy,
                    color=tcmd.color,
                    width=uw,
                    style=ustyle,
                ),
            )

        if (tcmd.strike or tcmd.dstrike) and tcmd.width_pt > 0:
            sy = tcmd.y_baseline - tcmd.ascender_pt * 0.35
            sw = max(0.5, tcmd.font_size_pt * 0.05)
            self._draw_line_cmd(
                page,
                DrawLineCmd(
                    x0=tcmd.x,
                    y0=sy,
                    x1=tcmd.x + tcmd.width_pt,
                    y1=sy,
                    color=tcmd.color,
                    width=sw,
                    style="double" if tcmd.dstrike else "single",
                ),
            )

    def _draw_llm_ruler(self, page: fitz.Page, pg: PageLayout) -> None:
        """Draw a subtle top horizontal tab/inch ruler and margin guide lines for LLM spatial awareness."""
        guide_rgb = (0.75, 0.82, 0.92)
        ruler_rgb = (0.35, 0.45, 0.65)
        ml = pg.margin_left_pt
        mr = pg.width_pt - pg.margin_right_pt
        mt = pg.margin_top_pt
        mb = pg.height_pt - pg.margin_bottom_pt

        # Margin boundary dashed box
        page.draw_rect(
            fitz.Rect(ml, mt, mr, mb),
            color=guide_rgb,
            width=0.4,
            dashes="[3 3] 0",
        )

        # Top ruler bar at y = 20pt showing 0.5-inch (1 standard tab = 36pt) ticks from left margin
        ruler_y = 20.0
        page.draw_line(fitz.Point(ml, ruler_y), fitz.Point(mr, ruler_y), color=ruler_rgb, width=0.6)
        tab_idx = 0
        curr_x = ml
        while curr_x <= mr + 0.5:
            is_inch = tab_idx % 2 == 0
            tick_h = 5.0 if is_inch else 3.0
            page.draw_line(
                fitz.Point(curr_x, ruler_y - tick_h),
                fitz.Point(curr_x, ruler_y + 2.0),
                color=ruler_rgb,
                width=0.6 if is_inch else 0.4,
            )
            # Label tab number (T1, T2, ...) and inch
            label = f"{tab_idx // 2}\"" if is_inch else f"T{tab_idx}"
            page.insert_text(
                fitz.Point(curr_x - 3.0, ruler_y - tick_h - 2.0),
                label,
                fontsize=6.0,
                fontname="helv",
                color=ruler_rgb,
            )
            # Subtle vertical tab grid line down the page
            if tab_idx > 0:
                page.draw_line(
                    fitz.Point(curr_x, mt),
                    fitz.Point(curr_x, mb),
                    color=(0.90, 0.93, 0.97),
                    width=0.3,
                    dashes="[1 4] 0",
                )
            curr_x += 36.0
            tab_idx += 1

    def _draw_llm_bboxes(self, page: fitz.Page, pg: PageLayout) -> None:
        """Draw color-coded bounding boxes around paragraphs, tables, and list markers."""
        colors = {
            "paragraph": (0.20, 0.50, 0.85),
            "table": (0.85, 0.35, 0.20),
            "table_cell": (0.90, 0.60, 0.30),
            "list_marker": (0.15, 0.70, 0.35),
            "image": (0.65, 0.25, 0.80),
        }
        for bb in pg.bboxes:
            if bb.kind not in colors:
                continue
            c = colors[bb.kind]
            page.draw_rect(
                fitz.Rect(bb.x0, bb.y0, bb.x1, bb.y1),
                color=c,
                width=0.4,
                dashes="[2 2] 0" if bb.kind == "paragraph" else None,
            )


def render_docx_to_png(
    docx_path: str | Path,
    output_dir: str | Path,
    prefix: str = "page",
    dpi: int = 150,
    show_ruler: bool = False,
    show_bboxes: bool = False,
    export_layout_json: bool = False,
    extra_font_dirs: Optional[List[str | Path]] = None,
    pages: PageSpec = None,
) -> List[Path]:
    """
    High-level convenience pipeline:
    Parse .docx -> Resolve Styles/Themes/Numbering -> Run Word Box/Flow Layout -> Render PNGs.

    `pages` limits which pages are rasterized -- ``"3"``, ``"1-4"``, ``"1,3,5-7"``, ``"8-"``, or an
    iterable of page numbers; None (the default) writes every page. The whole document is always
    parsed and laid out, so pagination, headers/footers and PAGE / NUMPAGES fields are identical to
    a full render; only the slow rasterization step is skipped for pages you did not ask for.
    """
    fm = FontManager(extra_font_dirs=extra_font_dirs)
    parsed_doc = DocxParser(docx_path).parse()
    laid_out = LayoutEngine(parsed_doc, font_manager=fm).layout_document()
    wanted = set(parse_page_spec(pages, len(laid_out)))
    selected = [pg for pg in laid_out if pg.page_number in wanted]
    renderer = CanvasRenderer(font_manager=fm)
    return renderer.render_to_png_files(
        selected,
        output_dir=output_dir,
        prefix=prefix,
        dpi=dpi,
        show_ruler=show_ruler,
        show_bboxes=show_bboxes,
        export_layout_json=export_layout_json,
    )
