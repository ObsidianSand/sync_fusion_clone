"""MS Word Flow & Box Layout Engine (Twips/EMUs -> Points Canvas Coordinates)."""

from __future__ import annotations
import math
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Union

from .fonts import FontManager
from .model import (
    Borders,
    BorderSpec,
    DrawingItem,
    DrawImageCmd,
    DrawLineCmd,
    DrawRectCmd,
    DrawTextCmd,
    ElementBBox,
    HeaderFooterContent,
    PageLayout,
    ParagraphBlock,
    ParagraphStyle,
    ParsedDocument,
    RunItem,
    RunStyle,
    SectionProps,
    TableBlock,
    TableCell,
    TableRow,
    TabStop,
)


@dataclass
class _InlineAtom:
    """Smallest indivisible inline layout unit (word, whitespace, tab, inline drawing, break)."""
    kind: str                       # 'word', 'space', 'tab', 'drawing', 'line_break', 'page_break', 'column_break'
    text: str = ""
    width_pt: float = 0.0
    ascender_pt: float = 0.0
    descender_pt: float = 0.0
    line_height_1x_pt: float = 0.0
    style: RunStyle = field(default_factory=RunStyle)
    drawing: Optional[DrawingItem] = None
    xml_path: str = ""
    # Populated when kind == 'tab'
    tab_leader: str = "none"
    tab_align: str = "left"


@dataclass
class _LaidOutBox:
    """Laid-out block container (used for both main pages and table cells / textboxes)."""
    height_pt: float = 0.0
    rects: List[DrawRectCmd] = field(default_factory=list)
    lines: List[DrawLineCmd] = field(default_factory=list)
    images: List[DrawImageCmd] = field(default_factory=list)
    texts: List[DrawTextCmd] = field(default_factory=list)
    bboxes: List[ElementBBox] = field(default_factory=list)

    def offset_by(self, dx: float, dy: float) -> None:
        for r in self.rects:
            r.x += dx
            r.y += dy
        for ln in self.lines:
            ln.x0 += dx
            ln.x1 += dx
            ln.y0 += dy
            ln.y1 += dy
        for im in self.images:
            im.x += dx
            im.y += dy
        for tx in self.texts:
            tx.x += dx
            tx.y_baseline += dy
        for bb in self.bboxes:
            bb.x0 += dx
            bb.x1 += dx
            bb.y0 += dy
            bb.y1 += dy


@dataclass
class _ExclusionZone:
    """Floating object wrap exclusion zone on the current page."""
    x0: float
    y0: float
    x1: float
    y1: float
    wrap_type: str = "square"       # 'square', 'tight', 'topAndBottom'


class LayoutEngine:
    """Converts a ParsedDocument into a list of PageLayout objects with exact coordinates."""

    def __init__(self, doc: ParsedDocument, font_manager: Optional[FontManager] = None) -> None:
        self.doc = doc
        self.fm = font_manager or FontManager()
        self.pages: List[PageLayout] = []
        self.curr_sect: SectionProps = doc.final_sect_pr
        self.curr_page: Optional[PageLayout] = None
        self.cursor_y: float = 0.0
        self.prev_space_after_pt: float = 0.0
        self.prev_style_id: Optional[str] = None
        self.prev_contextual_spacing: bool = False
        #: True at the top of the current *flow area* -- a new page, or a new column on it.
        self.is_top_of_page: bool = True
        self.exclusions: List[_ExclusionZone] = []
        self.sect_first_page: bool = True
        #: (x0, width) per text column of the current section, and which one we are filling.
        self.col_bounds: List[Tuple[float, float]] = self.curr_sect.column_bounds()
        self.col_index: int = 0
        #: y at which the current column set starts on this page. Usually the top margin, but a
        #: continuous section break opens a column set part-way down, and every column of that set
        #: starts level with it rather than at the top of the page.
        self.col_top_y: float = self.curr_sect.margin_top_pt
        #: How far down each column of the current set content reached, for balancing.
        self.col_bottoms: List[float] = [self.curr_sect.margin_top_pt]
        #: While balancing, the y no column on `_balance_page_idx` may run past.
        self._balance_cap: Optional[float] = None
        self._balance_page_idx: Optional[int] = None

    def layout_document(self) -> List[PageLayout]:
        """Run full pagination and layout pass over the document."""
        # Group blocks by sections (paragraphs with sect_pr end a section)
        sections: List[Tuple[SectionProps, List[Union[ParagraphBlock, TableBlock]]]] = []
        curr_blocks: List[Union[ParagraphBlock, TableBlock]] = []
        for blk in self.doc.blocks:
            curr_blocks.append(blk)
            if isinstance(blk, ParagraphBlock) and blk.sect_pr is not None:
                sections.append((blk.sect_pr, curr_blocks))
                curr_blocks = []
        sections.append((self.doc.final_sect_pr, curr_blocks))

        for s_idx, (sect_pr, blocks) in enumerate(sections):
            self.curr_sect = sect_pr
            self.sect_first_page = True
            if s_idx == 0 or sect_pr.break_type != "continuous":
                self._new_page()
            else:
                # A continuous break keeps the page but starts a fresh column set on it, which is
                # how a document switches between one and two columns mid-page. The space-after
                # still pending from the previous section is consumed here, by the boundary, so
                # that every column of the new set starts level; leaving it pending would indent
                # only whichever column happened to be filled first.
                #
                # Resume below the *tallest* column of the set being closed. The cursor sits at the
                # bottom of whichever column was filled last, which is usually the shortest -- and
                # after balancing, reliably so -- and starting there would run the new section back
                # up over the bottom of the first column.
                self.cursor_y = max([self.cursor_y] + self.col_bottoms)
                self.cursor_y += self.prev_space_after_pt
                self.prev_space_after_pt = 0.0
                self.col_bounds = sect_pr.column_bounds()
                self.col_index = 0
                self.col_top_y = self.cursor_y
                self.col_bottoms = [self.cursor_y] * len(self.col_bounds)
            # Word balances a multi-column section's last page when the section after it begins
            # with a continuous break. Checkpoint each block so that set can be re-flowed.
            next_sect = sections[s_idx + 1][0] if s_idx + 1 < len(sections) else None
            balancing = (
                sect_pr.cols_num > 1
                and next_sect is not None
                and next_sect.break_type == "continuous"
            )
            marks: List[Tuple[int, dict]] = []

            for b_idx, blk in enumerate(blocks):
                if balancing:
                    marks.append((b_idx, self._checkpoint()))
                next_blk = blocks[b_idx + 1] if b_idx + 1 < len(blocks) else None
                if isinstance(blk, ParagraphBlock):
                    self._layout_paragraph_on_page(blk, next_blk)
                elif isinstance(blk, TableBlock):
                    self._layout_table_on_page(blk)

            if balancing:
                self._balance_final_column_set(blocks, marks)

        if not self.pages:
            self._new_page()

        # Render headers & footers now that total page count is known
        total_pages = len(self.pages)
        self._render_all_headers_footers(total_pages)
        return self.pages

    # -------------------------------------------------------------------------
    # Page Management
    # -------------------------------------------------------------------------
    def _new_page(self) -> PageLayout:
        pg_num = len(self.pages) + 1
        sp = self.curr_sect
        pg = PageLayout(
            page_number=pg_num,
            width_pt=sp.page_width_pt,
            height_pt=sp.page_height_pt,
            margin_top_pt=sp.margin_top_pt,
            margin_bottom_pt=sp.margin_bottom_pt,
            margin_left_pt=sp.margin_left_pt,
            margin_right_pt=sp.margin_right_pt,
        )
        # Attach section metadata for post-pass header/footer rendering
        setattr(pg, "_sect_pr", sp)
        setattr(pg, "_is_first_of_sect", self.sect_first_page)
        self.sect_first_page = False
        self.pages.append(pg)
        self.curr_page = pg
        self.col_bounds = sp.column_bounds()
        self.col_index = 0
        self.col_top_y = sp.margin_top_pt
        self.col_bottoms = [sp.margin_top_pt] * len(self.col_bounds)
        self.cursor_y = sp.margin_top_pt
        self.prev_space_after_pt = 0.0
        self.prev_style_id = None
        self.prev_contextual_spacing = False
        self.is_top_of_page = True
        self.exclusions = []
        return pg

    def _advance_flow(self) -> None:
        """Out of vertical room: start the next column, or a new page after the last one.

        This is the "ran out of space" move, as distinct from `_new_page()`, which is the explicit
        one (a manual page break, `w:pageBreakBefore`, a section start). In a single-column section
        the two are the same thing.
        """
        if self.col_index + 1 >= len(self.col_bounds):
            self._new_page()
            return
        self.col_bottoms[self.col_index] = max(self.col_bottoms[self.col_index], self.cursor_y)
        self.col_index += 1
        self.cursor_y = self.col_top_y
        self.prev_space_after_pt = 0.0
        self.prev_style_id = None
        self.prev_contextual_spacing = False
        self.is_top_of_page = True
        # Floating objects are anchored to the page, not to a column, so their exclusion zones
        # still apply to the column we are moving into. Deliberately not cleared.

    # -------------------------------------------------------------------------
    # Column Balancing
    # -------------------------------------------------------------------------
    def _checkpoint(self) -> dict:
        """Snapshot enough state to undo everything laid out from this point on.

        Layout only ever appends to a page's primitive lists, so rewinding is truncation.
        """
        pg = self.curr_page
        assert pg is not None
        return {
            "n_pages": len(self.pages),
            "page": pg,
            "lens": (len(pg.rects), len(pg.lines), len(pg.images), len(pg.texts), len(pg.bboxes)),
            "cursor_y": self.cursor_y,
            "col_index": self.col_index,
            "col_top_y": self.col_top_y,
            "col_bottoms": list(self.col_bottoms),
            "col_bounds": list(self.col_bounds),
            "prev_space_after_pt": self.prev_space_after_pt,
            "prev_style_id": self.prev_style_id,
            "prev_contextual_spacing": self.prev_contextual_spacing,
            "is_top_of_page": self.is_top_of_page,
            "exclusions": list(self.exclusions),
            "sect_first_page": self.sect_first_page,
        }

    def _rewind_to(self, cp: dict) -> None:
        """Undo layout back to a `_checkpoint()`."""
        del self.pages[cp["n_pages"]:]
        pg = cp["page"]
        n_rects, n_lines, n_images, n_texts, n_bboxes = cp["lens"]
        del pg.rects[n_rects:]
        del pg.lines[n_lines:]
        del pg.images[n_images:]
        del pg.texts[n_texts:]
        del pg.bboxes[n_bboxes:]
        self.curr_page = pg
        self.cursor_y = cp["cursor_y"]
        self.col_index = cp["col_index"]
        self.col_top_y = cp["col_top_y"]
        self.col_bottoms = list(cp["col_bottoms"])
        self.col_bounds = list(cp["col_bounds"])
        self.prev_space_after_pt = cp["prev_space_after_pt"]
        self.prev_style_id = cp["prev_style_id"]
        self.prev_contextual_spacing = cp["prev_contextual_spacing"]
        self.is_top_of_page = cp["is_top_of_page"]
        self.exclusions = list(cp["exclusions"])
        self.sect_first_page = cp["sect_first_page"]

    def _lay_out_blocks(self, blocks: List[Union[ParagraphBlock, TableBlock]], start: int = 0) -> None:
        for i in range(start, len(blocks)):
            blk = blocks[i]
            nxt = blocks[i + 1] if i + 1 < len(blocks) else None
            if isinstance(blk, ParagraphBlock):
                self._layout_paragraph_on_page(blk, nxt)
            elif isinstance(blk, TableBlock):
                self._layout_table_on_page(blk)

    def _balance_final_column_set(
        self,
        blocks: List[Union[ParagraphBlock, TableBlock]],
        marks: List[Tuple[int, dict]],
    ) -> None:
        """Even out the columns on the last page of a multi-column section.

        Word fills columns to the bottom margin, except on the final page of a section that is
        followed by a continuous break -- there it evens the columns up. That trailing continuous
        break is exactly what the familiar "insert a continuous section break to balance columns"
        advice adds to a document.

        Balancing is done by re-flowing the content of that last column set under a cap on how far
        down any column may run, then searching for the shortest cap the content still fits under.
        A cap of exactly `used / columns` is the ideal, but lines are indivisible, so the search
        walks back up from there until the content fits.
        """
        n_cols = len(self.col_bounds)
        if n_cols < 2 or self.curr_page is None:
            return

        final_page_idx = len(self.pages) - 1
        top = self.col_top_y
        bottoms = list(self.col_bottoms)
        bottoms[self.col_index] = max(bottoms[self.col_index], self.cursor_y)
        used = sum(max(0.0, b - top) for b in bottoms)
        if used <= 0.0:
            return

        # Pick the earliest block the final column set can be re-flowed from. Best case there is a
        # block that began the set at the top of its first column. Otherwise fall back to the block
        # that crossed onto this page: it is checkpointed on the page before, so re-laying it
        # reproduces the break and only its remainder lands under the cap.
        crossed = None
        clean = None
        for b_idx, cp in marks:
            if cp["n_pages"] - 1 < final_page_idx:
                crossed = (b_idx, cp)
            elif clean is None and (
                cp["col_index"] == 0
                and abs(cp["cursor_y"] - top) < 0.01
                and abs(cp["col_top_y"] - top) < 0.01
            ):
                clean = (b_idx, cp)
        start = clean or crossed
        if start is None:
            return

        b_idx, cp = start
        full = self.page_bottom_raw - top
        if full <= 0.0:
            return

        def fits(cap_height: float) -> bool:
            self._rewind_to(cp)
            self._balance_cap = top + cap_height
            self._balance_page_idx = final_page_idx
            try:
                self._lay_out_blocks(blocks, b_idx)
            finally:
                self._balance_cap = None
                self._balance_page_idx = None
            return len(self.pages) - 1 == final_page_idx

        lo, hi = used / n_cols, full
        best = full                      # the unbalanced layout always fits, by construction
        for _ in range(12):
            if hi - lo < 0.5:
                break
            mid = (lo + hi) / 2.0
            if fits(mid):
                best, hi = mid, mid
            else:
                lo = mid
        fits(best)

    @property
    def page_bottom_limit(self) -> float:
        limit = self.page_bottom_raw
        if self._balance_cap is not None and len(self.pages) - 1 == self._balance_page_idx:
            return min(limit, self._balance_cap)
        return limit

    @property
    def page_bottom_raw(self) -> float:
        """The bottom margin, ignoring any balancing cap in force."""
        return self.curr_sect.page_height_pt - self.curr_sect.margin_bottom_pt

    @property
    def col_x(self) -> float:
        """Left edge of the column currently being filled."""
        return self.col_bounds[self.col_index][0]

    @property
    def col_w(self) -> float:
        """Width of the column currently being filled."""
        return self.col_bounds[self.col_index][1]

    def _avail_line_bounds(self, y_top: float, y_bottom: float, base_x0: float, base_x1: float) -> Tuple[float, float, Optional[ float]]:
        """
        Check floating exclusion zones overlapping [y_top, y_bottom].
        Returns (adjusted_x0, adjusted_x1, jump_to_y_if_blocked).
        """
        x0, x1 = base_x0, base_x1
        for ex in self.exclusions:
            if ex.y1 <= y_top or ex.y0 >= y_bottom:
                continue
            if ex.wrap_type == "topAndBottom":
                return x0, x1, ex.y1
            # Square/tight wrap: narrow left or right margin
            mid_ex = (ex.x0 + ex.x1) / 2.0
            mid_col = (base_x0 + base_x1) / 2.0
            if mid_ex <= mid_col:
                x0 = max(x0, ex.x1)
            else:
                x1 = min(x1, ex.x0)
            if x1 - x0 < 36.0:
                return base_x0, base_x1, ex.y1
        return x0, x1, None

    # -------------------------------------------------------------------------
    # Paragraph Layout (Main Page Flow)
    # -------------------------------------------------------------------------
    def _layout_paragraph_on_page(
        self, p: ParagraphBlock, next_blk: Optional[Union[ParagraphBlock, TableBlock]] = None
    ) -> None:
        assert self.curr_page is not None
        ps = p.style

        if ps.page_break_before and not self.is_top_of_page:
            self._new_page()

        # Process floating anchored drawings attached to this paragraph
        for r in p.runs:
            if r.kind == "drawing" and r.drawing is not None and r.drawing.is_floating:
                self._place_floating_drawing(r.drawing, self.curr_page, self.cursor_y)

        # Compute vertical gap before paragraph
        space_before = ps.space_before_pt if ps.space_before_pt is not None else 0.0
        space_after = ps.space_after_pt if ps.space_after_pt is not None else 0.0

        if self.is_top_of_page:
            gap_above = 0.0
        elif (
            ps.contextual_spacing
            and self.prev_contextual_spacing
            and p.style_id is not None
            and p.style_id == self.prev_style_id
        ):
            gap_above = 0.0
        else:
            gap_above = self.prev_space_after_pt + space_before

        col_x = self.col_x
        col_w = self.col_w

        # Build lines for this paragraph
        lines_data = self._break_paragraph_into_lines(
            p,
            col_x=col_x,
            col_w=col_w,
            page_num=self.curr_page.page_number,
            total_pages=max(1, len(self.pages)),
            use_exclusions=True,
            start_y=self.cursor_y + gap_above,
        )

        # If keep_lines or keep_next, check if whole paragraph fits on current page
        total_p_height = sum(ld["line_height"] for ld in lines_data)
        if not self.is_top_of_page and lines_data:
            need_h = gap_above + total_p_height
            if ps.keep_next and next_blk is not None:
                next_min_h = 24.0
                if isinstance(next_blk, ParagraphBlock):
                    for nr in next_blk.runs:
                        if nr.kind == "drawing" and nr.drawing is not None and not nr.drawing.is_floating:
                            next_min_h = max(next_min_h, nr.drawing.height_pt + 12.0)
                need_h += next_min_h
            if (ps.keep_lines or ps.keep_next) and (self.cursor_y + need_h > self.page_bottom_limit):
                self._advance_flow()
                gap_above = 0.0
                col_x, col_w = self.col_x, self.col_w
                lines_data = self._break_paragraph_into_lines(
                    p,
                    col_x=col_x,
                    col_w=col_w,
                    page_num=self.curr_page.page_number,
                    total_pages=max(1, len(self.pages)),
                    use_exclusions=True,
                    start_y=self.cursor_y,
                )

        self.cursor_y += gap_above
        p_start_y = self.cursor_y
        p_start_page = self.curr_page

        for l_idx, ld in enumerate(lines_data):
            if ld.get("explicit_column_break_before"):
                self._advance_flow()
                p_start_y = self.cursor_y
                p_start_page = self.curr_page
                col_x, col_w = self.col_x, self.col_w

            if ld.get("explicit_page_break_before"):
                self._new_page()
                p_start_y = self.cursor_y
                p_start_page = self.curr_page
                col_x, col_w = self.col_x, self.col_w

            lh = ld["line_height"]
            if not self.is_top_of_page and (self.cursor_y + lh > self.page_bottom_limit):
                self._advance_flow()
                p_start_y = self.cursor_y
                p_start_page = self.curr_page
                col_x, col_w = self.col_x, self.col_w

            y_top = self.cursor_y
            # If floating exclusion pushed this line down, sync cursor_y
            if ld.get("y_override") is not None and ld["y_override"] > y_top and self.curr_page == p_start_page:
                y_top = ld["y_override"]
                self.cursor_y = y_top

            self._emit_line_to_target(
                self.curr_page,
                p,
                ld,
                y_top=y_top,
                col_x=col_x,
                col_w=col_w,
                is_first_line=(l_idx == 0),
                is_last_line=(l_idx == len(lines_data) - 1),
            )
            self.cursor_y += lh
            self.is_top_of_page = False

            if ld.get("explicit_page_break_after"):
                self._new_page()
                p_start_y = self.cursor_y
                p_start_page = self.curr_page
                col_x, col_w = self.col_x, self.col_w

            if ld.get("explicit_column_break_after"):
                self._advance_flow()
                p_start_y = self.cursor_y
                p_start_page = self.curr_page
                col_x, col_w = self.col_x, self.col_w

        # Record bounding box for the paragraph on its current page
        assert self.curr_page is not None
        ind_l = ps.indent_left_pt or 0.0
        ind_r = ps.indent_right_pt or 0.0
        p_end_y = self.cursor_y
        if p_end_y > p_start_y:
            summary_txt = "".join(r.text for r in p.runs if r.kind == "text")[:80]
            self.curr_page.bboxes.append(
                ElementBBox(
                    kind="paragraph",
                    xml_path=p.xml_path,
                    x0=col_x + ind_l + min(0.0, ps.indent_first_line_pt or 0.0),
                    y0=p_start_y,
                    x1=col_x + col_w - ind_r,
                    y1=p_end_y,
                    style_id=p.style_id,
                    summary=summary_txt,
                )
            )

        self.prev_space_after_pt = space_after
        self.prev_style_id = p.style_id
        self.prev_contextual_spacing = bool(ps.contextual_spacing)

    # -------------------------------------------------------------------------
    # Sub-Box Layout (For Table Cells, Textboxes, Headers, Footers)
    # -------------------------------------------------------------------------
    def _layout_blocks_in_box(
        self,
        blocks: List[Union[ParagraphBlock, TableBlock]],
        box_width_pt: float,
        page_num: int = 1,
        total_pages: int = 1,
        in_table_cell: bool = False,
    ) -> _LaidOutBox:
        box = _LaidOutBox()
        curr_y = 0.0
        prev_after = 0.0
        prev_style: Optional[str] = None
        prev_ctx = False

        for b_idx, blk in enumerate(blocks):
            if isinstance(blk, ParagraphBlock):
                ps = blk.style
                space_before = ps.space_before_pt if ps.space_before_pt is not None else 0.0
                space_after = ps.space_after_pt if ps.space_after_pt is not None else (0.0 if in_table_cell else 8.0)

                if b_idx == 0:
                    gap = 0.0 if in_table_cell else space_before
                elif ps.contextual_spacing and prev_ctx and blk.style_id and blk.style_id == prev_style:
                    gap = 0.0
                else:
                    gap = prev_after + space_before

                curr_y += gap
                p_y0 = curr_y
                lines_data = self._break_paragraph_into_lines(
                    blk,
                    col_x=0.0,
                    col_w=box_width_pt,
                    page_num=page_num,
                    total_pages=total_pages,
                    use_exclusions=False,
                    start_y=curr_y,
                )
                for l_idx, ld in enumerate(lines_data):
                    lh = ld["line_height"]
                    self._emit_line_to_target(
                        box,
                        blk,
                        ld,
                        y_top=curr_y,
                        col_x=0.0,
                        col_w=box_width_pt,
                        is_first_line=(l_idx == 0),
                        is_last_line=(l_idx == len(lines_data) - 1),
                    )
                    curr_y += lh

                if curr_y > p_y0:
                    summary_txt = "".join(r.text for r in blk.runs if r.kind == "text")[:60]
                    box.bboxes.append(
                        ElementBBox(
                            kind="paragraph",
                            xml_path=blk.xml_path,
                            x0=ps.indent_left_pt or 0.0,
                            y0=p_y0,
                            x1=box_width_pt - (ps.indent_right_pt or 0.0),
                            y1=curr_y,
                            style_id=blk.style_id,
                            summary=summary_txt,
                        )
                    )
                prev_after = space_after
                prev_style = blk.style_id
                prev_ctx = bool(ps.contextual_spacing)

            elif isinstance(blk, TableBlock):
                tbl_box = self._layout_table_to_box(blk, box_width_pt, page_num, total_pages)
                tbl_box.offset_by(0.0, curr_y)
                box.rects.extend(tbl_box.rects)
                box.lines.extend(tbl_box.lines)
                box.images.extend(tbl_box.images)
                box.texts.extend(tbl_box.texts)
                box.bboxes.extend(tbl_box.bboxes)
                curr_y += tbl_box.height_pt
                prev_after = 0.0

        box.height_pt = curr_y
        return box

    # -------------------------------------------------------------------------
    # Core Line-Breaking, Indentation & Tab-Stop Engine
    # -------------------------------------------------------------------------
    def _tokenize_paragraph(
        self, p: ParagraphBlock, page_num: int, total_pages: int
    ) -> List[_InlineAtom]:
        atoms: List[_InlineAtom] = []
        base_rs = p.style.run_style

        for r in p.runs:
            rs = base_rs.merge_from(r.style)
            fam = rs.font_family or "Calibri"
            fsize = rs.font_size_pt or 11.0
            if rs.vert_align in ("superscript", "subscript"):
                fsize *= 0.65
            elif rs.small_caps:
                fsize *= 0.85
            bold = bool(rs.bold)
            italic = bool(rs.italic)
            cspace = rs.char_spacing_pt or 0.0
            asc, desc, lh1 = self.fm.get_metrics(fam, fsize, bold, italic)

            if r.kind == "page_num":
                txt = str(page_num)
                w = self.fm.measure_text(txt, fam, fsize, bold, italic, cspace)
                atoms.append(
                    _InlineAtom(
                        kind="word",
                        text=txt,
                        width_pt=w,
                        ascender_pt=asc,
                        descender_pt=desc,
                        line_height_1x_pt=lh1,
                        style=rs,
                        xml_path=r.xml_path,
                    )
                )
            elif r.kind == "num_pages":
                txt = str(total_pages)
                w = self.fm.measure_text(txt, fam, fsize, bold, italic, cspace)
                atoms.append(
                    _InlineAtom(
                        kind="word",
                        text=txt,
                        width_pt=w,
                        ascender_pt=asc,
                        descender_pt=desc,
                        line_height_1x_pt=lh1,
                        style=rs,
                        xml_path=r.xml_path,
                    )
                )
            elif r.kind == "tab":
                atoms.append(
                    _InlineAtom(
                        kind="tab",
                        text="\t",
                        width_pt=0.0,
                        ascender_pt=asc,
                        descender_pt=desc,
                        line_height_1x_pt=lh1,
                        style=rs,
                        xml_path=r.xml_path,
                    )
                )
            elif r.kind == "break":
                if r.break_type == "page":
                    bkind = "page_break"
                elif r.break_type == "column":
                    bkind = "column_break"
                else:
                    bkind = "line_break"
                atoms.append(
                    _InlineAtom(
                        kind=bkind,
                        ascender_pt=asc,
                        descender_pt=desc,
                        line_height_1x_pt=lh1,
                        style=rs,
                        xml_path=r.xml_path,
                    )
                )
            elif r.kind == "drawing" and r.drawing is not None and not r.drawing.is_floating:
                drw = r.drawing
                atoms.append(
                    _InlineAtom(
                        kind="drawing",
                        width_pt=drw.width_pt,
                        ascender_pt=drw.height_pt,
                        descender_pt=0.0,
                        line_height_1x_pt=drw.height_pt,
                        style=rs,
                        drawing=drw,
                        xml_path=r.xml_path,
                    )
                )
            elif r.kind == "text" and r.text:
                # Split into alternating word and space tokens while preserving exact spaces
                buf = ""
                in_space: Optional[bool] = None
                for ch in r.text:
                    is_sp = ch == " "
                    if in_space is None:
                        buf = ch
                        in_space = is_sp
                    elif is_sp == in_space:
                        buf += ch
                    else:
                        w = self.fm.measure_text(buf, fam, fsize, bold, italic, cspace)
                        atoms.append(
                            _InlineAtom(
                                kind="space" if in_space else "word",
                                text=buf,
                                width_pt=w,
                                ascender_pt=asc,
                                descender_pt=desc,
                                line_height_1x_pt=lh1,
                                style=rs,
                                xml_path=r.xml_path,
                            )
                        )
                        buf = ch
                        in_space = is_sp
                if buf:
                    w = self.fm.measure_text(buf, fam, fsize, bold, italic, cspace)
                    atoms.append(
                        _InlineAtom(
                            kind="space" if in_space else "word",
                            text=buf,
                            width_pt=w,
                            ascender_pt=asc,
                            descender_pt=desc,
                            line_height_1x_pt=lh1,
                            style=rs,
                            xml_path=r.xml_path,
                        )
                    )
        return atoms

    def _resolve_next_tab_stop(
        self, rel_x: float, custom_tabs: List[TabStop], right_edge_rel: float
    ) -> TabStop:
        """Find the next tab stop strictly after `rel_x` (relative to paragraph column start)."""
        for t in custom_tabs:
            if t.val != "bar" and t.pos_pt > rel_x + 0.5:
                return t
        # Fallback to Word's default tab interval (36pt = 0.5 inch)
        dts = max(9.0, self.doc.default_tab_stop_pt)
        next_pos = (math.floor((rel_x + 0.5) / dts) + 1) * dts
        return TabStop(pos_pt=min(next_pos, max(next_pos, right_edge_rel)), val="left", leader="none")

    def _break_paragraph_into_lines(
        self,
        p: ParagraphBlock,
        col_x: float,
        col_w: float,
        page_num: int,
        total_pages: int,
        use_exclusions: bool = False,
        start_y: float = 0.0,
    ) -> List[dict]:
        ps = p.style
        base_rs = ps.run_style
        fam = base_rs.font_family or "Calibri"
        fsize = base_rs.font_size_pt or 11.0
        asc_def, desc_def, lh1_def = self.fm.get_metrics(fam, fsize, bool(base_rs.bold), bool(base_rs.italic))

        ind_left = ps.indent_left_pt or 0.0
        ind_right = ps.indent_right_pt or 0.0
        ind_fl = ps.indent_first_line_pt or 0.0

        atoms = self._tokenize_paragraph(p, page_num, total_pages)

        # Empty paragraph still occupies one blank line in MS Word
        if not atoms and not p.list_marker:
            lh = self._compute_line_height(ps, lh1_def, asc_def + desc_def)
            return [
                {
                    "segments": [],
                    "ascender": asc_def,
                    "descender": desc_def,
                    "line_height": lh,
                    "rel_x_start": ind_left + ind_fl,
                    "rel_x_end": max(ind_left + ind_fl + 10.0, col_w - ind_right),
                    "list_marker_info": None,
                }
            ]

        # Prepare list marker info for line 0 if present
        marker_info = None
        line0_text_rel_x = ind_left + ind_fl
        if p.list_marker:
            mrs = p.list_marker_style or base_rs
            mfam = mrs.font_family or fam
            msize = mrs.font_size_pt or fsize
            mw = self.fm.measure_text(p.list_marker, mfam, msize, bool(mrs.bold), bool(mrs.italic))
            marker_rel_x = ind_left + ind_fl
            tab_target = p.list_tab_pt if p.list_tab_pt is not None else ind_left
            line0_text_rel_x = max(tab_target, marker_rel_x + mw + 4.0)
            marker_info = {
                "text": p.list_marker,
                "style": mrs,
                "rel_x": marker_rel_x,
                "width": mw,
            }

        lines: List[dict] = []
        idx = 0
        n_atoms = len(atoms)
        curr_y = start_y

        while idx < n_atoms or (not lines and marker_info is not None):
            is_first = len(lines) == 0
            rel_start = line0_text_rel_x if is_first else ind_left
            rel_end = max(rel_start + 36.0, col_w - ind_right)
            y_override: Optional[float] = None

            if use_exclusions and self.exclusions:
                adj_x0, adj_x1, jump_y = self._avail_line_bounds(
                    curr_y, curr_y + lh1_def, col_x + rel_start, col_x + rel_end
                )
                while jump_y is not None:
                    curr_y = jump_y
                    y_override = curr_y
                    adj_x0, adj_x1, jump_y = self._avail_line_bounds(
                        curr_y, curr_y + lh1_def, col_x + rel_start, col_x + rel_end
                    )
                rel_start = max(rel_start, adj_x0 - col_x)
                rel_end = min(rel_end, adj_x1 - col_x)

            # Segment-based horizontal layout on this line:
            # Each segment is separated by a <w:tab/> so right/center/decimal tabs and leaders work accurately
            segments: List[dict] = []
            curr_seg_atoms: List[_InlineAtom] = []
            curr_seg_start_x = rel_start
            curr_seg_align = "left"
            curr_seg_target_x = rel_start
            curr_seg_leader = "none"
            curr_seg_prev_end_x = rel_start

            max_asc = asc_def
            max_desc = desc_def
            max_lh1 = lh1_def
            explicit_pb_before = False
            explicit_pb_after = False
            explicit_cb_before = False
            explicit_cb_after = False

            def flush_segment() -> float:
                """Finalize current segment's horizontal start_x and return its right edge."""
                nonlocal curr_seg_atoms
                # Trim trailing spaces at end of segment before a tab or line wrap
                content_w = sum(a.width_pt for a in curr_seg_atoms)
                if curr_seg_align == "right":
                    seg_x = max(curr_seg_prev_end_x, curr_seg_target_x - content_w)
                elif curr_seg_align == "center":
                    seg_x = max(curr_seg_prev_end_x, curr_seg_target_x - content_w / 2.0)
                elif curr_seg_align == "decimal":
                    dec_w = 0.0
                    for a in curr_seg_atoms:
                        if "." in a.text:
                            prefix = a.text.split(".", 1)[0]
                            dec_w += self.fm.measure_text(
                                prefix,
                                a.style.font_family or fam,
                                a.style.font_size_pt or fsize,
                                bool(a.style.bold),
                                bool(a.style.italic),
                            )
                            break
                        dec_w += a.width_pt
                    seg_x = max(curr_seg_prev_end_x, curr_seg_target_x - dec_w)
                else:
                    seg_x = curr_seg_start_x

                segments.append(
                    {
                        "x": seg_x,
                        "width": content_w,
                        "atoms": list(curr_seg_atoms),
                        "leader": curr_seg_leader,
                        "leader_x0": curr_seg_prev_end_x,
                        "leader_x1": seg_x,
                        "align": curr_seg_align,
                    }
                )
                return seg_x + content_w

            while idx < n_atoms:
                atom = atoms[idx]
                if atom.kind == "page_break":
                    idx += 1
                    if not curr_seg_atoms and not segments:
                        explicit_pb_before = True
                    else:
                        explicit_pb_after = True
                        break
                    continue

                if atom.kind == "column_break":
                    idx += 1
                    if not curr_seg_atoms and not segments:
                        explicit_cb_before = True
                    else:
                        explicit_cb_after = True
                        break
                    continue

                if atom.kind == "line_break":
                    idx += 1
                    break

                if atom.kind == "tab":
                    seg_end_x = flush_segment()
                    curr_seg_atoms = []
                    tstop = self._resolve_next_tab_stop(seg_end_x, ps.tabs, rel_end)
                    curr_seg_prev_end_x = seg_end_x
                    curr_seg_leader = tstop.leader
                    curr_seg_align = tstop.val
                    curr_seg_target_x = tstop.pos_pt
                    curr_seg_start_x = tstop.pos_pt
                    idx += 1
                    continue

                # Skip leading spaces at the very start of a wrapped line
                if atom.kind == "space" and not curr_seg_atoms and not segments and not is_first:
                    idx += 1
                    continue

                # Estimate if adding `atom` overflows the line's right boundary. Word measures tab
                # stops from the left margin and honors them even when they sit outside the right
                # indent, so a segment anchored to such a stop may legitimately reach past
                # `rel_end` (TOC entries do this: the right dot-leader stop frequently clears the
                # entry style's own right indent). Only the page column itself is a hard limit.
                tentative_w = sum(a.width_pt for a in curr_seg_atoms) + atom.width_pt
                seg_rel_end = rel_end
                if curr_seg_align == "left":
                    projected_right = curr_seg_start_x + tentative_w
                elif curr_seg_align == "right":
                    projected_right = max(curr_seg_target_x, curr_seg_prev_end_x + tentative_w)
                    seg_rel_end = max(rel_end, min(curr_seg_target_x, col_w))
                else:
                    projected_right = curr_seg_start_x + tentative_w
                    seg_rel_end = max(rel_end, min(curr_seg_target_x, col_w))

                if projected_right > seg_rel_end + 0.5 and (curr_seg_atoms or segments):
                    if atom.kind == "space":
                        idx += 1
                    break

                curr_seg_atoms.append(atom)
                max_asc = max(max_asc, atom.ascender_pt)
                max_desc = max(max_desc, atom.descender_pt)
                max_lh1 = max(max_lh1, atom.line_height_1x_pt)
                idx += 1

            if curr_seg_atoms or not segments:
                flush_segment()

            lh = self._compute_line_height(ps, max_lh1, max_asc + max_desc)
            lines.append(
                {
                    "segments": segments,
                    "ascender": max_asc,
                    "descender": max_desc,
                    "line_height": lh,
                    "rel_x_start": rel_start,
                    "rel_x_end": rel_end,
                    "list_marker_info": marker_info if is_first else None,
                    "explicit_page_break_before": explicit_pb_before,
                    "explicit_page_break_after": explicit_pb_after,
                    "explicit_column_break_before": explicit_cb_before,
                    "explicit_column_break_after": explicit_cb_after,
                    "y_override": y_override,
                }
            )
            curr_y += lh
            if not atoms:
                break

        return lines

    def _compute_line_height(self, ps: ParagraphStyle, max_lh1: float, raw_glyph_h: float) -> float:
        lrule = ps.line_rule or "auto"
        lsp = ps.line_spacing if ps.line_spacing is not None else 1.0
        if lrule == "exact":
            return max(1.0, lsp)
        if lrule == "atLeast":
            return max(raw_glyph_h, lsp)
        # 'auto' multiplier (where 1.0 = single line height)
        return max(raw_glyph_h, max_lh1 * lsp)

    # -------------------------------------------------------------------------
    # Emitting a Laid-out Line to Page or Box
    # -------------------------------------------------------------------------
    def _emit_line_to_target(
        self,
        target: Union[PageLayout, _LaidOutBox],
        p: ParagraphBlock,
        ld: dict,
        y_top: float,
        col_x: float,
        col_w: float,
        is_first_line: bool,
        is_last_line: bool,
    ) -> None:
        ps = p.style
        lh = ld["line_height"]
        asc = ld["ascender"]
        desc = ld["descender"]
        y_baseline = y_top + asc

        ind_l = ps.indent_left_pt or 0.0
        ind_r = ps.indent_right_pt or 0.0
        rel_start = ld["rel_x_start"]
        rel_end = ld["rel_x_end"]
        avail_w = max(1.0, rel_end - rel_start)

        # 1. Paragraph background shading
        if ps.shading_fill:
            target.rects.append(
                DrawRectCmd(
                    x=col_x + ind_l,
                    y=y_top,
                    width=max(1.0, col_w - ind_l - ind_r),
                    height=lh,
                    fill_color=ps.shading_fill,
                    z_order=-1,
                    xml_path=p.xml_path,
                )
            )

        # 2. Paragraph borders (top on first line, bottom on last line, left/right on all lines)
        if ps.borders is not None:
            bx0 = col_x + ind_l
            bx1 = col_x + col_w - ind_r
            if is_first_line and ps.borders.top.is_visible:
                target.lines.append(
                    DrawLineCmd(
                        x0=bx0,
                        y0=y_top,
                        x1=bx1,
                        y1=y_top,
                        color=ps.borders.top.color or "000000",
                        width=ps.borders.top.width_pt,
                        style=ps.borders.top.val,
                        xml_path=p.xml_path,
                    )
                )
            if is_last_line and ps.borders.bottom.is_visible:
                by = y_top + lh + ps.borders.bottom.space_pt
                target.lines.append(
                    DrawLineCmd(
                        x0=bx0,
                        y0=by,
                        x1=bx1,
                        y1=by,
                        color=ps.borders.bottom.color or "000000",
                        width=ps.borders.bottom.width_pt,
                        style=ps.borders.bottom.val,
                        xml_path=p.xml_path,
                    )
                )
            if ps.borders.left.is_visible:
                target.lines.append(
                    DrawLineCmd(
                        x0=bx0,
                        y0=y_top,
                        x1=bx0,
                        y1=y_top + lh,
                        color=ps.borders.left.color or "000000",
                        width=ps.borders.left.width_pt,
                        style=ps.borders.left.val,
                        xml_path=p.xml_path,
                    )
                )
            if ps.borders.right.is_visible:
                target.lines.append(
                    DrawLineCmd(
                        x0=bx1,
                        y0=y_top,
                        x1=bx1,
                        y1=y_top + lh,
                        color=ps.borders.right.color or "000000",
                        width=ps.borders.right.width_pt,
                        style=ps.borders.right.val,
                        xml_path=p.xml_path,
                    )
                )

        # 3. List marker on line 0
        minfo = ld.get("list_marker_info")
        if minfo is not None:
            mrs: RunStyle = minfo["style"]
            mfam = mrs.font_family or "Calibri"
            msize = mrs.font_size_pt or 11.0
            masc, mdesc, _ = self.fm.get_metrics(mfam, msize, bool(mrs.bold), bool(mrs.italic))
            mx = col_x + minfo["rel_x"]
            target.texts.append(
                DrawTextCmd(
                    x=mx,
                    y_baseline=y_baseline,
                    text=minfo["text"],
                    font_family=mfam,
                    font_size_pt=msize,
                    bold=bool(mrs.bold),
                    italic=bool(mrs.italic),
                    color=mrs.color or "000000",
                    width_pt=minfo["width"],
                    ascender_pt=masc,
                    descender_pt=mdesc,
                    xml_path=p.xml_path,
                )
            )
            target.bboxes.append(
                ElementBBox(
                    kind="list_marker",
                    xml_path=p.xml_path,
                    x0=mx,
                    y0=y_baseline - masc,
                    x1=mx + minfo["width"],
                    y1=y_baseline + mdesc,
                    style_id=p.style_id,
                    summary=minfo["text"],
                )
            )

        segments: List[dict] = ld["segments"]
        if not segments:
            return

        # 4. Paragraph alignment offset (when there is a single segment without custom tabs)
        align = ps.alignment or "left"
        has_tabs = len(segments) > 1
        for seg_idx, seg in enumerate(segments):
            atoms: List[_InlineAtom] = seg["atoms"]
            # Trim trailing space atoms for alignment & justification math
            vis_atoms = list(atoms)
            while vis_atoms and vis_atoms[-1].kind == "space":
                vis_atoms.pop()
            seg_vis_w = sum(a.width_pt for a in vis_atoms)

            # Render tab leader (e.g., dot leader '........' in TOCs) if present
            leader_type = seg.get("leader", "none")
            lx0 = col_x + seg.get("leader_x0", 0.0) + 2.0
            lx1 = col_x + seg.get("leader_x1", 0.0) - 2.0
            if leader_type not in ("none", "") and lx1 - lx0 > 4.0:
                l_rs = atoms[0].style if atoms else ps.run_style
                lfam = l_rs.font_family or "Calibri"
                lsize = l_rs.font_size_pt or 11.0
                lch = "." if leader_type in ("dot", "middleDot") else ("-" if leader_type == "hyphen" else "_")
                step_str = lch
                step_w = max(1.5, self.fm.measure_text(step_str, lfam, lsize, False, False))
                count = max(1, int((lx1 - lx0) / step_w))
                leader_str = step_str * count
                target.texts.append(
                    DrawTextCmd(
                        x=lx0,
                        y_baseline=y_baseline,
                        text=leader_str,
                        font_family=lfam,
                        font_size_pt=lsize,
                        bold=False,
                        italic=False,
                        color=l_rs.color or "000000",
                        width_pt=lx1 - lx0,
                        ascender_pt=asc,
                        descender_pt=desc,
                        xml_path=p.xml_path,
                    )
                )
                target.bboxes.append(
                    ElementBBox(
                        kind="tab_leader",
                        xml_path=p.xml_path,
                        x0=lx0,
                        y0=y_baseline - asc * 0.5,
                        x1=lx1,
                        y1=y_baseline,
                        style_id=p.style_id,
                        summary=leader_type,
                    )
                )

            if not vis_atoms:
                continue

            # Determine starting x and space expansion (for justified text)
            extra_space_pt = 0.0
            if not has_tabs and seg_idx == 0:
                if align == "center":
                    curr_x = col_x + rel_start + max(0.0, (avail_w - seg_vis_w) / 2.0)
                elif align == "right":
                    curr_x = col_x + rel_end - seg_vis_w
                elif align == "both" and not is_last_line:
                    curr_x = col_x + rel_start
                    num_spaces = sum(1 for a in vis_atoms if a.kind == "space")
                    if num_spaces > 0 and avail_w > seg_vis_w:
                        extra_space_pt = (avail_w - seg_vis_w) / num_spaces
                else:
                    curr_x = col_x + seg["x"]
            else:
                curr_x = col_x + seg["x"]

            for a in vis_atoms:
                if a.kind == "drawing" and a.drawing is not None:
                    drw = a.drawing
                    img_y = y_baseline - drw.height_pt
                    self._emit_drawing_primitives(target, drw, curr_x, img_y)
                    curr_x += drw.width_pt
                else:
                    rs = a.style
                    fam = rs.font_family or "Calibri"
                    fsize = rs.font_size_pt or 11.0
                    by_offset = 0.0
                    if rs.vert_align == "superscript":
                        by_offset = -fsize * 0.45
                        fsize *= 0.65
                    elif rs.vert_align == "subscript":
                        by_offset = fsize * 0.20
                        fsize *= 0.65
                    elif rs.small_caps:
                        fsize *= 0.85

                    w = a.width_pt + (extra_space_pt if a.kind == "space" else 0.0)
                    target.texts.append(
                        DrawTextCmd(
                            x=curr_x,
                            y_baseline=y_baseline + by_offset,
                            text=a.text,
                            font_family=fam,
                            font_size_pt=fsize,
                            bold=bool(rs.bold),
                            italic=bool(rs.italic),
                            color=rs.color or "000000",
                            underline=rs.underline,
                            strike=bool(rs.strike),
                            dstrike=bool(rs.dstrike),
                            highlight=rs.highlight,
                            shading_fill=rs.shading_fill,
                            width_pt=w,
                            ascender_pt=a.ascender_pt,
                            descender_pt=a.descender_pt,
                            xml_path=a.xml_path,
                        )
                    )
                    curr_x += w

    # -------------------------------------------------------------------------
    # Drawings (Inline & Floating Figures / Shapes / Textboxes)
    # -------------------------------------------------------------------------
    def _emit_drawing_primitives(
        self, target: Union[PageLayout, _LaidOutBox], drw: DrawingItem, x: float, y: float
    ) -> None:
        z = -1 if drw.behind_doc else 0
        if drw.shape_fill or drw.shape_stroke:
            target.rects.append(
                DrawRectCmd(
                    x=x,
                    y=y,
                    width=drw.width_pt,
                    height=drw.height_pt,
                    fill_color=drw.shape_fill,
                    stroke_color=drw.shape_stroke,
                    stroke_width=drw.shape_stroke_pt or (0.75 if drw.shape_stroke else 0.0),
                    z_order=z,
                    xml_path=drw.xml_path,
                )
            )
        if drw.image_bytes:
            target.images.append(
                DrawImageCmd(
                    x=x,
                    y=y,
                    width=drw.width_pt,
                    height=drw.height_pt,
                    image_bytes=drw.image_bytes,
                    z_order=z,
                    xml_path=drw.xml_path,
                )
            )
        if drw.textbox_blocks:
            pad = 5.4
            tb_box = self._layout_blocks_in_box(
                drw.textbox_blocks, max(18.0, drw.width_pt - 2 * pad), in_table_cell=True
            )
            tb_box.offset_by(x + pad, y + pad)
            target.rects.extend(tb_box.rects)
            target.lines.extend(tb_box.lines)
            target.images.extend(tb_box.images)
            target.texts.extend(tb_box.texts)
            target.bboxes.extend(tb_box.bboxes)

        target.bboxes.append(
            ElementBBox(
                kind="image",
                xml_path=drw.xml_path,
                x0=x,
                y0=y,
                x1=x + drw.width_pt,
                y1=y + drw.height_pt,
                summary=f"{drw.width_pt:.1f}x{drw.height_pt:.1f}pt",
            )
        )

    def _place_floating_drawing(self, drw: DrawingItem, page: PageLayout, para_y: float) -> None:
        sp = self.curr_sect
        # Resolve Horizontal X
        if drw.pos_h_relative_from == "page":
            ref_x0, ref_w = 0.0, sp.page_width_pt
        elif drw.pos_h_relative_from == "column":
            # "column" means the text column, which is the whole text area only when there is one.
            ref_x0, ref_w = self.col_x, self.col_w
        else:
            ref_x0, ref_w = sp.margin_left_pt, sp.content_width_pt

        if drw.pos_h_align == "center":
            x = ref_x0 + (ref_w - drw.width_pt) / 2.0
        elif drw.pos_h_align == "right":
            x = ref_x0 + ref_w - drw.width_pt
        elif drw.pos_h_offset_pt is not None:
            x = ref_x0 + drw.pos_h_offset_pt
        else:
            x = ref_x0

        # Resolve Vertical Y
        if drw.pos_v_relative_from == "page":
            ref_y0, ref_h = 0.0, sp.page_height_pt
        elif drw.pos_v_relative_from == "margin":
            ref_y0, ref_h = sp.margin_top_pt, sp.page_height_pt - sp.margin_top_pt - sp.margin_bottom_pt
        else:
            ref_y0, ref_h = para_y, sp.page_height_pt - para_y - sp.margin_bottom_pt

        if drw.pos_v_align == "center":
            y = ref_y0 + (ref_h - drw.height_pt) / 2.0
        elif drw.pos_v_align == "bottom":
            y = ref_y0 + ref_h - drw.height_pt
        elif drw.pos_v_offset_pt is not None:
            y = ref_y0 + drw.pos_v_offset_pt
        else:
            y = ref_y0

        # Clamp within page bounds
        x = max(0.0, min(sp.page_width_pt - drw.width_pt, x))
        y = max(0.0, min(sp.page_height_pt - drw.height_pt, y))

        self._emit_drawing_primitives(page, drw, x, y)

        if drw.wrap_type in ("square", "tight", "topAndBottom"):
            self.exclusions.append(
                _ExclusionZone(
                    x0=x - drw.dist_l_pt,
                    y0=y - drw.dist_t_pt,
                    x1=x + drw.width_pt + drw.dist_r_pt,
                    y1=y + drw.height_pt + drw.dist_b_pt,
                    wrap_type=drw.wrap_type,
                )
            )

    # -------------------------------------------------------------------------
    # Table Layout Engine (Grid Resolution, Merged Cells, Borders, Pagination)
    # -------------------------------------------------------------------------
    def _resolve_table_column_widths(self, tb: TableBlock, avail_width_pt: float) -> List[float]:
        # 1. Determine total number of grid columns across all rows
        max_cols = len(tb.grid_cols_pt)
        for row in tb.rows:
            c_count = row.grid_before + sum(max(1, c.grid_span) for c in row.cells)
            max_cols = max(max_cols, c_count)
        if max_cols == 0:
            return []

        col_widths = list(tb.grid_cols_pt) + [0.0] * max(0, max_cols - len(tb.grid_cols_pt))

        # 2. Fill missing or zero column widths from single-span cells
        for row in tb.rows:
            g_idx = row.grid_before
            for cell in row.cells:
                span = max(1, cell.grid_span)
                if span == 1 and g_idx < max_cols:
                    if cell.width_type == "dxa" and cell.width_pt and cell.width_pt > col_widths[g_idx]:
                        col_widths[g_idx] = cell.width_pt
                    elif cell.width_type == "pct" and cell.width_pct:
                        w_est = avail_width_pt * cell.width_pct
                        if w_est > col_widths[g_idx]:
                            col_widths[g_idx] = w_est
                g_idx += span

        # 3. For any columns still 0.0, distribute remaining width evenly
        zero_indices = [i for i, w in enumerate(col_widths) if w <= 1.0]
        known_sum = sum(w for w in col_widths if w > 1.0)
        target_total = avail_width_pt
        if tb.width_type == "dxa" and tb.width_pt and tb.width_pt > 0:
            target_total = tb.width_pt
        elif tb.width_type == "pct" and tb.width_pct:
            target_total = avail_width_pt * tb.width_pct

        if zero_indices:
            rem = max(36.0 * len(zero_indices), target_total - known_sum)
            each = rem / len(zero_indices)
            for i in zero_indices:
                col_widths[i] = each

        # 4. Scale if table has explicit total width or overflows container width
        total_w = sum(col_widths)
        if total_w > 0:
            if (tb.width_type in ("dxa", "pct") and abs(total_w - target_total) > 2.0) or (
                total_w > avail_width_pt + 2.0
            ):
                scale_target = min(avail_width_pt, target_total) if total_w > avail_width_pt else target_total
                scale = scale_target / total_w
                col_widths = [w * scale for w in col_widths]

        return col_widths

    def _prepare_laid_out_rows(
        self, tb: TableBlock, col_widths: List[float], page_num: int, total_pages: int
    ) -> List[dict]:
        """Measure and lay out all cells and rows, including horizontal and vertical merges."""
        num_rows = len(tb.rows)
        num_cols = len(col_widths)
        row_records: List[dict] = []

        # First pass: lay out each cell's inner box and record grid positions
        grid_matrix: List[List[Optional[dict]]] = [[None] * num_cols for _ in range(num_rows)]

        for r_idx, row in enumerate(tb.rows):
            g_col = row.grid_before
            cell_entries: List[dict] = []
            for cell in row.cells:
                span = min(max(1, cell.grid_span), max(1, num_cols - g_col))
                cell_w = sum(col_widths[g_col : g_col + span])
                x_off = sum(col_widths[:g_col])

                pad_t = cell.margin_top_pt if cell.margin_top_pt is not None else tb.cell_margin_top_pt
                pad_b = cell.margin_bottom_pt if cell.margin_bottom_pt is not None else tb.cell_margin_bottom_pt
                pad_l = cell.margin_left_pt if cell.margin_left_pt is not None else tb.cell_margin_left_pt
                pad_r = cell.margin_right_pt if cell.margin_right_pt is not None else tb.cell_margin_right_pt

                inner_w = max(8.0, cell_w - pad_l - pad_r)
                if cell.v_merge == "continue":
                    inner_box = _LaidOutBox(height_pt=0.0)
                else:
                    inner_box = self._layout_blocks_in_box(
                        cell.blocks, inner_w, page_num=page_num, total_pages=total_pages, in_table_cell=True
                    )

                entry = {
                    "cell": cell,
                    "r_idx": r_idx,
                    "g_col": g_col,
                    "span": span,
                    "x_off": x_off,
                    "width": cell_w,
                    "pad_t": pad_t,
                    "pad_b": pad_b,
                    "pad_l": pad_l,
                    "pad_r": pad_r,
                    "inner_box": inner_box,
                    "req_height": inner_box.height_pt + pad_t + pad_b,
                    "v_span_rows": 1,
                }
                cell_entries.append(entry)
                for c in range(g_col, min(num_cols, g_col + span)):
                    grid_matrix[r_idx][c] = entry
                g_col += span

            # Base row height from non-vertically-merged cells. Word sizes a row to its tallest
            # cell and imposes no minimum of its own -- `w:trHeight` below is the only floor. A
            # fixed minimum here silently pads every row (a single line of 10pt Arial measures
            # 11.5pt), which compounds down a long table into whole extra pages.
            content_heights = [e["req_height"] for e in cell_entries if e["cell"].v_merge is None]
            natural_h = max(content_heights) if content_heights else 0.0
            if row.height_rule == "exact" and row.height_pt > 0:
                row_h = row.height_pt
            elif row.height_pt > 0:
                row_h = max(natural_h, row.height_pt)
            else:
                row_h = natural_h

            row_records.append({"row": row, "height": row_h, "cells": cell_entries})

        # Second pass: resolve vertical merges (vMerge="restart" followed by vMerge="continue")
        for r_idx in range(num_rows):
            for entry in row_records[r_idx]["cells"]:
                if entry["cell"].v_merge == "restart":
                    g_col = entry["g_col"]
                    span_rows = 1
                    for next_r in range(r_idx + 1, num_rows):
                        below = grid_matrix[next_r][g_col]
                        if below is not None and below["cell"].v_merge == "continue":
                            span_rows += 1
                        else:
                            break
                    entry["v_span_rows"] = span_rows
                    curr_span_h = sum(row_records[rr]["height"] for rr in range(r_idx, r_idx + span_rows))
                    if entry["req_height"] > curr_span_h:
                        row_records[r_idx + span_rows - 1]["height"] += entry["req_height"] - curr_span_h

        return row_records

    def _emit_table_row(
        self,
        target: Union[PageLayout, _LaidOutBox],
        tb: TableBlock,
        row_rec: dict,
        all_row_recs: List[dict],
        table_x: float,
        row_y: float,
    ) -> None:
        row_h = row_rec["height"]
        for entry in row_rec["cells"]:
            cell: TableCell = entry["cell"]
            if cell.v_merge == "continue":
                continue

            r_idx = entry["r_idx"]
            v_span = entry["v_span_rows"]
            cell_h = sum(all_row_recs[rr]["height"] for rr in range(r_idx, r_idx + v_span))
            cx = table_x + entry["x_off"]
            cw = entry["width"]

            # 1. Cell background shading
            if cell.shading_fill:
                target.rects.append(
                    DrawRectCmd(
                        x=cx,
                        y=row_y,
                        width=cw,
                        height=cell_h,
                        fill_color=cell.shading_fill,
                        z_order=-1,
                        xml_path=cell.xml_path,
                    )
                )

            # 2. Vertically align cell inner box
            ibox: _LaidOutBox = entry["inner_box"]
            pad_t = entry["pad_t"]
            pad_b = entry["pad_b"]
            pad_l = entry["pad_l"]
            avail_inner_h = max(0.0, cell_h - pad_t - pad_b)

            if cell.v_align == "center":
                content_y = row_y + pad_t + max(0.0, (avail_inner_h - ibox.height_pt) / 2.0)
            elif cell.v_align == "bottom":
                content_y = row_y + cell_h - pad_b - ibox.height_pt
            else:
                content_y = row_y + pad_t

            # Clone or offset inner box primitives
            ibox_copy = _LaidOutBox(
                height_pt=ibox.height_pt,
                rects=[DrawRectCmd(**r.__dict__) for r in ibox.rects],
                lines=[DrawLineCmd(**ln.__dict__) for ln in ibox.lines],
                images=[DrawImageCmd(**im.__dict__) for im in ibox.images],
                texts=[DrawTextCmd(**tx.__dict__) for tx in ibox.texts],
                bboxes=[ElementBBox(**bb.__dict__) for bb in ibox.bboxes],
            )
            ibox_copy.offset_by(cx + pad_l, content_y)
            target.rects.extend(ibox_copy.rects)
            target.lines.extend(ibox_copy.lines)
            target.images.extend(ibox_copy.images)
            target.texts.extend(ibox_copy.texts)
            target.bboxes.extend(ibox_copy.bboxes)

            # 3. Cell borders (for vertically merged cells, use bottom border of the last merged row)
            bdr_bottom = cell.borders.bottom
            if v_span > 1:
                last_r_cells = all_row_recs[r_idx + v_span - 1]["cells"]
                for le in last_r_cells:
                    if le["g_col"] == entry["g_col"]:
                        bdr_bottom = le["cell"].borders.bottom
                        break

            for side, spec, (x0, y0, x1, y1) in (
                ("top", cell.borders.top, (cx, row_y, cx + cw, row_y)),
                ("bottom", bdr_bottom, (cx, row_y + cell_h, cx + cw, row_y + cell_h)),
                ("left", cell.borders.left, (cx, row_y, cx, row_y + cell_h)),
                ("right", cell.borders.right, (cx + cw, row_y, cx + cw, row_y + cell_h)),
            ):
                if spec.is_visible:
                    target.lines.append(
                        DrawLineCmd(
                            x0=x0,
                            y0=y0,
                            x1=x1,
                            y1=y1,
                            color=spec.color or "000000",
                            width=spec.width_pt,
                            style=spec.val,
                            z_order=1,
                            xml_path=cell.xml_path,
                        )
                    )

            target.bboxes.append(
                ElementBBox(
                    kind="table_cell",
                    xml_path=cell.xml_path,
                    x0=cx,
                    y0=row_y,
                    x1=cx + cw,
                    y1=row_y + cell_h,
                    summary=f"r{r_idx + 1}c{entry['g_col'] + 1}",
                )
            )

    def _layout_table_on_page(self, tb: TableBlock) -> None:
        assert self.curr_page is not None
        # A table in a multi-column section is laid out inside its column, not across the page.
        avail_w = self.col_w
        flow_x = self.col_x
        col_widths = self._resolve_table_column_widths(tb, avail_w)
        if not col_widths:
            return

        total_tbl_w = sum(col_widths)
        if tb.alignment == "center":
            table_x = flow_x + (avail_w - total_tbl_w) / 2.0
        elif tb.alignment == "right":
            table_x = flow_x + avail_w - total_tbl_w
        else:
            table_x = flow_x + tb.indent_left_pt

        row_recs = self._prepare_laid_out_rows(
            tb, col_widths, page_num=self.curr_page.page_number, total_pages=max(1, len(self.pages))
        )
        if not row_recs:
            return

        # Identify repeating header rows at top of table
        header_recs: List[dict] = []
        for rr in row_recs:
            if rr["row"].is_header:
                header_recs.append(rr)
            else:
                break

        if not self.is_top_of_page:
            self.cursor_y += self.prev_space_after_pt

        tbl_start_y = self.cursor_y

        for r_idx, rr in enumerate(row_recs):
            rh = rr["height"]
            if not self.is_top_of_page and (self.cursor_y + rh > self.page_bottom_limit):
                # Record table bbox on current page before breaking
                if self.cursor_y > tbl_start_y:
                    self.curr_page.bboxes.append(
                        ElementBBox(
                            kind="table",
                            xml_path=tb.xml_path,
                            x0=table_x,
                            y0=tbl_start_y,
                            x1=table_x + total_tbl_w,
                            y1=self.cursor_y,
                            style_id=tb.style_id,
                        )
                    )
                self._advance_flow()
                tbl_start_y = self.cursor_y
                # Follow the flow into the next column: the rows already measured keep their
                # widths, so the table only shifts sideways by the column delta.
                table_x += self.col_x - flow_x
                flow_x = self.col_x
                # Repeat header rows on new page if we are past the header rows
                if r_idx >= len(header_recs):
                    for h_rr in header_recs:
                        self._emit_table_row(self.curr_page, tb, h_rr, row_recs, table_x, self.cursor_y)
                        self.cursor_y += h_rr["height"]
                        self.is_top_of_page = False

            self._emit_table_row(self.curr_page, tb, rr, row_recs, table_x, self.cursor_y)
            self.cursor_y += rh
            self.is_top_of_page = False

        if self.cursor_y > tbl_start_y:
            self.curr_page.bboxes.append(
                ElementBBox(
                    kind="table",
                    xml_path=tb.xml_path,
                    x0=table_x,
                    y0=tbl_start_y,
                    x1=table_x + total_tbl_w,
                    y1=self.cursor_y,
                    style_id=tb.style_id,
                )
            )
        self.prev_space_after_pt = 6.0
        self.prev_style_id = None
        self.prev_contextual_spacing = False

    def _layout_table_to_box(
        self, tb: TableBlock, avail_width_pt: float, page_num: int, total_pages: int
    ) -> _LaidOutBox:
        box = _LaidOutBox()
        col_widths = self._resolve_table_column_widths(tb, avail_width_pt)
        if not col_widths:
            return box
        total_tbl_w = sum(col_widths)
        if tb.alignment == "center":
            table_x = (avail_width_pt - total_tbl_w) / 2.0
        elif tb.alignment == "right":
            table_x = avail_width_pt - total_tbl_w
        else:
            table_x = tb.indent_left_pt

        row_recs = self._prepare_laid_out_rows(tb, col_widths, page_num, total_pages)
        curr_y = 0.0
        for rr in row_recs:
            self._emit_table_row(box, tb, rr, row_recs, table_x, curr_y)
            curr_y += rr["height"]
        box.height_pt = curr_y
        return box

    # -------------------------------------------------------------------------
    # Headers & Footers Post-Pass
    # -------------------------------------------------------------------------
    def _render_all_headers_footers(self, total_pages: int) -> None:
        for pg in self.pages:
            sp: Optional[SectionProps] = getattr(pg, "_sect_pr", None)
            if sp is None:
                continue
            is_first: bool = getattr(pg, "_is_first_of_sect", False)
            is_even = pg.page_number % 2 == 0

            # Select active header
            hdr: Optional[HeaderFooterContent] = None
            if is_first and sp.title_pg and sp.header_first is not None:
                hdr = sp.header_first
            elif is_even and self.doc.even_and_odd_headers and sp.header_even is not None:
                hdr = sp.header_even
            else:
                hdr = sp.header_default

            if hdr is not None and hdr.blocks:
                hbox = self._layout_blocks_in_box(
                    hdr.blocks, sp.content_width_pt, page_num=pg.page_number, total_pages=total_pages, in_table_cell=True
                )
                hbox.offset_by(sp.margin_left_pt, sp.header_dist_pt)
                pg.rects.extend(hbox.rects)
                pg.lines.extend(hbox.lines)
                pg.images.extend(hbox.images)
                pg.texts.extend(hbox.texts)

            # Select active footer
            ftr: Optional[HeaderFooterContent] = None
            if is_first and sp.title_pg and sp.footer_first is not None:
                ftr = sp.footer_first
            elif is_even and self.doc.even_and_odd_headers and sp.footer_even is not None:
                ftr = sp.footer_even
            else:
                ftr = sp.footer_default

            if ftr is not None and ftr.blocks:
                fbox = self._layout_blocks_in_box(
                    ftr.blocks, sp.content_width_pt, page_num=pg.page_number, total_pages=total_pages, in_table_cell=True
                )
                fy = max(
                    sp.page_height_pt - sp.margin_bottom_pt,
                    sp.page_height_pt - sp.footer_dist_pt - fbox.height_pt,
                )
                fbox.offset_by(sp.margin_left_pt, fy)
                pg.rects.extend(fbox.rects)
                pg.lines.extend(fbox.lines)
                pg.images.extend(fbox.images)
                pg.texts.extend(fbox.texts)
