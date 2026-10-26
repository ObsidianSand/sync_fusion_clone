"""Data models for OpenXML document elements, resolved styles, and laid-out primitives."""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple, Union


@dataclass
class BorderSpec:
    """Single border specification (top, bottom, left, right, insideH, insideV)."""
    val: str = "none"          # 'single', 'double', 'dashed', 'dotted', 'thick', 'nil', 'none'
    width_pt: float = 0.0      # in points
    color: Optional[str] = "000000"  # 6-char hex or None
    space_pt: float = 0.0

    @property
    def is_visible(self) -> bool:
        return self.val not in ("none", "nil", "") and self.width_pt > 0.0


@dataclass
class Borders:
    """Collection of borders for a paragraph or table/cell."""
    top: BorderSpec = field(default_factory=BorderSpec)
    bottom: BorderSpec = field(default_factory=BorderSpec)
    left: BorderSpec = field(default_factory=BorderSpec)
    right: BorderSpec = field(default_factory=BorderSpec)
    insideH: BorderSpec = field(default_factory=BorderSpec)
    insideV: BorderSpec = field(default_factory=BorderSpec)


@dataclass
class TabStop:
    """Custom tab stop from <w:tabs><w:tab .../></w:tabs>."""
    pos_pt: float              # Position in points relative to paragraph column left
    val: str = "left"          # 'left', 'center', 'right', 'decimal', 'bar', 'clear', 'num'
    leader: str = "none"       # 'none', 'dot', 'hyphen', 'underscore', 'middleDot'


@dataclass
class RunStyle:
    """Effective (resolved) character/run formatting."""
    font_family: Optional[str] = None
    font_size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[str] = None     # 'single', 'double', 'dotted', 'none', etc.
    strike: Optional[bool] = None
    dstrike: Optional[bool] = None
    small_caps: Optional[bool] = None
    all_caps: Optional[bool] = None
    vert_align: Optional[str] = None    # 'superscript', 'subscript', 'baseline'
    color: Optional[str] = None         # 6-char hex
    highlight: Optional[str] = None     # 6-char hex
    shading_fill: Optional[str] = None  # 6-char hex
    char_spacing_pt: Optional[float] = None  # w:spacing in twips -> pt

    def merge_from(self, other: Optional[RunStyle]) -> RunStyle:
        """Return a new RunStyle where non-None attributes in `other` override `self`."""
        if other is None:
            return RunStyle(**self.__dict__)
        merged = RunStyle(**self.__dict__)
        for k, v in other.__dict__.items():
            if v is not None:
                setattr(merged, k, v)
        return merged


@dataclass
class ParagraphStyle:
    """Effective (resolved) paragraph formatting."""
    alignment: Optional[str] = None         # 'left', 'center', 'right', 'both'
    indent_left_pt: Optional[float] = None
    indent_right_pt: Optional[float] = None
    indent_first_line_pt: Optional[float] = None  # positive = firstLine, negative = hanging
    space_before_pt: Optional[float] = None
    space_after_pt: Optional[float] = None
    line_spacing: Optional[float] = None    # multiplier if line_rule=='auto', else pt
    line_rule: Optional[str] = None         # 'auto', 'exact', 'atLeast'
    keep_next: Optional[bool] = None
    keep_lines: Optional[bool] = None
    page_break_before: Optional[bool] = None
    contextual_spacing: Optional[bool] = None
    shading_fill: Optional[str] = None
    borders: Optional[Borders] = None
    tabs: List[TabStop] = field(default_factory=list)
    num_id: Optional[int] = None
    num_ilvl: Optional[int] = None
    outline_lvl: Optional[int] = None
    run_style: RunStyle = field(default_factory=RunStyle)

    def merge_from(self, other: Optional[ParagraphStyle]) -> ParagraphStyle:
        """Return a new ParagraphStyle where `other` overrides `self`."""
        if other is None:
            return ParagraphStyle(
                **{k: v for k, v in self.__dict__.items() if k not in ("tabs", "run_style", "borders")},
                tabs=list(self.tabs),
                borders=self.borders,
                run_style=self.run_style.merge_from(None),
            )
        merged = ParagraphStyle(
            **{k: v for k, v in self.__dict__.items() if k not in ("tabs", "run_style", "borders")},
            tabs=list(self.tabs),
            borders=self.borders,
            run_style=self.run_style.merge_from(other.run_style),
        )
        for k, v in other.__dict__.items():
            if k in ("tabs", "run_style", "borders"):
                continue
            if v is not None:
                setattr(merged, k, v)
        if other.borders is not None:
            if merged.borders is None:
                merged.borders = Borders(
                    top=other.borders.top,
                    bottom=other.borders.bottom,
                    left=other.borders.left,
                    right=other.borders.right,
                    insideH=other.borders.insideH,
                    insideV=other.borders.insideV,
                )
            else:
                merged.borders = Borders(
                    top=other.borders.top if other.borders.top.val != "none" else merged.borders.top,
                    bottom=other.borders.bottom if other.borders.bottom.val != "none" else merged.borders.bottom,
                    left=other.borders.left if other.borders.left.val != "none" else merged.borders.left,
                    right=other.borders.right if other.borders.right.val != "none" else merged.borders.right,
                    insideH=other.borders.insideH if other.borders.insideH.val != "none" else merged.borders.insideH,
                    insideV=other.borders.insideV if other.borders.insideV.val != "none" else merged.borders.insideV,
                )
        if other.tabs:
            # Merge tabs: 'clear' removes an existing stop at that position; otherwise override/append
            tab_map: Dict[float, TabStop] = {round(t.pos_pt, 2): t for t in merged.tabs}
            for t in other.tabs:
                key = round(t.pos_pt, 2)
                if t.val == "clear":
                    tab_map.pop(key, None)
                else:
                    tab_map[key] = t
            merged.tabs = sorted(tab_map.values(), key=lambda x: x.pos_pt)
        return merged


@dataclass
class DrawingItem:
    """Inline or anchored image/shape from <w:drawing>."""
    image_bytes: Optional[bytes] = None
    content_type: Optional[str] = None
    width_pt: float = 0.0
    height_pt: float = 0.0
    is_floating: bool = False
    # Floating anchor positioning
    pos_h_relative_from: str = "column"     # 'page', 'margin', 'column', 'character'
    pos_h_offset_pt: Optional[float] = None
    pos_h_align: Optional[str] = None       # 'left', 'center', 'right'
    pos_v_relative_from: str = "paragraph"  # 'page', 'margin', 'paragraph', 'line'
    pos_v_offset_pt: Optional[float] = None
    pos_v_align: Optional[str] = None       # 'top', 'center', 'bottom'
    wrap_type: str = "inline"               # 'inline', 'square', 'tight', 'topAndBottom', 'none'
    behind_doc: bool = False
    dist_t_pt: float = 0.0
    dist_b_pt: float = 0.0
    dist_l_pt: float = 9.0
    dist_r_pt: float = 9.0
    # Optional shape / textbox content
    shape_fill: Optional[str] = None
    shape_stroke: Optional[str] = None
    shape_stroke_pt: float = 0.0
    textbox_blocks: List[Union[ParagraphBlock, TableBlock]] = field(default_factory=list)
    xml_path: str = ""


@dataclass
class RunItem:
    """A single content item inside a paragraph."""
    kind: str = "text"                      # 'text', 'tab', 'break', 'drawing', 'page_num', 'num_pages'
    text: str = ""
    break_type: str = "textWrapping"        # 'textWrapping', 'page', 'column'
    drawing: Optional[DrawingItem] = None
    style: RunStyle = field(default_factory=RunStyle)
    xml_path: str = ""


@dataclass
class ParagraphBlock:
    """A parsed paragraph with its resolved style and run items."""
    style_id: Optional[str] = None
    style: ParagraphStyle = field(default_factory=ParagraphStyle)
    runs: List[RunItem] = field(default_factory=list)
    list_marker: Optional[str] = None
    list_marker_style: Optional[RunStyle] = None
    list_tab_pt: Optional[float] = None
    sect_pr: Optional[SectionProps] = None
    xml_path: str = ""


@dataclass
class TableCell:
    """A single table cell <w:tc>."""
    width_pt: Optional[float] = None
    width_type: str = "dxa"                 # 'dxa', 'pct', 'auto', 'nil'
    width_pct: Optional[float] = None       # 0.0 to 1.0
    grid_span: int = 1
    v_merge: Optional[str] = None           # None, 'restart', 'continue'
    v_align: str = "top"                    # 'top', 'center', 'bottom'
    shading_fill: Optional[str] = None
    borders: Borders = field(default_factory=Borders)
    margin_top_pt: Optional[float] = None
    margin_bottom_pt: Optional[float] = None
    margin_left_pt: Optional[float] = None
    margin_right_pt: Optional[float] = None
    no_wrap: bool = False
    blocks: List[Union[ParagraphBlock, TableBlock]] = field(default_factory=list)
    xml_path: str = ""


@dataclass
class TableRow:
    """A table row <w:tr>."""
    height_pt: float = 0.0
    height_rule: str = "auto"               # 'auto', 'atLeast', 'exact'
    is_header: bool = False
    cant_split: bool = False
    grid_before: int = 0
    w_before_pt: float = 0.0
    cells: List[TableCell] = field(default_factory=list)
    xml_path: str = ""


@dataclass
class TableBlock:
    """A parsed table <w:tbl>."""
    style_id: Optional[str] = None
    alignment: str = "left"                 # 'left', 'center', 'right'
    indent_left_pt: float = 0.0
    width_pt: Optional[float] = None
    width_type: str = "auto"                # 'dxa', 'pct', 'auto'
    width_pct: Optional[float] = None
    layout_type: str = "autofit"            # 'fixed', 'autofit'
    grid_cols_pt: List[float] = field(default_factory=list)
    borders: Borders = field(default_factory=Borders)
    cell_margin_top_pt: float = 0.0
    cell_margin_bottom_pt: float = 0.0
    cell_margin_left_pt: float = 5.4        # Word default = 108 twips = 5.4 pt
    cell_margin_right_pt: float = 5.4       # Word default = 108 twips = 5.4 pt
    shading_fill: Optional[str] = None
    rows: List[TableRow] = field(default_factory=list)
    xml_path: str = ""


@dataclass
class HeaderFooterContent:
    """Parsed blocks for a header or footer part."""
    blocks: List[Union[ParagraphBlock, TableBlock]] = field(default_factory=list)


@dataclass
class SectionProps:
    """Page setup and section properties <w:sectPr>."""
    page_width_pt: float = 612.0            # 8.5 in = 12240 twips = 612 pt
    page_height_pt: float = 792.0           # 11.0 in = 15840 twips = 792 pt
    orientation: str = "portrait"
    margin_top_pt: float = 72.0             # 1.0 in
    margin_bottom_pt: float = 72.0
    margin_left_pt: float = 72.0
    margin_right_pt: float = 72.0
    header_dist_pt: float = 36.0            # 0.5 in
    footer_dist_pt: float = 36.0            # 0.5 in
    cols_num: int = 1
    cols_space_pt: float = 36.0
    cols_equal_width: bool = True
    #: Per-column (width_pt, space_after_pt) from explicit <w:col> children, used when the section
    #: declares unequal columns. Empty for the usual equal-width case.
    cols_explicit: List[Tuple[float, float]] = field(default_factory=list)
    title_pg: bool = False                  # Different first page header/footer
    break_type: str = "nextPage"            # 'nextPage', 'continuous', 'evenPage', 'oddPage'
    header_default: Optional[HeaderFooterContent] = None
    header_first: Optional[HeaderFooterContent] = None
    header_even: Optional[HeaderFooterContent] = None
    footer_default: Optional[HeaderFooterContent] = None
    footer_first: Optional[HeaderFooterContent] = None
    footer_even: Optional[HeaderFooterContent] = None

    @property
    def content_width_pt(self) -> float:
        return max(36.0, self.page_width_pt - self.margin_left_pt - self.margin_right_pt)

    def column_bounds(self) -> List[Tuple[float, float]]:
        """``(x0, width)`` of each text column of this section, left to right, in points.

        Word divides the text area into `w:num` columns separated by `w:space`, so an equal-width
        column measures ``(content_width - (num - 1) * space) / num``. A section declaring unequal
        columns carries an explicit ``<w:col w:w=".." w:space=".."/>`` per column instead, and the
        last one's space is unused.
        """
        n = max(1, self.cols_num)
        if n == 1:
            return [(self.margin_left_pt, self.content_width_pt)]

        if self.cols_explicit and len(self.cols_explicit) >= n:
            bounds: List[Tuple[float, float]] = []
            x = self.margin_left_pt
            for width, space in self.cols_explicit[:n]:
                bounds.append((x, max(1.0, width)))
                x += width + space
            return bounds

        width = max(1.0, (self.content_width_pt - self.cols_space_pt * (n - 1)) / n)
        return [
            (self.margin_left_pt + i * (width + self.cols_space_pt), width)
            for i in range(n)
        ]


@dataclass
class ParsedDocument:
    """Complete parsed and style-resolved OpenXML document."""
    blocks: List[Union[ParagraphBlock, TableBlock]] = field(default_factory=list)
    final_sect_pr: SectionProps = field(default_factory=SectionProps)
    default_tab_stop_pt: float = 36.0       # Word default = 720 twips = 36 pt (0.5 in)
    even_and_odd_headers: bool = False


# ============================================================================
# Layout Primitives (Output of the Layout Engine, Input to the Canvas Renderer)
# ============================================================================

@dataclass
class DrawTextCmd:
    x: float
    y_baseline: float
    text: str
    font_family: str
    font_size_pt: float
    bold: bool
    italic: bool
    color: str
    underline: Optional[str] = None
    strike: bool = False
    dstrike: bool = False
    highlight: Optional[str] = None
    shading_fill: Optional[str] = None
    width_pt: float = 0.0
    ascender_pt: float = 0.0
    descender_pt: float = 0.0
    xml_path: str = ""


@dataclass
class DrawRectCmd:
    x: float
    y: float
    width: float
    height: float
    fill_color: Optional[str] = None
    stroke_color: Optional[str] = None
    stroke_width: float = 0.0
    dash_pattern: Optional[str] = None      # 'dashed', 'dotted', 'double'
    z_order: int = 0                        # -1 behind text, 0 normal, 1 foreground
    xml_path: str = ""


@dataclass
class DrawLineCmd:
    x0: float
    y0: float
    x1: float
    y1: float
    color: str = "000000"
    width: float = 0.5
    style: str = "single"                   # 'single', 'double', 'dashed', 'dotted', 'thick'
    z_order: int = 1
    xml_path: str = ""


@dataclass
class DrawImageCmd:
    x: float
    y: float
    width: float
    height: float
    image_bytes: bytes
    z_order: int = 0
    xml_path: str = ""


@dataclass
class ElementBBox:
    """Bounding box metadata for LLM visual grounding & inspection."""
    kind: str                               # 'paragraph', 'table', 'table_cell', 'image', 'list_marker', 'tab_leader'
    xml_path: str
    x0: float
    y0: float
    x1: float
    y1: float
    style_id: Optional[str] = None
    summary: str = ""


@dataclass
class PageLayout:
    """A single laid-out page ready for canvas rendering."""
    page_number: int
    width_pt: float
    height_pt: float
    margin_top_pt: float
    margin_bottom_pt: float
    margin_left_pt: float
    margin_right_pt: float
    rects: List[DrawRectCmd] = field(default_factory=list)
    lines: List[DrawLineCmd] = field(default_factory=list)
    images: List[DrawImageCmd] = field(default_factory=list)
    texts: List[DrawTextCmd] = field(default_factory=list)
    bboxes: List[ElementBBox] = field(default_factory=list)
