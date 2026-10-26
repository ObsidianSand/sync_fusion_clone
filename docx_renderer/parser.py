"""Direct OpenXML (.docx) Parser and Style/Theme/Numbering Cascade Resolver."""

from __future__ import annotations
import posixpath
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union
from lxml import etree

from .model import (
    Borders,
    BorderSpec,
    DrawingItem,
    HeaderFooterContent,
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
from .units import (
    HIGHLIGHT_COLORS,
    NS,
    apply_tint_or_shade,
    eighths_pt_to_pt,
    emu_to_pt,
    half_pts_to_pt,
    parse_bool_attr,
    qn,
    twips_to_pt,
)

# Private-Use Area (PUA) Wingdings/Symbol bullet codepoints used by MS Word -> Unicode equivalents
PUA_BULLET_MAP = {
    "\uf0b7": "•",
    "\uf0a7": "▪",
    "\uf0a8": "▫",
    "\uf0d8": "➢",
    "\uf0fc": "✓",
    "\uf076": "❖",
    "\uf06e": "■",
    "\uf075": "◆",
    "\uf0e0": "➔",
    "o": "○",
    "§": "▪",
    "¶": "•",
}


def _int_to_roman(n: int, upper: bool = False) -> str:
    if n <= 0:
        return str(n)
    vals = [
        (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"),
        (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
        (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i"),
    ]
    out = []
    for v, s in vals:
        while n >= v:
            out.append(s)
            n -= v
    res = "".join(out)
    return res.upper() if upper else res


def _int_to_alpha(n: int, upper: bool = False) -> str:
    if n <= 0:
        return str(n)
    res = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res = chr(97 + rem) + res
    return res.upper() if upper else res


@dataclass
class ThemeData:
    major_font: str = "Calibri Light"
    minor_font: str = "Calibri"
    colors: Dict[str, str] = field(default_factory=lambda: {
        "dk1": "000000",
        "lt1": "FFFFFF",
        "dk2": "44546A",
        "lt2": "E7E6E6",
        "accent1": "4472C4",
        "accent2": "ED7D31",
        "accent3": "A5A5A5",
        "accent4": "FFC000",
        "accent5": "5B9BD5",
        "accent6": "70AD47",
        "hlink": "0563C1",
        "folHlink": "954F72",
        "tx1": "000000",
        "tx2": "44546A",
        "bg1": "FFFFFF",
        "bg2": "E7E6E6",
    })


@dataclass
class StyleDef:
    style_id: str
    style_type: str                     # 'paragraph', 'character', 'table', 'numbering'
    name: str = ""
    based_on: Optional[str] = None
    is_default: bool = False
    p_style: ParagraphStyle = field(default_factory=ParagraphStyle)
    r_style: RunStyle = field(default_factory=RunStyle)
    # Table style specific properties
    tbl_borders: Optional[Borders] = None
    tbl_cell_margins: Dict[str, float] = field(default_factory=dict)
    tbl_shading: Optional[str] = None
    # Conditional formatting: key -> (ParagraphStyle, RunStyle, Optional[Borders], Optional[str shading])
    cond_formats: Dict[str, Tuple[ParagraphStyle, RunStyle, Optional[Borders], Optional[str]]] = field(
        default_factory=dict
    )


@dataclass
class NumLevelDef:
    ilvl: int
    start: int = 1
    num_fmt: str = "decimal"            # 'decimal', 'bullet', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman', 'none'
    lvl_text: str = "%1."
    lvl_jc: str = "left"
    suff: str = "tab"                   # 'tab', 'space', 'nothing'
    p_style: ParagraphStyle = field(default_factory=ParagraphStyle)
    r_style: RunStyle = field(default_factory=RunStyle)


class DocxParser:
    """Parses a .docx file into a fully style-resolved ParsedDocument."""

    def __init__(self, docx_path: str | Path) -> None:
        self.docx_path = Path(docx_path)
        self.zf = zipfile.ZipFile(self.docx_path, "r")
        self.theme = ThemeData()
        self.doc_default_p = ParagraphStyle(
            alignment="left",
            indent_left_pt=0.0,
            indent_right_pt=0.0,
            indent_first_line_pt=0.0,
            space_before_pt=0.0,
            space_after_pt=0.0,
            line_spacing=1.0,
            line_rule="auto",
        )
        self.doc_default_r = RunStyle(
            font_family="Calibri",
            font_size_pt=11.0,
            bold=False,
            italic=False,
            color="000000",
        )
        self.styles: Dict[str, StyleDef] = {}
        self.default_Style_id_by_type: Dict[str, str] = {}
        self.abstract_nums: Dict[int, Dict[int, NumLevelDef]] = {}
        self.num_to_abstract: Dict[int, int] = {}
        self.num_level_overrides: Dict[Tuple[int, int], NumLevelDef] = {}
        self.num_Override_start: Dict[Tuple[int, int], int] = {}
        self.num_counters: Dict[Tuple[int, int], int] = {}
        self.default_tab_stop_pt: float = 36.0
        self.even_and_odd_headers: bool = False

    def close(self) -> None:
        self.zf.close()

    def _read_xml(self, part_path: str) -> Optional[etree._Element]:
        norm = part_path.lstrip("/")
        if norm not in self.zf.namelist():
            return None
        data = self.zf.read(norm)
        return etree.fromstring(data)

    def _read_rels(self, part_path: str) -> Dict[str, str]:
        """Read relationships for a given part (e.g., 'word/document.xml' -> 'word/_rels/document.xml.rels')."""
        norm = part_path.lstrip("/")
        dirname, basename = posixpath.split(norm)
        rels_path = posixpath.join(dirname, "_rels", f"{basename}.rels")
        root = self._read_xml(rels_path)
        if root is None:
            return {}
        rels: Dict[str, str] = {}
        for rel in root:
            rid = rel.get("Id")
            target = rel.get("Target")
            tmode = rel.get("TargetMode", "")
            if rid and target and tmode != "External":
                if target.startswith("/"):
                    resolved = target.lstrip("/")
                else:
                    resolved = posixpath.normpath(posixpath.join(dirname, target))
                rels[rid] = resolved
        return rels

    def parse(self) -> ParsedDocument:
        """Execute full parsing pipeline and return ParsedDocument."""
        try:
            self._parse_theme()
            self._parse_settings()
            self._parse_styles()
            self._parse_numbering()
            return self._parse_document()
        finally:
            self.close()

    # -------------------------------------------------------------------------
    # 1. Theme & Color Resolution
    # -------------------------------------------------------------------------
    def _parse_theme(self) -> None:
        root = self._read_xml("word/theme/theme1.xml")
        if root is None:
            return
        # Color scheme
        clr_scheme = root.find(".//a:clrScheme", NS)
        if clr_scheme is not None:
            for child in clr_scheme:
                local_name = etree.QName(child).localname
                srgb = child.find("a:srgbClr", NS)
                sys_clr = child.find("a:sysClr", NS)
                hex_val = None
                if srgb is not None:
                    hex_val = srgb.get("val")
                elif sys_clr is not None:
                    hex_val = sys_clr.get("lastClr") or sys_clr.get("val")
                if hex_val and len(hex_val) == 6:
                    self.theme.colors[local_name] = hex_val.upper()
            # Map tx1/tx2/bg1/bg2 aliases
            self.theme.colors["tx1"] = self.theme.colors.get("dk1", "000000")
            self.theme.colors["tx2"] = self.theme.colors.get("dk2", "44546A")
            self.theme.colors["bg1"] = self.theme.colors.get("lt1", "FFFFFF")
            self.theme.colors["bg2"] = self.theme.colors.get("lt2", "E7E6E6")

        # Font scheme
        font_scheme = root.find(".//a:fontScheme", NS)
        if font_scheme is not None:
            maj = font_scheme.find("a:majorFont/a:latin", NS)
            if maj is not None and maj.get("typeface"):
                self.theme.major_font = maj.get("typeface")
            min_f = font_scheme.find("a:minorFont/a:latin", NS)
            if min_f is not None and min_f.get("typeface"):
                self.theme.minor_font = min_f.get("typeface")

    def _resolve_color_elem(self, elem: Optional[etree._Element], attr_val: str = "w:val") -> Optional[str]:
        if elem is None:
            return None
        val = elem.get(qn(attr_val))
        theme_color = elem.get(qn("w:themeColor")) or elem.get(qn("w:themeFill"))
        theme_tint = elem.get(qn("w:themeTint")) or elem.get(qn("w:themeFillTint"))
        theme_shade = elem.get(qn("w:themeShade")) or elem.get(qn("w:themeFillShade"))

        base_hex: Optional[str] = None
        if theme_color:
            alias_map = {
                "dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2",
                "text1": "tx1", "text2": "tx2", "background1": "bg1", "background2": "bg2",
                "hyperlink": "hlink", "followedHyperlink": "folHlink",
            }
            tkey = alias_map.get(theme_color, theme_color)
            base_hex = self.theme.colors.get(tkey)

        if not base_hex and val:
            low = val.lower()
            if low in ("auto", "none", "nil"):
                # "auto" defers the choice to the consumer: Word resolves automatic *text* to
                # black, but an automatic *fill* is the absence of shading, not a black cell.
                return "000000" if low == "auto" and attr_val != "w:fill" else None
            if len(val) == 6:
                base_hex = val.upper()

        if base_hex:
            return apply_tint_or_shade(base_hex, theme_tint=theme_tint, theme_shade=theme_shade)
        return None

    def _resolve_font_name(self, rfonts_elem: Optional[etree._Element]) -> Optional[str]:
        if rfonts_elem is None:
            return None
        theme_attr = rfonts_elem.get(qn("w:asciiTheme")) or rfonts_elem.get(qn("w:hAnsiTheme"))
        if theme_attr:
            if "major" in theme_attr.lower():
                return self.theme.major_font
            if "minor" in theme_attr.lower():
                return self.theme.minor_font
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            val = rfonts_elem.get(qn(attr))
            if val:
                if val.startswith("+mj"):
                    return self.theme.major_font
                if val.startswith("+mn"):
                    return self.theme.minor_font
                return val
        return None

    # -------------------------------------------------------------------------
    # 2. Settings
    # -------------------------------------------------------------------------
    def _parse_settings(self) -> None:
        root = self._read_xml("word/settings.xml")
        if root is None:
            return
        dts = root.find("w:defaultTabStop", NS)
        if dts is not None:
            val = twips_to_pt(dts.get(qn("w:val")))
            if val > 0:
                self.default_tab_stop_pt = val
        eao = root.find("w:evenAndOddHeaders", NS)
        if eao is not None:
            self.even_and_odd_headers = parse_bool_attr(eao)

    # -------------------------------------------------------------------------
    # 3. Properties Parsing Helpers (rPr, pPr, Borders)
    # -------------------------------------------------------------------------
    def _parse_border_spec(self, elem: Optional[etree._Element]) -> BorderSpec:
        if elem is None:
            return BorderSpec()
        val = elem.get(qn("w:val"), "none")
        if val in ("nil", "none"):
            return BorderSpec(val="none", width_pt=0.0, color=None)
        sz = eighths_pt_to_pt(elem.get(qn("w:sz")))
        if sz <= 0.0:
            sz = 0.5
        space = float(elem.get(qn("w:space") or "0", 0))
        color = self._resolve_color_elem(elem) or "000000"
        return BorderSpec(val=val, width_pt=sz, color=color, space_pt=space)

    def _parse_borders(self, bdr_container: Optional[etree._Element]) -> Optional[Borders]:
        if bdr_container is None:
            return None
        left_el = bdr_container.find("w:left", NS)
        if left_el is None:
            left_el = bdr_container.find("w:start", NS)
        right_el = bdr_container.find("w:right", NS)
        if right_el is None:
            right_el = bdr_container.find("w:end", NS)
        return Borders(
            top=self._parse_border_spec(bdr_container.find("w:top", NS)),
            bottom=self._parse_border_spec(bdr_container.find("w:bottom", NS)),
            left=self._parse_border_spec(left_el),
            right=self._parse_border_spec(right_el),
            insideH=self._parse_border_spec(bdr_container.find("w:insideH", NS)),
            insideV=self._parse_border_spec(bdr_container.find("w:insideV", NS)),
        )

    def _parse_rpr(self, rpr: Optional[etree._Element]) -> RunStyle:
        rs = RunStyle()
        if rpr is None:
            return rs
        rfonts = rpr.find("w:rFonts", NS)
        if rfonts is not None:
            rs.font_family = self._resolve_font_name(rfonts)

        sz = rpr.find("w:sz", NS)
        if sz is not None and sz.get(qn("w:val")) is not None:
            rs.font_size_pt = half_pts_to_pt(sz.get(qn("w:val")))

        b = rpr.find("w:b", NS)
        if b is not None:
            rs.bold = parse_bool_attr(b)
        i = rpr.find("w:i", NS)
        if i is not None:
            rs.italic = parse_bool_attr(i)
        u = rpr.find("w:u", NS)
        if u is not None:
            uval = u.get(qn("w:val"), "single")
            rs.underline = None if uval in ("none", "nil") else uval
        strike = rpr.find("w:strike", NS)
        if strike is not None:
            rs.strike = parse_bool_attr(strike)
        dstrike = rpr.find("w:dstrike", NS)
        if dstrike is not None:
            rs.dstrike = parse_bool_attr(dstrike)
        sc = rpr.find("w:smallCaps", NS)
        if sc is not None:
            rs.small_caps = parse_bool_attr(sc)
        caps = rpr.find("w:caps", NS)
        if caps is not None:
            rs.all_caps = parse_bool_attr(caps)
        va = rpr.find("w:vertAlign", NS)
        if va is not None:
            rs.vert_align = va.get(qn("w:val"))

        clr = rpr.find("w:color", NS)
        if clr is not None:
            rs.color = self._resolve_color_elem(clr)

        hl = rpr.find("w:highlight", NS)
        if hl is not None:
            hval = hl.get(qn("w:val"), "")
            rs.highlight = HIGHLIGHT_COLORS.get(hval)

        shd = rpr.find("w:shd", NS)
        if shd is not None:
            rs.shading_fill = self._resolve_color_elem(shd, "w:fill")

        sp = rpr.find("w:spacing", NS)
        if sp is not None and sp.get(qn("w:val")) is not None:
            rs.char_spacing_pt = twips_to_pt(sp.get(qn("w:val")))

        return rs

    def _parse_ppr(self, ppr: Optional[etree._Element]) -> ParagraphStyle:
        ps = ParagraphStyle()
        if ppr is None:
            return ps

        jc = ppr.find("w:jc", NS)
        if jc is not None:
            val = jc.get(qn("w:val"), "left")
            if val in ("start", "left"):
                ps.alignment = "left"
            elif val in ("end", "right"):
                ps.alignment = "right"
            elif val in ("both", "distribute"):
                ps.alignment = "both"
            elif val == "center":
                ps.alignment = "center"

        ind = ppr.find("w:ind", NS)
        if ind is not None:
            left = ind.get(qn("w:left")) or ind.get(qn("w:start"))
            right = ind.get(qn("w:right")) or ind.get(qn("w:end"))
            fl = ind.get(qn("w:firstLine"))
            hang = ind.get(qn("w:hanging"))
            if left is not None:
                ps.indent_left_pt = twips_to_pt(left)
            if right is not None:
                ps.indent_right_pt = twips_to_pt(right)
            if hang is not None:
                ps.indent_first_line_pt = -twips_to_pt(hang)
            elif fl is not None:
                ps.indent_first_line_pt = twips_to_pt(fl)

        spc = ppr.find("w:spacing", NS)
        if spc is not None:
            bef = spc.get(qn("w:before"))
            aft = spc.get(qn("w:after"))
            line = spc.get(qn("w:line"))
            lrule = spc.get(qn("w:lineRule"), "auto")
            if spc.get(qn("w:beforeAutospacing")) in ("1", "true"):
                ps.space_before_pt = 5.0
            elif bef is not None:
                ps.space_before_pt = twips_to_pt(bef)
            if spc.get(qn("w:afterAutospacing")) in ("1", "true"):
                ps.space_after_pt = 5.0
            elif aft is not None:
                ps.space_after_pt = twips_to_pt(aft)
            if line is not None:
                ps.line_rule = lrule
                if lrule == "auto":
                    try:
                        ps.line_spacing = float(line) / 240.0
                    except ValueError:
                        ps.line_spacing = 1.08
                else:
                    ps.line_spacing = twips_to_pt(line)

        kn = ppr.find("w:keepNext", NS)
        if kn is not None:
            ps.keep_next = parse_bool_attr(kn)
        kl = ppr.find("w:keepLines", NS)
        if kl is not None:
            ps.keep_lines = parse_bool_attr(kl)
        pbb = ppr.find("w:pageBreakBefore", NS)
        if pbb is not None:
            ps.page_break_before = parse_bool_attr(pbb)
        cs = ppr.find("w:contextualSpacing", NS)
        if cs is not None:
            ps.contextual_spacing = parse_bool_attr(cs)

        shd = ppr.find("w:shd", NS)
        if shd is not None:
            ps.shading_fill = self._resolve_color_elem(shd, "w:fill")

        pbdr = ppr.find("w:pBdr", NS)
        if pbdr is not None:
            ps.borders = self._parse_borders(pbdr)

        tabs_el = ppr.find("w:tabs", NS)
        if tabs_el is not None:
            for tab_el in tabs_el.findall("w:tab", NS):
                pos = twips_to_pt(tab_el.get(qn("w:pos")))
                val = tab_el.get(qn("w:val"), "left")
                leader = tab_el.get(qn("w:leader"), "none")
                ps.tabs.append(TabStop(pos_pt=pos, val=val, leader=leader))

        num_pr = ppr.find("w:numPr", NS)
        if num_pr is not None:
            ilvl_el = num_pr.find("w:ilvl", NS)
            numid_el = num_pr.find("w:numId", NS)
            if ilvl_el is not None and ilvl_el.get(qn("w:val")) is not None:
                try:
                    ps.num_ilvl = int(ilvl_el.get(qn("w:val")))
                except ValueError:
                    pass
            if numid_el is not None and numid_el.get(qn("w:val")) is not None:
                try:
                    ps.num_id = int(numid_el.get(qn("w:val")))
                except ValueError:
                    pass

        ol = ppr.find("w:outlineLvl", NS)
        if ol is not None and ol.get(qn("w:val")) is not None:
            try:
                ps.outline_lvl = int(ol.get(qn("w:val")))
            except ValueError:
                pass

        rpr = ppr.find("w:rPr", NS)
        if rpr is not None:
            ps.run_style = self._parse_rpr(rpr)

        return ps

    # -------------------------------------------------------------------------
    # 4. Styles Cascade
    # -------------------------------------------------------------------------
    def _parse_styles(self) -> None:
        # Update default font from theme minor font
        self.doc_default_r.font_family = self.theme.minor_font
        root = self._read_xml("word/styles.xml")
        if root is None:
            return

        # 1. docDefaults
        doc_defs = root.find("w:docDefaults", NS)
        if doc_defs is not None:
            rpr_def = doc_defs.find("w:rPrDefault/w:rPr", NS)
            if rpr_def is not None:
                self.doc_default_r = self.doc_default_r.merge_from(self._parse_rpr(rpr_def))
            ppr_def = doc_defs.find("w:pPrDefault/w:pPr", NS)
            if ppr_def is not None:
                self.doc_default_p = self.doc_default_p.merge_from(self._parse_ppr(ppr_def))

        # 2. Individual styles
        for s_el in root.findall("w:style", NS):
            sid = s_el.get(qn("w:styleId"))
            stype = s_el.get(qn("w:type"), "paragraph")
            if not sid:
                continue
            is_def = s_el.get(qn("w:default")) in ("1", "true")
            name_el = s_el.find("w:name", NS)
            name = name_el.get(qn("w:val"), "") if name_el is not None else ""
            based_el = s_el.find("w:basedOn", NS)
            based_on = based_el.get(qn("w:val")) if based_el is not None else None

            ppr = self._parse_ppr(s_el.find("w:pPr", NS))
            rpr = self._parse_rpr(s_el.find("w:rPr", NS))
            ppr.run_style = ppr.run_style.merge_from(rpr)

            sdef = StyleDef(
                style_id=sid,
                style_type=stype,
                name=name,
                based_on=based_on,
                is_default=is_def,
                p_style=ppr,
                r_style=rpr,
            )

            # Table properties on table styles
            tbl_pr = s_el.find("w:tblPr", NS)
            if tbl_pr is not None:
                sdef.tbl_borders = self._parse_borders(tbl_pr.find("w:tblBorders", NS))
                shd = tbl_pr.find("w:shd", NS)
                if shd is not None:
                    sdef.tbl_shading = self._resolve_color_elem(shd, "w:fill")
                tcm = tbl_pr.find("w:tblCellMar", NS)
                if tcm is not None:
                    for side in ("top", "bottom", "left", "right", "start", "end"):
                        m_el = tcm.find(f"w:{side}", NS)
                        if m_el is not None and m_el.get(qn("w:w")) is not None:
                            key = "left" if side == "start" else ("right" if side == "end" else side)
                            sdef.tbl_cell_margins[key] = twips_to_pt(m_el.get(qn("w:w")))

            # Conditional table style properties (<w:tblStylePr>)
            for cond_el in s_el.findall("w:tblStylePr", NS):
                ctype = cond_el.get(qn("w:type"))
                if not ctype:
                    continue
                cp = self._parse_ppr(cond_el.find("w:pPr", NS))
                cr = self._parse_rpr(cond_el.find("w:rPr", NS))
                tc_pr = cond_el.find("w:tcPr", NS)
                cbdr = self._parse_borders(tc_pr.find("w:tcBorders", NS)) if tc_pr is not None else None
                cshd = (
                    self._resolve_color_elem(tc_pr.find("w:shd", NS), "w:fill")
                    if tc_pr is not None and tc_pr.find("w:shd", NS) is not None
                    else None
                )
                sdef.cond_formats[ctype] = (cp, cr, cbdr, cshd)

            self.styles[sid] = sdef
            if is_def:
                self.default_Style_id_by_type[stype] = sid

    def resolve_paragraph_style(self, style_id: Optional[str]) -> ParagraphStyle:
        """Resolve full inheritance chain: docDefaults -> default paragraph style -> basedOn chain."""
        base = self.doc_default_p.merge_from(None)
        base.run_style = self.doc_default_r.merge_from(None)

        target_id = style_id or self.default_Style_id_by_type.get("paragraph", "Normal")
        chain: List[StyleDef] = []
        visited = set()
        curr_id = target_id
        while curr_id and curr_id in self.styles and curr_id not in visited:
            visited.add(curr_id)
            sdef = self.styles[curr_id]
            chain.append(sdef)
            curr_id = sdef.based_on

        for sdef in reversed(chain):
            base = base.merge_from(sdef.p_style)
            base.run_style = base.run_style.merge_from(sdef.r_style)

        # Built-in fallback heuristics if a TOC or Heading style wasn't explicitly defined in styles.xml
        if style_id and style_id not in self.styles:
            low = style_id.lower()
            if low.startswith("toc") and len(low) >= 4 and low[3:].strip().isdigit():
                lvl = int(low[3:].strip())
                base.indent_left_pt = (lvl - 1) * 11.0
                base.space_before_pt = 0.0
                base.space_after_pt = 4.0
        return base

    def is_toc_entry_style(self, style_id: Optional[str]) -> bool:
        """True if `style_id` (or any style it is based on) formats a TOC / table-of-figures entry.

        Word writes <w:rStyle w:val="Hyperlink"/> on every entry of a hyperlinked TOC field
        (``TOC \\h``) but renders those entries with the entry paragraph style's own color, not
        blue + underlined. The same applies to the lists a ``TOC \\c "Table"`` / ``\\c "Figure"``
        field produces, which Word styles with ``table of figures`` (styleId ``TableofFigures``)
        rather than ``toc N``, so the style name is checked alongside the styleId.
        """
        curr_id = style_id
        visited = set()
        while curr_id and curr_id not in visited:
            visited.add(curr_id)
            sdef = self.styles.get(curr_id)
            for label in (curr_id, sdef.name if sdef is not None else ""):
                low = "".join(label.lower().split())
                if low in ("tableoffigures", "tableofauthorities"):
                    return True
                if low.startswith("toc") and (not low[3:] or low[3:].isdigit()):
                    return True
            curr_id = sdef.based_on if sdef is not None else None
        return False

    def resolve_character_style(self, style_id: Optional[str]) -> RunStyle:
        rs = RunStyle()
        if not style_id:
            return rs
        if style_id not in self.styles and style_id.lower() == "hyperlink":
            return RunStyle(color=self.theme.colors.get("hlink", "0563C1"), underline="single")
        chain: List[StyleDef] = []
        visited = set()
        curr_id = style_id
        while curr_id and curr_id in self.styles and curr_id not in visited:
            visited.add(curr_id)
            sdef = self.styles[curr_id]
            chain.append(sdef)
            curr_id = sdef.based_on
        for sdef in reversed(chain):
            rs = rs.merge_from(sdef.r_style)
        return rs

    def resolve_table_style(self, style_id: Optional[str]) -> Optional[StyleDef]:
        target_id = style_id or self.default_Style_id_by_type.get("table")
        if not target_id or target_id not in self.styles:
            return None
        chain: List[StyleDef] = []
        visited = set()
        curr_id: Optional[str] = target_id
        while curr_id and curr_id in self.styles and curr_id not in visited:
            visited.add(curr_id)
            sdef = self.styles[curr_id]
            chain.append(sdef)
            curr_id = sdef.based_on

        merged = StyleDef(style_id=target_id, style_type="table")
        for sdef in reversed(chain):
            merged.p_style = merged.p_style.merge_from(sdef.p_style)
            merged.r_style = merged.r_style.merge_from(sdef.r_style)
            if sdef.tbl_borders is not None:
                merged.tbl_borders = sdef.tbl_borders
            if sdef.tbl_shading is not None:
                merged.tbl_shading = sdef.tbl_shading
            merged.tbl_cell_margins.update(sdef.tbl_cell_margins)
            merged.cond_formats.update(sdef.cond_formats)
        return merged

    # -------------------------------------------------------------------------
    # 5. Numbering & Lists
    # -------------------------------------------------------------------------
    def _parse_numbering(self) -> None:
        root = self._read_xml("word/numbering.xml")
        if root is None:
            return

        for abs_el in root.findall("w:abstractNum", NS):
            aid_str = abs_el.get(qn("w:abstractNumId"))
            if aid_str is None:
                continue
            aid = int(aid_str)
            levels: Dict[int, NumLevelDef] = {}
            for lvl_el in abs_el.findall("w:lvl", NS):
                levels[int(lvl_el.get(qn("w:ilvl"), "0"))] = self._parse_num_lvl(lvl_el)
            self.abstract_nums[aid] = levels

        for num_el in root.findall("w:num", NS):
            nid_str = num_el.get(qn("w:numId"))
            abs_ref = num_el.find("w:abstractNumId", NS)
            if nid_str is None or abs_ref is None or abs_ref.get(qn("w:val")) is None:
                continue
            nid = int(nid_str)
            aid = int(abs_ref.get(qn("w:val")))
            self.num_to_abstract[nid] = aid
            for ovr_el in num_el.findall("w:lvlOverride", NS):
                ilvl = int(ovr_el.get(qn("w:ilvl"), "0"))
                st_ovr = ovr_el.find("w:startOverride", NS)
                if st_ovr is not None and st_ovr.get(qn("w:val")) is not None:
                    self.num_Override_start[(nid, ilvl)] = int(st_ovr.get(qn("w:val")))
                lvl_el = ovr_el.find("w:lvl", NS)
                if lvl_el is not None:
                    self.num_level_overrides[(nid, ilvl)] = self._parse_num_lvl(lvl_el)

    def _parse_num_lvl(self, lvl_el: etree._Element) -> NumLevelDef:
        ilvl = int(lvl_el.get(qn("w:ilvl"), "0"))
        st_el = lvl_el.find("w:start", NS)
        start = int(st_el.get(qn("w:val"), "1")) if st_el is not None else 1
        fmt_el = lvl_el.find("w:numFmt", NS)
        num_fmt = fmt_el.get(qn("w:val"), "decimal") if fmt_el is not None else "decimal"
        txt_el = lvl_el.find("w:lvlText", NS)
        lvl_text = txt_el.get(qn("w:val"), "%1.") if txt_el is not None else "%1."
        jc_el = lvl_el.find("w:lvlJc", NS)
        lvl_jc = jc_el.get(qn("w:val"), "left") if jc_el is not None else "left"
        suff_el = lvl_el.find("w:suff", NS)
        suff = suff_el.get(qn("w:val"), "tab") if suff_el is not None else "tab"
        p_style = self._parse_ppr(lvl_el.find("w:pPr", NS))
        r_style = self._parse_rpr(lvl_el.find("w:rPr", NS))
        return NumLevelDef(
            ilvl=ilvl,
            start=start,
            num_fmt=num_fmt,
            lvl_text=lvl_text,
            lvl_jc=lvl_jc,
            suff=suff,
            p_style=p_style,
            r_style=r_style,
        )

    def _resolve_list_item(
        self, num_id: int, ilvl: int, base_p_style: ParagraphStyle
    ) -> Tuple[ParagraphStyle, Optional[str], Optional[RunStyle], Optional[float]]:
        """
        Advance list counters and return:
        (updated_paragraph_style, formatted_marker_string, marker_run_style, list_tab_pt).
        """
        if num_id <= 0 or num_id not in self.num_to_abstract:
            return base_p_style, None, None, None
        aid = self.num_to_abstract[num_id]
        lvl_def = self.num_level_overrides.get((num_id, ilvl)) or self.abstract_nums.get(aid, {}).get(ilvl)
        if lvl_def is None:
            return base_p_style, None, None, None

        # Merge numbering level paragraph formatting BEHIND direct formatting,
        # so lvl_def's indent applies unless overridden by paragraph pPr
        merged_p = lvl_def.p_style.merge_from(base_p_style)

        # Reset deeper levels when an outer level increments
        for deeper in range(ilvl + 1, 10):
            self.num_counters.pop((num_id, deeper), None)

        start_val = self.num_Override_start.get((num_id, ilvl), lvl_def.start)
        curr_val = self.num_counters.get((num_id, ilvl), start_val - 1) + 1
        self.num_counters[(num_id, ilvl)] = curr_val

        # Format marker string
        if lvl_def.num_fmt == "none":
            marker = ""
        elif lvl_def.num_fmt == "bullet":
            raw = lvl_def.lvl_text or "•"
            marker = "".join(PUA_BULLET_MAP.get(ch, ch) for ch in raw)
            if not marker.strip():
                marker = "•"
        else:
            marker = lvl_def.lvl_text
            for l_idx in range(ilvl + 1):
                l_def = (
                    self.num_level_overrides.get((num_id, l_idx))
                    or self.abstract_nums.get(aid, {}).get(l_idx)
                    or lvl_def
                )
                l_start = self.num_Override_start.get((num_id, l_idx), l_def.start)
                val = self.num_counters.get((num_id, l_idx), l_start)
                fmt = l_def.num_fmt
                if fmt == "lowerLetter":
                    s_val = _int_to_alpha(val, upper=False)
                elif fmt == "upperLetter":
                    s_val = _int_to_alpha(val, upper=True)
                elif fmt == "lowerRoman":
                    s_val = _int_to_roman(val, upper=False)
                elif fmt == "upperRoman":
                    s_val = _int_to_roman(val, upper=True)
                else:
                    s_val = str(val)
                marker = marker.replace(f"%{l_idx + 1}", s_val)

        marker_r_style = merged_p.run_style.merge_from(lvl_def.r_style)
        # If bullet font was Symbol/Wingdings, switch to Calibri/Arial since we mapped PUA to Unicode
        if marker_r_style.font_family and marker_r_style.font_family.lower() in ("symbol", "wingdings", "wingdings 2", "wingdings 3", "courier new"):
            if lvl_def.num_fmt == "bullet":
                marker_r_style.font_family = "Calibri"

        list_tab_pt = merged_p.indent_left_pt or 18.0
        if lvl_def.p_style.tabs:
            list_tab_pt = lvl_def.p_style.tabs[0].pos_pt

        return merged_p, marker, marker_r_style, list_tab_pt

    # -------------------------------------------------------------------------
    # 6. Document Body, Sections, Headers/Footers, Paragraphs, Drawings, Tables
    # -------------------------------------------------------------------------
    def _parse_document(self) -> ParsedDocument:
        root = self._read_xml("word/document.xml")
        if root is None:
            raise ValueError(f"Could not find word/document.xml in {self.docx_path}")

        rels = self._read_rels("word/document.xml")
        body = root.find("w:body", NS)
        if body is None:
            return ParsedDocument()

        blocks = self._parse_block_container(body, rels, "word/document.xml:/w:body")
        sect_el = body.find("w:sectPr", NS)
        final_sect = self._parse_sect_pr(sect_el, rels) if sect_el is not None else SectionProps()

        return ParsedDocument(
            blocks=blocks,
            final_sect_pr=final_sect,
            default_tab_stop_pt=self.default_tab_stop_pt,
            even_and_odd_headers=self.even_and_odd_headers,
        )

    def _parse_sect_pr(self, sect_el: etree._Element, doc_rels: Dict[str, str]) -> SectionProps:
        sp = SectionProps()
        pg_sz = sect_el.find("w:pgSz", NS)
        if pg_sz is not None:
            w = twips_to_pt(pg_sz.get(qn("w:w")))
            h = twips_to_pt(pg_sz.get(qn("w:h")))
            if w > 0:
                sp.page_width_pt = w
            if h > 0:
                sp.page_height_pt = h
            sp.orientation = pg_sz.get(qn("w:orient"), "portrait")

        pg_mar = sect_el.find("w:pgMar", NS)
        if pg_mar is not None:
            for attr, prop in (
                ("w:top", "margin_top_pt"),
                ("w:bottom", "margin_bottom_pt"),
                ("w:left", "margin_left_pt"),
                ("w:right", "margin_right_pt"),
                ("w:header", "header_dist_pt"),
                ("w:footer", "footer_dist_pt"),
            ):
                val = pg_mar.get(qn(attr))
                if val is not None:
                    setattr(sp, prop, abs(twips_to_pt(val)))

        cols = sect_el.find("w:cols", NS)
        if cols is not None:
            num = cols.get(qn("w:num"))
            if num and num.isdigit():
                sp.cols_num = max(1, int(num))
            space = cols.get(qn("w:space"))
            if space is not None:
                sp.cols_space_pt = twips_to_pt(space)

        tpg = sect_el.find("w:titlePg", NS)
        if tpg is not None:
            sp.title_pg = parse_bool_attr(tpg)

        tp = sect_el.find("w:type", NS)
        if tp is not None and tp.get(qn("w:val")):
            sp.break_type = tp.get(qn("w:val"))

        # Headers and footers
        for hf_el in sect_el.findall("w:headerReference", NS):
            htype = hf_el.get(qn("w:type"), "default")
            rid = hf_el.get(qn("r:id"))
            if rid and rid in doc_rels:
                content = self._parse_header_footer_part(doc_rels[rid])
                if htype == "first":
                    sp.header_first = content
                elif htype == "even":
                    sp.header_even = content
                else:
                    sp.header_default = content

        for ff_el in sect_el.findall("w:footerReference", NS):
            ftype = ff_el.get(qn("w:type"), "default")
            rid = ff_el.get(qn("r:id"))
            if rid and rid in doc_rels:
                content = self._parse_header_footer_part(doc_rels[rid])
                if ftype == "first":
                    sp.footer_first = content
                elif ftype == "even":
                    sp.footer_even = content
                else:
                    sp.footer_default = content

        return sp

    def _parse_header_footer_part(self, part_path: str) -> Optional[HeaderFooterContent]:
        root = self._read_xml(part_path)
        if root is None:
            return None
        rels = self._read_rels(part_path)
        blocks = self._parse_block_container(root, rels, f"{part_path}:/root")
        return HeaderFooterContent(blocks=blocks)

    def _parse_block_container(
        self,
        container: etree._Element,
        rels: Dict[str, str],
        base_path: str,
        table_cond_p: Optional[ParagraphStyle] = None,
        table_cond_r: Optional[RunStyle] = None,
    ) -> List[Union[ParagraphBlock, TableBlock]]:
        blocks: List[Union[ParagraphBlock, TableBlock]] = []
        p_idx = 0
        t_idx = 0
        for child in container:
            tag = child.tag
            if tag == qn("w:p"):
                p_idx += 1
                blocks.append(
                    self._parse_paragraph(
                        child, rels, f"{base_path}/w:p[{p_idx}]", table_cond_p, table_cond_r
                    )
                )
            elif tag == qn("w:tbl"):
                t_idx += 1
                blocks.append(self._parse_table(child, rels, f"{base_path}/w:tbl[{t_idx}]"))
            elif tag == qn("w:sdt"):
                # Unwrap Structured Document Tags (used for TOCs, content controls, etc.)
                sdt_content = child.find("w:sdtContent", NS)
                if sdt_content is not None:
                    blocks.extend(
                        self._parse_block_container(
                            sdt_content, rels, f"{base_path}/w:sdt/w:sdtContent", table_cond_p, table_cond_r
                        )
                    )
            elif tag == qn("mc:AlternateContent"):
                choice = child.find("mc:Choice", NS)
                if choice is None:
                    choice = child.find("mc:Fallback", NS)
                if choice is not None:
                    blocks.extend(
                        self._parse_block_container(
                            choice, rels, f"{base_path}/mc:AlternateContent", table_cond_p, table_cond_r
                        )
                    )
        return blocks

    def _parse_paragraph(
        self,
        p_el: etree._Element,
        rels: Dict[str, str],
        xml_path: str,
        table_cond_p: Optional[ParagraphStyle] = None,
        table_cond_r: Optional[RunStyle] = None,
    ) -> ParagraphBlock:
        ppr_el = p_el.find("w:pPr", NS)
        style_id: Optional[str] = None
        sect_pr: Optional[SectionProps] = None
        if ppr_el is not None:
            ps_el = ppr_el.find("w:pStyle", NS)
            if ps_el is not None:
                style_id = ps_el.get(qn("w:val"))
            s_el = ppr_el.find("w:sectPr", NS)
            if s_el is not None:
                sect_pr = self._parse_sect_pr(s_el, rels)

        # Cascade: docDefaults -> style chain -> table conditional style -> direct pPr
        resolved_p = self.resolve_paragraph_style(style_id)
        if table_cond_p is not None:
            resolved_p = resolved_p.merge_from(table_cond_p)
            # In Word, paragraphs inside tables default to 0 space_after unless explicitly set on style/pPr
            if style_id is None:
                resolved_p.space_after_pt = table_cond_p.space_after_pt or 0.0
                resolved_p.line_spacing = table_cond_p.line_spacing or 1.0
        if table_cond_r is not None:
            resolved_p.run_style = resolved_p.run_style.merge_from(table_cond_r)

        direct_p = self._parse_ppr(ppr_el)
        resolved_p = resolved_p.merge_from(direct_p)

        # Resolve numbering if present
        list_marker: Optional[str] = None
        list_marker_style: Optional[RunStyle] = None
        list_tab_pt: Optional[float] = None
        if resolved_p.num_id is not None and resolved_p.num_id > 0:
            ilvl = resolved_p.num_ilvl or 0
            resolved_p, list_marker, list_marker_style, list_tab_pt = self._resolve_list_item(
                resolved_p.num_id, ilvl, resolved_p
            )

        # Parse inline runs, hyperlinks, SDTs, and fields
        runs = self._parse_inline_container(
            p_el, rels, resolved_p.run_style, xml_path, is_toc_para=self.is_toc_entry_style(style_id)
        )

        return ParagraphBlock(
            style_id=style_id,
            style=resolved_p,
            runs=runs,
            list_marker=list_marker,
            list_marker_style=list_marker_style,
            list_tab_pt=list_tab_pt,
            sect_pr=sect_pr,
            xml_path=xml_path,
        )

    def _parse_inline_container(
        self,
        container: etree._Element,
        rels: Dict[str, str],
        base_run_style: RunStyle,
        xml_path: str,
        is_toc_para: bool = False,
    ) -> List[RunItem]:
        items: List[RunItem] = []
        # Track complex field state: 'outside', 'instr', 'result'
        fld_state = "outside"
        curr_instr = ""
        r_idx = 0

        def process_children(elem: etree._Element, inherited_r: RunStyle, path_prefix: str) -> None:
            nonlocal fld_state, curr_instr, r_idx
            for child in elem:
                tag = child.tag
                if tag == qn("w:r"):
                    r_idx += 1
                    r_path = f"{path_prefix}/w:r[{r_idx}]"
                    rpr_el = child.find("w:rPr", NS)
                    r_style_id = None
                    if rpr_el is not None:
                        rs_el = rpr_el.find("w:rStyle", NS)
                        if rs_el is not None:
                            r_style_id = rs_el.get(qn("w:val"))
                    char_style = self.resolve_character_style(r_style_id)
                    # In TOC hyperlinks, Word suppresses the blue underline of the 'Hyperlink' character style
                    if is_toc_para and r_style_id and r_style_id.lower() == "hyperlink":
                        char_style.color = None
                        char_style.underline = None
                    direct_r = self._parse_rpr(rpr_el)
                    eff_r = inherited_r.merge_from(char_style).merge_from(direct_r)

                    for rchild in child:
                        rtag = rchild.tag
                        if rtag == qn("w:fldChar"):
                            fctype = rchild.get(qn("w:fldCharType"))
                            if fctype == "begin":
                                fld_state = "instr"
                                curr_instr = ""
                            elif fctype == "separate":
                                fld_state = "result"
                            elif fctype == "end":
                                fld_state = "outside"
                                curr_instr = ""
                        elif rtag == qn("w:instrText"):
                            curr_instr += (rchild.text or "")
                        elif rtag == qn("w:t"):
                            if fld_state == "instr":
                                continue
                            txt = rchild.text or ""
                            if eff_r.all_caps:
                                txt = txt.upper()
                            if fld_state == "result":
                                up_instr = curr_instr.upper()
                                if "PAGE" in up_instr and "NUMPAGES" not in up_instr and "PAGEREF" not in up_instr:
                                    items.append(RunItem(kind="page_num", text=txt, style=eff_r, xml_path=r_path))
                                    continue
                                if "NUMPAGES" in up_instr:
                                    items.append(RunItem(kind="num_pages", text=txt, style=eff_r, xml_path=r_path))
                                    continue
                            if txt:
                                items.append(RunItem(kind="text", text=txt, style=eff_r, xml_path=r_path))
                        elif rtag in (qn("w:tab"), qn("w:ptab")):
                            if fld_state != "instr":
                                items.append(RunItem(kind="tab", style=eff_r, xml_path=r_path))
                        elif rtag in (qn("w:br"), qn("w:cr")):
                            btype = rchild.get(qn("w:type"), "textWrapping")
                            items.append(RunItem(kind="break", break_type=btype, style=eff_r, xml_path=r_path))
                        elif rtag == qn("w:noBreakHyphen"):
                            items.append(RunItem(kind="text", text="-", style=eff_r, xml_path=r_path))
                        elif rtag == qn("w:sym"):
                            ch_hex = rchild.get(qn("w:char"), "")
                            sym_char = "•"
                            if ch_hex:
                                try:
                                    code = int(ch_hex, 16)
                                    sym_char = PUA_BULLET_MAP.get(chr(code), chr(code) if code < 0xF000 else "•")
                                except ValueError:
                                    pass
                            items.append(RunItem(kind="text", text=sym_char, style=eff_r, xml_path=r_path))
                        elif rtag == qn("w:drawing"):
                            drw = self._parse_drawing(rchild, rels, r_path)
                            if drw is not None:
                                items.append(RunItem(kind="drawing", drawing=drw, style=eff_r, xml_path=r_path))
                        elif rtag == qn("mc:AlternateContent"):
                            choice = rchild.find("mc:Choice", NS) or rchild.find("mc:Fallback", NS)
                            if choice is not None:
                                d_el = choice.find("w:drawing", NS)
                                if d_el is not None:
                                    drw = self._parse_drawing(d_el, rels, r_path)
                                    if drw is not None:
                                        items.append(RunItem(kind="drawing", drawing=drw, style=eff_r, xml_path=r_path))

                elif tag == qn("w:hyperlink"):
                    hl_style = inherited_r
                    if not is_toc_para:
                        hl_style = inherited_r.merge_from(self.resolve_character_style("Hyperlink"))
                    process_children(child, hl_style, f"{path_prefix}/w:hyperlink")
                elif tag == qn("w:fldSimple"):
                    instr = (child.get(qn("w:instr")) or "").upper()
                    len_before = len(items)
                    process_children(child, inherited_r, f"{path_prefix}/w:fldSimple")
                    if "PAGE" in instr and "NUMPAGES" not in instr and "PAGEREF" not in instr:
                        if len(items) == len_before:
                            items.append(RunItem(kind="page_num", text="1", style=inherited_r, xml_path=path_prefix))
                        else:
                            for it in items[len_before:]:
                                it.kind = "page_num"
                    elif "NUMPAGES" in instr:
                        if len(items) == len_before:
                            items.append(RunItem(kind="num_pages", text="1", style=inherited_r, xml_path=path_prefix))
                        else:
                            for it in items[len_before:]:
                                it.kind = "num_pages"
                elif tag == qn("w:sdt"):
                    sdt_c = child.find("w:sdtContent", NS)
                    if sdt_c is not None:
                        process_children(sdt_c, inherited_r, f"{path_prefix}/w:sdt/w:sdtContent")
                elif tag == qn("w:ins"):
                    # Render inserted text cleanly
                    process_children(child, inherited_r, f"{path_prefix}/w:ins")

        process_children(container, base_run_style, xml_path)
        return items

    def _parse_drawing(
        self, drawing_el: etree._Element, rels: Dict[str, str], xml_path: str
    ) -> Optional[DrawingItem]:
        inline = drawing_el.find("wp:inline", NS)
        anchor = drawing_el.find("wp:anchor", NS)
        target = inline if inline is not None else anchor
        if target is None:
            return None

        is_floating = anchor is not None
        ext = target.find("wp:extent", NS)
        if ext is None:
            ext = target.find(".//a:ext", NS)
        width_pt = emu_to_pt(ext.get("cx")) if ext is not None else 72.0
        height_pt = emu_to_pt(ext.get("cy")) if ext is not None else 72.0

        item = DrawingItem(
            width_pt=max(1.0, width_pt),
            height_pt=max(1.0, height_pt),
            is_floating=is_floating,
            xml_path=xml_path,
        )

        if is_floating and anchor is not None:
            item.behind_doc = anchor.get("behindDoc") in ("1", "true")
            item.dist_t_pt = emu_to_pt(anchor.get("distT"))
            item.dist_b_pt = emu_to_pt(anchor.get("distB"))
            item.dist_l_pt = emu_to_pt(anchor.get("distL") or "114300")
            item.dist_r_pt = emu_to_pt(anchor.get("distR") or "114300")

            pos_h = anchor.find("wp:positionH", NS)
            if pos_h is not None:
                item.pos_h_relative_from = pos_h.get("relativeFrom", "column")
                po = pos_h.find("wp:posOffset", NS)
                al = pos_h.find("wp:align", NS)
                if po is not None and po.text:
                    item.pos_h_offset_pt = emu_to_pt(po.text)
                if al is not None and al.text:
                    item.pos_h_align = al.text.strip()

            pos_v = anchor.find("wp:positionV", NS)
            if pos_v is not None:
                item.pos_v_relative_from = pos_v.get("relativeFrom", "paragraph")
                po = pos_v.find("wp:posOffset", NS)
                al = pos_v.find("wp:align", NS)
                if po is not None and po.text:
                    item.pos_v_offset_pt = emu_to_pt(po.text)
                if al is not None and al.text:
                    item.pos_v_align = al.text.strip()

            if anchor.find("wp:wrapTopAndBottom", NS) is not None:
                item.wrap_type = "topAndBottom"
            elif anchor.find("wp:wrapSquare", NS) is not None:
                item.wrap_type = "square"
            elif anchor.find("wp:wrapTight", NS) is not None or anchor.find("wp:wrapThrough", NS) is not None:
                item.wrap_type = "tight"
            elif anchor.find("wp:wrapNone", NS) is not None:
                item.wrap_type = "none"
            else:
                item.wrap_type = "square"

        # Extract embedded image bytes (<a:blip r:embed="..."/>)
        blip = target.find(".//a:blip", NS)
        if blip is not None:
            rid = blip.get(qn("r:embed"))
            if rid and rid in rels:
                media_path = rels[rid].lstrip("/")
                if media_path in self.zf.namelist():
                    item.image_bytes = self.zf.read(media_path)

        # Extract optional WordprocessingShape (<wps:wsp>) fill, border, and textbox content (<wps:txbx>)
        wsp = target.find(".//wps:wsp", NS)
        if wsp is not None:
            sp_pr = wsp.find("wps:spPr", NS)
            if sp_pr is not None:
                sfill = sp_pr.find("a:solidFill/a:srgbClr", NS)
                if sfill is not None and sfill.get("val"):
                    item.shape_fill = sfill.get("val")
                ln = sp_pr.find("a:ln", NS)
                if ln is not None:
                    item.shape_stroke_pt = emu_to_pt(ln.get("w") or "9525")
                    lfill = ln.find("a:solidFill/a:srgbClr", NS)
                    if lfill is not None and lfill.get("val"):
                        item.shape_stroke = lfill.get("val")
            txbx = wsp.find(".//w:txbxContent", NS)
            if txbx is not None:
                item.textbox_blocks = self._parse_block_container(txbx, rels, f"{xml_path}/w:txbxContent")

        return item

    # -------------------------------------------------------------------------
    # 7. Table Parsing & Conditional Style Resolution
    # -------------------------------------------------------------------------
    def _parse_table(self, tbl_el: etree._Element, rels: Dict[str, str], xml_path: str) -> TableBlock:
        tbl_pr = tbl_el.find("w:tblPr", NS)
        style_id: Optional[str] = None
        if tbl_pr is not None:
            ts = tbl_pr.find("w:tblStyle", NS)
            if ts is not None:
                style_id = ts.get(qn("w:val"))

        tbl_style_def = self.resolve_table_style(style_id)
        tb = TableBlock(style_id=style_id, xml_path=xml_path)

        # Apply base table style properties
        if tbl_style_def is not None:
            if tbl_style_def.tbl_borders is not None:
                tb.borders = tbl_style_def.tbl_borders
            if tbl_style_def.tbl_shading is not None:
                tb.shading_fill = tbl_style_def.tbl_shading
            tb.cell_margin_top_pt = tbl_style_def.tbl_cell_margins.get("top", tb.cell_margin_top_pt)
            tb.cell_margin_bottom_pt = tbl_style_def.tbl_cell_margins.get("bottom", tb.cell_margin_bottom_pt)
            tb.cell_margin_left_pt = tbl_style_def.tbl_cell_margins.get("left", tb.cell_margin_left_pt)
            tb.cell_margin_right_pt = tbl_style_def.tbl_cell_margins.get("right", tb.cell_margin_right_pt)

        # Look flags for conditional formatting
        look_first_row = True
        look_last_row = False
        look_first_col = True
        look_last_col = False
        look_no_hband = False
        look_no_vband = True

        if tbl_pr is not None:
            jc = tbl_pr.find("w:jc", NS)
            if jc is not None:
                val = jc.get(qn("w:val"), "left")
                tb.alignment = "right" if val in ("right", "end") else ("center" if val == "center" else "left")

            ind = tbl_pr.find("w:tblInd", NS)
            if ind is not None and ind.get(qn("w:w")) is not None:
                tb.indent_left_pt = twips_to_pt(ind.get(qn("w:w")))

            tw = tbl_pr.find("w:tblW", NS)
            if tw is not None:
                wtype = tw.get(qn("w:type"), "auto")
                wval = tw.get(qn("w:w"), "0")
                tb.width_type = wtype
                if wtype == "dxa":
                    tb.width_pt = twips_to_pt(wval)
                elif wtype == "pct":
                    try:
                        pct_raw = float(wval.rstrip("%"))
                        tb.width_pct = (pct_raw / 100.0) if "%" in str(wval) else (pct_raw / 5000.0)
                    except ValueError:
                        tb.width_pct = 1.0

            tlay = tbl_pr.find("w:tblLayout", NS)
            if tlay is not None and tlay.get(qn("w:type")):
                tb.layout_type = tlay.get(qn("w:type"))

            tbdr = tbl_pr.find("w:tblBorders", NS)
            if tbdr is not None:
                direct_bdr = self._parse_borders(tbdr)
                if direct_bdr is not None:
                    for side in ("top", "bottom", "left", "right", "insideH", "insideV"):
                        spec = getattr(direct_bdr, side)
                        if spec.val != "none" or tbdr.find(f"w:{side}", NS) is not None:
                            setattr(tb.borders, side, spec)

            tshd = tbl_pr.find("w:shd", NS)
            if tshd is not None:
                tb.shading_fill = self._resolve_color_elem(tshd, "w:fill") or tb.shading_fill

            tcm = tbl_pr.find("w:tblCellMar", NS)
            if tcm is not None:
                for side, attr_name in (
                    ("top", "cell_margin_top_pt"),
                    ("bottom", "cell_margin_bottom_pt"),
                    ("left", "cell_margin_left_pt"),
                    ("start", "cell_margin_left_pt"),
                    ("right", "cell_margin_right_pt"),
                    ("end", "cell_margin_right_pt"),
                ):
                    m_el = tcm.find(f"w:{side}", NS)
                    if m_el is not None and m_el.get(qn("w:w")) is not None:
                        setattr(tb, attr_name, twips_to_pt(m_el.get(qn("w:w"))))

            tlook = tbl_pr.find("w:tblLook", NS)
            if tlook is not None:
                if tlook.get(qn("w:firstRow")) is not None:
                    look_first_row = tlook.get(qn("w:firstRow")) in ("1", "true")
                if tlook.get(qn("w:lastRow")) is not None:
                    look_last_row = tlook.get(qn("w:lastRow")) in ("1", "true")
                if tlook.get(qn("w:firstColumn")) is not None:
                    look_first_col = tlook.get(qn("w:firstColumn")) in ("1", "true")
                if tlook.get(qn("w:lastColumn")) is not None:
                    look_last_col = tlook.get(qn("w:lastColumn")) in ("1", "true")
                if tlook.get(qn("w:noHBand")) is not None:
                    look_no_hband = tlook.get(qn("w:noHBand")) in ("1", "true")
                if tlook.get(qn("w:noVBand")) is not None:
                    look_no_vband = tlook.get(qn("w:noVBand")) in ("1", "true")

        # Table grid columns (<w:tblGrid>)
        tgrid = tbl_el.find("w:tblGrid", NS)
        if tgrid is not None:
            for gcol in tgrid.findall("w:gridCol", NS):
                tb.grid_cols_pt.append(twips_to_pt(gcol.get(qn("w:w"))))

        # Collect <w:tr> rows (unwrapping SDTs if any)
        tr_elems: List[etree._Element] = []
        for child in tbl_el:
            if child.tag == qn("w:tr"):
                tr_elems.append(child)
            elif child.tag == qn("w:sdt"):
                sc = child.find("w:sdtContent", NS)
                if sc is not None:
                    tr_elems.extend(sc.findall("w:tr", NS))

        num_rows = len(tr_elems)
        for r_idx, tr_el in enumerate(tr_elems):
            tr_path = f"{xml_path}/w:tr[{r_idx + 1}]"
            tr_pr = tr_el.find("w:trPr", NS)
            row = TableRow(xml_path=tr_path)
            if tr_pr is not None:
                tr_h = tr_pr.find("w:trHeight", NS)
                if tr_h is not None:
                    row.height_pt = twips_to_pt(tr_h.get(qn("w:val")))
                    row.height_rule = tr_h.get(qn("w:hRule"), "auto")
                if tr_pr.find("w:tblHeader", NS) is not None:
                    row.is_header = parse_bool_attr(tr_pr.find("w:tblHeader", NS))
                if tr_pr.find("w:cantSplit", NS) is not None:
                    row.cant_split = parse_bool_attr(tr_pr.find("w:cantSplit", NS))
                gb = tr_pr.find("w:gridBefore", NS)
                if gb is not None and gb.get(qn("w:val")):
                    row.grid_before = int(gb.get(qn("w:val")))
                wb = tr_pr.find("w:wBefore", NS)
                if wb is not None and wb.get(qn("w:w")):
                    row.w_before_pt = twips_to_pt(wb.get(qn("w:w")))

            tc_elems: List[etree._Element] = []
            for cchild in tr_el:
                if cchild.tag == qn("w:tc"):
                    tc_elems.append(cchild)
                elif cchild.tag == qn("w:sdt"):
                    sc = cchild.find("w:sdtContent", NS)
                    if sc is not None:
                        tc_elems.extend(sc.findall("w:tc", NS))

            num_cols = len(tc_elems)
            for c_idx, tc_el in enumerate(tc_elems):
                tc_path = f"{tr_path}/w:tc[{c_idx + 1}]"
                cell = self._parse_table_cell(
                    tc_el,
                    rels,
                    tc_path,
                    tb,
                    tbl_style_def,
                    r_idx,
                    num_rows,
                    c_idx,
                    num_cols,
                    look_first_row or row.is_header,
                    look_last_row,
                    look_first_col,
                    look_last_col,
                    look_no_hband,
                    look_no_vband,
                )
                row.cells.append(cell)

            tb.rows.append(row)

        return tb

    def _parse_table_cell(
        self,
        tc_el: etree._Element,
        rels: Dict[str, str],
        tc_path: str,
        tb: TableBlock,
        tbl_style_def: Optional[StyleDef],
        r_idx: int,
        num_rows: int,
        c_idx: int,
        num_cols: int,
        look_first_row: bool,
        look_last_row: bool,
        look_first_col: bool,
        look_last_col: bool,
        look_no_hband: bool,
        look_no_vband: bool,
    ) -> TableCell:
        cell = TableCell(xml_path=tc_path)

        # Default borders from table borders (outer vs insideH/insideV)
        cell.borders = Borders(
            top=tb.borders.top if r_idx == 0 else tb.borders.insideH,
            bottom=tb.borders.bottom if r_idx == num_rows - 1 else tb.borders.insideH,
            left=tb.borders.left if c_idx == 0 else tb.borders.insideV,
            right=tb.borders.right if c_idx == num_cols - 1 else tb.borders.insideV,
        )
        cell.shading_fill = tb.shading_fill

        # Resolve conditional table style formatting
        cond_p = ParagraphStyle()
        cond_r = RunStyle()
        if tbl_style_def is not None:
            cond_p = cond_p.merge_from(tbl_style_def.p_style)
            cond_r = cond_r.merge_from(tbl_style_def.r_style)

            active_conds: List[str] = []
            if not look_no_hband:
                data_r = r_idx - (1 if look_first_row else 0)
                if data_r >= 0:
                    active_conds.append("band1Horz" if data_r % 2 == 0 else "band2Horz")
            if not look_no_vband:
                data_c = c_idx - (1 if look_first_col else 0)
                if data_c >= 0:
                    active_conds.append("band1Vert" if data_c % 2 == 0 else "band2Vert")
            if look_first_col and c_idx == 0:
                active_conds.append("firstCol")
            if look_last_col and c_idx == num_cols - 1:
                active_conds.append("lastCol")
            if look_first_row and r_idx == 0:
                active_conds.append("firstRow")
            if look_last_row and r_idx == num_rows - 1:
                active_conds.append("lastRow")

            for cname in active_conds:
                if cname in tbl_style_def.cond_formats:
                    cp, cr, cbdr, cshd = tbl_style_def.cond_formats[cname]
                    cond_p = cond_p.merge_from(cp)
                    cond_r = cond_r.merge_from(cr)
                    if cshd is not None:
                        cell.shading_fill = cshd
                    if cbdr is not None:
                        for side in ("top", "bottom", "left", "right"):
                            spec = getattr(cbdr, side)
                            if spec.val != "none":
                                setattr(cell.borders, side, spec)

        # Direct cell properties (<w:tcPr>)
        tc_pr = tc_el.find("w:tcPr", NS)
        if tc_pr is not None:
            tcw = tc_pr.find("w:tcW", NS)
            if tcw is not None:
                wtype = tcw.get(qn("w:type"), "dxa")
                wval = tcw.get(qn("w:w"), "0")
                cell.width_type = wtype
                if wtype == "dxa":
                    cell.width_pt = twips_to_pt(wval)
                elif wtype == "pct":
                    try:
                        pct_raw = float(str(wval).rstrip("%"))
                        cell.width_pct = (pct_raw / 100.0) if "%" in str(wval) else (pct_raw / 5000.0)
                    except ValueError:
                        pass

            gs = tc_pr.find("w:gridSpan", NS)
            if gs is not None and gs.get(qn("w:val")):
                try:
                    cell.grid_span = max(1, int(gs.get(qn("w:val"))))
                except ValueError:
                    pass

            vm = tc_pr.find("w:vMerge", NS)
            if vm is not None:
                cell.v_merge = vm.get(qn("w:val"), "continue")

            va = tc_pr.find("w:vAlign", NS)
            if va is not None and va.get(qn("w:val")):
                cell.v_align = va.get(qn("w:val"))

            shd = tc_pr.find("w:shd", NS)
            if shd is not None:
                sfill = self._resolve_color_elem(shd, "w:fill")
                if sfill is not None:
                    cell.shading_fill = sfill

            tcbdr = tc_pr.find("w:tcBorders", NS)
            if tcbdr is not None:
                dbdr = self._parse_borders(tcbdr)
                if dbdr is not None:
                    for side in ("top", "bottom", "left", "right"):
                        spec = getattr(dbdr, side)
                        if spec.val != "none" or tcbdr.find(f"w:{side}", NS) is not None:
                            setattr(cell.borders, side, spec)

            tcmar = tc_pr.find("w:tcMar", NS)
            if tcmar is not None:
                for side, attr_name in (
                    ("top", "margin_top_pt"),
                    ("bottom", "margin_bottom_pt"),
                    ("left", "margin_left_pt"),
                    ("start", "margin_left_pt"),
                    ("right", "margin_right_pt"),
                    ("end", "margin_right_pt"),
                ):
                    m_el = tcmar.find(f"w:{side}", NS)
                    if m_el is not None and m_el.get(qn("w:w")) is not None:
                        setattr(cell, attr_name, twips_to_pt(m_el.get(qn("w:w"))))

            if tc_pr.find("w:noWrap", NS) is not None:
                cell.no_wrap = parse_bool_attr(tc_pr.find("w:noWrap", NS))

        cell.blocks = self._parse_block_container(tc_el, rels, tc_path, cond_p, cond_r)
        return cell
