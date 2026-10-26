"""Word's own record of how it laid a .docx out, read straight out of the file.

No Word install needed: Word leaves three different traces of its last layout in the package, and
they are *not* equally trustworthy. In decreasing order:

1. ``recorded_page_count()`` -- ``docProps/app.xml`` ``<Pages>``. Word's total page count as of its
   last save.
2. ``pageref_targets()`` -- every TOC / Table of Tables / Table of Figures entry carries a
   ``PAGEREF`` field whose cached result is the page number Word resolved that heading to at the
   last field update. That is one assertable anchor per heading, table caption and figure caption --
   on this repo's sample, 118 of them spread over the whole document.
3. ``rendered_page_starts()`` -- the ``<w:lastRenderedPageBreak/>`` markers Word writes into the run
   that begins each page. **Stale in general**, and stale in this repo's sample: it yields 83 pages
   where ``<Pages>`` and the ``PAGEREF`` results both say 88. Word only refreshes these when it
   actually repaginates a region, so they can predate the last field update. Treat them as a hint,
   never as a score, and cross-check their count against ``recorded_page_count()`` first.

Marker attribution has one more trap, which ``page_break_only_labels()`` exists to normalize: Word
puts the marker on the first *run* of the new page. A paragraph whose only content is a manual page
break has no run after the break, so the marker lands in the *following* paragraph even though the
break paragraph's own mark is the first line on the new page. Comparing raw marker positions
against the first element of each page therefore reports a one-paragraph disagreement on every such
paragraph -- 42 of them here -- that is not a layout difference at all.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Set

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

#: bboxes for body-level blocks carry exactly this prefix; cell paragraphs are deeper.
BODY_PREFIX = "word/document.xml:/w:body/"

TOC_ENTRY_STYLE = re.compile(r"(?i)^(toc\d|tableof(figures|authorities))")

_PAGE_BR = ".//{0}br[@{0}type='page']".format(W)


def _body(docx: Path) -> etree._Element:
    root = etree.fromstring(zipfile.ZipFile(docx).read("word/document.xml"))
    return root.find(W + "body")


def body_labels(body: etree._Element) -> Dict[etree._Element, str]:
    """Map each body-level paragraph/table to its ``w:p[n]`` / ``w:tbl[n]`` label (1-based)."""
    labels: Dict[etree._Element, str] = {}
    p_i = t_i = 0
    for child in body:
        if child.tag == W + "p":
            p_i += 1
            labels[child] = "w:p[{}]".format(p_i)
        elif child.tag == W + "tbl":
            t_i += 1
            labels[child] = "w:tbl[{}]".format(t_i)
    return labels


def body_label_sequence(docx: Path) -> List[str]:
    """Every body-level block label in document order."""
    body = _body(docx)
    labels = body_labels(body)
    return [labels[child] for child in body if child in labels]


def recorded_page_count(docx: Path) -> Optional[int]:
    """Word's own page count from ``docProps/app.xml``, or None if absent."""
    try:
        xml = zipfile.ZipFile(docx).read("docProps/app.xml")
    except KeyError:
        return None
    m = re.search(rb"<Pages>(\d+)</Pages>", xml)
    return int(m.group(1)) if m else None


@dataclass
class PageRefTarget:
    """One TOC entry: the page Word said its target is on, and which body block that target is."""
    entry_label: str                # the TOC entry paragraph itself, e.g. 'w:p[7]'
    style_id: str                   # 'TOC1', 'TableofFigures', ...
    bookmark: str                   # '_Toc214561493'
    word_page: int                  # page number Word cached in the PAGEREF field result
    target_label: Optional[str]     # body block holding the bookmark, e.g. 'w:p[133]'
    text: str                       # entry text, for reporting


def _cached_field_result(para: etree._Element) -> str:
    """Text of the last complete field result in a paragraph (between 'separate' and 'end')."""
    collecting, out = False, []
    for e in para.iter():
        if e.tag == W + "fldChar":
            ft = e.get(W + "fldCharType")
            if ft == "separate":
                collecting, out = True, []
            elif ft == "end":
                collecting = False
        elif e.tag == W + "t" and collecting:
            out.append(e.text or "")
    return "".join(out)


def pageref_targets(docx: Path) -> List[PageRefTarget]:
    """Every TOC entry's PAGEREF bookmark paired with the page number Word resolved it to."""
    body = _body(docx)
    labels = body_labels(body)

    bookmark_label: Dict[str, str] = {}
    for child, label in labels.items():
        for bs in child.iter(W + "bookmarkStart"):
            name = bs.get(W + "name")
            if name and name.startswith("_Toc"):
                bookmark_label.setdefault(name, label)

    targets: List[PageRefTarget] = []
    for child, label in labels.items():
        if child.tag != W + "p":
            continue
        pPr = child.find(W + "pPr")
        style_el = pPr.find(W + "pStyle") if pPr is not None else None
        style_id = style_el.get(W + "val") if style_el is not None else ""
        if not TOC_ENTRY_STYLE.match(style_id or ""):
            continue
        instr = "".join(e.text or "" for e in child.iter(W + "instrText"))
        m = re.search(r"PAGEREF\s+(_Toc\d+)", instr)
        if not m:
            continue
        text = "".join(e.text or "" for e in child.iter(W + "t"))
        # The entry's PAGEREF field result is the page number; fall back to the trailing digits of
        # the entry text for entries that keep it only in a plain run.
        num = re.search(r"(\d+)\s*$", _cached_field_result(child).strip() or text.strip())
        if not num:
            continue
        bookmark = m.group(1)
        targets.append(PageRefTarget(
            entry_label=label,
            style_id=style_id or "",
            bookmark=bookmark,
            word_page=int(num.group(1)),
            target_label=bookmark_label.get(bookmark),
            text=text,
        ))
    return targets


def rendered_page_starts(docx: Path) -> List[str]:
    """Labels of the body blocks Word marked as starting a page. Stale in general -- see module doc."""
    body = _body(docx)
    labels = body_labels(body)
    return [
        labels[child]
        for child in body
        if child in labels
        for _ in child.iter(W + "lastRenderedPageBreak")
    ]


def page_break_only_labels(docx: Path) -> Set[str]:
    """Labels of paragraphs whose only content is a manual page break.

    Word's page-start marker skips these (it needs a run to live in), so a marker comparison has to
    count 'our page starts here' and 'Word's marker is in the next paragraph' as agreement.
    """
    body = _body(docx)
    out: Set[str] = set()
    for child, label in body_labels(body).items():
        if child.tag != W + "p":
            continue
        if not child.findall(_PAGE_BR):
            continue
        if "".join(t.text or "" for t in child.iter(W + "t")).strip():
            continue
        if child.find(".//" + W + "drawing") is not None:
            continue
        out.add(label)
    return out


def our_page_of_label(pages) -> Dict[str, int]:
    """Map each body-level block label to the page number we laid it out on."""
    out: Dict[str, int] = {}
    for pg in pages:
        for bb in pg.bboxes:
            label = _body_label(bb.xml_path)
            if label is not None:
                out.setdefault(label, pg.page_number)
    return out


def our_page_starts(pages) -> List[Optional[str]]:
    """Label of the first body-level block on each page after the first."""
    starts: List[Optional[str]] = []
    for pg in pages[1:]:
        label = None
        for bb in pg.bboxes:
            label = _body_label(bb.xml_path)
            if label is not None:
                break
        starts.append(label)
    return starts


def _body_label(xml_path: Optional[str]) -> Optional[str]:
    """The ``w:p[n]`` / ``w:tbl[n]`` label of a body-level bbox path, else None.

    Matching the whole path matters: a table cell's first paragraph is also ``w:p[1]``, so keying on
    the last path segment alone silently confuses cell paragraphs with body paragraph 1.
    """
    if not xml_path or not xml_path.startswith(BODY_PREFIX):
        return None
    label = xml_path[len(BODY_PREFIX):]
    return label if "/" not in label else None
