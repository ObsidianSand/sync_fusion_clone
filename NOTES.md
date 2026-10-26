# Notes

## Resolved — the "we render 88 pages, Word renders 83" gap did not exist

**Status:** closed. Our pagination of the sample document matches Word's own record of it exactly.
There was no 5-page gap and no height overestimate to chase; the measurement was wrong.

### What the number actually was

The old `tools_compare_word_pagination.py` counted `<w:lastRenderedPageBreak/>` markers and got 83
pages. Two better records of Word's own layout live in the same file and both say 88:

| Word's record | says | vintage |
|---|---|---|
| `docProps/app.xml` `<Pages>` | **88 pages** | Word's last save |
| 118 TOC / ToT / ToF `PAGEREF` field results | **every anchor on the page we put it on** | Word's last field update |
| `<w:lastRenderedPageBreak/>` markers | 83 pages | stale — predates the field update |

We render 88 pages, and all 118 headings, table captions and figure captions land on exactly the
page Word recorded for them — including the last one, "Table 50 → page 88". Word refreshes the
`lastRenderedPageBreak` markers only when it actually repaginates a region, so they can predate the
last field update, and here they do.

### Why the boundary score was also misleading

The old tool scored 23/78 body boundaries. Of the 55 "disagreements":

* **42 were one artifact.** Word puts its page-start marker on the first *run* of the new page. A
  paragraph whose only content is a manual page break (`w:p[170]`, and 51 more like it) has no run
  after the break, so the marker lands in the *next* paragraph while the break paragraph's own mark
  is genuinely the first line on the new page. All 42 off-by-one-paragraph cases are exactly this,
  verified individually — not one was a layout difference.
* **1 was the tool's own bug.** It identified a page's first element by the last segment of the bbox
  XPath, so a table cell's first paragraph (`.../w:tbl[22]/w:tr[1]/w:tc[1]/w:p[1]`) was reported as
  body paragraph `w:p[1]`. `word_truth._body_label()` now matches the whole path.
* **the rest were index shift.** The score zipped the two sequences positionally, so one extra entry
  in the stale marker list made every later position disagree.

Normalizing the page-break artifact, 80 of Word's 82 stale markers are also page starts of ours.
The two that are not (`w:tbl[20]`, `w:tbl[26]`) sit in the stale data and cannot be scored.

### How to measure it now

```bash
.venv\Scripts\python.exe tools_compare_word_pagination.py
.venv\Scripts\python.exe inspect_render.py --pages 1      # same thing as Check 5
```

`word_truth.py` reads all three of Word's records and documents which to trust. **Score against the
`PAGEREF` anchors, not the markers.** The anchor score is a sharp regression detector: one extra
page anywhere in the front matter shifts every later anchor, so it collapses from 118/118 to 0/118.
Verified sensitive — a +1 pt bump to `_compute_line_height` gives 90 pages and 0/118. (A +0.5 pt
bump still gives 88 pages and 118/118: most pages in this document end on a manual page break with
300–500 pt of unused space, so they do not respond to small height changes at all.)

### Dead ends — do not re-litigate without a measurement

* **"Word compresses inter-paragraph spacing at a page boundary."** Implemented in
  `_layout_paragraph_on_page` and measured: it moved **zero** page boundaries. Not committed.
* **"A few points of per-page height overestimate is pushing a trailing paragraph over."** This was
  the standing hypothesis while the gap was believed real. It is wrong for this document: the pages
  that break are nowhere near full. Measured unused space at the bottom of the 87 non-final pages:
  min 7.9 pt, **median 357.9 pt**, max 512.3 pt — 72 of 87 pages end with over 200 pt to spare, and
  only 7 are within 30 pt of full. They break on manual page breaks, not on overflow, so no
  plausible per-page height error could move them.
* **Multi-column sections.** Also ruled out as a cause here: every `w:sectPr` in the sample has a
  bare `<w:cols w:space="720"/>` with no `w:num`, i.e. one column. ("2 Columns" in the filename
  refers to the tables.)

---

## Open — multi-column sections are parsed but not laid out

`SectionProps.cols_num` and `cols_space_pt` are parsed in `parser.py:876` and never read by
`layout.py`. A document with a genuine `<w:cols w:num="2"/>` section would silently lay out
single-column, at the wrong measure and the wrong page count. The sample document does not exercise
this, so it is latent, not a live bug — and nothing in the checks would catch it today.

## Related

`CanvasRenderer.render_to_png_files` numbers output by `page.page_number`. If pagination changes,
rendered filenames change with it — any saved reference PNGs need regenerating.
