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

## Done — multi-column sections

`<w:cols>` is now laid out, not just parsed. `SectionProps.column_bounds()` returns the `(x0, width)`
of each column, and the engine fills them one at a time.

* **Geometry.** Equal columns measure `(content_width - (num - 1) * space) / num`; a section with
  explicit `<w:col w:w=".." w:space=".."/>` children uses those widths instead. The parser reads
  `w:equalWidth` and the `<w:col>` children, which are also the authority on the column count —
  Word omits `w:num` when it writes them.
* **Flow.** `_advance_flow()` replaces `_new_page()` everywhere the engine means "ran out of room":
  it moves to the next column and only starts a page after the last one. `_new_page()` still handles
  the explicit breaks (manual page break, `w:pageBreakBefore`, section start). Paragraphs and tables
  both flow; a table taller than a column continues in the next one.
* **`<w:br w:type="column"/>`** was being treated as a line break. It now advances the flow.
* **Column set top.** A continuous section break opens a column set part-way down a page, and every
  column of that set starts level with it rather than at the top margin — otherwise the right column
  floats up above the single-column heading that preceded it on the same page. The space-after still
  pending from the previous section is consumed at the boundary so the columns align.
* **Floating objects** positioned `relativeFrom="column"` now resolve against the column. That was
  indistinguishable from `margin` while everything was single-column.
* **Balancing.** Word fills columns to the bottom margin except on the final page of a section that
  is *followed* by a continuous break, where it evens them up — that trailing break is exactly what
  the familiar "insert a continuous section break to balance columns" advice adds. The engine
  checkpoints each block, and at the end of such a section rewinds the final column set and
  re-flows it under a cap on how far down any column may run, binary-searching for the shortest cap
  the content still fits under. `used / columns` is the ideal, but lines are indivisible, so the
  search walks back up from there.
* **Re-measuring across unequal columns.** A paragraph is broken into lines against the column it
  starts in. When it flows into a column of a *different* width — a section with explicit `<w:col>`
  children — the lines not yet placed are broken again against the column they are actually going
  into. `_tokenize_paragraph` does not depend on the width, so each line records the atom it began
  at and the remainder is re-broken from there; the continuation drops the first-line indent and
  the list marker, which belong to the line the paragraph started on. Without this the narrow
  column kept the wide column's measure and ran its text clean off the page (x=677 on a 612pt
  sheet). Equal-width columns never trigger it, so the common case is untouched.
* **Resuming after a column set.** Content following a column set starts below the *tallest*
  column. The cursor sits at the bottom of whichever column was filled last — usually the shortest,
  and after balancing reliably so — so resuming there ran the next section back up over the bottom
  of the first column. The overlap was a couple of points and invisible in a baseline comparison;
  it took comparing line boxes to see it.

Covered by `tests/test_columns.py` (13 tests, which build their own fixtures — no sample in the repo
has a multi-column section). Verified non-vacuous: forcing `column_bounds()` back to a single
full-width column fails every behavioral test and leaves only the single-column baselines passing.

### Known limitations

* **Balancing is skipped when there is no rewind point.** Re-flowing needs a block to restart from:
  one that opened the final column set, or the block that crossed onto its page. A section with
  neither (its whole last set spilled in from the page before, with no block boundary on it) is
  left filled rather than balanced. The cap search is also bounded at 12 iterations, so a result
  can sit a line off the theoretical even split.
* **Tables are not re-measured across unequal columns.** Paragraphs are (see above), but a table's
  column widths are resolved once, against the column it starts in; continuing into a narrower one
  shifts it sideways without re-fitting. Equal-width columns are unaffected.
* **No Word-recorded ground truth.** Unlike pagination, these tests assert spec arithmetic and
  reading order, because a python-docx fixture carries no `<Pages>` and no `PAGEREF` results. To get
  real ground truth, open `tests`' generated fixture in Word, save it, and the records described
  above become available.

## Related

`CanvasRenderer.render_to_png_files` numbers output by `page.page_number`. If pagination changes,
rendered filenames change with it — any saved reference PNGs need regenerating.
