# Open work

## Pagination fidelity — we render 88 pages, Word renders 83

**Status:** open. Four bugs fixed (see git log on `fix/toc-rendering-and-page-selection`); this is
what is left.

### Where it stands

After removing the hardcoded 14 pt minimum table row height, the sample document
(`NDA BLA XXXXXX - Safety Tables and Figures - 2 Columns RD.docx`) went from 91 pages to 88 with no
blank pages, and body page boundaries agreeing with Word went from 1/78 to 23/78. The remaining
gap is roughly 5 pages.

### How to measure it

```bash
.venv\Scripts\python.exe tools_compare_word_pagination.py
```

The `.docx` records Word's own pagination: Word writes a `<w:lastRenderedPageBreak/>` into the run
where each page began, the last time it laid the document out. That is real ground truth — no Word
install needed. The script lines those markers up against the first element on each of our
`PageLayout`s and prints page count, blank pages, and the N/78 boundary-agreement score. Use that
score as the objective function for any further layout change.

**Caveat:** it only scores from `w:p[146]` onward (the `--from` default). Word's cached markers
predate the last TOC regeneration, so the front matter — pages 1–9, the TOC / Table of Tables /
Table of Figures — does not line up and tells you nothing.

### The remaining symptom

The boundaries that still disagree are all off by exactly one paragraph, in the same direction —
our page starts one paragraph earlier than Word's:

| Word | ours |
|---|---|
| `w:p[171]` | `w:p[170]` |
| `w:p[216]` | `w:p[215]` |
| `w:p[222]` | `w:p[221]` |
| `w:p[242]` | `w:p[241]` |
| `w:p[256]` | `w:p[255]` |

and so on, re-syncing at `w:p[178]`, `189`, `204`, `231`, `297`, `332`, `367`. That pattern reads
like one more systematic height overestimate of a few points per page, pushing a short trailing
paragraph over each time. Prime suspects: line height for the `Caption` / `TableNote` styles, the
`Heading3` space-before, or table row heights still being marginally tall.

### Dead end already explored

Word compresses inter-paragraph spacing at a page boundary — a gap that would push an
otherwise-fitting first line onto the next page gets truncated at the bottom margin instead of
carried over. This was implemented in `_layout_paragraph_on_page` and measured: it changed **zero**
page boundaries on this document. It may still be correct Word behavior, but it is not what is
causing the drift here, so it was not committed. Don't re-litigate it without a test case that
actually moves the score.

### Related

`CanvasRenderer.render_to_png_files` numbers output by `page.page_number`. If pagination changes,
rendered filenames change with it — any saved reference PNGs need regenerating.
