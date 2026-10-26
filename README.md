# `docx_renderer` — Offline OpenXML (`.docx`) to PNG Layout & Rendering Engine

A self-contained, pure-Python OpenXML (`.docx`) layout and 2D canvas rendering engine designed for air-gapped Linux environments (such as SEMOSS MCP) without MS Word or LibreOffice.

Built to give LLMs a high-fidelity visual feedback loop when editing raw OpenXML (`word/document.xml`, `styles.xml`, `numbering.xml`).

---

## Features
- **Direct OpenXML Parsing (`lxml`)**: Parses `document.xml`, `styles.xml`, `theme1.xml`, `numbering.xml`, `settings.xml`, headers/footers, and `<w:sdt>` Structured Document Tags (Table of Contents).
- **Full Style & Theme Cascade**: Resolves `docDefaults` $\rightarrow$ `basedOn` style chains $\rightarrow$ conditional table styles (`<w:tblStylePr>`) $\rightarrow$ list levels $\rightarrow$ direct `pPr`/`rPr` formatting, including HSL `themeTint`/`themeShade` modifiers.
- **Exact Word Box & Flow Layout (Twips & EMUs $\rightarrow$ Points)**:
  - Page sizes, margins, section breaks, and dynamic `PAGE` / `NUMPAGES` header/footer fields.
  - Paragraph indentations (`left`, `right`, `firstLine`, `hanging`) and ECMA-376 spec spacing rules.
  - Custom `<w:tabs>` (`left`, `center`, `right`, `decimal`) and default `0.5"` (`36 pt`) tab stops.
  - Contiguous dot, hyphen, and underscore tab leaders (`leader="dot"`) for Tables of Contents.
  - Multi-level numbered and bulleted lists (`numbering.xml`) with automatic PUA Wingdings/Symbol bullet mapping.
  - Complex `<w:tbl>` tables with `<w:tblGrid>` column resolution, horizontal spanning (`gridSpan`), vertical merging (`vMerge`), cell margins, shading, borders, and repeating header rows.
  - Inline (`<wp:inline>`) and floating (`<wp:anchor>`) `<w:drawing>` figures and textboxes.
- **LLM Visual Grounding Mode**:
  - `--ruler`: Draws a top `0.5"` tab-stop ruler (`0"`, `T1`, `1"`, `T3`, ...) and vertical tab grid lines across the page so an LLM can immediately spot indentation/tab-stop mismatches.
  - `--bboxes`: Draws color-coded bounding boxes around paragraphs, tables, cells, list markers, and images.
  - `--json`: Exports a companion JSON file mapping every XML element XPath to its exact page bounding box and indentation in `0.5"` tab units.

---

## Installation

```bash
python -m venv .venv
# Windows
.\.venv\Scripts\python.exe -m pip install -e .
# Linux / macOS
./.venv/bin/pip install -e .
```

---

## Usage

### Command Line
```bash
# Standard clean render (150 DPI PNGs)
python -m docx_renderer "Table of Contents template example.docx" -o output_pages --prefix page --dpi 150

# Only the pages you care about: '3', '1-4', '1,3,5-7', '8-' (page 8 to the end)
python -m docx_renderer "report.docx" -o output_pages --pages 1-4,9

# LLM Visual Inspection Mode (with 0.5" tab ruler, bounding boxes, and XPath layout JSON)
python -m docx_renderer "Table of Contents template example.docx" -o output_pages --prefix page_llm --dpi 150 --ruler --bboxes --json
```

### Python API
```python
from docx_renderer import render_docx_to_png

png_paths = render_docx_to_png(
    docx_path="Table of Contents template example.docx",
    output_dir="output_pages",
    prefix="page",
    dpi=150,
    show_ruler=True,
    show_bboxes=False,
    export_layout_json=True,
    pages="1-4,9",   # str spec, int, iterable of ints, or None for every page
)
```

`pages` only skips rasterization — the slow step. The document is always parsed and laid out in
full, so pagination, headers/footers and `PAGE` / `NUMPAGES` fields are identical to a full render,
and each PNG is named after the page it actually is (`page_9.png` is page 9).

---

## Inspecting a render

Two companion entry points render a document and then assert on the engine's own layout model
(resolved run colors, emitted tab leaders) rather than on pixels, so the checks survive a change
of DPI.

```bash
# Notebook: renders, shows pages inline, runs the checks.
# Open inspect_render.ipynb in VS Code and select the project's .venv as the kernel.
# (VS Code will offer to install ipykernel — the only package beyond the venv that it needs.)

# Script: same checks, no Jupyter required.
.venv\Scripts\python.exe inspect_render.py                      # every page + checks
.venv\Scripts\python.exe inspect_render.py --pages 1-4          # just the front matter
.venv\Scripts\python.exe inspect_render.py --pages 1,4,9-10 --open 1 4
.venv\Scripts\python.exe inspect_render.py --pages 1-4 --ruler --bboxes --json
.venv\Scripts\python.exe inspect_render.py --docx other.docx --dpi 150
```

The checks cover Word behaviors that are easy to regress:

- **TOC entries are not hyperlinks.** Word writes `<w:rStyle w:val="Hyperlink"/>` on every entry
  of a `TOC \h` field but renders the entry in its paragraph style's color. `DocxParser.is_toc_entry_style()`
  recognizes the whole TOC family — `toc 1`..`toc 9`, `table of figures`, `table of authorities`,
  and any style based on them — so a `TOC \c "Table"` list does not come out blue and underlined.
- **TOC dot leaders survive.** TOC styles routinely put their right dot-leader tab stop *outside*
  their own right indent. Word measures tab stops from the left margin and honors them anyway;
  the line breaker does the same, instead of wrapping the page number onto its own line.
- **Table cells are not filled black.** `<w:shd w:fill="auto"/>` is what Word writes for an
  unshaded cell. `auto` defers the choice to the renderer and the two shading attributes resolve
  differently: automatic *text* (`w:color`) is black, an automatic *fill* (`w:fill`) is no fill
  at all.
- **Table rows are not padded to a minimum.** Word sizes a row to its tallest cell, with
  `w:trHeight` as the only floor. A fixed minimum silently pads every row -- a single line of 10pt
  Arial measures 11.5pt -- and over a long table that compounds into whole extra pages, stranding
  trailing paragraphs on a page of their own.
