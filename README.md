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
)
```
