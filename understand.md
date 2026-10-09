# Technical Overview - DocTeX Converter

## Project Goal
The goal of this project is to provide a simple, local tool to convert engineering reports between Microsoft Word (.docx) and LaTeX (.tex) formats. It avoids complex cloud dependencies and focuses on the core structure of academic documents.

## Technical Depth

### 1. Document Parsing
- **DOCX Parsing**: Uses `python-docx` to iterate through the XML-based structure of Word files. It walks the document body for paragraphs (`w:p`), tables (`w:tbl`) and content controls (`w:sdt`, e.g. the table of contents). Inside a paragraph it reads runs, hyperlinks, tracked insertions and equations (`m:oMath`) in order. Footnotes are read straight from `footnotes.xml` because python-docx has no API for them.
- **Equations (OMML)**: Word stores equations as "Office Math Markup". `omml.py` walks that XML recursively; each element maps to one LaTeX construct (`m:f` → `\frac`, `m:sSup` → `^{}`, `m:nary` → `\sum`/`\int`, `m:d` → `\left( \right)` ...).
- **LaTeX Parsing**: Walks the source character by character, reading nested `{...}` groups and matching `\begin`/`\end` pairs. Block-level things (sections, lists, tables, figures, display math) are handled first, then the text inside each paragraph is rendered into Word runs with a recursive inline parser that tracks bold/italic/etc.
  - Before parsing, `\input` files are pasted in and simple macros (`\newcommand`, `\def`) are expanded.
  - The document is built **twice**, like running LaTeX twice: the first pass only collects label numbers and the citation order, so `\ref{}` and `\cite{}` can show the right numbers in the second pass.

### 2. Conversion Logic
- **Styles**: Maps Word heading styles (Heading 1, 2, 3) to LaTeX section levels (`\section`, `\subsection`, etc.).
- **Formatting**: Handles inline formatting (bold, italic, underline) by wrapping text in corresponding LaTeX commands.
- **Tables**: Converts Word tables into LaTeX `tabular` environments inside a `table` float, automatically calculating column alignments.
- **Lists**: Word list paragraphs (list styles or Word numbering) are grouped into `itemize`/`enumerate`, nested by indent level.
- **Images**: Extracts embedded images from DOCX files into an `images/` folder next to the `.tex` file (named `<document>-imageN`), and includes them using the `graphicx` package at the width they had in Word.
- **Captions**: A Word "Caption" paragraph is attached to the figure/table right before it (or the table right after it).
- **Numbering (LaTeX → Word)**: Counters for sections, figures, tables and equations follow LaTeX's rules (per chapter in reports, letters after `\appendix`).
- **Bibliography**: `.bib` files are read with a small BibTeX parser and formatted roughly like the `plain` style.
- **Symbols**: `symbols.py` is the shared Unicode ↔ LaTeX table (α ↔ `\alpha`, ≤ ↔ `\leq` ...).

### 3. Orchestration
The `DocTeXConverter` class acts as the central hub. It:
- Detects the conversion direction based on file extensions.
- Manages `ConversionOptions` (margins, font size, etc.).
- Expands folders into input files, collects warnings from the generators, and can compile the result to PDF with `latexmk`/`pdflatex`.

## Application Flow

1. **Input Stage**: The user provides a file via CLI or the Web UI.
2. **Setup**: The converter initializes and validates options.
3. **Execution**:
   - For **DOCX to LaTeX**: The `LatexGenerator` walks the document, extracts text/runs, escapes special characters, and builds a .tex file.
   - For **LaTeX to DOCX**: The `DocxGenerator` reads the .tex file and builds a Word document using high-level `python-docx` calls.
4. **Post-Processing**: (Optional) Bibliography extraction, image optimization and PDF compilation.
5. **Output**: The final file is saved to the specified path.

## Design Decisions
- **Loose Coupling**: The conversion logic is separated from the interface (CLI/Web), making it easy to embed in other tools.
- **Error Handling**: Custom exception classes (`ConversionError`, `InvalidFileFormatError`) provide meaningful feedback to the user.
- **Fully Tested**: Every line of the package is run by the test suite (`pytest --cov`), and CI enforces it.
- **Casual Codebase**: Comments are written in a casual, student-friendly style to make the codebase approachable for peers.
