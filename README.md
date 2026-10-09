# doc2tex - DOCX ↔ LaTeX Converter

[![tests](https://github.com/IOST-ASCOL/doc2tex/actions/workflows/tests.yml/badge.svg)](https://github.com/IOST-ASCOL/doc2tex/actions/workflows/tests.yml)
[![release](https://img.shields.io/github/v/release/IOST-ASCOL/doc2tex)](https://github.com/IOST-ASCOL/doc2tex/releases)
![python](https://img.shields.io/badge/python-3.10%2B-blue)
[![license](https://img.shields.io/badge/license-MIT-green)](LICENSE)

doc2tex converts documents between Microsoft Word (.docx) and LaTeX (.tex) while keeping headings, formatting, lists, tables, images, links, **equations, footnotes, captions, citations and cross references**. It's built for engineering/lab reports and theses, and it runs entirely on your own computer.

## Features
- **DOCX → LaTeX**
  - Title (+ author from the file properties), subtitle, headings
    - `report`/`thesis`/`book`: Heading 1 becomes `\chapter`
    - numbers you typed into headings ("2.1 Method") are removed, LaTeX numbers them for you
  - Bold, italic, underline, strikethrough, small caps, super/subscript, code fonts → `\texttt`
  - **Word equations → real LaTeX math** (fractions, roots, sums/integrals, matrices, accents, brackets, aligned equations ...)
  - **Footnotes** → `\footnote{}`
  - Bullet and numbered lists (including nested lists)
  - Tables (booktabs style, merged cells → `\multicolumn`, column alignment from Word, long text wraps in `p{}` columns)
  - Images are extracted to `images/` and keep the size they had in Word; a picture on its own line becomes a `figure`
  - **Captions**: Word captions become `\caption` on the matching figure/table ("Figure 1:" is stripped, LaTeX numbers it)
  - Table of contents → `\tableofcontents`, page breaks → `\clearpage`, quotes → `quote`
  - Symbols like α, ≤, →, ° are written so `pdflatex` can compile them
  - Optional: "References" section → `thebibliography` (`--extract-bib`)
  - Tracked changes: insertions are kept, deletions dropped (= "accept all")
- **LaTeX → DOCX**
  - `\section`, `\section*`, `\chapter`, `\appendix` ... **numbered like LaTeX** (1, 1.1, A.1, Chapter 2)
  - **Cross references work**: `\ref`, `\eqref`, `\autoref`, `\cref` show the right numbers
  - Numbered figure/table captions ("Figure 2.1: ...") and equation numbers, including `align` per line, `\nonumber`, `\tag`
  - **Citations**: `\cite`, `\citep`, `\citet` ... become [1], [2] using `thebibliography` **or your `.bib` file** (`\bibliography{refs}` / `\addbibresource`), with a formatted reference list
  - Your own macros: `\newcommand`, `\renewcommand`, `\def`, `\DeclareMathOperator` (with arguments)
  - `siunitx`: `\SI{3}{\kilo\metre\per\second}` → 3 km/s, `\si`, `\num`, `\qty`, `\ang`
  - `\tableofcontents` → a real Word table of contents (Word fills it in when you open the file)
  - Page size, margins (`geometry`), font size (`\documentclass[11pt]`) and line spacing (`setspace`) come from your .tex
  - Every `enumerate` restarts at 1 (Word normally keeps counting)
  - Nested formatting, `itemize`/`enumerate`/`description`, `tabular`/`longtable` with `\multicolumn`, figures with `\includegraphics` (and `\graphicspath`), inline/display math as readable text, `\href`, `\url`, footnotes, `verbatim`/`lstlisting`, `\input{}`
- Warnings for anything that didn't convert cleanly (missing images, unknown references ...)
- Local processing (no data sent to cloud), CLI + Web UI + Python API

## Installation
You'll need Python 3.10 or newer.

```bash
pip install .
```

This installs two commands: `doc2tex` (terminal) and `doc2tex-web` (browser UI).

If you'd rather not install it, you can also just install the dependencies and run the scripts directly:
```bash
pip install -r requirements.txt
python cli.py my_document.docx
python web.py
```

## How to use

### 1. Command Line (CLI)
To convert a file:
```bash
doc2tex my_document.docx
```
This will create `my_document.tex` (and an `images/` folder if the document has pictures) in the same folder.

Some more examples:
```bash
doc2tex report.tex -o report.docx               # LaTeX -> Word
doc2tex *.docx -d converted/                    # convert a whole bunch of files
doc2tex reports/ -d converted/                  # every .docx/.tex in a folder
doc2tex thesis.docx --doc-type thesis --pdf     # Word -> LaTeX -> PDF in one go
doc2tex report.docx --body-only                 # just the content, to \input into your own template
doc2tex report.docx --extract-bib               # turn the References section into \bibitem entries
doc2tex report.docx --spacing onehalf --font-size 11pt --package siunitx
doc2tex --help                                  # see every option
```

`--pdf` needs a LaTeX installation (`latexmk` or `pdflatex`, e.g. TeX Live, MiKTeX or MacTeX).

### 2. Web Interface
If you prefer a UI, just run:
```bash
doc2tex-web
```
Then open `http://localhost:5000` in your browser. When a Word file has pictures, you get a `.zip` with the `.tex` file and its `images/` folder. Anything that didn't convert cleanly is listed under the download button.

By default the server only listens on your own computer. To let others on your network use it:
```bash
DOC2TEX_HOST=0.0.0.0 DOC2TEX_PORT=8080 doc2tex-web
```
(`DOC2TEX_DEBUG=1` turns on Flask's debug mode. Only use that on your own machine, never with `DOC2TEX_HOST=0.0.0.0`.)

### 3. From Python
```python
import doc2tex

doc2tex.convert("thesis.docx", document_type="report")   # -> thesis.tex

from doc2tex import DocTeXConverter, ConversionOptions, DocumentType

converter = DocTeXConverter(ConversionOptions(document_type=DocumentType.REPORT))
converter.convert("thesis.docx")                 # -> thesis.tex
print(converter.warnings)                        # anything that didn't convert cleanly
converter.batch_convert(["a.docx", "b.tex"], output_dir="out")
converter.compile_pdf("thesis.tex")              # needs latexmk or pdflatex
```

## Running the tests
```bash
pip install ".[dev]"
python -m pytest --cov=doc2tex
```
The test suite covers every line of the package (CI fails below 100%).

## Project Structure
- `doc2tex/`: Core logic
  - `latex.py`: DOCX to LaTeX code
  - `omml.py`: Word equations to LaTeX math
  - `docx.py`: LaTeX to DOCX code (also macros, `.bib` files, siunitx)
  - `symbols.py`: Unicode ↔ LaTeX symbol tables
  - `converter.py`: Orchestrator (+ batch, folders, PDF compile)
  - `cli.py`: The terminal tool
  - `web.py`, `templates/`, `static/`: The Flask web UI
- `cli.py`, `web.py`: Small shortcuts so `python cli.py` / `python web.py` still work
- `tests/`: Tests (run with pytest)
- `CHANGELOG.md`: What changed in each version

## Notes
- This is a student project, so it won't handle very complex LaTeX (TikZ, custom environments, packages that redefine everything).
- `--doc-type thesis` writes the standard `report` class (there's no `thesis` class in plain LaTeX), and `letter` writes `article`.
- For LaTeX → DOCX, images need to be next to the `.tex` file or in a `\graphicspath` folder.
- LaTeX → DOCX math is converted to readable text, not to editable Word equations.
- Settings in your .tex (font size, margins, spacing) win over the command line options for LaTeX → DOCX.

## Contributing
Bug reports, test documents and pull requests are very welcome - see [CONTRIBUTING.md](CONTRIBUTING.md).
Please follow the [Code of Conduct](CODE_OF_CONDUCT.md), and report security issues as described in [SECURITY.md](SECURITY.md).

## License
[MIT](LICENSE) © IOST - Initiative for OpenSource Technology

Hope this helps with your reports! 🚀
