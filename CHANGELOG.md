# Changelog

All notable changes to this project are documented here.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [2.0.0] - 2026-10-09

### DOCX → LaTeX
- Word equations are converted to LaTeX math (they used to be dropped). Fractions, roots,
  super/subscripts, sums/integrals, brackets, accents, matrices and aligned equations.
- Footnotes become `\footnote{}` (they used to be dropped).
- Captions are attached to their figure/table as `\caption`; pictures on their own line become `figure` floats.
- Images keep their size from Word, and images inside tables no longer break the table.
- Tables: column alignment taken from Word, long text wraps in `p{}` columns, nested tables flattened.
- `report`/`thesis`/`book` use `\chapter` for Heading 1. `thesis` writes the `report` class and
  `letter` writes `article`, because `\documentclass{thesis}` didn't compile.
- Manually typed heading numbers ("2.1 Method") are removed.
- `--extract-bib` now actually does something: the References section becomes `thebibliography`.
- Symbols like α ≤ → are written so pdflatex can compile them.
- Strikethrough, small caps, code fonts, subtitles, quotes, page breaks, line breaks,
  table of contents, content controls, internal links and tracked insertions are handled.
- Image files are named after the document (`report-image1.png`), so converting several
  documents into one folder no longer overwrites pictures.

### LaTeX → DOCX
- Sections, figures, tables and equations are numbered like LaTeX (turn off with `--no-numbering`).
- `\ref`, `\eqref`, `\autoref`, `\cref` show numbers instead of label names.
- Citations become [1], [2] from `thebibliography` or from `.bib` files, with a reference list.
- `\newcommand`, `\renewcommand`, `\def`, `\DeclareMathOperator` macros are expanded.
- `siunitx` commands (`\SI`, `\si`, `\num`, `\qty`, `\ang`, ranges).
- `\tableofcontents` inserts a real Word table of contents.
- Font size, margins and line spacing are taken from the .tex file (and from the options).
- Every `enumerate` restarts at 1.
- `\appendix` (A, B, ...), `flalign`/`alignat`, `\tag`, `\nonumber`, more citation commands.
- Fixed: a `.tex` file ending in a lone `\` crashed the converter; `\input` loops are cut off with a warning.

### CLI / Web / API
- Folders as input, `--pdf`, `--body-only`, `--image-width`, `--language`, `--package`, `--no-numbering`.
- `-o` with several inputs is now an error instead of being silently ignored.
- Warnings are printed by the CLI, returned by the web API and shown in the web page.
- Web UI: book/letter types, body-only and numbering options, version in the footer.
- Fixed: uploads named like `???.docx` were converted in the wrong direction.
- `doc2tex.convert(path, **options)` one-liner, `converter.warnings`, `compile_pdf()`, `find_inputs()`.
- `InvalidFileFormatError` is now a `ConversionError`.

### Other
- Fixed `unescape_latex`, which had tab characters where `\t...` commands were meant.
- 115 tests, 100% line coverage, enforced in CI; tested on Python 3.10 to 3.14.
- Proper packaging (`pyproject.toml`, `doc2tex` / `doc2tex-web` commands), contributing guide,
  code of conduct, security policy and issue/PR templates.

## [1.1.0]
- Package layout, CLI batch mode and the web UI.

[2.0.0]: https://github.com/IOST-ASCOL/doc2tex/releases/tag/v2.0.0
