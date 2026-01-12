# doc2tex - Simple DOCX ↔ LaTeX Converter

Hey! This is a simple document converter I built for my engineering project. It helps convert Microsoft Word (.docx) files to LaTeX (.tex) and vice versa. It's meant to be run locally.

## Features
- Convert DOCX to LaTeX (preserves headings, bold/italic, tables)
- Convert LaTeX back to DOCX
- Handle simple tables and images
- Local processing (no data sent to cloud)
- Both CLI and Web interfaces

## Requirements
You'll need Python 3 and a few libraries:
- `python-docx` (for word docs)
- `PIL` (for images)
- `Flask` (for the web UI)

Install them using:
```bash
pip install -r requirements.txt
```

## How to use

### 1. Simple Command Line (CLI)
To convert a file:
```bash
python cli.py my_document.docx
```
This will create `my_document.tex` in the same folder.

### 2. Web Interface
If you prefer a UI, just run:
```bash
python web.py
```
Then open `http://localhost:5000` in your browser.

## Project Structure
- `doc2tex/`: Core logic
  - `latex.py`: DOCX to LaTeX code
  - `docx.py`: LaTeX to DOCX code
  - `converter.py`: Orchestrator
- `web.py`: The Flask server
- `cli.py`: The terminal tool
- `templates/` & `static/`: Frontend for the web UI

## Notes
- This is a student project, so it might not handle very complex LaTeX packages.
- Images need to be in the folder where the .tex file is.
- For bibliography, it tries to extract entries but it's basic.

Hope this helps with your reports! 🚀
