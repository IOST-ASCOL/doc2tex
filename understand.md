# Technical Overview - DocTeX Converter

## Project Goal
The goal of this project is to provide a simple, local tool to convert engineering reports between Microsoft Word (.docx) and LaTeX (.tex) formats. It avoids complex cloud dependencies and focuses on the core structure of academic documents.

## Technical Depth

### 1. Document Parsing
- **DOCX Parsing**: Uses `python-docx` to iterate through the XML-based structure of Word files. It specifically looks for `CT_P` (Paragraphs) and `CT_Tbl` (Tables) in the document body.
- **LaTeX Parsing**: Uses regular expressions and string analysis to identify common LaTeX environments like `\section`, `\textbf`, and `tabular`.

### 2. Conversion Logic
- **Styles**: Maps Word heading styles (Heading 1, 2, 3) to LaTeX section levels (`\section`, `\subsection`, etc.).
- **Formatting**: Handles inline formatting (bold, italic, underline) by wrapping text in corresponding LaTeX commands.
- **Tables**: Converts Word tables into LaTeX `tabular` environments inside a `table` float, automatically calculating column alignments.
- **Images**: Extracts embedded images from DOCX files, saves them to a local directory, and includes them using the `graphicx` package.

### 3. Orchestration
The `DocTeXConverter` class acts as the central hub. It:
- Detects the conversion direction based on file extensions.
- Manages `ConversionOptions` (margins, font size, etc.).
- Handles temporary files and directories.

## Application Flow

1. **Input Stage**: The user provides a file via CLI or the Web UI.
2. **Setup**: The converter initializes and validates options.
3. **Execution**:
   - For **DOCX to LaTeX**: The `LatexGenerator` walks the document, extracts text/runs, escapes special characters, and builds a .tex file.
   - For **LaTeX to DOCX**: The `DocxGenerator` reads the .tex file and builds a Word document using high-level `python-docx` calls.
4. **Post-Processing**: (Optional) Bibliography extraction and image optimization.
5. **Output**: The final file is saved to the specified path.

## Design Decisions
- **Loose Coupling**: The conversion logic is separated from the interface (CLI/Web), making it easy to embed in other tools.
- **Error Handling**: Custom exception classes (`ConversionError`, `InvalidFileFormatError`) provide meaningful feedback to the user.
- **Casual Codebase**: Comments are written in a casual, student-friendly style to make the codebase approachable for peers.
