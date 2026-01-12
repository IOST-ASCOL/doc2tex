# doc2tex - Word to LaTeX conversion logic
# This is the part that actually writes the .tex file.
# Note: Word's structure is a mess compared to LaTeX, so we have to do some guessing.

import os
import re
from typing import List, Dict, Optional, Tuple, Any
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

from .options import ConversionOptions, LineSpacing
from .utils import (
    escape_latex, logger, ensure_directory, 
    optimize_image, sanitize_filename, get_temp_dir
)
from .errors import ConversionError, ImageProcessingError


class LatexGenerator:
    """
    This is my main class for turning a Word doc into LaTeX.
    I tried to make the output code clean so you don't have to fix everything 
    manually like you do with Pandoc sometimes.
    """
    
    def __init__(self, options: ConversionOptions):
        self.options = options
        self.bib_list = [] # Stores bibliography entries we find
        self.img_idx = 0   # Keeps track of how many images we've saved
        self.footer_idx = 0
        self.temp_workspace = None
        
    def convert(self, docx_path: str, tex_path: str) -> str:
        # The main function called by the converter
        try:
            logger.info(f"Wait, converting {docx_path} now...")
            
            # Load the actual word file
            my_doc = Document(docx_path)
            
            # If the user wants images, we need a place to put them
            if self.options.preserve_images:
                self.temp_workspace = get_temp_dir()
            
            # Build the latex string
            full_latex = self._build_document(my_doc, tex_path)
            
            # Save it to the file
            with open(tex_path, 'w', encoding=self.options.output_encoding) as target:
                target.write(full_latex)
            
            # Handle bibliography if we found any entries
            if self.options.extract_bibliography and self.bib_list:
                self._write_bib_file(tex_path)
            
            return tex_path
            
        except Exception as err:
            # If something breaks, at least tell the user why
            logger.error(f"Ugh, something broke: {err}")
            raise ConversionError(f"Conversion failed mid-way: {err}")
    
    def _build_document(self, doc: Document, path: str) -> str:
        # Puts together the preamble and the body
        output_chunks = []
        
        # 1. The Preamble (all the setup stuff)
        if self.options.include_preamble and self.options.standalone_document:
            output_chunks.append(self._make_preamble())
        
        # 2. Start of document
        if self.options.standalone_document:
            output_chunks.append("\\begin{document}\n")
        
        # 3. The actual content
        # I iterate through the body elements to keep the order correct
        body_tex = self._parse_body(doc, path)
        output_chunks.append(body_tex)
        
        # 4. Wrap it up
        if self.options.standalone_document:
            output_chunks.append("\n\\end{document}")
        
        return '\n'.join(output_chunks)
    
    def _make_preamble(self) -> str:
        # This is where we set up the LaTeX packages.
        # I added some extra ones that usually help with engineering reports.
        lines = []
        
        d_type = self.options.document_type.value
        f_size = self.options.font_size.value
        lines.append(f"\\documentclass[{f_size}]{{{d_type}}}\n")
        
        # Basic character support
        if self.options.unicode_support:
            lines.append("% Support for non-english characters")
            lines.append("\\usepackage[T1]{fontenc}")
            lines.append("\\usepackage[utf8]{inputenc}")
        
        # Layout and Margins
        lines.append(f"\\usepackage[{self.options.page_margins}]{{geometry}}")
        
        # Images - we stick them in an 'images' subfolder for neatness
        if self.options.preserve_images:
            lines.append("\\usepackage{graphicx}")
            lines.append("\\graphicspath{{./images/}}")
        
        # Better links (clickable)
        lines.append("\\usepackage{hyperref}")
        lines.append("\\hypersetup{colorlinks=true, linkcolor=blue, urlcolor=cyan}")
        
        # Standard math and table packages
        lines.append("\\usepackage{amsmath, amssymb, amsfonts}")
        lines.append("\\usepackage{booktabs} % For nice professional tables")
        lines.append("\\usepackage{longtable} % In case tables are huge")
        lines.append("\\usepackage{array}")
        
        # Line spacing (single, double, etc)
        if self.options.line_spacing != LineSpacing.SINGLE:
            lines.append("\\usepackage{setspace}")
            if self.options.line_spacing == LineSpacing.ONE_HALF:
                lines.append("\\onehalfspacing")
            elif self.options.line_spacing == LineSpacing.DOUBLE:
                lines.append("\\doublespacing")
        
        # Bibliography setup
        if self.options.extract_bibliography:
            lines.append("\\usepackage{natbib}")
            lines.append(f"\\bibliographystyle{{{self.options.bibliography_style}}}")
        
        # Any other random packages the user asked for
        for pkg in self.options.custom_packages:
            lines.append(f"\\usepackage{{{pkg}}}")
        
        return '\n'.join(lines) + '\n'
    
    def _parse_body(self, doc: Document, path: str) -> str:
        # Loop over every item in the document body
        # Paragraphs and Tables are the main things here
        final_lines = []
        
        for el in doc.element.body:
            # Check if it's a paragraph
            if isinstance(el, CT_P):
                p_obj = Paragraph(el, doc)
                p_tex = self._handle_paragraph(p_obj)
                if p_tex:
                    final_lines.append(p_tex)
            
            # Check if it's a table
            elif isinstance(el, CT_Tbl):
                t_obj = Table(el, doc)
                t_tex = self._handle_table(t_obj)
                if t_tex:
                    final_lines.append(t_tex)
        
        return '\n\n'.join(final_lines)
    
    def _handle_paragraph(self, para: Paragraph) -> str:
        # Turns a single line of text into LaTeX
        if not para.text.strip():
            return "" # Ignore empty lines
        
        # If it's a heading, handle it separately
        if para.style.name.startswith('Heading'):
            return self._handle_heading(para)
        
        # Break the paragraph into 'runs' (bits with different formatting)
        tex_pieces = []
        for run in para.runs:
            # Escape LaTeX symbols like % and &
            txt = escape_latex(run.text)
            
            # Apply common formatting
            # Note: I'm combining them so you can have bold AND italic
            if run.bold:
                txt = f"\\textbf{{{txt}}}"
            if run.italic:
                txt = f"\\textit{{{txt}}}"
            if run.underline:
                txt = f"\\underline{{{txt}}}"
            
            # Try to catch hyperlinks (though docx library is limited here)
            if hasattr(run, 'hyperlink') and run.hyperlink:
                link = run.hyperlink.address if hasattr(run.hyperlink, 'address') else ''
                if link:
                    txt = f"\\href{{{link}}}{{{txt}}}"
            
            tex_pieces.append(txt)
        
        clean_text = ''.join(tex_pieces)
        
        # Handle alignment (Center/Right)
        # Word calls them 'CENTER' and 'RIGHT'
        try:
            align = para.alignment
            if align == WD_PARAGRAPH_ALIGNMENT.CENTER:
                clean_text = f"\\begin{{center}}\n{clean_text}\n\\end{{center}}"
            elif align == WD_PARAGRAPH_ALIGNMENT.RIGHT:
                clean_text = f"\\begin{{flushright}}\n{clean_text}\n\\end{{flushright}}"
        except:
            # Sometimes alignment is 'None' or some weird value, just ignore it
            pass
            
        return clean_text
    
    def _handle_heading(self, para: Paragraph) -> str:
        # Maps Word headings to LaTeX sections
        txt = escape_latex(para.text)
        s_name = para.style.name
        
        # I added some 'Smart' detection for sections that usually start on new pages
        prefix = ""
        if "Heading 1" in s_name and self.options.document_type.value in ["report", "thesis"]:
             # reports usually start Heading 1 on a fresh page
             prefix = "\\clearpage\n"
             
        if 'Heading 1' in s_name:
            return f"{prefix}\\section{{{txt}}}"
        elif 'Heading 2' in s_name:
            return f"\\subsection{{{txt}}}"
        elif 'Heading 3' in s_name:
            return f"\\subsubsection{{{txt}}}"
        elif 'Heading 4' in s_name:
            return f"\\paragraph{{{txt}}}"
        else:
            return f"\\subparagraph{{{txt}}}"
    
    def _handle_table(self, tbl: Table) -> str:
        # Reconstructs tables. This is always a bit messy.
        if not tbl.rows:
            return ""
        
        # Count how many columns we need
        col_count = len(tbl.rows[0].cells)
        
        bits = []
        bits.append("\\begin{table}[h!]") # [h!] helps with positioning
        bits.append("\\centering")
        
        # Define columns (centered by default for neatness)
        col_def = "|" + "|".join(["c"] * col_count) + "|"
        bits.append(f"\\begin{{tabular}}{{{col_def}}}")
        bits.append("\\toprule") # booktabs style
        
        for i, row in enumerate(tbl.rows):
            cells = []
            for cell in row.cells:
                # Clean up text in each cell
                val = escape_latex(cell.text.strip())
                cells.append(val)
            
            # Join with & and end with \\
            row_str = " & ".join(cells) + " \\\\"
            bits.append(row_str)
            
            # Add fancy lines for headers
            if i == 0:
                bits.append("\\midrule")
            else:
                bits.append("\\hline")
        
        bits.append("\\bottomrule")
        bits.append("\\end{tabular}")
        
        # Add a placeholder caption since we can't always find it in Word
        bits.append("\\caption{Automated Table from Word}")
        bits.append("\\end{table}")
        
        return '\n'.join(bits)
    
    def _write_bib_file(self, tex_path: str) -> None:
        # Writes the references to a separate .bib file
        bib_file = tex_path.replace('.tex', '.bib')
        
        try:
            with open(bib_file, 'w', encoding=self.options.output_encoding) as b:
                for entry in self.bib_list:
                    b.write(entry + '\n\n')
            logger.info(f"Cool, created bib file at {bib_file}")
        except:
             logger.warning("Couldnt write the bib file for some reason.")
