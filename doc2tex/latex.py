# doc2tex - Word to LaTeX conversion logic
# This is the part that actually writes the .tex file.
# Note: Word's structure is a mess compared to LaTeX, so we have to do some guessing.

import os
import re
from typing import Dict, Iterator, List, Optional, Tuple
from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .omml import M_NS, omml_to_latex
from .options import ConversionOptions, DocumentType, LineSpacing
from .utils import escape_latex, logger, optimize_image
from .errors import ConversionError

# Folder (next to the .tex file) where extracted images are saved
IMAGE_DIR = 'images'

# Word numbering formats that mean "bullet list"; everything else is numbered
_BULLET_FORMATS = {'bullet', 'none'}

# Our document types -> real LaTeX classes ("thesis" isn't a standard class,
# and the letter class has no \section, so those fall back to something that compiles)
_LATEX_CLASS = {'thesis': 'report', 'letter': 'article'}

# Headings that start the reference list (only used with extract_bibliography)
_REFERENCE_HEADINGS = re.compile(
    r'^\s*(references?|bibliography|works cited|literature cited|sources)\s*:?\s*$', re.I)

# "Figure 3: ..." / "Table 2. ..." prefixes Word puts in captions (LaTeX numbers them itself)
_CAPTION_PREFIX = re.compile(r'^\s*(figure|fig\.?|table|tab\.?)\s*\d+(?:[.\-]\d+)*\s*[:.\-–—]?\s*', re.I)

# Manual numbering typed into headings, e.g. "2.1 Method" or "3) Results"
_HEADING_NUMBER = re.compile(r'^\s*(?:\d+(?:\.\d+)*|[IVX]+)[.)]?\s+(?=\S)')

_OMATH = f'{{{M_NS}}}oMath'
_OMATH_PARA = f'{{{M_NS}}}oMathPara'
_W = lambda tag: qn('w:' + tag)  # noqa: E731


class LatexGenerator:
    """
    This is my main class for turning a Word doc into LaTeX.
    I tried to make the output code clean so you don't have to fix everything
    manually like you do with Pandoc sometimes.
    """

    def __init__(self, options: ConversionOptions):
        self.options = options
        self.bib_list: List[str] = []  # Reference entries we find
        self.img_idx = 0    # Keeps track of how many images we've saved
        self.images_saved: List[str] = []
        self.warnings: List[str] = []
        self.temp_workspace = None  # kept for backwards compatibility
        self._doc: Optional[DocumentObject] = None
        self._image_dir: Optional[str] = None
        self._image_prefix = ''
        self._footnotes: Dict[str, List] = {}
        self._uses_chapters = False

    def convert(self, docx_path: str, tex_path: str) -> str:
        # The main function called by the converter
        try:
            logger.info(f"Converting {docx_path} to LaTeX...")

            # Load the actual word file
            self._doc = Document(docx_path)
            self._footnotes = self._load_footnotes()
            self._uses_chapters = self._latex_class() in ('report', 'book')

            # Images go into an 'images' folder right next to the .tex file,
            # so the output compiles as-is.
            if self.options.preserve_images:
                out_dir = os.path.dirname(os.path.abspath(tex_path))
                self._image_dir = os.path.join(out_dir, IMAGE_DIR)
                # Named after the document so several conversions can share
                # one output folder without overwriting each other's pictures
                stem = re.sub(r'[^A-Za-z0-9_-]+', '_', os.path.splitext(os.path.basename(tex_path))[0])
                self._image_prefix = (stem.strip('_') or 'doc') + '-'

            # Build the latex string
            full_latex = self._build_document(self._doc)

            # Save it to the file
            with open(tex_path, 'w', encoding=self.options.output_encoding) as target:
                target.write(full_latex)

            if self.images_saved:
                logger.info(f"Saved {len(self.images_saved)} image(s) to {self._image_dir}")

            return tex_path

        except ConversionError:
            raise
        except Exception as err:
            logger.error(f"DOCX to LaTeX failed: {err}")
            raise ConversionError(f"Conversion failed mid-way: {err}") from err

    def _warn(self, msg: str) -> None:
        self.warnings.append(msg)
        logger.warning(msg)

    def _latex_class(self) -> str:
        d_type = self.options.document_type.value
        return _LATEX_CLASS.get(d_type, d_type)

    def _build_document(self, doc: DocumentObject) -> str:
        # Puts together the preamble and the body
        output_chunks = []

        # Body first: it decides some things the preamble needs to know
        body = self._parse_body(doc)

        # 1. The Preamble (all the setup stuff)
        if self.options.include_preamble and self.options.standalone_document:
            output_chunks.append(self._make_preamble())

        # 2. Start of document
        if self.options.standalone_document:
            output_chunks.append("\\begin{document}\n")

        # 3. The actual content
        output_chunks.append(body)

        # 4. Wrap it up
        if self.options.standalone_document:
            output_chunks.append("\n\\end{document}\n")

        return '\n'.join(output_chunks)

    def _make_preamble(self) -> str:
        # This is where we set up the LaTeX packages.
        # I added some extra ones that usually help with engineering reports.
        lines = []

        if self.options.document_type in (DocumentType.THESIS, DocumentType.LETTER):
            lines.append(f"% '{self.options.document_type.value}' is written as the "
                         f"{self._latex_class()} class so it compiles everywhere")
        f_size = self.options.font_size.value
        lines.append(f"\\documentclass[{f_size}]{{{self._latex_class()}}}\n")

        # Basic character support
        if self.options.unicode_support:
            lines.append("% Support for non-english characters")
            lines.append("\\usepackage[T1]{fontenc}")
            lines.append("\\usepackage[utf8]{inputenc}")
            lines.append("\\usepackage{textcomp}")
        if self.options.language and self.options.language.lower() != 'english':
            lines.append(f"\\usepackage[{self.options.language.lower()}]{{babel}}")

        # Layout and Margins
        lines.append(f"\\usepackage[{self.options.page_margins}]{{geometry}}")

        # Images - we stick them in an 'images' subfolder for neatness
        if self.options.preserve_images:
            lines.append("\\usepackage{graphicx}")
            lines.append(f"\\graphicspath{{{{./{IMAGE_DIR}/}}}}")

        # Standard math and table packages
        lines.append("\\usepackage{amsmath, amssymb, amsfonts}")
        lines.append("\\usepackage{booktabs} % For nice professional tables")
        lines.append("\\usepackage{longtable} % In case tables are huge")
        lines.append("\\usepackage{array}")
        lines.append("\\usepackage[normalem]{ulem} % \\sout for strikethrough")

        # Line spacing (single, double, etc)
        if self.options.line_spacing != LineSpacing.SINGLE:
            lines.append("\\usepackage{setspace}")
            if self.options.line_spacing == LineSpacing.ONE_HALF:
                lines.append("\\onehalfspacing")
            elif self.options.line_spacing == LineSpacing.DOUBLE:
                lines.append("\\doublespacing")

        # Any other random packages the user asked for
        for pkg in self.options.custom_packages:
            lines.append(f"\\usepackage{{{pkg}}}")

        # hyperref should be loaded last, it patches a lot of other packages
        lines.append("\\usepackage{hyperref}")
        lines.append("\\hypersetup{colorlinks=true, linkcolor=blue, urlcolor=cyan, citecolor=blue}")

        return '\n'.join(lines) + '\n'

    # ------------------------------------------------------------------
    # Body
    # ------------------------------------------------------------------

    def _iter_blocks(self, container) -> Iterator:
        # Yields paragraphs, tables and TOC markers in order. Content controls
        # (<w:sdt>) wrap things like the cover page or table of contents.
        for el in container:
            if el.tag == _W('p'):
                yield Paragraph(el, self._doc)
            elif el.tag == _W('tbl'):
                yield Table(el, self._doc)
            elif el.tag == _W('sdt'):
                gallery = el.find(f".//{_W('docPartGallery')}")
                if gallery is not None and 'Table of Contents' in (gallery.get(_W('val')) or ''):
                    yield 'toc'
                    continue
                content = el.find(_W('sdtContent'))
                if content is not None:
                    yield from self._iter_blocks(content)

    def _parse_body(self, doc: DocumentObject) -> str:
        # Loop over every item in the document body
        # Paragraphs and Tables are the main things here.
        # Consecutive list paragraphs get collected and turned into one list.
        final_lines: List[str] = []
        pending_list: List[Tuple[Paragraph, int, str]] = []
        pending_caption: Optional[str] = None  # a caption written above a table
        last_kind = None  # 'figure' / 'table' / None: what the last block was
        in_refs = False

        def flush_list():
            if pending_list:
                final_lines.append(self._render_list(pending_list))
                pending_list.clear()

        def flush_refs():
            nonlocal in_refs
            if in_refs:
                in_refs = False
                if self.bib_list:
                    final_lines.append(self._render_bibliography())

        blocks = list(self._iter_blocks(doc.element.body))
        for idx, block in enumerate(blocks):
            if block == 'toc':
                flush_list()
                final_lines.append("\\tableofcontents\n\\clearpage")
                last_kind = None
                continue

            if isinstance(block, Table):
                flush_list()
                t_tex = self._handle_table(block, pending_caption)
                pending_caption = None
                if t_tex:
                    final_lines.append(t_tex)
                    last_kind = 'table'
                continue

            p_obj = block
            style_name = p_obj.style.name if p_obj.style is not None else ''

            # Paragraphs inside the "References" section become \bibitem entries
            if in_refs and not style_name.startswith('Heading'):
                entry = self._render_inline(p_obj).strip()
                if entry:
                    self.bib_list.append(re.sub(r'^\s*(\\?\[\d+\\?\]|\d+\.)\s*', '', entry))
                continue

            list_info = self._list_info(p_obj)
            if list_info:
                pending_list.append((p_obj, *list_info))
                last_kind = None
                continue
            flush_list()

            if style_name.startswith('Caption'):
                caption = _CAPTION_PREFIX.sub('', self._render_inline(p_obj).strip())
                if last_kind in ('figure', 'table') and final_lines:
                    env = last_kind
                    final_lines[-1] = final_lines[-1].replace(
                        f"\\end{{{env}}}", f"\\caption{{{caption}}}\n\\end{{{env}}}", 1)
                    last_kind = None
                elif idx + 1 < len(blocks) and isinstance(blocks[idx + 1], Table):
                    pending_caption = caption
                else:
                    final_lines.append(f"\\begin{{center}}\n\\textit{{{caption}}}\n\\end{{center}}")
                continue

            if style_name.startswith('Heading'):
                flush_refs()
                if self.options.extract_bibliography and _REFERENCE_HEADINGS.match(p_obj.text):
                    in_refs = True
                    continue

            p_tex = self._handle_paragraph(p_obj)
            if p_tex:
                final_lines.append(p_tex)
                last_kind = 'figure' if p_tex.startswith('\\begin{figure}') else None

        flush_list()
        flush_refs()
        return '\n\n'.join(final_lines)

    def _render_bibliography(self) -> str:
        lines = ["\\begin{thebibliography}{99}"]
        for i, entry in enumerate(self.bib_list, start=1):
            lines.append(f"\\bibitem{{ref{i}}} {entry}")
        lines.append("\\end{thebibliography}")
        return '\n'.join(lines)

    # ------------------------------------------------------------------
    # Paragraphs and inline text
    # ------------------------------------------------------------------

    def _handle_paragraph(self, para: Paragraph) -> str:
        # Turns a single paragraph into LaTeX
        style_name = para.style.name if para.style is not None else ''

        # If it's a heading/title, handle it separately
        if style_name == 'Title':
            title = escape_latex(para.text.strip())
            if not title:
                return ""
            author = (self._doc.core_properties.author or '').strip()
            author = '' if author.lower() in ('python-docx', 'author') else escape_latex(author)
            return f"\\title{{{title}}}\n\\author{{{author}}}\n\\date{{}}\n\\maketitle"
        if style_name == 'Subtitle':
            text = self._render_inline(para).strip()
            return f"\\begin{{center}}\n\\large {text}\n\\end{{center}}" if text else ""
        if style_name.startswith('Heading'):
            return self._handle_heading(para, style_name)

        clean_text = self._clean_breaks(self._render_inline(para))
        if not clean_text:
            return ""  # Ignore empty lines

        # A paragraph that is only picture(s) becomes a figure, so it can get a caption
        if '\\includegraphics' in clean_text and not re.sub(
                r'\\includegraphics\[[^\]]*\]\{[^}]*\}|\\newline|\s', '', clean_text):
            graphics = '\n'.join(re.findall(r'\\includegraphics\[[^\]]*\]\{[^}]*\}', clean_text))
            return f"\\begin{{figure}}[htbp]\n\\centering\n{graphics}\n\\end{{figure}}"

        if style_name in ('Quote', 'Intense Quote'):
            return f"\\begin{{quote}}\n{clean_text}\n\\end{{quote}}"

        # Handle alignment (Center/Right)
        align = para.alignment
        if align == WD_PARAGRAPH_ALIGNMENT.CENTER:
            clean_text = f"\\begin{{center}}\n{clean_text}\n\\end{{center}}"
        elif align == WD_PARAGRAPH_ALIGNMENT.RIGHT:
            clean_text = f"\\begin{{flushright}}\n{clean_text}\n\\end{{flushright}}"

        return clean_text

    @staticmethod
    def _clean_breaks(tex: str) -> str:
        # \newline at the very start/end of a paragraph is a LaTeX error
        tex = tex.strip()
        tex = re.sub(r'^(\\newline\s*)+', '', tex)
        tex = re.sub(r'(\s*\\newline)+$', '', tex)
        return tex.strip()

    def _inline_items(self, container, para: Paragraph) -> Iterator:
        # Runs, links and equations of a paragraph in document order.
        # Tracked insertions (<w:ins>) count as normal text, deletions are dropped.
        for el in container:
            tag = el.tag
            if tag == _W('r'):
                yield Run(el, para)
            elif tag == _W('hyperlink'):
                yield Hyperlink(el, para)
            elif tag in (_OMATH, _OMATH_PARA):
                yield el
            elif tag in (_W('ins'), _W('smartTag'), _W('fldSimple'), _W('customXml'), _W('bdo'), _W('dir')):
                yield from self._inline_items(el, para)
            elif tag == _W('sdt'):
                content = el.find(_W('sdtContent'))
                if content is not None:
                    yield from self._inline_items(content, para)

    def _render_inline(self, para: Paragraph) -> str:
        # Walks the runs, hyperlinks and equations of a paragraph in document order
        pieces = []
        for item in self._inline_items(para._p, para):
            if isinstance(item, Hyperlink):
                inner = ''.join(self._render_run(r) for r in item.runs)
                address = item.address
                if address:
                    pieces.append(f"\\href{{{self._escape_url(address)}}}{{{inner}}}")
                elif item.fragment:
                    pieces.append(f"\\hyperlink{{{item.fragment}}}{{{inner}}}")
                else:
                    pieces.append(inner)
            elif isinstance(item, Run):
                pieces.append(self._render_run(item))
            else:
                math = omml_to_latex(item)
                if not math:
                    continue
                if item.tag == _OMATH_PARA:
                    pieces.append(f"\n\\[\n{math}\n\\]\n")
                else:
                    pieces.append(f"${math}$")
        return ''.join(pieces)

    def _render_run(self, run: Run) -> str:
        # One run = a piece of text with the same formatting
        r = run._r
        txt = escape_latex(run.text).replace('\t', ' ').replace('\n', '\\newline ')

        if txt.strip():
            # Note: I'm combining them so you can have bold AND italic
            if run.bold:
                txt = f"\\textbf{{{txt}}}"
            if run.italic:
                txt = f"\\textit{{{txt}}}"
            if run.underline:
                txt = f"\\underline{{{txt}}}"
            if run.font.strike:
                txt = f"\\sout{{{txt}}}"
            if run.font.small_caps:
                txt = f"\\textsc{{{txt}}}"
            if run.font.superscript:
                txt = f"\\textsuperscript{{{txt}}}"
            elif run.font.subscript:
                txt = f"\\textsubscript{{{txt}}}"
            if self._is_code_font(run):
                txt = f"\\texttt{{{txt}}}"

        # Page breaks
        for br in r.iter(_W('br')):
            if br.get(_W('type')) == 'page':
                txt += "\n\\clearpage\n"

        # Footnotes
        for ref in r.iter(_W('footnoteReference')):
            note = self._footnote_text(ref.get(_W('id')))
            if note:
                txt += f"\\footnote{{{note}}}"

        # Images live inside runs as <w:drawing> elements
        if self.options.preserve_images:
            for blip in r.iter(qn('a:blip')):
                rid = blip.get(qn('r:embed'))
                if rid:
                    txt += self._save_image(rid, self._drawing_width(r))

        return txt

    @staticmethod
    def _is_code_font(run: Run) -> bool:
        name = (run.font.name or '').lower()
        return any(f in name for f in ('courier', 'consolas', 'mono', 'menlo', 'source code'))

    @staticmethod
    def _escape_url(url: str) -> str:
        # Inside \href only a few characters need escaping
        return url.replace('\\', '/').replace('%', '\\%').replace('#', '\\#')

    def _handle_heading(self, para: Paragraph, s_name: str) -> str:
        # Maps Word headings to LaTeX sections
        raw = _HEADING_NUMBER.sub('', para.text.strip())
        txt = escape_latex(raw)
        if not txt:
            return ""

        m = re.search(r'(\d+)', s_name)
        level = int(m.group(1)) if m else 1

        # Reports/theses/books are organised in chapters
        if self._uses_chapters:
            commands = {1: 'chapter', 2: 'section', 3: 'subsection', 4: 'subsubsection',
                        5: 'paragraph'}
        else:
            commands = {1: 'section', 2: 'subsection', 3: 'subsubsection', 4: 'paragraph'}
        return f"\\{commands.get(level, 'subparagraph')}{{{txt}}}"

    # ------------------------------------------------------------------
    # Footnotes
    # ------------------------------------------------------------------

    def _load_footnotes(self) -> Dict[str, List]:
        # python-docx has no footnote API, so read footnotes.xml ourselves
        for rel in self._doc.part.rels.values():
            if rel.reltype == RT.FOOTNOTES and not rel.is_external:
                try:
                    root = parse_xml(rel.target_part.blob)
                except Exception as e:
                    logger.debug(f"Could not read footnotes: {e}")
                    return {}
                return {fn.get(_W('id')): fn.findall(_W('p')) for fn in root.findall(_W('footnote'))}
        return {}

    def _footnote_text(self, fid: Optional[str]) -> str:
        paras = self._footnotes.get(fid) or []
        texts = []
        for p in paras:
            # Reuse the normal run rendering so bold/italic/links survive
            texts.append(self._render_inline(Paragraph(p, self._doc)).strip())
        return ' '.join(t for t in texts if t)

    # ------------------------------------------------------------------
    # Lists
    # ------------------------------------------------------------------

    def _list_info(self, para: Paragraph) -> Optional[Tuple[int, str]]:
        # Returns (level, 'itemize'|'enumerate') if the paragraph is a list item
        style_name = para.style.name if para.style is not None else ''
        if style_name.startswith('Heading'):
            return None  # numbered headings are still headings

        pPr = para._p.pPr
        numPr = pPr.numPr if pPr is not None else None
        if numPr is not None and numPr.numId is not None and numPr.numId.val != 0:
            level = numPr.ilvl.val if numPr.ilvl is not None else 0
            fmt = self._numbering_format(numPr.numId.val, level)
            if fmt is None:
                fmt = 'enumerate' if 'Number' in style_name else 'itemize'
            return level, fmt

        # Fall back to the built-in list styles ("List Bullet 2", "List Number" ...)
        m = re.match(r'List (Bullet|Number)(?: (\d))?$', style_name)
        if m:
            level = int(m.group(2)) - 1 if m.group(2) else 0
            return level, 'itemize' if m.group(1) == 'Bullet' else 'enumerate'
        return None

    def _numbering_format(self, num_id: int, level: int) -> Optional[str]:
        # Looks up whether a Word numbering definition is bullets or numbers
        try:
            numbering = self._doc.part.numbering_part.element
        except (KeyError, NotImplementedError, AttributeError):
            return None

        num = numbering.find(f"{qn('w:num')}[@{qn('w:numId')}='{num_id}']")
        if num is None:
            return None
        abstract_ref = num.find(qn('w:abstractNumId'))
        if abstract_ref is None:
            return None
        abstract_id = abstract_ref.get(qn('w:val'))
        abstract = numbering.find(
            f"{qn('w:abstractNum')}[@{qn('w:abstractNumId')}='{abstract_id}']"
        )
        if abstract is None:
            return None
        lvl = abstract.find(f"{qn('w:lvl')}[@{qn('w:ilvl')}='{level}']")
        if lvl is None:
            return None
        num_fmt = lvl.find(qn('w:numFmt'))
        if num_fmt is None:
            return None
        return 'itemize' if num_fmt.get(qn('w:val')) in _BULLET_FORMATS else 'enumerate'

    def _render_list(self, items: List[Tuple[Paragraph, int, str]]) -> str:
        # Turns a run of list paragraphs into (possibly nested) itemize/enumerate
        lines = []
        stack: List[Tuple[int, str]] = []  # open environments: (level, kind)

        for para, level, kind in items:
            # Close deeper lists, or a same-level list of a different kind
            while stack and (stack[-1][0] > level or
                             (stack[-1][0] == level and stack[-1][1] != kind)):
                lines.append('  ' * (len(stack) - 1) + f"\\end{{{stack.pop()[1]}}}")
            if not stack or stack[-1][0] < level:
                lines.append('  ' * len(stack) + f"\\begin{{{kind}}}")
                stack.append((level, kind))
            text = self._clean_breaks(self._render_inline(para))
            lines.append('  ' * len(stack) + f"\\item {text}")

        while stack:
            lines.append('  ' * (len(stack) - 1) + f"\\end{{{stack.pop()[1]}}}")
        return '\n'.join(lines)

    # ------------------------------------------------------------------
    # Images
    # ------------------------------------------------------------------

    def _drawing_width(self, r) -> Optional[str]:
        # Uses the size the picture has in Word, as a fraction of the text width
        extent = r.find(f".//{qn('wp:extent')}")
        if extent is None:
            return None
        try:
            cx = int(extent.get('cx'))
            section = self._doc.sections[0]
            text_width = section.page_width - section.left_margin - section.right_margin
            frac = min(cx / text_width, 1.0)
        except (TypeError, ValueError, ZeroDivisionError, AttributeError):
            return None
        return f"{frac:.2f}\\textwidth" if frac > 0.05 else None

    def _save_image(self, rid: str, width: Optional[str] = None) -> str:
        # Writes an embedded picture to the images folder and returns the
        # \includegraphics command for it
        try:
            part = self._doc.part.related_parts[rid]
        except KeyError:
            self._warn(f"Image reference {rid} not found in document")
            return ""

        ext = os.path.splitext(str(part.partname))[1].lower() or '.png'
        if ext not in ('.png', '.jpg', '.jpeg', '.pdf'):
            # pdflatex can't read things like .emf/.wmf/.gif - convert to png if we can
            ext = '.png'

        os.makedirs(self._image_dir, exist_ok=True)
        self.img_idx += 1
        name = f"{self._image_prefix}image{self.img_idx}{ext}"
        path = os.path.join(self._image_dir, name)

        with open(path, 'wb') as f:
            f.write(part.blob)

        if ext == '.png' and not part.partname.lower().endswith('.png'):
            try:
                from PIL import Image
                with Image.open(path) as img:
                    img.save(path, 'PNG')
            except Exception as e:
                self._warn(f"Could not convert {part.partname} to PNG ({e}); "
                           f"{name} may need converting by hand")

        if self.options.optimize_images and ext in ('.png', '.jpg', '.jpeg'):
            jpg_name = f"{self._image_prefix}image{self.img_idx}.jpg"
            jpg_path = os.path.join(self._image_dir, jpg_name)
            tmp_path = jpg_path + '.tmp'
            try:
                optimize_image(path, tmp_path, max_width=1600,
                               quality=self.options.image_quality)
                os.remove(path)
                os.replace(tmp_path, jpg_path)
                name, path = jpg_name, jpg_path
            except Exception as e:
                logger.warning(f"Image optimization failed, keeping original: {e}")

        self.images_saved.append(path)
        width = width or self.options.image_max_width
        return f"\\includegraphics[width={width}]{{{name}}}"

    # ------------------------------------------------------------------
    # Tables
    # ------------------------------------------------------------------

    def _handle_table(self, tbl: Table, caption: Optional[str] = None) -> str:
        # Reconstructs tables (booktabs style, merged cells -> \multicolumn)
        rows = []
        aligns: List[List[Optional[str]]] = []
        for row in tbl.rows:
            cells = []
            row_aligns = []
            prev_tc = None
            for cell in row.cells:
                # python-docx repeats a merged cell once per grid column
                if prev_tc is not None and cell._tc is prev_tc:
                    cells[-1][1] += 1
                    continue
                prev_tc = cell._tc
                parts = []
                for block in self._iter_blocks(cell._tc):
                    if isinstance(block, Table):
                        self._warn("Nested table flattened into its cell")
                        parts.append(' '.join(c.text for r in block.rows for c in r.cells))
                    elif isinstance(block, Paragraph):
                        parts.append(self._clean_breaks(self._render_inline(block)))
                text = ' \\newline '.join(p for p in parts if p)
                cells.append([text, 1])
                para_align = cell.paragraphs[0].alignment if cell.paragraphs else None
                row_aligns.append({WD_PARAGRAPH_ALIGNMENT.CENTER: 'c',
                                   WD_PARAGRAPH_ALIGNMENT.RIGHT: 'r'}.get(para_align, 'l'))
            rows.append(cells)
            aligns.append(row_aligns)

        if not rows:
            return ""

        col_count = max(sum(span for _, span in r) for r in rows)
        if col_count == 0:
            return ""

        spec = self._column_spec(rows, aligns, col_count)
        bits = [
            "\\begin{table}[htbp]",
            "\\centering",
        ]
        if caption is not None:
            bits.append(f"\\caption{{{caption}}}")
        bits += [
            f"\\begin{{tabular}}{{{spec}}}",
            "\\toprule",
        ]

        for i, cells in enumerate(rows):
            rendered = []
            for text, span in cells:
                if span > 1:
                    rendered.append(f"\\multicolumn{{{span}}}{{c}}{{{text}}}")
                else:
                    rendered.append(text)
            # Pad short rows so LaTeX doesn't complain
            missing = col_count - sum(span for _, span in cells)
            rendered.extend([''] * missing)
            bits.append(" & ".join(rendered) + " \\\\")
            if i == 0 and len(rows) > 1:
                bits.append("\\midrule")

        bits.append("\\bottomrule")
        bits.append("\\end{tabular}")
        bits.append("\\end{table}")

        return '\n'.join(bits)

    @staticmethod
    def _column_spec(rows, aligns, col_count: int) -> str:
        # Picks l/c/r from the cell alignment in Word. If the text is long,
        # switch to fixed-width p{} columns so the table wraps instead of
        # running off the page.
        lengths = [0] * col_count
        col_align: List[Optional[str]] = [None] * col_count
        for r_idx, (cells, row_aligns) in enumerate(zip(rows, aligns)):
            c = 0
            for (text, span), al in zip(cells, row_aligns):
                if span == 1 and c < col_count:
                    plain = re.sub(r'\\[A-Za-z]+|[{}]', '', text)
                    lengths[c] = max(lengths[c], len(plain))
                    # the header row is often centred; the data rows say more
                    if r_idx > 0 or len(rows) == 1:
                        col_align[c] = col_align[c] or al
                c += span
        col_align = [a or 'l' for a in col_align]

        if sum(lengths) <= 80:
            return ''.join(col_align)

        total = sum(max(n, 4) for n in lengths)
        usable = 0.95 - 0.02 * col_count  # leave room for the column padding
        specs = []
        for n, al in zip(lengths, col_align):
            frac = max(n, 4) / total * usable
            if n <= 12:
                specs.append(al)  # short columns can stay natural width
            else:
                specs.append(f">{{\\raggedright\\arraybackslash}}p{{{frac:.2f}\\linewidth}}")
        return ''.join(specs)
