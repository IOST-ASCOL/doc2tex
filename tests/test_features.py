# Tests for the v2.0 features (equations, footnotes, captions, references,
# numbering, macros, bibliographies, CLI/web extras) and the helper functions.
# Run with: python -m pytest

import io
import os
import re
import shutil
import subprocess
import zipfile

import pytest
from docx import Document
from docx.enum.text import WD_BREAK, WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt
from PIL import Image

import doc2tex
from doc2tex import (ConversionError, ConversionOptions, DocTeXConverter, DocumentType,
                     InvalidFileFormatError, LatexCompilationError, LineSpacing)
from doc2tex import utils
from doc2tex.cli import main as cli_main
from doc2tex.docx import (expand_macros, find_env_end, format_bib_entry, parse_bib,
                          read_arg, read_group, read_optional, si_unit_text,
                          split_top_level, strip_comments)
from doc2tex.errors import ImageProcessingError, UnicodeHandlingError
from doc2tex.omml import omml_to_latex

M = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'


def omml(inner: str, para: bool = False):
    xml = f'<m:oMath {M}>{inner}</m:oMath>'
    if para:
        xml = f'<m:oMathPara {M}>{xml}</m:oMathPara>'
    return parse_xml(xml)


def r(text, plain=False):
    sty = '<m:rPr><m:sty m:val="p"/></m:rPr>' if plain else ''
    return f'<m:r>{sty}<m:t>{text}</m:t></m:r>'


def tex_of(path):
    return open(path, encoding='utf-8').read()


@pytest.fixture
def conv():
    return DocTeXConverter(ConversionOptions())


# ======================================================================
# Word equations -> LaTeX
# ======================================================================

@pytest.mark.parametrize('inner, expected', [
    (r('x'), 'x'),
    (f'<m:f><m:num>{r("a+b")}</m:num><m:den>{r("2")}</m:den></m:f>', r'\frac{a+b}{2}'),
    (f'<m:f><m:fPr><m:type m:val="lin"/></m:fPr><m:num>{r("a")}</m:num><m:den>{r("b")}</m:den></m:f>', 'a/b'),
    (f'<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>{r("n")}</m:num><m:den>{r("k")}</m:den></m:f>',
     r'\binom{n}{k}'),
    (f'<m:sSup><m:e>{r("x")}</m:e><m:sup>{r("2")}</m:sup></m:sSup>', 'x^{2}'),
    (f'<m:sSub><m:e>{r("a")}</m:e><m:sub>{r("ij")}</m:sub></m:sSub>', 'a_{ij}'),
    (f'<m:sSubSup><m:e>{r("x")}</m:e><m:sub>{r("1")}</m:sub><m:sup>{r("2")}</m:sup></m:sSubSup>', 'x_{1}^{2}'),
    (f'<m:sPre><m:sub>{r("1")}</m:sub><m:sup>{r("2")}</m:sup><m:e>{r("X")}</m:e></m:sPre>', '{}_{1}^{2}X'),
    (f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{r("x")}</m:e></m:rad>', r'\sqrt{x}'),
    (f'<m:rad><m:deg>{r("3")}</m:deg><m:e>{r("x")}</m:e></m:rad>', r'\sqrt[3]{x}'),
    (f'<m:d><m:e>{r("a")}</m:e></m:d>', r'\left( a \right)'),
    (f'<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val=""/></m:dPr><m:e>{r("a")}</m:e><m:e>{r("b")}</m:e></m:d>',
     r'\left[ a \mid b \right.'),
    (f'<m:d><m:dPr><m:sepChr m:val=","/></m:dPr><m:e>{r("a")}</m:e><m:e>{r("b")}</m:e></m:d>',
     r'\left( a , b \right)'),
    (f'<m:nary><m:naryPr><m:chr m:val="∑"/></m:naryPr><m:sub>{r("i=1")}</m:sub>'
     f'<m:sup>{r("n")}</m:sup><m:e>{r("i")}</m:e></m:nary>', r'\sum_{i=1}^{n} i'),
    (f'<m:nary><m:naryPr><m:subHide m:val="1"/><m:supHide m:val="1"/></m:naryPr>'
     f'<m:sub/><m:sup/><m:e>{r("f")}</m:e></m:nary>', r'\int f'),
    (f'<m:func><m:fName>{r("sin", True)}</m:fName><m:e>{r("x")}</m:e></m:func>', r'\sin x'),
    (f'<m:func><m:fName>{r("foo")}</m:fName><m:e>{r("x")}</m:e></m:func>', 'foo x'),
    (f'<m:limLow><m:e>{r("lim", True)}</m:e><m:lim>{r("x→0")}</m:lim></m:limLow>', r'\lim_{x\rightarrow 0}'),
    (f'<m:limUpp><m:e>{r("x")}</m:e><m:lim>{r("def")}</m:lim></m:limUpp>', r'\overset{def}{x}'),
    (f'<m:acc><m:accPr><m:chr m:val="⃗"/></m:accPr><m:e>{r("v")}</m:e></m:acc>', r'\vec{v}'),
    (f'<m:acc><m:e>{r("x")}</m:e></m:acc>', r'\hat{x}'),
    (f'<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>{r("z")}</m:e></m:bar>', r'\overline{z}'),
    (f'<m:bar><m:e>{r("z")}</m:e></m:bar>', r'\underline{z}'),
    (f'<m:groupChr><m:e>{r("ab")}</m:e></m:groupChr>', r'\underbrace{ab}'),
    (f'<m:groupChr><m:groupChrPr><m:chr m:val="⏞"/></m:groupChrPr><m:e>{r("ab")}</m:e></m:groupChr>',
     r'\overbrace{ab}'),
    (f'<m:groupChr><m:groupChrPr><m:chr m:val="→"/></m:groupChrPr><m:e>{r("ab")}</m:e></m:groupChr>', 'ab'),
    (f'<m:borderBox><m:e>{r("E")}</m:e></m:borderBox>', r'\boxed{E}'),
    (f'<m:box><m:e>{r("E")}</m:e></m:box>', 'E'),
    (f'<m:eqArr><m:e>{r("x&amp;=1")}</m:e><m:e>{r("y&amp;=2")}</m:e></m:eqArr>',
     r'\begin{aligned} x&=1 \\ y&=2 \end{aligned}'),
    (f'<m:m><m:mr><m:e>{r("1")}</m:e><m:e>{r("0")}</m:e></m:mr><m:mr><m:e>{r("0")}</m:e><m:e>{r("1")}</m:e></m:mr></m:m>',
     r'\begin{matrix} 1 & 0 \\ 0 & 1 \end{matrix}'),
    (r('α≤β, 50% {a} &amp;'), r'\alpha \leq \beta , 50\% \{a\} \&'),
    (r('kg', True), r'\mathrm{kg}'),
])
def test_omml(inner, expected):
    assert omml_to_latex(omml(inner)) == expected


def test_omml_paragraphs_and_odd_children():
    assert omml_to_latex(omml(r('E=mc'), para=True)) == 'E=mc'
    two = parse_xml(f'<m:oMathPara {M}><m:oMath>{r("a")}</m:oMath><m:oMath>{r("b")}</m:oMath></m:oMathPara>')
    assert omml_to_latex(two) == r'\begin{gathered} a \\ b \end{gathered}'
    assert omml_to_latex(parse_xml(f'<m:oMathPara {M}/>')) == ''
    # a stray Word run and an unknown Word element inside an equation
    odd = parse_xml(f'<m:oMath {M} {nsdecls("w")}><w:r><w:t>q</w:t></w:r><w:bookmarkStart/>'
                    f'<!-- note --><m:f><m:num>{r("1")}</m:num></m:f></m:oMath>')
    assert omml_to_latex(odd) == r'q\frac{1}{}'
    # function name that came through as \mathrm{...}
    f = omml(f'<m:func><m:fName>{r("cosh", True)}</m:fName><m:e>{r("t")}</m:e></m:func>')
    assert omml_to_latex(f) == r'\cosh t'
    f = omml(f'<m:func><m:fName><m:r><m:rPr><m:nor/></m:rPr><m:t>max</m:t></m:r></m:fName><m:e>{r("ab")}</m:e></m:func>')
    assert omml_to_latex(f) == r'\max {ab}'


# ======================================================================
# DOCX -> LaTeX
# ======================================================================

def add_footnotes(docx_path, notes):
    # python-docx can't create footnotes, so patch the zip by hand
    with zipfile.ZipFile(docx_path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    body = ''.join(
        f'<w:footnote w:id="{i}"><w:p><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p></w:footnote>'
        for i, t in notes.items())
    files['word/footnotes.xml'] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {nsdecls("w")}>'
        f'{body}</w:footnotes>').encode()
    rels = files['word/_rels/document.xml.rels'].decode()
    files['word/_rels/document.xml.rels'] = rels.replace(
        '</Relationships>',
        '<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/footnotes" Target="footnotes.xml"/></Relationships>').encode()
    ct = files['[Content_Types].xml'].decode()
    files['[Content_Types].xml'] = ct.replace(
        '</Types>',
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-'
        'officedocument.wordprocessingml.footnotes+xml"/></Types>').encode()
    with zipfile.ZipFile(docx_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, data in files.items():
            z.writestr(n, data)


def png(color='red', size=(40, 20)):
    buf = io.BytesIO()
    Image.new('RGB', size, color).save(buf, 'PNG')
    buf.seek(0)
    return buf


@pytest.fixture
def rich_docx(tmp_path):
    d = Document()
    d.core_properties.author = 'Ada Lovelace'
    d.add_paragraph('Big Report', style='Title')
    d.add_paragraph('a subtitle', style='Subtitle')

    # Table of contents content control
    toc = parse_xml(f'<w:sdt {nsdecls("w")}><w:sdtPr><w:docPartObj><w:docPartGallery '
                    f'w:val="Table of Contents"/></w:docPartObj></w:sdtPr><w:sdtContent/></w:sdt>')
    d.element.body.insert(len(d.element.body) - 1, toc)
    # Some other content control whose content must not be lost
    other = parse_xml(f'<w:sdt {nsdecls("w")}><w:sdtContent><w:p><w:r><w:t>Inside a control'
                      f'</w:t></w:r></w:p></w:sdtContent></w:sdt>')
    d.element.body.insert(len(d.element.body) - 1, other)

    d.add_heading('2.1 Methods', 1)
    p = d.add_paragraph('Energy is ')
    p._p.append(omml(f'<m:sSup><m:e>{r("mc")}</m:e><m:sup>{r("2")}</m:sup></m:sSup>'))
    p.add_run(' and α ≤ β.')
    run = p.add_run(' see note')
    run._r.append(parse_xml(f'<w:footnoteReference {nsdecls("w")} w:id="1"/>'))
    run = p.add_run('old')
    run.font.strike = True
    run = p.add_run('code')
    run.font.name = 'Courier New'
    run = p.add_run('caps')
    run.font.small_caps = True
    # tracked insertion: kept; deletion: dropped
    p._p.append(parse_xml(f'<w:ins {nsdecls("w")} w:id="9" w:author="x"><w:r><w:t> inserted</w:t></w:r></w:ins>'))
    p._p.append(parse_xml(f'<w:del {nsdecls("w")} w:id="8" w:author="x"><w:r><w:delText>gone</w:delText></w:r></w:del>'))
    # internal link to a bookmark
    p._p.append(parse_xml(f'<w:hyperlink {nsdecls("w")} w:anchor="sec1"><w:r><w:t>jump</w:t></w:r></w:hyperlink>'))

    disp = d.add_paragraph()
    disp._p.append(omml(f'<m:f><m:num>{r("1")}</m:num><m:den>{r("2")}</m:den></m:f>', para=True))

    br = d.add_paragraph('line one')
    br.runs[0].add_break()
    br.add_run('line two')
    d.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    d.add_paragraph('Wise words.', style='Quote')
    right = d.add_paragraph('signed')
    right.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    centered = d.add_paragraph('middle')
    centered.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Figure with a caption below it
    d.add_picture(png(size=(400, 200)))
    d.add_paragraph('Figure 1: A red box', style='Caption')

    # Table with the caption above it, long text, aligned numbers
    d.add_paragraph('Table 1. Measurements', style='Caption')
    t = d.add_table(rows=3, cols=3)
    for c, v in enumerate(['Name', 'Description', 'Value']):
        t.cell(0, c).text = v
    long = 'A rather long description that would never fit on one line of a page ' * 2
    for row in (1, 2):
        t.cell(row, 0).text = f'item {row}'
        t.cell(row, 1).text = long
        t.cell(row, 2).text = str(row * 10)
        t.cell(row, 2).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    # caption after a table
    t2 = d.add_table(rows=1, cols=2)
    t2.cell(0, 0).text = 'x'
    t2.cell(0, 1).paragraphs[0].add_run().add_picture(png(), width=Pt(20))
    t2.cell(0, 1).add_table(rows=1, cols=1).cell(0, 0).text = 'nested'
    d.add_paragraph('Second table', style='Caption')
    # a caption with nothing to attach to
    d.add_paragraph('Lonely caption', style='Caption')

    d.add_heading('References', 1)
    d.add_paragraph('[1] Knuth, D. The TeXbook. 1984.')
    d.add_paragraph('Lamport, L. LaTeX. 1994.', style='List Number')
    d.add_paragraph('')
    d.add_heading('Appendix', 2)
    d.add_paragraph('The end.')
    path = tmp_path / 'rich.docx'
    d.save(path)
    add_footnotes(path, {1: 'A footnote with 50%'})
    return path


def test_docx_rich_features(rich_docx):
    c = DocTeXConverter(ConversionOptions(extract_bibliography=True))
    tex = tex_of(c.convert(str(rich_docx)))

    assert '\\title{Big Report}\n\\author{Ada Lovelace}' in tex
    assert '\\large a subtitle' in tex
    assert '\\tableofcontents' in tex
    assert 'Inside a control' in tex
    assert '\\section{Methods}' in tex  # typed "2.1" numbering removed
    assert 'Energy is ${mc}^{2}$ and \\ensuremath{\\alpha} \\ensuremath{\\leq} \\ensuremath{\\beta}.' in tex
    assert '\\footnote{A footnote with 50\\%}' in tex
    assert '\\sout{old}' in tex and '\\texttt{code}' in tex and '\\textsc{caps}' in tex
    assert ' inserted' in tex and 'gone' not in tex
    assert '\\hyperlink{sec1}{jump}' in tex
    assert '\\[\n\\frac{1}{2}\n\\]' in tex
    assert 'line one\\newline line two' in tex
    assert '\\clearpage' in tex
    assert '\\begin{quote}\nWise words.\n\\end{quote}' in tex
    assert '\\begin{flushright}\nsigned' in tex and '\\begin{center}\nmiddle' in tex

    assert re.search(r'\\begin\{figure\}\[htbp\]\n\\centering\n\\includegraphics\[width=0\.\d\d\\textwidth\]'
                     r'\{rich-image1\.png\}\n\\caption\{A red box\}\n\\end\{figure\}', tex)
    assert '\\caption{Measurements}\n\\begin{tabular}{l>{\\raggedright\\arraybackslash}p{' in tex
    assert re.search(r'\\linewidth\}r\}', tex)  # number column right aligned
    assert '\\caption{Second table}\n\\end{table}' in tex
    assert 'nested' in tex and any('Nested table' in w for w in c.warnings)
    assert '\\textit{Lonely caption}' in tex

    assert '\\begin{thebibliography}{99}' in tex
    assert '\\bibitem{ref1} Knuth, D. The TeXbook. 1984.' in tex
    assert '\\bibitem{ref2} Lamport, L. LaTeX. 1994.' in tex
    assert '\\section{References}' not in tex
    assert '\\subsection{Appendix}' in tex and 'The end.' in tex
    assert tex.index('\\end{thebibliography}') < tex.index('\\subsection{Appendix}')


def test_docx_references_kept_without_extract(rich_docx, conv):
    tex = tex_of(conv.convert(str(rich_docx)))
    assert '\\section{References}' in tex and 'thebibliography' not in tex


def test_docx_document_types(rich_docx, tmp_path):
    for doc_type, cls, heading in [('report', 'report', '\\chapter{Methods}'),
                                   ('thesis', 'report', '\\chapter{Methods}'),
                                   ('book', 'book', '\\chapter{Methods}'),
                                   ('letter', 'article', '\\section{Methods}')]:
        c = DocTeXConverter(ConversionOptions(document_type=DocumentType(doc_type)))
        tex = tex_of(c.convert(str(rich_docx), str(tmp_path / f'{doc_type}.tex')))
        assert f'\\documentclass[12pt]{{{cls}}}' in tex
        assert heading in tex


def test_docx_preamble_options(rich_docx, tmp_path):
    opts = ConversionOptions(line_spacing=LineSpacing.DOUBLE, language='German',
                             custom_packages=['siunitx'], unicode_support=False)
    tex = tex_of(DocTeXConverter(opts).convert(str(rich_docx), str(tmp_path / 'p.tex')))
    assert '\\doublespacing' in tex and '\\usepackage[german]{babel}' in tex
    assert '\\usepackage{siunitx}' in tex and 'inputenc' not in tex
    opts = ConversionOptions(line_spacing=LineSpacing.ONE_HALF)
    tex = tex_of(DocTeXConverter(opts).convert(str(rich_docx), str(tmp_path / 'q.tex')))
    assert '\\onehalfspacing' in tex


def test_docx_body_only(rich_docx, tmp_path):
    opts = ConversionOptions(include_preamble=False, standalone_document=False)
    tex = tex_of(DocTeXConverter(opts).convert(str(rich_docx), str(tmp_path / 'b.tex')))
    assert '\\documentclass' not in tex and '\\begin{document}' not in tex
    assert '\\section{Methods}' in tex


def test_docx_optimize_images(rich_docx, tmp_path):
    opts = ConversionOptions(optimize_images=True, image_quality=50)
    tex = tex_of(DocTeXConverter(opts).convert(str(rich_docx), str(tmp_path / 'o.tex')))
    assert '{o-image1.jpg}' in tex
    assert (tmp_path / 'images' / 'o-image1.jpg').is_file()
    assert not (tmp_path / 'images' / 'o-image1.png').exists()


def test_docx_gif_and_broken_images(tmp_path, conv):
    d = Document()
    gif = io.BytesIO()
    Image.new('RGB', (10, 10), 'green').save(gif, 'GIF')
    gif.seek(0)
    d.add_picture(gif)
    d.add_paragraph('text').add_run()._r.append(parse_xml(
        f'<a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId999"/>'))
    path = tmp_path / 'g.docx'
    d.save(path)
    tex = tex_of(conv.convert(str(path)))
    assert '{g-image1.png}' in tex
    with Image.open(tmp_path / 'images' / 'g-image1.png') as img:
        assert img.format == 'PNG'
    assert any('rId999' in w for w in conv.warnings)


def test_docx_unconvertible_image_and_optimize_failure(tmp_path, monkeypatch):
    d = Document()
    d.add_picture(png())
    path = tmp_path / 'x.docx'
    d.save(path)
    # pretend the image is an .emf that PIL can't read
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    media = next(n for n in files if n.startswith('word/media/'))
    files[media.replace('.png', '.emf')] = b'not an image'
    del files[media]
    for n in list(files):
        if n.endswith('.rels') or n == '[Content_Types].xml':
            files[n] = files[n].replace(b'image1.png', b'image1.emf').replace(
                b'Extension="png"', b'Extension="emf"')
    with zipfile.ZipFile(path, 'w') as z:
        for n, data in files.items():
            z.writestr(n, data)
    c = DocTeXConverter(ConversionOptions(optimize_images=True))
    monkeypatch.setattr('doc2tex.latex.optimize_image',
                        lambda *a, **k: (_ for _ in ()).throw(ImageProcessingError('nope')))
    tex = tex_of(c.convert(str(path)))
    assert '{x-image1.png}' in tex
    assert any('may need converting' in w for w in c.warnings)


def test_docx_lists_with_word_numbering(tmp_path, conv):
    d = Document()
    d.add_paragraph('bullet', style='List Bullet')
    d.add_paragraph('number', style='List Number')
    d.add_paragraph('Heading with numbering', style='Heading 1')
    p = d.add_paragraph('custom numbering')
    num_pr = p._p.get_or_add_pPr().get_or_add_numPr()
    num_pr.get_or_add_numId().val = 999   # numId that doesn't exist
    num_pr.get_or_add_ilvl().val = 0
    path = tmp_path / 'l.docx'
    d.save(path)
    tex = tex_of(conv.convert(str(path)))
    assert '\\begin{itemize}\n  \\item bullet\n\\end{itemize}\n\\begin{enumerate}' in tex
    assert '\\section{Heading with numbering}' in tex
    assert '\\item custom numbering' in tex


def test_docx_numbering_lookup(tmp_path):
    from doc2tex.latex import LatexGenerator
    d = Document()
    d.add_paragraph('x', style='List Number')
    path = tmp_path / 'n.docx'
    d.save(path)
    gen = LatexGenerator(ConversionOptions())
    gen._doc = Document(path)
    num_id = gen._doc.styles['List Number'].element.pPr.numPr.numId.val
    assert gen._numbering_format(num_id, 0) == 'enumerate'
    assert gen._numbering_format(num_id, 42) is None
    assert gen._numbering_format(12345, 0) is None
    gen._doc = Document()
    gen._doc.part.numbering_part.element.clear()
    assert gen._numbering_format(1, 0) is None


def test_docx_empty_table_and_bad_file(tmp_path, conv):
    d = Document()
    d.add_table(rows=1, cols=1)
    d.add_paragraph('', style='Title')
    d.add_heading('', 1)
    path = tmp_path / 'e.docx'
    d.save(path)
    tex = tex_of(conv.convert(str(path)))
    assert '\\begin{tabular}{l}' in tex and '\\title' not in tex

    bad = tmp_path / 'bad.docx'
    bad.write_bytes(b'this is not a zip')
    with pytest.raises(ConversionError):
        conv.convert(str(bad))


def test_docx_table_column_spec_helpers():
    from doc2tex.latex import LatexGenerator
    spec = LatexGenerator._column_spec([[['a', 1], ['b', 1]]], [['c', 'r']], 2)
    assert spec == 'cr'
    assert LatexGenerator._escape_url('http://x.com/a%20b#top') == 'http://x.com/a\\%20b\\#top'
    assert LatexGenerator._clean_breaks('\\newline a \\newline ') == 'a'


# ======================================================================
# LaTeX -> DOCX
# ======================================================================

FULL_TEX = r"""
\documentclass[11pt]{report}
\usepackage[margin=2cm,left=3cm]{geometry}
\onehalfspacing
\newcommand{\R}{\mathbb{R}}
\newcommand{\vect}[1]{\mathbf{#1}}
\newcommand{\greet}[2][Hello]{#1, #2!}
\def\course{EE101}
\DeclareMathOperator{\sgn}{sgn}
\begin{document}
\tableofcontents
\chapter{Intro}\label{ch:intro}
\greet{world} \greet[Hi]{you} in \course. $x \in \R$, $\vect{v}$, $\sgn(x)$.
See Section~\ref{sec:m}, \autoref{fig:a}, \cref{tab:t}, \eqref{eq:e} and \ref{missing}.
Cited \cite{knuth} and \citep[see][p.~4]{lamport, knuth} and \citet{lamport}; \citeauthor{many} \citeyear{knuth}.
\nocite{unused}
\section{Method}\label{sec:m}
Speed \SI{3.5}{\kilo\metre\per\second} and \si{kg.m/s^2}, \num{6.02e23}, \ang{90},
\qty{2}{\square\metre}, \SIrange{1}{5}{\volt} and \unit{\watt\hour\squared}.
\begin{equation}
E = mc^2 \label{eq:e}
\end{equation}
\begin{align}
a &= 1 \\
b &= 2 \nonumber \\
c &= 3 \tag{*}
\end{align}
\begin{figure}[h]
\includegraphics[width=4cm]{pic}
\caption{A picture}\label{fig:a}
\end{figure}
\begin{table}[h]
\caption{Data}\label{tab:t}
\begin{tabular}{lc}
a & b \\
\end{tabular}
\end{table}
\begin{enumerate}
\item first
\end{enumerate}
\begin{enumerate}
\item again one
\end{enumerate}
\section*{Unnumbered}
\appendix
\chapter{Extra}
\begin{figure}
\caption{In appendix}
\end{figure}
\bibliographystyle{plain}
\bibliography{refs}
\end{document}
"""

BIB = r"""
% a comment
@book{knuth,
  author = {Donald E. Knuth},
  title = {The {\TeX}book},
  publisher = {Addison-Wesley},
  year = 1984
}
@article{lamport, author = "Lamport, Leslie and Smith, J.", title = {Paper},
  journal = {J. Things}, volume = {3}, number = {2}, pages = {1--10}, year = {1994},
  doi = {10.1/xyz}}
@misc{many, author = {A One and B Two and C Three}, title = {Web}, url = {https://ex.com}}
@comment{ignored}
@misc{unused, title = {Never cited}}
"""


@pytest.fixture
def full_tex(tmp_path):
    (tmp_path / 'refs.bib').write_text(BIB, encoding='utf-8')
    Image.new('RGB', (40, 20), 'blue').save(tmp_path / 'pic.png')
    path = tmp_path / 'full.tex'
    path.write_text(FULL_TEX, encoding='utf-8')
    return path


def test_latex_to_docx_numbering_refs_and_bib(full_tex, conv):
    d = Document(conv.convert(str(full_tex)))
    paras = [(p.style.name, p.text) for p in d.paragraphs]
    texts = [t for _, t in paras]
    joined = '\n'.join(texts)

    # TOC field + "update fields on open"
    assert ('TOC Heading', 'Contents') in paras
    assert d.element.xpath('.//w:instrText[contains(text(), "TOC")]')
    assert d.settings.element.find(qn('w:updateFields')) is not None

    # chapters push sections down a level and get numbered
    assert ('Heading 1', 'Chapter 1\nIntro') in paras
    assert ('Heading 2', '1.1\u2003Method') in paras
    assert ('Heading 2', 'Unnumbered') in paras
    assert ('Heading 1', 'Appendix A\nExtra') in paras

    # macros
    assert 'Hello, world! Hi, you! in EE101.' in joined
    assert 'x∈R' in joined and 'sgn(x)' in joined

    # cross references
    assert 'See Section\u00a01.1, Figure\u00a01.1, Table\u00a01.1, (1.1) and ??.' in joined
    assert any('Undefined reference: missing' in w for w in conv.warnings)

    # citations numbered by order of first use
    assert 'Cited [1] and [see 2, 1, p.\u00a04] and Lamport and Smith [2]; One et al. 1984.' in joined

    # siunitx
    assert '3.5\u00a0km/s and kg·m/s², 6.02×10²³, 90°,' in joined
    assert '2\u00a0m², 1\u00a0V to 5\u00a0V and W·h².' in joined

    # equations
    assert any(t.endswith('(1.1)') and 'E' in t for t in texts)
    align = next(t for t in texts if 'a=1' in t.replace(' ', ''))
    assert '(1.2)' in align and '(*)' in align and '(1.3)' not in align

    # captions
    assert ('Caption', 'Figure 1.1: A picture') in paras
    assert ('Caption', 'Table 1.1: Data') in paras
    assert ('Caption', 'Figure A.1: In appendix') in paras

    # every enumerate restarts at 1 (its own w:num with a startOverride)
    nums = [p._p.pPr.numPr.numId.val for p in d.paragraphs
            if p.text in ('first', 'again one')]
    assert len(set(nums)) == 2
    numbering = d.part.numbering_part.element
    for n in nums:
        assert numbering.num_having_numId(n).find(qn('w:lvlOverride')) is not None

    # bibliography from refs.bib, only cited entries (+ \nocite)
    assert ('Heading 1', 'References') in paras
    refs = texts[texts.index('References') + 1:]
    assert refs[0] == '[1] Donald E. Knuth. The TeXbook. Addison-Wesley. 1984.'
    assert refs[1] == '[2] Leslie Lamport and J. Smith. Paper. J. Things, 3(2), pp. 1–10. 1994. https://doi.org/10.1/xyz'
    assert refs[2] == '[3] A One, B Two, and C Three. Web. https://ex.com'
    assert refs[3] == '[4] Never cited.'

    # layout from the .tex: 11pt, 1.5 spacing, margins
    normal = d.styles['Normal']
    assert normal.font.size == Pt(11)
    assert normal.paragraph_format.line_spacing == 1.5
    s = d.sections[0]
    assert round(s.left_margin.cm, 1) == 3.0 and round(s.top_margin.cm, 1) == 2.0


def test_latex_to_docx_no_numbering(full_tex):
    c = DocTeXConverter(ConversionOptions(number_sections=False))
    d = Document(c.convert(str(full_tex)))
    assert ('Heading 2', 'Method') in [(p.style.name, p.text) for p in d.paragraphs]


def test_latex_to_docx_layout_variants(tmp_path, conv):
    for cmd, expected in [('\\doublespacing', 2.0), ('\\singlespacing', None),
                          ('\\linespread{1.3}', 1.49), ('\\setstretch{1.2}', 1.2)]:
        path = tmp_path / 'l.tex'
        path.write_text(f'\\documentclass{{article}}\\geometry{{top=1in}}{cmd}'
                        f'\\begin{{document}}x\\end{{document}}', encoding='utf-8')
        d = Document(conv.convert(str(path)))
        spacing = d.styles['Normal'].paragraph_format.line_spacing
        assert spacing == expected or abs(spacing - expected) < 0.01
        assert round(d.sections[0].top_margin.inches, 2) == 1.0
    # options spacing used when the .tex says nothing
    c = DocTeXConverter(ConversionOptions(line_spacing=LineSpacing.ONE_HALF,
                                          page_margins='margin=20mm,vmargin=1in,bmargin=72pt'))
    path.write_text('hi', encoding='utf-8')
    d = Document(c.convert(str(path)))
    assert d.styles['Normal'].paragraph_format.line_spacing == 1.5
    assert round(d.sections[0].left_margin.mm) == 20
    assert round(d.sections[0].top_margin.inches, 2) == 1.0


def test_latex_thebibliography_numbers(tmp_path, conv):
    path = tmp_path / 'b.tex'
    path.write_text(r"""\begin{document}
As \cite{b} and \cite{a} and \cite{zz}.
\begin{thebibliography}{9}
\bibitem{a} First.
\bibitem[X]{b} Second.
\end{thebibliography}
\end{document}""", encoding='utf-8')
    d = Document(conv.convert(str(path)))
    texts = [p.text for p in d.paragraphs]
    assert 'As [2] and [1] and [zz].' in texts
    assert '[1] First.' in texts and '[2] Second.' in texts
    assert any('zz' in w for w in conv.warnings)


def test_latex_missing_bib_and_input(tmp_path, conv):
    path = tmp_path / 'm.tex'
    path.write_text('\\input{nothere}\\bibliography{nope}Text \\printbibliography', encoding='utf-8')
    d = Document(conv.convert(str(path)))
    assert [p.text for p in d.paragraphs] == ['Text']
    assert any('nope.bib' in w for w in conv.warnings)
    assert any('nothere' in w for w in conv.warnings)


def test_latex_inputs_and_misc_environments(tmp_path, conv):
    (tmp_path / 'part.tex').write_text('Included \\textsc{text}.', encoding='utf-8')
    path = tmp_path / 'misc.tex'
    path.write_text(r"""\title{T}\author{Me}\date{\today}
\begin{document}
\maketitle
\begin{abstract}Short.\end{abstract}
\input{part}
\begin{quote}Quoted\end{quote}
\begin{flushright}R\end{flushright}
\begin{flushleft}L\end{flushleft}
\begin{description}\item[Term] meaning\end{description}
\begin{itemize}\item\end{itemize}
\begin{minipage}{0.5\textwidth}Mini\end{minipage}
\begin{lstlisting}[language=Python]
print(1)
\end{lstlisting}
$$ \frac{a+b}{c} + \sqrt{x} + x^\alpha $$
\[ \alpha \]
\begin{math} y \end{math}
\newpage
Text with \textcolor{red}{color}, \footnote{note}, \url{http://a.b}, \LaTeX{},
``quotes'', a--b, a---b, \'e, \unknowncmd{kept}, \relax x\\[2pt]y, \& \% \, \; \ \S.
{\bfseries bold \itshape both} \textnormal{plain} \includegraphics{missing}
\begin{table}\begin{tabular*}{\textwidth}{l}\hline\multirow{2}{*}{m} \\ \cline{1-1}\end{tabular*}\end{table}
\begin{tabularx}{\textwidth}{X}x\end{tabularx}
\end{document}""", encoding='utf-8')
    d = Document(conv.convert(str(path)))
    texts = [p.text for p in d.paragraphs]
    joined = '\n'.join(texts)
    assert texts[0] == 'T' and 'Me' in texts
    assert 'Abstract' in texts and 'Short.' in texts
    assert 'Included text.' in texts
    for t in ('Quoted', 'R', 'L', 'Term meaning', 'Mini', 'print(1)', 'y'):
        assert t in joined
    assert '(a+b)/c+√x+xα' in joined  # spacing in math isn't kept and 'α' in texts
    assert 'color' in joined and '(note)' in joined and 'http://a.b' in joined
    assert '“quotes”' in joined and 'a–b' in joined and 'a—b' in joined and 'é' in joined
    assert 'kept' in joined and 'LaTeX' in joined and '& %' in joined and '§' in joined
    assert '[Image file not found: missing]' in joined
    assert [c.text for c in d.tables[0].rows[0].cells] == ['m']
    assert [c.text for c in d.tables[1].rows[0].cells] == ['x']


def test_latex_bad_image(tmp_path, conv):
    (tmp_path / 'broken.png').write_bytes(b'not really a png')
    path = tmp_path / 'i.tex'
    path.write_text('\\includegraphics[width=0.3\\linewidth]{broken.png}', encoding='utf-8')
    d = Document(conv.convert(str(path)))
    assert 'error loading' in d.paragraphs[0].text


def test_latex_image_widths(tmp_path):
    from doc2tex.docx import DocxGenerator
    gen = DocxGenerator(ConversionOptions())
    gen.word_doc = Document()
    full = gen._image_width('width=\\textwidth')
    assert gen._image_width('width=0.5\\textwidth') == full // 2
    assert gen._image_width('width=2in') == 2 * 914400
    assert gen._image_width('width=100cm') == full
    assert gen._image_width(None) == int(full * 0.8)


def test_latex_generator_wraps_errors(tmp_path, conv, monkeypatch):
    path = tmp_path / 'x.tex'
    path.write_text('x', encoding='utf-8')
    monkeypatch.setattr('doc2tex.docx.DocxGenerator._parse_and_build',
                        lambda *a: (_ for _ in ()).throw(RuntimeError('boom')))
    with pytest.raises(ConversionError, match='boom'):
        conv.convert(str(path))


# ----------------------------------------------------------------------
# LaTeX parsing helpers
# ----------------------------------------------------------------------

def test_scanning_helpers():
    assert read_group('{a{b}c}d', 0) == ('a{b}c', 7)
    assert read_group('{open', 0) == ('open', 5)
    assert read_arg('  {x}', 0) == ('x', 5)
    assert read_arg('\\alpha b', 0) == ('\\alpha', 6)
    assert read_arg('z', 0) == ('z', 1)
    assert read_arg('', 0) == (None, 0)
    assert read_arg('}', 0) == (None, 0)
    assert read_optional('[a{]}]b', 0) == ('a{]}', 6)
    assert read_optional('x', 0) == (None, 0)
    assert read_optional('[never', 0) == (None, 0)
    s = 'a\\begin{x}b\\begin{x}c\\end{x}\\end{x}d'
    assert find_env_end(s, 10, 'x') == (s.rindex('\\end{x}'), len(s) - 1)
    assert find_env_end('nothing', 0, 'x') == (7, 7)
    assert split_top_level('a & {b & c} & \\begin{x}d & e\\end{x}', '&') == ['a ', ' {b & c} ', ' \\begin{x}d & e\\end{x}']
    assert split_top_level('\\item a \\itemsep \\item b', '\\item') == ['', ' a \\itemsep ', ' b']
    assert strip_comments('a % c\nb \\% d\n\\begin{verbatim}%x\\end{verbatim}') == \
        'a b \\% d\n\\begin{verbatim}%x\\end{verbatim}'


def test_expand_macros():
    assert expand_macros('no macros') == 'no macros'
    tex = ('\\newcommand*{\\a}{A}\\renewcommand\\b{\\a B}\\providecommand{\\c}[1]{<#1>}'
           '\\a \\b \\c{x} \\abc')
    assert expand_macros(tex) == '{A} {A} B <x> \\abc'
    assert expand_macros('\\newcommand{\\broken}') == '\\newcommand{\\broken}'


def test_parse_and_format_bib():
    entries = parse_bib(BIB)
    assert set(entries) == {'knuth', 'lamport', 'many', 'unused'}
    assert entries['knuth']['year'] == '1984'
    assert entries['lamport']['author'] == 'Lamport, Leslie and Smith, J.'
    assert format_bib_entry({'title': 'T', 'booktitle': 'Conf', 'year': '2000'}) == 'T. \\textit{Conf}. 2000.'
    assert format_bib_entry({'editor': 'E. Ditor', 'title': 'Book'}) == 'E. Ditor. Book.'
    assert parse_bib('@misc{k, title = "unterminated') == {'k': {'ENTRYTYPE': 'misc', 'title': 'unterminated'}}
    assert parse_bib('@misc{k, year = }')['k']['year'] == ''


def test_si_units():
    assert si_unit_text('\\kilo\\metre\\per\\second\\squared') == 'km/s²'
    assert si_unit_text('\\cubic\\centi\\metre') == 'cm³'
    assert si_unit_text('m^{-1}') == 'm⁻¹'
    assert si_unit_text('\\metre^{2}') == 'm²'
    assert si_unit_text('') == ''
    assert si_unit_text(None) == ''


# ======================================================================
# Converter, API, CLI, web
# ======================================================================

def test_top_level_convert(tmp_path):
    path = tmp_path / 'a.tex'
    path.write_text('Hi', encoding='utf-8')
    out = doc2tex.convert(str(path), str(tmp_path / 'b.docx'), document_type='report', font_size='11pt')
    assert out.endswith('b.docx') and os.path.isfile(out)
    assert doc2tex.__version__ == '2.0.0'


def test_converter_errors_and_batch(tmp_path, conv):
    with pytest.raises(ConversionError):
        conv.convert(__file__, direction='sideways')
    txt = tmp_path / 'a.txt'
    txt.write_text('x')
    with pytest.raises(InvalidFileFormatError):
        conv.convert(str(txt), direction='to_latex')
    with pytest.raises(InvalidFileFormatError):
        conv.convert(str(txt), direction='to_docx')
    tex = tmp_path / 'ok.tex'
    tex.write_text('x', encoding='utf-8')
    assert conv.batch([str(tex), str(txt)], output_dir=str(tmp_path / 'o')) == [
        str(tmp_path / 'o' / 'ok.docx'), None]
    assert conv.batch_convert([str(tex)])[0] == str(tmp_path / 'ok.docx')


def test_converter_verbose(tmp_path):
    c = DocTeXConverter(ConversionOptions(verbose=True))
    assert c.settings.verbose
    utils.logger.setLevel('INFO')


def test_find_inputs(tmp_path):
    (tmp_path / 'a.docx').write_bytes(b'')
    (tmp_path / 'b.TEX').write_text('')
    (tmp_path / '~$lock.docx').write_bytes(b'')
    (tmp_path / 'c.txt').write_text('')
    (tmp_path / 'sub').mkdir()
    found = DocTeXConverter.find_inputs([str(tmp_path), 'x.docx'])
    assert [os.path.basename(f) for f in found] == ['a.docx', 'b.TEX', 'x.docx']


class FakeProc:
    def __init__(self, code, out=''):
        self.returncode, self.stdout = code, out


def test_compile_pdf(tmp_path, monkeypatch):
    tex = tmp_path / 'r.tex'
    tex.write_text('x')
    calls = []

    monkeypatch.setattr(shutil, 'which', lambda name: None)
    with pytest.raises(LatexCompilationError, match='No LaTeX'):
        DocTeXConverter.compile_pdf(str(tex))

    monkeypatch.setattr(shutil, 'which', lambda name: '/bin/' + name if name == 'pdflatex' else None)
    monkeypatch.setattr(subprocess, 'run', lambda cmd, **kw: calls.append(cmd) or FakeProc(0))
    assert DocTeXConverter.compile_pdf(str(tex)) == str(tmp_path / 'r.pdf')
    assert len(calls) == 2 and calls[0][0] == 'pdflatex'

    monkeypatch.setattr(shutil, 'which', lambda name: '/bin/' + name)
    monkeypatch.setattr(subprocess, 'run', lambda cmd, **kw: FakeProc(1, 'junk\n! Undefined control sequence.\n'))
    with pytest.raises(LatexCompilationError, match='Undefined control sequence'):
        DocTeXConverter.compile_pdf(str(tex))

    def timeout(cmd, **kw):
        raise subprocess.TimeoutExpired(cmd, 1)
    monkeypatch.setattr(subprocess, 'run', timeout)
    with pytest.raises(LatexCompilationError, match='longer than'):
        DocTeXConverter.compile_pdf(str(tex), timeout=1)


@pytest.fixture
def two_docx(tmp_path):
    folder = tmp_path / 'in'
    folder.mkdir()
    for name, color in (('one', 'red'), ('two', 'blue')):
        d = Document()
        d.add_paragraph(name)
        d.add_picture(png(color))
        d.save(folder / f'{name}.docx')
    return folder


def test_cli_folder_input_keeps_images_apart(two_docx, tmp_path, capsys):
    out = tmp_path / 'out'
    assert cli_main([str(two_docx), '-d', str(out), '--body-only', '--package', 'siunitx',
                     '--language', 'french', '--image-width', '5cm', '--no-numbering']) == 0
    assert {p.name for p in (out / 'images').iterdir()} == {'one-image1.png', 'two-image1.png'}
    assert '\\documentclass' not in (out / 'one.tex').read_text()
    assert 'OK' in capsys.readouterr().out


def test_cli_errors(two_docx, tmp_path, capsys, monkeypatch):
    empty = tmp_path / 'empty'
    empty.mkdir()
    assert cli_main([str(empty)]) == 1
    with pytest.raises(SystemExit):
        cli_main([str(two_docx), '-o', 'x.tex'])
    assert cli_main([str(tmp_path / 'nope.docx')]) == 1
    bad = tmp_path / 'bad.docx'
    bad.write_bytes(b'junk')
    assert cli_main([str(bad), str(two_docx / 'one.docx'), '-d', str(tmp_path / 'o')]) == 1
    assert 'FAIL' in capsys.readouterr().out
    monkeypatch.setattr('doc2tex.cli.DocTeXConverter.convert',
                        lambda *a, **k: (_ for _ in ()).throw(RuntimeError('weird')))
    assert cli_main([str(two_docx / 'one.docx')]) == 1


def test_cli_pdf_and_warnings(two_docx, tmp_path, monkeypatch, capsys):
    monkeypatch.setattr(DocTeXConverter, 'compile_pdf', staticmethod(lambda tex: tex[:-4] + '.pdf'))
    assert cli_main([str(two_docx / 'one.docx'), '-o', str(tmp_path / 'one.tex'), '--pdf']) == 0
    assert 'PDF:' in capsys.readouterr().out
    assert cli_main([str(two_docx / 'one.docx'), '-o', str(tmp_path / 'b.tex'), '--pdf', '--body-only']) == 1

    tex = tmp_path / 'w.tex'
    tex.write_text('\\includegraphics{gone}', encoding='utf-8')
    assert cli_main([str(tex)]) == 0
    assert 'warning: Image not found: gone' in capsys.readouterr().out


@pytest.fixture
def client(tmp_path):
    from doc2tex.web import app
    app.config['UPLOAD_FOLDER'] = str(tmp_path / 'uploads')
    os.makedirs(app.config['UPLOAD_FOLDER'])
    app.config['TESTING'] = True
    return app.test_client()


def test_web_home_and_options(client):
    page = client.get('/')
    assert page.status_code == 200 and b'v2.0.0' in page.data and b'bodyOnly' in page.data
    resp = client.post('/convert', data={
        'file': (io.BytesIO(b'\\includegraphics{x}'), 'a.tex'), 'number_sections': 'false'})
    data = resp.get_json()
    assert data['success'] and data['warnings'] == ['Image not found: x']


def test_web_body_only_and_odd_names(client):
    d = Document()
    d.add_paragraph('hello')
    buf = io.BytesIO()
    d.save(buf)
    buf.seek(0)
    resp = client.post('/convert', data={'file': (buf, '???.docx'), 'body_only': 'true'})
    data = resp.get_json()
    assert data['output_filename'] == 'document.tex'
    assert b'\\documentclass' not in client.get(data['download_url']).data


def test_web_conversion_failures(client, monkeypatch):
    resp = client.post('/convert', data={'file': (io.BytesIO(b'junk'), 'a.docx')})
    assert resp.status_code == 422
    monkeypatch.setattr('doc2tex.web.DocTeXConverter.convert',
                        lambda *a, **k: (_ for _ in ()).throw(RuntimeError('x')))
    resp = client.post('/convert', data={'file': (io.BytesIO(b'x'), 'a.tex')})
    assert resp.status_code == 500


def test_web_cleanup_old_jobs(client, tmp_path):
    from doc2tex import web
    root = web.app.config['UPLOAD_FOLDER']
    old = os.path.join(root, 'a' * 32)
    new = os.path.join(root, 'b' * 32)
    os.makedirs(old)
    os.makedirs(new)
    open(os.path.join(root, 'stray.txt'), 'w').close()
    os.utime(old, (0, 0))
    web.cleanup_old_jobs()
    assert not os.path.exists(old) and os.path.exists(new)


def test_web_start_server(monkeypatch, capsys):
    from doc2tex import web
    seen = {}
    monkeypatch.setattr(web.app, 'run', lambda **kw: seen.update(kw))
    monkeypatch.setenv('DOC2TEX_PORT', '8123')
    monkeypatch.setenv('DOC2TEX_DEBUG', '1')
    web.start_server()
    assert seen == {'host': '127.0.0.1', 'port': 8123, 'debug': True}
    assert 'localhost:8123' in capsys.readouterr().out
    utils.logger.setLevel('INFO')


# ======================================================================
# Options and utils
# ======================================================================

def test_options_roundtrip_and_validation():
    opts = ConversionOptions(document_type=DocumentType.BOOK, line_spacing=LineSpacing.DOUBLE,
                             extract_bibliography=True, custom_packages=['tikz', 'graphicx'])
    again = ConversionOptions.from_dict(opts.to_dict())
    assert again == opts
    pkgs = opts.get_latex_packages()
    assert 'setspace' in pkgs and 'natbib' in pkgs and pkgs.count('graphicx') == 1 and 'tikz' in pkgs
    assert 'fontenc' not in ConversionOptions(unicode_support=False).get_latex_packages()
    with pytest.raises(ValueError):
        ConversionOptions(image_quality=0).validate()
    with pytest.raises(ValueError):
        ConversionOptions(output_encoding='').validate()


def test_utils(tmp_path):
    assert utils.escape_latex('') == '' and utils.unescape_latex('') == ''
    assert utils.escape_latex('5° α\u00a0x') == '5\\textdegree{} \\ensuremath{\\alpha}~x'
    assert utils.sanitize_filename(' a<b>:c?. ') == 'a_b__c_'
    assert len(utils.sanitize_filename('x' * 300 + '.tex')) == 255
    d = utils.ensure_directory(str(tmp_path / 'x' / 'y'))
    assert d.is_dir()
    tmp = utils.get_temp_dir()
    assert os.path.isdir(tmp)
    utils.cleanup_temp_dir(tmp)
    assert not os.path.exists(tmp)
    utils.cleanup_temp_dir(tmp)  # already gone: no error
    f = tmp_path / 'h.txt'
    f.write_bytes(b'abc')
    assert utils.get_file_hash(str(f)) == '900150983cd24fb0d6963f7d28e17f72'
    assert utils.normalize_whitespace('  a   b\n\n\n\nc ') == 'a b\n\nc'
    assert utils.handle_unicode(b'caf\xc3\xa9') == 'café' and utils.handle_unicode('x') == 'x'
    with pytest.raises(UnicodeHandlingError):
        utils.handle_unicode(b'x', encoding='no-such-codec')
    assert utils.extract_extension('A.DOCX') == 'docx'
    assert utils.is_valid_file(str(f), ['TXT']) and not utils.is_valid_file('nope.txt', ['txt'])
    assert utils.format_file_size(512) == '512.00 B'
    assert utils.format_file_size(2048) == '2.00 KB'
    assert utils.format_file_size(1024 ** 4) == '1.00 TB'
    info = utils.get_file_info(str(f))
    assert info['name'] == 'h.txt' and info['size'] == 3 and info['extension'] == 'txt'


def test_utils_failing_cleanup(monkeypatch, tmp_path):
    monkeypatch.setattr(utils.shutil, 'rmtree', lambda p: (_ for _ in ()).throw(OSError('busy')))
    utils.cleanup_temp_dir(str(tmp_path))  # logs a warning, doesn't raise


def test_image_utils(tmp_path):
    rgba = tmp_path / 'a.png'
    Image.new('RGBA', (300, 100), (255, 0, 0, 128)).save(rgba)
    out, w, h = utils.optimize_image(str(rgba), str(tmp_path / 'a.jpg'), max_width=150, quality=70)
    assert (w, h) == (150, 50) and os.path.isfile(out)
    gray = tmp_path / 'g.png'
    Image.new('L', (10, 10)).save(gray)
    assert utils.optimize_image(str(gray), str(tmp_path / 'g.jpg'))[1:] == (10, 10)
    with pytest.raises(ImageProcessingError):
        utils.optimize_image(str(tmp_path / 'missing.png'), str(tmp_path / 'm.jpg'))
    assert utils.get_image_dimensions(str(rgba)) == (300, 100)
    assert utils.get_image_dimensions(str(tmp_path / 'missing.png')) == (0, 0)
    assert utils.setup_logger('doc2tex-test', verbose=True).level == 10


def test_script_shortcuts():
    import runpy
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    ns = runpy.run_path(os.path.join(root, 'web.py'))
    assert 'app' in ns and 'start_server' in ns
    ns = runpy.run_path(os.path.join(root, 'cli.py'))
    assert 'main' in ns


# ======================================================================
# Edge cases (so every branch gets exercised at least once)
# ======================================================================

def test_omml_function_typed_as_italic_text():
    f = omml(f'<m:func><m:fName>{r("sin")}</m:fName><m:e>{r("x")}</m:e></m:func>')
    assert omml_to_latex(f) == r'\sin x'


def test_docx_inline_edge_cases(tmp_path, conv):
    d = Document()
    p = d.add_paragraph()
    for text, attr in (('u', 'underline'), ('up', 'superscript'), ('down', 'subscript')):
        run = p.add_run(text)
        if attr == 'underline':
            run.underline = True
        else:
            setattr(run.font, attr, True)
    rel = d.part.relate_to('https://example.com/a b', 'http://schemas.openxmlformats.org/'
                           'officeDocument/2006/relationships/hyperlink', is_external=True)
    p._p.append(parse_xml(f'<w:hyperlink {nsdecls("w", "r")} r:id="{rel}"><w:r><w:t>ext</w:t></w:r></w:hyperlink>'))
    p._p.append(parse_xml(f'<w:hyperlink {nsdecls("w")}><w:r><w:t> plain</w:t></w:r></w:hyperlink>'))
    p._p.append(parse_xml(f'<w:sdt {nsdecls("w")}><w:sdtContent><w:r><w:t> in sdt</w:t></w:r></w:sdtContent></w:sdt>'))
    p._p.append(parse_xml(f'<w:sdt {nsdecls("w")}/>'))
    p._p.append(omml(''))  # empty equation: dropped
    # picture without size information + a footnote id that doesn't exist
    run = p.add_run()
    run._r.append(parse_xml(f'<w:footnoteReference {nsdecls("w")} w:id="77"/>'))
    path = tmp_path / 'e.docx'
    d.save(path)
    tex = tex_of(conv.convert(str(path)))
    assert '\\underline{u}\\textsuperscript{up}\\textsubscript{down}' in tex
    assert '\\href{https://example.com/a b}{ext} plain in sdt' in tex


def test_docx_broken_footnotes_and_numbering(tmp_path, conv):
    d = Document()
    d.add_paragraph('x', style='List Number')
    path = tmp_path / 'f.docx'
    d.save(path)
    add_footnotes(path, {1: 'n'})
    with zipfile.ZipFile(path) as z:
        files = {n: z.read(n) for n in z.namelist()}
    files['word/footnotes.xml'] = b'<not xml'
    # a numbering definition with pieces missing
    files['word/numbering.xml'] = (
        f'<w:numbering {nsdecls("w")}><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"/></w:abstractNum>'
        f'<w:num w:numId="1"/><w:num w:numId="2"><w:abstractNumId w:val="9"/></w:num>'
        f'<w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num></w:numbering>').encode()
    with zipfile.ZipFile(path, 'w') as z:
        for n, data in files.items():
            z.writestr(n, data)
    from doc2tex.latex import LatexGenerator
    gen = LatexGenerator(ConversionOptions())
    gen._doc = Document(path)
    assert gen._load_footnotes() == {}
    assert [gen._numbering_format(i, 0) for i in (1, 2, 3)] == [None, None, None]


def test_docx_without_numbering_part(tmp_path):
    from doc2tex.latex import LatexGenerator
    gen = LatexGenerator(ConversionOptions())

    class NoNumbering:
        @property
        def numbering_part(self):
            raise NotImplementedError

    class FakeDoc:
        part = NoNumbering()
    gen._doc = FakeDoc()
    assert gen._numbering_format(1, 0) is None


def test_docx_drawing_width_and_tables(tmp_path):
    from doc2tex.latex import LatexGenerator
    gen = LatexGenerator(ConversionOptions())
    gen._doc = Document()
    bad = parse_xml(f'<w:r {nsdecls("w", "wp")}><wp:extent cx="oops" cy="1"/></w:r>')
    assert gen._drawing_width(bad) is None
    tiny = parse_xml(f'<w:r {nsdecls("w", "wp")}><wp:extent cx="10" cy="1"/></w:r>')
    assert gen._drawing_width(tiny) is None
    t = gen._doc.add_table(rows=0, cols=0)
    assert gen._handle_table(t) == ''
    t = gen._doc.add_table(rows=1, cols=0)
    assert gen._handle_table(t) == ''


def test_docx_reraises_conversion_error(tmp_path, conv, monkeypatch, sample_docx_path):
    monkeypatch.setattr('doc2tex.latex.LatexGenerator._build_document',
                        lambda *a: (_ for _ in ()).throw(ConversionError('inner')))
    with pytest.raises(ConversionError, match='inner'):
        conv.convert(sample_docx_path)


@pytest.fixture
def sample_docx_path(tmp_path):
    d = Document()
    d.add_paragraph('x')
    d.save(tmp_path / 's.docx')
    return str(tmp_path / 's.docx')


def test_latex_reraises_conversion_error(tmp_path, conv, monkeypatch):
    path = tmp_path / 'x.tex'
    path.write_text('x', encoding='utf-8')
    monkeypatch.setattr('doc2tex.docx.DocxGenerator._parse_and_build',
                        lambda *a: (_ for _ in ()).throw(ConversionError('inner')))
    with pytest.raises(ConversionError, match='inner'):
        conv.convert(str(path))


def test_latex_inline_edge_cases(tmp_path, conv):
    (tmp_path / 'loop.tex').write_text('L\\input{loop}', encoding='utf-8')
    path = tmp_path / 'edge.tex'
    path.write_text(r"""\begin{document}
\input{loop}

\listoffigures
\vspace{1em}\label{top}Start {\em em} `single' $a & b$ $\mathbf{v}$ \mathit{q} \in \S
\begin{center}Centered\end{center}
\begin{tabular}{l}
cell \caption{inline cap} \\
\end{tabular}
\begin{tabular}{ll}
a & b & c \\
\end{tabular}
\begin{equation*}
x \\ y
\end{equation*}
\begin{equation}\label{e}
z
\end{equation}
\cref{e} and \cref{top} and \autoref{nolabel}.
\cite{*}
\textrm{rm} \unknown x \begin{x} \end{x} \textsl{sl}
\noindent{}text\
\end{document}""", encoding='utf-8')
    (tmp_path / 'edge.tex').write_text(path.read_text() + '\\', encoding='utf-8')
    d = Document(conv.convert(str(path)))
    texts = [p.text for p in d.paragraphs]
    joined = '\n'.join(texts)
    assert texts[0] == 'L' * 11  # self-input stops after a while
    assert any('nested too deeply' in w for w in conv.warnings)
    assert 'Start em ‘single’ ab v q ∈ §' in joined
    assert 'Centered' in texts
    assert [c.text for c in d.tables[0].rows[0].cells] == ['cell']
    assert [c.text for c in d.tables[1].rows[0].cells] == ['a', 'b', 'c']
    assert any('x' in t and 'y' in t and '\n' in t for t in texts)
    assert 'Equation (1) and ?? and ??. [*] rm x' in joined  # \label before any section
    assert 'sl text' in joined


def test_latex_restart_numbering_without_style(conv):
    from doc2tex.docx import DocxGenerator
    gen = DocxGenerator(ConversionOptions())
    gen.word_doc = Document()
    assert gen._restart_numbering('No Such Style') is None
    assert gen._restart_numbering('Normal') is None


def test_latex_bib_list_without_citations(tmp_path, conv):
    (tmp_path / 'r.bib').write_text('@misc{a, title={T}}', encoding='utf-8')
    path = tmp_path / 'b.tex'
    path.write_text('Nothing cited. \\bibliography{r}', encoding='utf-8')
    d = Document(conv.convert(str(path)))
    assert [p.text for p in d.paragraphs] == ['Nothing cited.']
    assert DocTeXConverter  # keeps flake quiet
    from doc2tex.docx import DocxGenerator
    assert DocxGenerator(ConversionOptions())._author_short('nokey') == 'nokey'


def test_latex_style_run_and_emit():
    from doc2tex.docx import DocxGenerator
    gen = DocxGenerator(ConversionOptions())
    gen.word_doc = Document()
    p = gen.word_doc.add_paragraph()
    gen._emit(p, '', {})
    assert p.runs == []
    gen._inline('\\underline{u}$x_2$\\texttt{m}', p, {})
    runs = {r.text: r for r in p.runs}
    assert runs['u'].underline and runs['2'].font.subscript and runs['m'].font.name == 'Courier New'
    assert DocxGenerator._accent("'", 'ab') == 'ab'
    gen._trim_runs(gen.word_doc.add_paragraph())  # empty paragraph: nothing to trim


def test_web_bad_download_id(client):
    assert client.get('/download/not-a-job-id/x.tex').status_code == 404


def test_web_cleanup_ignores_os_errors(client, monkeypatch):
    from doc2tex import web
    os.makedirs(os.path.join(web.app.config['UPLOAD_FOLDER'], 'c' * 32))
    monkeypatch.setattr(web.os.path, 'getmtime', lambda p: (_ for _ in ()).throw(OSError('gone')))
    web.cleanup_old_jobs()


@pytest.mark.filterwarnings('ignore::RuntimeWarning')
def test_modules_run_as_scripts(monkeypatch, capsys):
    import runpy
    import sys
    import flask
    monkeypatch.setattr(sys, 'argv', ['doc2tex', '--version'])
    with pytest.raises(SystemExit):
        runpy.run_module('doc2tex.cli', run_name='__main__')
    assert '2.0.0' in capsys.readouterr().out
    monkeypatch.setattr(flask.Flask, 'run', lambda self, **kw: None)
    runpy.run_module('doc2tex.web', run_name='__main__')
    utils.logger.setLevel('INFO')


def test_latex_remaining_edge_cases(tmp_path, conv):
    from doc2tex.docx import DocxGenerator
    # nested macro definitions and a macro used as another macro's argument
    assert expand_macros('\\newcommand{\\a}{\\newcommand{\\b}{x}}\\newcommand{\\c}[1]{(#1)}\\c{\\a}') == \
        '(\\newcommand{\\b}{x})'
    # @comment entries and an entry with no fields
    assert parse_bib('@comment{a, b}@misc{k, }') == {'k': {'ENTRYTYPE': 'misc'}}

    (tmp_path / 'r.bib').write_text('@misc{z, title={Zed}}', encoding='utf-8')
    path = tmp_path / 'x.tex'
    path.write_text('\\nocite{*}\\begin{tabular}{l}\\end{tabular}\n'
                    '\\begin{verbatim}\nline1\nline2\n\\end{verbatim}\n'
                    '\\section{A \\caption{c} \\begin{x} b}\n\\bibliography{r}\nend\\', encoding='utf-8')
    d = Document(conv.convert(str(path)))
    texts = [p.text for p in d.paragraphs]
    assert 'line1\nline2' in texts
    assert '1\u2003A b' in texts
    assert '[1] Zed.' in texts
    assert texts[-1] == 'end'
    assert len(d.tables) == 0

    gen = DocxGenerator(ConversionOptions())
    gen.word_doc = Document()
    p = gen.word_doc.add_paragraph()
    gen._inline('a}b\\', p, {})
    assert p.text == 'ab'
