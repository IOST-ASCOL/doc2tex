# Tests for doc2tex
# Run with: python -m pytest

import io
import os
import re
import zipfile

import pytest
from docx import Document
from PIL import Image

from doc2tex import ConversionError, ConversionOptions, DocTeXConverter
from doc2tex.cli import main as cli_main
from doc2tex.utils import escape_latex, unescape_latex


@pytest.fixture
def converter():
    return DocTeXConverter(ConversionOptions())


def png_bytes(color='red'):
    buf = io.BytesIO()
    Image.new('RGB', (40, 20), color).save(buf, 'PNG')
    buf.seek(0)
    return buf


@pytest.fixture
def sample_docx(tmp_path):
    d = Document()
    d.add_paragraph('My Report', style='Title')
    d.add_heading('Intro & Setup', 1)
    p = d.add_paragraph('Costs 50% of $10, path a\\b. ')
    r = p.add_run('bold')
    r.bold = True
    r = p.add_run('both')
    r.bold = True
    r.italic = True
    d.add_paragraph('first bullet', style='List Bullet')
    d.add_paragraph('nested bullet', style='List Bullet 2')
    d.add_paragraph('second bullet', style='List Bullet')
    d.add_paragraph('step one', style='List Number')
    t = d.add_table(rows=2, cols=2)
    t.cell(0, 0).text = 'A'
    t.cell(0, 1).text = 'B'
    t.cell(1, 0).text = '1'
    t.cell(1, 1).text = '2'
    d.add_picture(png_bytes())
    path = tmp_path / 'sample.docx'
    d.save(path)
    return path


# ----------------------------------------------------------------------
# Basics
# ----------------------------------------------------------------------

def test_options_init():
    options = ConversionOptions()
    assert options.font_size.value == '12pt'
    assert options.standalone_document


def test_detect_direction(converter):
    assert converter._detect_direction('test.docx') == 'to_latex'
    assert converter._detect_direction('test.tex') == 'to_docx'
    assert converter._detect_direction('test.LATEX') == 'to_docx'


def test_unknown_extension_rejected(converter, tmp_path):
    f = tmp_path / 'notes.txt'
    f.write_text('hi')
    with pytest.raises(Exception):
        converter.convert(str(f))


def test_missing_file(converter):
    with pytest.raises(ConversionError):
        converter.convert('does-not-exist.docx')


@pytest.mark.parametrize('text', [
    'plain', 'a\\b', '{braces}', '50% & $5 #1 x_y', '~^', '\\\\{}',
])
def test_escape_roundtrip(text):
    assert unescape_latex(escape_latex(text)) == text


def test_escape_backslash_is_valid_latex():
    # Used to come out as \textbackslash\{\} which doesn't compile
    assert escape_latex('a\\b') == 'a\\textbackslash{}b'


# ----------------------------------------------------------------------
# DOCX -> LaTeX
# ----------------------------------------------------------------------

def test_docx_to_latex(converter, sample_docx, tmp_path):
    out = converter.convert(str(sample_docx))
    tex = open(out, encoding='utf-8').read()

    assert '\\title{My Report}' in tex and '\\maketitle' in tex
    assert '\\section{Intro \\& Setup}' in tex
    assert 'Costs 50\\% of \\$10, path a\\textbackslash{}b.' in tex
    assert '\\textbf{bold}' in tex
    assert '\\textit{\\textbf{both}}' in tex
    # Lists, including nesting
    assert '\\begin{itemize}' in tex and '\\item first bullet' in tex
    assert tex.index('\\item nested bullet') < tex.index('\\item second bullet')
    assert tex.count('\\begin{itemize}') == 2
    assert '\\begin{enumerate}' in tex and '\\item step one' in tex
    # Tables use booktabs only (no vertical rules / \hline mixed in)
    assert '\\begin{tabular}{ll}' in tex
    assert 'A & B \\\\' in tex and '\\hline' not in tex
    # Images get extracted next to the .tex file
    # ...in a figure, at the size they had in Word
    assert re.search(r'\\begin\{figure\}\[htbp\]\n\\centering\n'
                     r'\\includegraphics\[width=0\.\d\d\\textwidth\]\{sample-image1\.png\}', tex)
    assert (tmp_path / 'images' / 'sample-image1.png').is_file()
    assert tex.rstrip().endswith('\\end{document}')


def test_docx_to_latex_without_images(sample_docx, tmp_path):
    c = DocTeXConverter(ConversionOptions(preserve_images=False))
    tex = open(c.convert(str(sample_docx)), encoding='utf-8').read()
    assert 'includegraphics' not in tex
    assert not (tmp_path / 'images').exists()


def test_docx_merged_cells(converter, tmp_path):
    d = Document()
    t = d.add_table(rows=2, cols=3)
    t.cell(0, 0).merge(t.cell(0, 2)).text = 'Wide'
    for i, v in enumerate('xyz'):
        t.cell(1, i).text = v
    path = tmp_path / 'merged.docx'
    d.save(path)
    tex = open(converter.convert(str(path)), encoding='utf-8').read()
    assert '\\multicolumn{3}{c}{Wide} \\\\' in tex
    assert 'x & y & z \\\\' in tex


# ----------------------------------------------------------------------
# LaTeX -> DOCX
# ----------------------------------------------------------------------

LATEX_SAMPLE = r"""
\documentclass{article}
\graphicspath{{./figs/}}
\title{Lab \textbf{Report}}
\author{A \and B}
\begin{document}
\maketitle
\section*{Intro}
Text right after the heading, \textbf{bold \textit{nested}} and $x^2$. % comment
See \href{https://example.com}{the site}.

\begin{itemize}
  \item one
  \begin{enumerate}
    \item inner
  \end{enumerate}
  \item two
\end{itemize}

\begin{table}[h]
\centering
\caption{Results}
\begin{tabular}{|l|c|}
\toprule
\multicolumn{2}{c}{Header} \\
\midrule
A \& co & {\bf B} \\
\bottomrule
\end{tabular}
\end{table}

\begin{figure}[h]
\includegraphics[width=0.5\textwidth]{plot}
\caption{A plot}
\end{figure}

\begin{verbatim}
x = 5 % kept
\end{verbatim}
\end{document}
"""


@pytest.fixture
def sample_tex(tmp_path):
    (tmp_path / 'figs').mkdir()
    Image.new('RGB', (40, 20), 'blue').save(tmp_path / 'figs' / 'plot.png')
    path = tmp_path / 'in.tex'
    path.write_text(LATEX_SAMPLE, encoding='utf-8')
    return path


def test_latex_to_docx(converter, sample_tex):
    out = converter.convert(str(sample_tex))
    d = Document(out)
    paras = [(p.style.name, p.text) for p in d.paragraphs]
    texts = [t for _, t in paras]

    assert ('Title', 'Lab Report') in paras
    assert 'A, B' in texts
    # \section* used to swallow the paragraph after it
    assert ('Heading 1', 'Intro') in paras
    body = next(p for p in d.paragraphs if p.text.startswith('Text right after'))
    assert 'comment' not in body.text
    assert body.text.startswith('Text right after the heading, bold nested and x2.')
    nested = next(r for r in body.runs if r.text == 'nested')
    assert nested.bold and nested.italic
    assert next(r for r in body.runs if r.text == '2').font.superscript
    link = next(p for p in d.paragraphs if 'the site' in p.text)
    assert link._p.xpath('.//w:hyperlink')

    assert ('List Bullet', 'one') in paras
    assert ('List Number 2', 'inner') in paras
    assert ('List Bullet', 'two') in paras

    assert ('Caption', 'Table 1: Results') in paras
    assert ('Caption', 'Figure 1: A plot') in paras
    assert '[h]' not in ' '.join(texts)
    assert any(p._p.xpath('.//a:blip') for p in d.paragraphs)
    assert 'x = 5 % kept' in texts

    assert len(d.tables) == 1
    rows = [[c.text for c in r.cells] for r in d.tables[0].rows]
    assert rows == [['Header', 'Header'], ['A & co', 'B']]
    assert d.tables[0].rows[1].cells[1].paragraphs[0].runs[0].bold


def test_latex_snippet_without_document_env(converter, tmp_path):
    path = tmp_path / 'snippet.tex'
    path.write_text('Just \\emph{some} text.', encoding='utf-8')
    d = Document(converter.convert(str(path)))
    assert [p.text for p in d.paragraphs] == ['Just some text.']


def test_bare_tabular(converter, tmp_path):
    path = tmp_path / 't.tex'
    path.write_text('\\begin{tabular}{cc}\na & b \\\\\nc & d\n\\end{tabular}', encoding='utf-8')
    d = Document(converter.convert(str(path)))
    assert [[c.text for c in r.cells] for r in d.tables[0].rows] == [['a', 'b'], ['c', 'd']]


def test_roundtrip_docx_tex_docx(converter, sample_docx, tmp_path):
    tex = converter.convert(str(sample_docx))
    back = converter.convert(tex, str(tmp_path / 'back.docx'))
    texts = [p.text for p in Document(back).paragraphs]
    assert 'My Report' in texts
    assert '1\u2003Intro & Setup' in texts
    assert 'Costs 50% of $10, path a\\b. boldboth' in texts
    assert 'nested bullet' in texts


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------

def test_cli_single(sample_docx, tmp_path):
    out = tmp_path / 'out' / 'result.tex'
    assert cli_main([str(sample_docx), '-o', str(out)]) == 0
    assert out.is_file()


def test_cli_batch(sample_docx, sample_tex, tmp_path):
    out_dir = tmp_path / 'batch'
    assert cli_main([str(sample_docx), str(sample_tex), '-d', str(out_dir)]) == 0
    assert (out_dir / 'sample.tex').is_file()
    assert (out_dir / 'in.docx').is_file()


def test_cli_batch_reports_failures(sample_docx, tmp_path):
    assert cli_main([str(sample_docx), str(tmp_path / 'missing.docx'),
                     '-d', str(tmp_path / 'o')]) == 1


# ----------------------------------------------------------------------
# Web
# ----------------------------------------------------------------------

@pytest.fixture
def client(tmp_path):
    from doc2tex.web import app
    app.config['UPLOAD_FOLDER'] = str(tmp_path / 'uploads')
    os.makedirs(app.config['UPLOAD_FOLDER'])
    app.config['TESTING'] = True
    return app.test_client()


def test_web_docx_with_images_returns_zip(client, sample_docx):
    with open(sample_docx, 'rb') as f:
        resp = client.post('/convert', data={'file': (f, 'sample.docx'), 'line_spacing': 'double'})
    data = resp.get_json()
    assert resp.status_code == 200 and data['success'], data
    assert data['output_filename'] == 'sample.zip'

    dl = client.get(data['download_url'])
    assert dl.status_code == 200
    names = zipfile.ZipFile(io.BytesIO(dl.data)).namelist()
    assert 'sample.tex' in names and 'images/sample-image1.png' in names


def test_web_tex_to_docx(client):
    resp = client.post('/convert', data={'file': (io.BytesIO(b'Hello \\textbf{world}'), 'a.tex')})
    data = resp.get_json()
    assert data['success'] and data['output_filename'] == 'a.docx'
    assert client.get(data['download_url']).status_code == 200


def test_web_rejects_bad_input(client):
    assert client.post('/convert', data={}).status_code == 400
    resp = client.post('/convert', data={'file': (io.BytesIO(b'x'), 'a.pdf')})
    assert resp.status_code == 400
    resp = client.post('/convert', data={'file': (io.BytesIO(b'x'), 'a.tex'), 'font_size': '99pt'})
    assert resp.status_code == 400
    assert client.get('/download/../../etc/passwd').status_code == 404
    assert client.get('/download/' + 'a' * 32 + '/nope.tex').status_code == 404
