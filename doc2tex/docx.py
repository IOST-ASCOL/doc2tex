# doc2tex - LaTeX to Word conversion logic
# This way is harder because LaTeX is basically a programming language.
# Instead of one regex per feature, this walks the source character by character
# and understands nested {braces} and \begin...\end environments, which
# handles normal documents a lot better (nested bold/italic, \section*, etc).

import os
import re
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.enum.text import WD_BREAK, WD_PARAGRAPH_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.run import Run
from docx.shared import Emu, Pt, RGBColor

from .options import ConversionOptions, LineSpacing
from .symbols import MATH_TO_UNICODE
from .utils import logger
from .errors import ConversionError

# Sectioning commands -> Word heading level
_HEADINGS = {
    'part': 0, 'chapter': 1, 'section': 1, 'subsection': 2,
    'subsubsection': 3, 'paragraph': 4, 'subparagraph': 5,
}

# Inline formatting commands that take one argument
_FORMAT_COMMANDS = {
    'textbf': {'bold': True},
    'textit': {'italic': True},
    'emph': {'italic': True},
    'textsl': {'italic': True},
    'underline': {'underline': True},
    'uline': {'underline': True},
    'texttt': {'mono': True},
    'textsc': {'smallcaps': True},
    'textsuperscript': {'superscript': True},
    'textsubscript': {'subscript': True},
    'textrm': {'mono': False, 'italic': False},
    'textnormal': {'bold': False, 'italic': False, 'mono': False},
    'mathrm': {'italic': False},
    'text': {'italic': False},
    'mathbf': {'bold': True, 'italic': False},
    'mathit': {'italic': True},
}

# Declarations like {\bfseries text} that change the rest of the group
_DECLARATIONS = {
    'bfseries': {'bold': True}, 'bf': {'bold': True},
    'itshape': {'italic': True}, 'it': {'italic': True}, 'em': {'italic': True},
    'ttfamily': {'mono': True}, 'tt': {'mono': True},
    'scshape': {'smallcaps': True},
    'normalfont': {'bold': False, 'italic': False, 'mono': False},
}

# Commands whose arguments we throw away entirely
_CITE_COMMANDS = {'cite', 'citep', 'citet', 'parencite', 'autocite', 'textcite',
                  'citealp', 'citealt', 'citeauthor', 'citeyear', 'footcite', 'Cite',
                  'Citep', 'Citet', 'Textcite', 'Parencite', 'Autocite', 'nocite'}

_DROP_WITH_ARG = {
    'vspace', 'hspace', 'index', 'bibliographystyle', 'pagestyle',
    'thispagestyle', 'setlength', 'addtocounter', 'setcounter', 'newcommand',
    'renewcommand', 'graphicspath', 'usepackage', 'documentclass', 'color',
    'pagenumbering', 'addcontentsline',
}

# Commands that produce plain text
_TEXT_COMMANDS = {
    'LaTeX': 'LaTeX', 'TeX': 'TeX', 'ldots': '…', 'dots': '…', 'textellipsis': '…',
    'today': '', 'textendash': '–', 'textemdash': '—', 'textbackslash': '\\',
    'textasciitilde': '~', 'textasciicircum': '^', 'textbar': '|', 'S': '§',
    'textdegree': '°', 'copyright': '©', 'quad': '\u2003', 'qquad': '\u2003\u2003',
    'noindent': '', 'centering': '', 'raggedright': '', 'raggedleft': '',
    'small': '', 'large': '', 'Large': '', 'LARGE': '', 'huge': '', 'Huge': '',
    'footnotesize': '', 'scriptsize': '', 'tiny': '', 'normalsize': '',
    'par': '', 'medskip': '', 'bigskip': '', 'smallskip': '', 'hfill': ' ',
    'protect': '', 'relax': '', 'displaystyle': '', 'left': '', 'right': '',
}

# Math commands -> readable unicode so equations stay readable in Word
_MATH_SYMBOLS = {
    **MATH_TO_UNICODE,
    ',': '\u2009', ':': '\u2009', ';': ' ', '!': '', ' ': ' ',
}

_NUMBERED_MATH = {'equation', 'align', 'gather', 'multline', 'eqnarray', 'flalign', 'alignat'}
_MATH_ENVS = {'flalign', 'flalign*', 'alignat', 'alignat*', 'equation', 'equation*', 'align', 'align*', 'gather', 'gather*',
              'multline', 'multline*', 'displaymath', 'eqnarray', 'eqnarray*', 'math'}
_LIST_ENVS = {'itemize', 'enumerate', 'description'}
_TABLE_ENVS = {'tabular', 'tabular*', 'tabularx', 'longtable', 'array'}
_VERBATIM_ENVS = {'verbatim', 'verbatim*', 'lstlisting', 'minted'}
_RULE_RE = re.compile(r'\\(?:hline|toprule|midrule|bottomrule|cline\{[^}]*\}|'
                      r'cmidrule(?:\([^)]*\))?\{[^}]*\}|endhead|endfirsthead|endfoot|endlastfoot)')
_BEGIN_RE = re.compile(r'\\begin\s*\{([^}]+)\}')
_IMAGE_EXTS = ['', '.png', '.jpg', '.jpeg', '.gif', '.bmp']


# ----------------------------------------------------------------------
# Low level scanning helpers
# ----------------------------------------------------------------------

def read_group(s: str, i: int) -> Tuple[str, int]:
    # Reads a {...} group starting at s[i] == '{' (nested braces allowed).
    # Returns (contents, index after the closing brace).
    depth = 0
    j = i
    while j < len(s):
        c = s[j]
        if c == '\\':
            j += 2
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
        j += 1
    return s[i + 1:], len(s)


def skip_spaces(s: str, i: int) -> int:
    while i < len(s) and s[i] in ' \t\n':
        i += 1
    return i


def read_arg(s: str, i: int) -> Tuple[Optional[str], int]:
    # Reads a mandatory argument: {group} or a single token
    j = skip_spaces(s, i)
    if j >= len(s):
        return None, i
    if s[j] == '{':
        return read_group(s, j)
    if s[j] == '\\':
        m = re.match(r'\\([A-Za-z]+|.)', s[j:])
        return m.group(0), j + m.end()
    if s[j] in '}]':
        return None, i
    return s[j], j + 1


def read_optional(s: str, i: int) -> Tuple[Optional[str], int]:
    # Reads an optional [argument] if there is one
    j = skip_spaces(s, i)
    if j < len(s) and s[j] == '[':
        depth = 0
        k = j
        while k < len(s):
            if s[k] == '{':
                depth += 1
            elif s[k] == '}':
                depth -= 1
            elif s[k] == ']' and depth == 0:
                return s[j + 1:k], k + 1
            k += 1
    return None, i


def find_env_end(s: str, i: int, name: str) -> Tuple[int, int]:
    # Finds the matching \end{name} for an environment whose body starts at i.
    # Returns (start of \end, index after \end{name}).
    pattern = re.compile(r'\\(begin|end)\s*\{' + re.escape(name) + r'\}')
    depth = 1
    for m in pattern.finditer(s, i):
        depth += 1 if m.group(1) == 'begin' else -1
        if depth == 0:
            return m.start(), m.end()
    return len(s), len(s)


def split_top_level(s: str, sep: str) -> List[str]:
    # Splits on a separator ('&', '\\\\' or '\\item') but only at the top level,
    # i.e. not inside {groups} or nested environments.
    parts = []
    depth = 0
    env_depth = 0
    start = 0
    i = 0
    while i < len(s):
        if s.startswith(sep, i) and depth == 0 and env_depth == 0:
            # \item must not be the start of e.g. \itemsep
            if not (sep[-1].isalpha() and i + len(sep) < len(s) and s[i + len(sep)].isalpha()):
                parts.append(s[start:i])
                i += len(sep)
                start = i
                continue
        c = s[i]
        if c == '\\':
            if s.startswith('\\begin', i):
                env_depth += 1
            elif s.startswith('\\end', i) and not s.startswith('\\endhead', i):
                env_depth = max(0, env_depth - 1)
            i += 2
            continue
        if c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
        i += 1
    parts.append(s[start:])
    return parts


def strip_comments(s: str) -> str:
    # Removes % comments, but leaves verbatim blocks and \% alone
    pieces = re.split(r'(\\begin\{(?:verbatim\*?|lstlisting|minted)\}.*?\\end\{(?:verbatim\*?|lstlisting|minted)\})',
                      s, flags=re.DOTALL)
    out = []
    for idx, piece in enumerate(pieces):
        if idx % 2 == 1:
            out.append(piece)
            continue
        # A comment swallows the line ending too (that's what TeX does)
        out.append(re.sub(r'(?<!\\)((?:\\\\)*)%[^\n]*\n?[ \t]*', r'\1', piece))
    return ''.join(out)


_MACRO_DEF_RE = re.compile(
    r'\\(?:(?:re)?newcommand|providecommand|DeclareRobustCommand)\*?\s*'
    r'(?:\{\s*\\([A-Za-z]+)\s*\}|\\([A-Za-z]+))'
    r'|\\def\s*\\([A-Za-z]+)\s*(?=\{)'
    r'|\\DeclareMathOperator\*?\s*\{\s*\\([A-Za-z]+)\s*\}')


def expand_macros(tex: str) -> str:
    # Handles simple \newcommand / \def / \DeclareMathOperator macros by
    # pasting their definition wherever they're used (#1, #2 ... included).
    macros: Dict[str, Tuple[int, Optional[str], str]] = {}
    out = []
    pos = 0
    for m in _MACRO_DEF_RE.finditer(tex):
        if m.start() < pos:
            continue
        name = next(g for g in m.groups() if g)
        j = m.end()
        nargs, default = 0, None
        if m.group(4):  # \DeclareMathOperator{\name}{text}
            body, j = read_arg(tex, j)
            body = '\\operatorname{' + (body or '') + '}'
        else:
            n, j = read_optional(tex, j)
            if n and n.strip().isdigit():
                nargs = int(n)
                default, j = read_optional(tex, j)
            body, j = read_arg(tex, j)
        if body is None:
            continue
        macros[name] = (nargs, default, body)
        out.append(tex[pos:m.start()])
        pos = j
    out.append(tex[pos:])
    tex = ''.join(out)
    if not macros:
        return tex

    use_re = re.compile(r'\\(' + '|'.join(sorted(map(re.escape, macros), key=len, reverse=True)) +
                        r')(?![A-Za-z])')
    for _ in range(10):  # macros can use other macros
        changed = False
        out = []
        pos = 0
        for m in use_re.finditer(tex):
            if m.start() < pos:
                continue
            nargs, default, body = macros[m.group(1)]
            j = m.end()
            args = []
            if default is not None:
                opt, j = read_optional(tex, j)
                args.append(default if opt is None else opt)
            while len(args) < nargs:
                arg, j = read_arg(tex, j)
                args.append(arg or '')
            expanded = body
            for k, a in enumerate(args, start=1):
                expanded = expanded.replace(f'#{k}', a)
            # keep "\foo bar" from turning into "...bar" glued to the macro text
            space = ''
            if nargs == 0 and j < len(tex) and tex[j] == ' ':
                space = ' '
                j += 1
            out.append(tex[pos:m.start()])
            if nargs == 0 and not body.startswith('\\'):
                expanded = '{' + expanded + '}'
            out.append(expanded + space)
            pos = j
            changed = True
        out.append(tex[pos:])
        tex = ''.join(out)
        if not changed:
            break
    return tex


def parse_bib(text: str) -> Dict[str, Dict[str, str]]:
    # A small BibTeX reader: @type{key, field = {value} or "value" or 2020, ...}
    entries: Dict[str, Dict[str, str]] = {}
    for m in re.finditer(r'@(\w+)\s*\{\s*([^,\s]+)\s*,', text):
        kind = m.group(1).lower()
        if kind in ('comment', 'preamble', 'string'):
            continue
        fields = {'ENTRYTYPE': kind}
        i = m.end()
        while i < len(text):
            fm = re.compile(r'\s*(\w+)\s*=\s*').match(text, i)
            if not fm:
                break
            i = fm.end()
            if i < len(text) and text[i] == '{':
                val, i = read_group(text, i)
            elif i < len(text) and text[i] == '"':
                end = text.find('"', i + 1)
                end = len(text) if end == -1 else end
                val, i = text[i + 1:end], end + 1
            else:
                vm = re.compile(r'[^,}\s]+').match(text, i)
                val = vm.group(0) if vm else ''
                i = vm.end() if vm else i
            fields[fm.group(1).lower()] = ' '.join(val.split())
            cm = re.compile(r'\s*,?').match(text, i)
            i = cm.end()
            if i < len(text) and text[i] == '}':
                break
        entries[m.group(2)] = fields
    return entries


def format_bib_entry(e: Dict[str, str]) -> str:
    # Plain-ish reference text, roughly like the "plain" bibliography style
    authors = [a.strip() for a in re.split(r'\s+and\s+', e.get('author') or e.get('editor') or '') if a.strip()]
    authors = [' '.join(reversed([p.strip() for p in a.split(',', 1)])) if ',' in a else a for a in authors]
    if len(authors) > 2:
        who = ', '.join(authors[:-1]) + ', and ' + authors[-1]
    else:
        who = ' and '.join(authors)
    parts = [who] if who else []
    if e.get('title'):
        parts.append(e['title'])
    venue = e.get('journal') or e.get('booktitle') or e.get('publisher') or e.get('school') \
        or e.get('institution') or e.get('howpublished') or ''
    if venue:
        extra = venue
        if e.get('volume'):
            extra += f", {e['volume']}"
            if e.get('number'):
                extra += f"({e['number']})"
        if e.get('pages'):
            extra += f", pp. {e['pages'].replace('--', '–')}"
        parts.append('\\textit{' + extra + '}' if e.get('journal') or e.get('booktitle') else extra)
    if e.get('year'):
        parts.append(e['year'])
    text = '. '.join(p.rstrip('.') for p in parts) + '.'
    if e.get('url'):
        text += ' \\url{' + e['url'] + '}'
    elif e.get('doi'):
        text += ' \\url{https://doi.org/' + e['doi'] + '}'
    return text


# siunitx unit macros -> symbols
_SI_UNITS = {
    'metre': 'm', 'meter': 'm', 'second': 's', 'kilogram': 'kg', 'gram': 'g',
    'ampere': 'A', 'kelvin': 'K', 'mole': 'mol', 'candela': 'cd', 'hertz': 'Hz',
    'newton': 'N', 'pascal': 'Pa', 'joule': 'J', 'watt': 'W', 'volt': 'V',
    'ohm': 'Ω', 'coulomb': 'C', 'farad': 'F', 'henry': 'H', 'tesla': 'T',
    'weber': 'Wb', 'siemens': 'S', 'celsius': '°C', 'degreeCelsius': '°C',
    'percent': '%', 'litre': 'L', 'liter': 'L', 'minute': 'min', 'hour': 'h',
    'day': 'd', 'degree': '°', 'radian': 'rad', 'steradian': 'sr', 'decibel': 'dB',
    'bel': 'B', 'byte': 'B', 'bit': 'bit', 'electronvolt': 'eV', 'bar': 'bar',
    'lumen': 'lm', 'lux': 'lx', 'becquerel': 'Bq', 'gray': 'Gy', 'sievert': 'Sv',
    'tonne': 't', 'angstrom': 'Å', 'arcminute': '′', 'arcsecond': '″',
}
_SI_PREFIXES = {
    'yocto': 'y', 'zepto': 'z', 'atto': 'a', 'femto': 'f', 'pico': 'p', 'nano': 'n',
    'micro': 'µ', 'milli': 'm', 'centi': 'c', 'deci': 'd', 'deca': 'da', 'deka': 'da',
    'hecto': 'h', 'kilo': 'k', 'mega': 'M', 'giga': 'G', 'tera': 'T', 'peta': 'P',
    'exa': 'E',
}
_SUPERSCRIPTS = str.maketrans('0123456789-+', '⁰¹²³⁴⁵⁶⁷⁸⁹⁻⁺')


def si_unit_text(spec: str) -> str:
    # \si{\kilo\metre\per\second\squared} -> km/s², \si{kg.m/s^2} -> kg·m/s²
    spec = spec or ''
    if '\\' not in spec:
        spec = re.sub(r'\^\{?(-?\d+)\}?', lambda m: m.group(1).translate(_SUPERSCRIPTS), spec)
        return spec.replace('.', '·').replace('~', ' ').replace('*', '·')
    out = ''
    pending_power = ''
    prefix = ''
    for name in re.findall(r'\\([A-Za-z]+)|(\^\{?-?\d+\}?)', spec):
        cmd, power = name
        if power:
            out += re.sub(r'[\^{}]', '', power).translate(_SUPERSCRIPTS)
            continue
        if cmd in _SI_PREFIXES:
            prefix += _SI_PREFIXES[cmd]
        elif cmd in ('per',):
            out += '/'
        elif cmd in ('square', 'cubic'):
            pending_power = '²' if cmd == 'square' else '³'
        elif cmd in ('squared', 'cubed'):
            out += '²' if cmd == 'squared' else '³'
        elif cmd in _SI_UNITS:
            if out and not out.endswith('/'):
                out += '·'
            out += prefix + _SI_UNITS[cmd] + pending_power
            prefix = pending_power = ''
    return out


def _letter(n: int) -> str:
    # 1 -> A, 2 -> B ... (appendix numbering)
    return chr(ord('A') + (n - 1) % 26) if n > 0 else 'A'


def strip_comments_bib(text: str) -> str:
    # In .bib files only lines starting with % are comments
    return re.sub(r'(?m)^\s*%.*$', '', text)


# ----------------------------------------------------------------------
# The generator
# ----------------------------------------------------------------------

class DocxGenerator:
    """
    My class for generating Word docs from LaTeX.
    It walks the LaTeX source and builds the Word document as it goes.
    """

    def __init__(self, options: ConversionOptions):
        self.options = options
        self.word_doc = None
        self.base_dir = '.'
        self.graphics_paths: List[str] = []
        self.meta: Dict[str, str] = {}
        self.warnings: List[str] = []
        # Cross references: label -> (kind, number), filled in a first pass
        self.labels: Dict[str, Tuple[str, str]] = {}
        self.citations: Dict[str, int] = {}   # cite key -> [n]
        self.bib_entries: Dict[str, Dict[str, str]] = {}  # from .bib files
        self._quiet = False

    def convert(self, tex_file: str, docx_file: str) -> str:
        # Tries to turn your .tex into a .docx
        try:
            logger.info(f"Reading {tex_file}...")
            self.base_dir = os.path.dirname(os.path.abspath(tex_file))

            raw_tex = expand_macros(self._read_with_inputs(tex_file))
            self._load_bib_files(raw_tex)
            self._number_bibitems(raw_tex)

            # LaTeX needs two runs to get \ref right, and so do we:
            # the first pass only collects label numbers and citation order.
            self._quiet = True
            self._build(raw_tex)
            self._quiet = False
            self._build(raw_tex)

            self.word_doc.save(docx_file)
            logger.info(f"Saved the Word doc to {docx_file}")
            return docx_file

        except ConversionError:
            raise
        except Exception as err:
            logger.error(f"LaTeX parsing failed: {err}")
            raise ConversionError(f"Something went wrong reading the LaTeX file: {err}") from err

    def _build(self, raw_tex: str) -> None:
        # Create a blank Word document and fill it
        self.word_doc = Document()
        self.meta = {}
        self.graphics_paths = []
        self._counters = {'sec': [0] * 6, 'figure': 0, 'table': 0, 'equation': 0}
        self._ref_target: Tuple[str, str] = ('', '')
        self._chapters = bool(re.search(r'\\chapter\b', raw_tex))
        self._appendix = False
        self._cite_order: List[str] = []
        self._apply_student_styles(raw_tex)
        self._parse_and_build(raw_tex)

    def _warn(self, msg: str) -> None:
        if self._quiet:
            return
        self.warnings.append(msg)
        logger.warning(msg)

    def _read_with_inputs(self, path: str, depth: int = 0) -> str:
        # Reads a .tex file and pastes in any \input{} / \include{} files
        with open(path, 'r', encoding=self.options.output_encoding, errors='replace') as f:
            text = strip_comments(f.read())
        if depth > 10:
            self._warn(f"\\input nested too deeply in {os.path.basename(path)}, stopped there")
            return re.sub(r'\\(input|include)\s*\{[^}]*\}', '', text)

        def replace(m):
            name = m.group(2)
            candidates = [name, name + '.tex']
            for c in candidates:
                full = os.path.join(self.base_dir, c)
                if os.path.isfile(full):
                    return self._read_with_inputs(full, depth + 1)
            self._warn(f"Could not find \\{m.group(1)}{{{name}}}")
            return ''

        return re.sub(r'\\(input|include)\s*\{([^}]+)\}', replace, text)

    def _apply_student_styles(self, tex: str = '') -> None:
        # Setup the document styles to look like a standard report.
        # Settings written in the .tex file itself (\documentclass[11pt],
        # geometry margins, \onehalfspacing) win over the defaults.
        style = self.word_doc.styles['Normal']
        f = style.font
        f.name = 'Times New Roman'
        size = self.options.font_size.value
        m = re.search(r'\\documentclass\s*\[([^\]]*)\]', tex)
        if m:
            sm = re.search(r'\b(\d\d)pt\b', m.group(1))
            size = sm.group(1) + 'pt' if sm else size
        f.size = Pt(int(size.replace('pt', '')))

        spacing = {LineSpacing.SINGLE: 1.0, LineSpacing.ONE_HALF: 1.5,
                   LineSpacing.DOUBLE: 2.0}.get(self.options.line_spacing, 1.0)
        if re.search(r'\\onehalfspacing\b', tex):
            spacing = 1.5
        elif re.search(r'\\doublespacing\b', tex):
            spacing = 2.0
        elif re.search(r'\\singlespacing\b', tex):
            spacing = 1.0
        else:
            sm = re.search(r'\\(?:linespread|setstretch)\s*\{([\d.]+)\}', tex)
            if sm:
                # \linespread{1.3} is what \onehalfspacing does at 12pt
                spacing = round(float(sm.group(1)) * (1.15 if 'linespread' in sm.group(0) else 1), 2)
        if spacing != 1.0:
            style.paragraph_format.line_spacing = spacing

        margins = self._parse_margins(self.options.page_margins)
        gm = re.search(r'\\usepackage\s*\[([^\]]*)\]\s*\{geometry\}', tex) or \
            re.search(r'\\geometry\s*\{([^}]*)\}', tex)
        if gm:
            margins.update(self._parse_margins(gm.group(1)))
        for section in self.word_doc.sections:
            for side, emu in margins.items():
                setattr(section, f'{side}_margin', Emu(emu))

    @staticmethod
    def _parse_margins(spec: str) -> Dict[str, int]:
        # "top=2.5cm,left=1in" / "margin=2cm" -> {'top': emu, ...}
        units = {'cm': 360000, 'mm': 36000, 'in': 914400, 'pt': 12700}
        groups = {'margin': ('top', 'bottom', 'left', 'right'),
                  'hmargin': ('left', 'right'), 'vmargin': ('top', 'bottom'),
                  'lmargin': ('left',), 'rmargin': ('right',), 'tmargin': ('top',),
                  'bmargin': ('bottom',), 'inner': ('left',), 'outer': ('right',)}
        general: Dict[str, int] = {}
        specific: Dict[str, int] = {}
        for key, val, unit in re.findall(r'(\w+)\s*=\s*([\d.]+)\s*(cm|mm|in|pt)', spec or ''):
            emu = int(float(val) * units[unit])
            if key in ('top', 'bottom', 'left', 'right'):
                specific[key] = emu
            for side in groups.get(key, ()):
                general[side] = emu
        return {**general, **specific}

    def _parse_and_build(self, content: str) -> None:
        # Pull metadata out of the preamble
        preamble_end = content.find('\\begin{document}')
        preamble = content[:preamble_end] if preamble_end != -1 else ''
        self._read_metadata(preamble)

        # We only really care about stuff inside \begin{document}
        if preamble_end != -1:
            body_start = preamble_end + len('\\begin{document}')
            end_m = re.search(r'\\end\{document\}', content[body_start:])
            body = content[body_start:body_start + end_m.start()] if end_m else content[body_start:]
        else:
            # Maybe it's just a snippet?
            body = content

        # \title etc can also live in the body
        self._read_metadata(body)
        self._render_blocks(body, {})

    def _read_metadata(self, text: str) -> None:
        for key in ('title', 'author', 'date'):
            m = re.search(r'\\' + key + r'\s*(?:\[[^\]]*\])?\s*\{', text)
            if m:
                self.meta[key], _ = read_group(text, m.end() - 1)
        for m in re.finditer(r'\\graphicspath\s*\{', text):
            group, _ = read_group(text, m.end() - 1)
            self.graphics_paths.extend(re.findall(r'\{([^}]*)\}', group))

    # ------------------------------------------------------------------
    # Block level
    # ------------------------------------------------------------------

    def _render_blocks(self, s: str, ctx: dict) -> None:
        # Walks a chunk of LaTeX, collecting paragraph text and dispatching
        # block-level things (headings, environments, display math)
        buf: List[str] = []

        def flush():
            text = ''.join(buf)
            buf.clear()
            if text.strip():
                self._add_paragraph(text, ctx)

        i = 0
        n = len(s)
        while i < n:
            # A blank line ends the paragraph
            m = re.match(r'[ \t]*\n[ \t]*\n\s*', s[i:])
            if m:
                flush()
                i += m.end()
                continue

            if s.startswith('\\[', i):
                flush()
                end = s.find('\\]', i + 2)
                end = n if end == -1 else end
                self._add_display_math(s[i + 2:end])
                i = end + 2
                continue
            if s.startswith('$$', i):
                flush()
                end = s.find('$$', i + 2)
                end = n if end == -1 else end
                self._add_display_math(s[i + 2:end])
                i = end + 2
                continue

            if s[i] == '\\':
                cm = re.match(r'\\([A-Za-z]+)(\*?)|\\.', s[i:], re.DOTALL)
                if cm is None:  # a lone backslash at the very end
                    i += 1
                    continue
                name = cm.group(1)
                after = i + cm.end()

                if name == 'begin':
                    env, j = read_arg(s, after)
                    env = (env or '').strip()
                    end_start, end_after = find_env_end(s, j, env)
                    inner = s[j:end_start]
                    if env in ('math',):
                        buf.append('$' + inner + '$')
                    else:
                        flush()
                        self._handle_environment(env, inner, ctx)
                    i = end_after
                    continue

                if name in _HEADINGS:
                    flush()
                    _, j = read_optional(s, after)
                    title, j = read_arg(s, j)
                    self._add_heading(title or '', name, numbered=not cm.group(2))
                    i = j
                    continue

                if name == 'maketitle':
                    flush()
                    self._add_title_block()
                    i = after
                    continue

                if name in ('newpage', 'clearpage', 'pagebreak', 'cleardoublepage'):
                    flush()
                    self.word_doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
                    i = after
                    continue

                if name == 'tableofcontents':
                    flush()
                    self._add_toc()
                    i = after
                    continue

                if name == 'appendix':
                    flush()
                    self._appendix = True
                    self._counters['sec'] = [0] * 6
                    i = after
                    continue

                if name in ('printbibliography', 'bibliography'):
                    flush()
                    j = after
                    if name == 'bibliography':
                        _, j = read_arg(s, j)
                    else:
                        _, j = read_optional(s, j)
                    self._add_bib_list()
                    i = j
                    continue

                if name in ('listoffigures', 'listoftables', 'frontmatter',
                            'mainmatter', 'backmatter'):
                    flush()
                    i = after
                    continue

                if name in ('title', 'author', 'date', 'caption', 'bibliographystyle',
                            'addbibresource'):
                    flush()
                    _, j = read_optional(s, after)
                    _, j = read_arg(s, j)
                    i = j
                    continue

                # Anything else is inline - keep the whole command token together
                buf.append(s[i:after])
                i = after
                continue

            buf.append(s[i])
            i += 1

        flush()

    def _handle_environment(self, env: str, inner: str, ctx: dict) -> None:
        base = env.rstrip('*')
        if env in _LIST_ENVS:
            self._add_list(env, inner, ctx)
        elif env in _TABLE_ENVS or base in _TABLE_ENVS:
            self._add_table(env, inner)
        elif base in ('table', 'figure', 'wrapfigure', 'subfigure'):
            self._add_float(inner, ctx)
        elif env in _MATH_ENVS:
            self._add_display_math(inner, env)
        elif env in _VERBATIM_ENVS:
            self._add_verbatim(inner)
        elif env == 'center':
            self._render_blocks(inner, {**ctx, 'align': WD_PARAGRAPH_ALIGNMENT.CENTER})
        elif env == 'flushright':
            self._render_blocks(inner, {**ctx, 'align': WD_PARAGRAPH_ALIGNMENT.RIGHT})
        elif env == 'flushleft':
            self._render_blocks(inner, {**ctx, 'align': WD_PARAGRAPH_ALIGNMENT.LEFT})
        elif env in ('quote', 'quotation'):
            self._render_blocks(inner, {**ctx, 'style': 'Quote'})
        elif env == 'abstract':
            p = self.word_doc.add_paragraph()
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p.add_run('Abstract').bold = True
            self._render_blocks(inner, ctx)
        elif env == 'thebibliography':
            self._add_bibliography(inner)
        else:
            # Unknown environment (minipage, document parts, ...): just render what's inside
            _, j = read_optional(inner, 0)
            if env in ('minipage', 'tabbing'):
                _, j = read_arg(inner, j)
            self._render_blocks(inner[j:], ctx)

    def _add_paragraph(self, text: str, ctx: dict, style: Optional[str] = None):
        p = self.word_doc.add_paragraph(style=style or ctx.get('style'))
        if ctx.get('align') is not None:
            p.alignment = ctx['align']
        self._apply_inline(text, p)
        self._trim_runs(p)
        # Things like a lone \centering or \label{} produce nothing visible
        if not p.text.strip() and not p._p.xpath('.//w:drawing | .//w:hyperlink | .//w:br'):
            p._p.getparent().remove(p._p)
            return None
        return p

    def _add_heading(self, title: str, command: str = 'section', numbered: bool = True) -> None:
        # With \chapter in the document, \section is one level further down
        level = _HEADINGS.get(command, 1)
        if self._chapters and command not in ('part', 'chapter'):
            level += 1
        level = min(level, 9)
        p = self.word_doc.add_heading('', level=level)

        # Number it like LaTeX would (1, 1.1, 1.1.1; chapters and appendices too)
        if command != 'part' and level >= 1:
            sec = self._counters['sec']
            sec[level - 1] += 1
            for k in range(level, len(sec)):
                sec[k] = 0
            if command == 'chapter':
                # figures, tables and equations count from 1 again in each chapter
                for kind in ('figure', 'table', 'equation'):
                    self._counters[kind] = 0
            if numbered and self.options.number_sections and level <= 3:
                nums = [str(n) for n in sec[:level]]
                if self._appendix:
                    nums[0] = _letter(sec[0])
                number = '.'.join(nums)
                if command == 'chapter':
                    label = ('Appendix ' if self._appendix else 'Chapter ') + number
                    p.add_run(label)
                    p.add_run().add_break()
                else:
                    p.add_run(number + '\u2003')
                self._ref_target = ('section', number)
        self._apply_inline(title, p)
        self._trim_runs(p)

    def _add_title_block(self) -> None:
        if self.meta.get('title'):
            p = self.word_doc.add_paragraph(style='Title')
            self._apply_inline(self.meta['title'], p)
            self._trim_runs(p)
        for key in ('author', 'date'):
            val = self.meta.get(key, '')
            if val.strip() and val.strip() != '\\today':
                text = re.sub(r'\s*\\and\b\s*', ', ', val)
                p = self.word_doc.add_paragraph()
                p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                self._apply_inline(text, p)
                self._trim_runs(p)

    def _add_list(self, env: str, inner: str, ctx: dict) -> None:
        # Bullet/numbered/description lists, with nesting
        level = ctx.get('list_level', 0) + 1
        suffix = '' if level == 1 else f' {min(level, 3)}'
        if env == 'enumerate':
            style = 'List Number' + suffix
        elif env == 'itemize':
            style = 'List Bullet' + suffix
        else:
            style = None

        # Word keeps counting across every "List Number" paragraph in the file,
        # so each enumerate gets its own numbering instance starting at 1
        numbering = self._restart_numbering(style) if env == 'enumerate' else None

        items = split_top_level(inner, '\\item')
        for item in items[1:]:  # anything before the first \item is setup junk
            label, j = read_optional(item, 0)
            body = item[j:]
            if env == 'description' and label:
                body = '\\textbf{' + label + '} ' + body
            item_ctx = {**ctx, 'list_level': level}
            # The first paragraph gets the bullet; later ones in the same item don't
            self._render_list_item(body, style, item_ctx, numbering)

    def _render_list_item(self, body: str, style: Optional[str], ctx: dict,
                          numbering: Optional[Tuple[int, int]] = None) -> None:
        start = len(self.word_doc.paragraphs)
        self._render_blocks(body, {**ctx, 'style': 'List Continue' if style else None})
        paras = self.word_doc.paragraphs[start:]
        if paras and style:
            first = paras[0]
            first.style = self.word_doc.styles[style]
        elif not paras:
            first = self.word_doc.add_paragraph(style=style)
        else:
            return
        if numbering is not None:
            num_pr = first._p.get_or_add_pPr().get_or_add_numPr()
            num_pr.get_or_add_ilvl().val = numbering[1]
            num_pr.get_or_add_numId().val = numbering[0]

    def _restart_numbering(self, style_name: Optional[str]) -> Optional[Tuple[int, int]]:
        # Makes a new <w:num> pointing at the same list definition as the
        # style, with the counter reset to 1. Returns (numId, ilvl).
        try:
            style = self.word_doc.styles[style_name]
            num_pr = style.element.pPr.numPr
            style_num_id = num_pr.numId.val
            ilvl = num_pr.ilvl.val if num_pr.ilvl is not None else 0
            numbering = self.word_doc.part.numbering_part.element
            num = numbering.num_having_numId(style_num_id)
            new_num = numbering.add_num(num.abstractNumId.val)
            new_num.add_lvlOverride(ilvl=ilvl).add_startOverride(1)
            return new_num.numId, ilvl
        except (KeyError, AttributeError, ValueError):
            return None

    def _add_float(self, inner: str, ctx: dict) -> None:
        # table / figure environments: images, tables and captions
        _, j = read_optional(inner, 0)  # placement like [htbp]
        inner = inner[j:]
        caption = None
        cm = re.search(r'\\caption\s*(?:\[[^\]]*\])?\s*\{', inner)
        if cm:
            caption, _ = read_group(inner, cm.end() - 1)

        is_table = bool(re.search(r'\\begin\s*\{(?:tabular|longtable|tabularx|array)', inner))
        kind = 'table' if is_table else 'figure'
        if caption is not None:
            # LaTeX gives the float its number at \caption
            self._counters[kind] += 1
            number = self._float_number(self._counters[kind])
            self._ref_target = (kind, number)
            caption = f"{kind.capitalize()} {number}: {caption}"

        # Word convention: table captions above, figure captions below
        if caption is not None and is_table:
            self._add_paragraph(caption, {}, style='Caption')

        self._render_blocks(inner, {**ctx, 'align': WD_PARAGRAPH_ALIGNMENT.CENTER, 'in_float': True})

        if caption is not None and not is_table:
            self._add_paragraph(caption, {}, style='Caption')

    def _float_number(self, n: int) -> str:
        # Figures are numbered per chapter in reports/books (Figure 2.3)
        if self._chapters and self._counters['sec'][0]:
            chap = self._counters['sec'][0]
            return f"{_letter(chap) if self._appendix else chap}.{n}"
        return str(n)

    def _resolve_image(self, path: str) -> Optional[str]:
        # Looks for the image next to the .tex file and in \graphicspath folders
        dirs = [''] + self.graphics_paths
        for d in dirs:
            for ext in _IMAGE_EXTS:
                candidate = os.path.join(self.base_dir, d, path + ext)
                if os.path.isfile(candidate):
                    return candidate
        return None

    def _image_width(self, opts: Optional[str]):
        # Understands width=0.5\textwidth, width=8cm, width=3in, scale=0.5 (sort of)
        section = self.word_doc.sections[0]
        text_width = section.page_width - section.left_margin - section.right_margin
        if opts:
            m = re.search(r'width\s*=\s*([\d.]*)\s*\\(?:text|line|column)width', opts)
            if m:
                return int(text_width * float(m.group(1) or 1))
            m = re.search(r'width\s*=\s*([\d.]+)\s*(cm|mm|in|pt)', opts)
            if m:
                val, unit = float(m.group(1)), m.group(2)
                emu = {'cm': 360000, 'mm': 36000, 'in': 914400, 'pt': 12700}[unit]
                return min(int(val * emu), text_width)
        return int(text_width * 0.8)

    def _add_table(self, env: str, inner: str) -> None:
        # Rebuilds a Word table from tabular/longtable source
        j = 0
        if env == 'tabular*':
            _, j = read_arg(inner, j)
        _, j = read_optional(inner, j)
        if env == 'tabularx':
            _, j = read_arg(inner, j)
        spec, j = read_arg(inner, j)
        body = _RULE_RE.sub('', inner[j:])
        body = re.sub(r'\\caption\s*\{[^}]*\}\s*(\\\\)?', '', body)

        rows = []
        for raw in split_top_level(body, '\\\\'):
            raw = re.sub(r'^\s*\[[^\]]*\]', '', raw)  # \\[2pt] spacing
            if not raw.strip():
                continue
            cells = []
            for cell in split_top_level(raw, '&'):
                cell = cell.strip()
                mc = re.match(r'\\multicolumn\s*\{(\d+)\}', cell)
                span = 1
                if mc:
                    span = int(mc.group(1))
                    _, k = read_arg(cell, mc.end())
                    cell, _ = read_arg(cell, k)
                    cell = (cell or '').strip()
                mr = re.match(r'\\multirow\s*\{[^}]*\}\s*\{[^}]*\}', cell)
                if mr:
                    cell, _ = read_arg(cell, mr.end())
                    cell = (cell or '').strip()
                cells.append((cell, span))
            rows.append(cells)

        if not rows:
            return

        spec_cols = len(re.findall(r'[lcrX]|[pmb]\s*\{', spec or ''))
        num_c = max([spec_cols] + [sum(sp for _, sp in r) for r in rows])

        t = self.word_doc.add_table(rows=len(rows), cols=num_c)
        t.style = 'Table Grid'
        for r_idx, cells in enumerate(rows):
            c_idx = 0
            for text, span in cells:
                target = t.cell(r_idx, c_idx)
                if span > 1:
                    end = min(c_idx + span, num_c) - 1
                    target = target.merge(t.cell(r_idx, end))
                para = target.paragraphs[0]
                self._apply_inline(text, para)
                self._trim_runs(para)
                c_idx += span
        # Leave a little gap after the table so the next paragraph doesn't stick to it
        self.word_doc.add_paragraph()

    def _add_display_math(self, math: str, env: str = 'displaymath') -> None:
        numbered = env in _NUMBERED_MATH
        # equation/multline get one number; align/gather/eqnarray one per line
        per_line = env.rstrip('*') in ('align', 'gather', 'eqnarray', 'flalign', 'alignat')
        lines = [ln for ln in split_top_level(math, '\\\\') if ln.strip()]
        if not per_line:
            lines = [' \\\\ '.join(lines)] if lines else []
        p = self.word_doc.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        for idx, line in enumerate(lines):
            if idx:
                p.add_run().add_break()
            tag = re.search(r'\\tag\*?\s*\{([^}]*)\}', line)
            skip = re.search(r'\\(nonumber|notag)\b', line)
            number = None
            if tag:
                number = tag.group(1)
            elif numbered and not skip:
                self._counters['equation'] += 1
                number = self._float_number(self._counters['equation'])
            for lm in re.finditer(r'\\label\s*\{([^}]*)\}', line):
                self._set_label(lm.group(1).strip(), ('equation', number or ''))
            line = re.sub(r'\\(label|tag|nonumber|notag)\b\*?(\s*\{[^}]*\})?', '', line)
            parts = split_top_level(line, '\\\\')
            for k, part in enumerate(parts):
                if k:
                    p.add_run().add_break()
                self._inline(part.replace('&', ''), p, {'italic': True, 'math': True})
            if number:
                p.add_run(f'\u2003\u2003({number})')
        self._trim_runs(p)

    def _add_verbatim(self, inner: str) -> None:
        inner = re.sub(r'^\s*(\[[^\]]*\])?(\{[^}]*\})?\n', '', inner, count=1)  # options/language
        p = self.word_doc.add_paragraph()
        for idx, line in enumerate(inner.strip('\n').split('\n')):
            if idx:
                p.add_run().add_break()
            run = p.add_run(line)
            run.font.name = 'Courier New'

    # ------------------------------------------------------------------
    # Cross references, citations, table of contents
    # ------------------------------------------------------------------

    def _set_label(self, key: str, target: Tuple[str, str]) -> None:
        if key and (self._quiet or key not in self.labels):
            self.labels[key] = target

    def _ref_text(self, command: str, key: str) -> str:
        kind, number = self.labels.get(key, ('', ''))
        if not number:
            if not self._quiet:
                self._warn(f"Undefined reference: {key}")
            return '??'
        if command == 'eqref':
            return f'({number})'
        if command in ('autoref', 'cref', 'Cref'):
            names = {'section': 'Section', 'figure': 'Figure', 'table': 'Table',
                     'equation': 'Equation', 'item': 'Item'}
            word = names.get(kind, '')
            if kind == 'equation':
                return f'{word} ({number})'
            return f'{word}\u00a0{number}' if word else number
        return number

    def _cite_text(self, command: str, keys: str, note1: Optional[str], note2: Optional[str]) -> str:
        key_list = [k.strip() for k in keys.split(',') if k.strip()]
        for k in key_list:
            if k not in self._cite_order and k != '*':
                self._cite_order.append(k)
            if k == '*':
                for bk in self.bib_entries:
                    if bk not in self._cite_order:
                        self._cite_order.append(bk)
        if command == 'nocite':
            return ''
        nums = []
        for k in key_list:
            n = self.citations.get(k)
            if n is None and not self._quiet:
                self._warn(f"Citation not found in bibliography: {k}")
            nums.append(str(n) if n is not None else k)
        # natbib: \cite[see][p. 4]{x} -> note1 is the prefix, note2 the page
        post, pre = (note2, note1) if note2 is not None else (note1, None)
        inner = ', '.join(nums)
        if pre:
            inner = pre.strip().replace('~', '\u00a0') + ' ' + inner
        if post:
            inner += ', ' + post.strip().replace('~', '\u00a0')
        if command in ('citeauthor',):
            return ', '.join(self._author_short(k) for k in key_list)
        if command in ('citeyear',):
            return ', '.join(self.bib_entries.get(k, {}).get('year', '?') for k in key_list)
        if command in ('citet', 'Citet', 'textcite', 'Textcite'):
            return ', '.join(f"{self._author_short(k)} [{self.citations.get(k, k)}]" for k in key_list)
        return f'[{inner}]'

    def _author_short(self, key: str) -> str:
        author = self.bib_entries.get(key, {}).get('author', '')
        names = [a.strip() for a in re.split(r'\s+and\s+', author) if a.strip()]
        if not names:
            return key
        last = [n.split(',')[0] if ',' in n else n.split()[-1] for n in names]
        if len(last) > 2:
            return last[0] + ' et al.'
        return ' and '.join(last)

    def _number_bibitems(self, tex: str) -> None:
        # \bibitem order decides the numbers when there's a thebibliography
        for idx, m in enumerate(re.finditer(r'\\bibitem\s*(?:\[[^\]]*\])?\s*\{([^}]*)\}', tex), 1):
            self.citations.setdefault(m.group(1).strip(), idx)

    def _load_bib_files(self, tex: str) -> None:
        # \bibliography{refs} / \addbibresource{refs.bib}
        names = []
        for m in re.finditer(r'\\(?:bibliography|addbibresource)\s*(?:\[[^\]]*\])?\s*\{([^}]*)\}', tex):
            names.extend(n.strip() for n in m.group(1).split(','))
        for name in names:
            path = os.path.join(self.base_dir, name if name.endswith('.bib') else name + '.bib')
            if not os.path.isfile(path):
                self._warn(f"Bibliography file not found: {os.path.basename(path)}")
                continue
            with open(path, encoding='utf-8', errors='replace') as f:
                self.bib_entries.update(parse_bib(strip_comments_bib(f.read())))

    def _add_bib_list(self) -> None:
        # References from a .bib file, numbered in the order they're cited
        if not self.bib_entries:
            return
        if self._quiet:
            for k in self._cite_order:
                if k in self.bib_entries:
                    self.citations.setdefault(k, len(self.citations) + 1)
        cited = sorted((n, k) for k, n in self.citations.items() if k in self.bib_entries)
        if not cited:
            return
        self.word_doc.add_heading('References', level=1)
        for n, key in cited:
            p = self.word_doc.add_paragraph(style='Bibliography' if 'Bibliography' in self.word_doc.styles else None)
            p.add_run(f'[{n}] ')
            self._apply_inline(format_bib_entry(self.bib_entries[key]), p)
            self._trim_runs(p)

    def _add_toc(self) -> None:
        # A real Word table of contents field. Word fills it in when the file
        # is opened (we ask it to update fields) or on right click > Update Field.
        self.word_doc.add_heading('Contents', level=1).style = self.word_doc.styles['TOC Heading'] \
            if 'TOC Heading' in self.word_doc.styles else None
        p = self.word_doc.add_paragraph()
        run = p.add_run()
        for kind, text in (('begin', None), ('instr', 'TOC \\o "1-3" \\h \\z \\u'),
                           ('separate', None), ('text', 'Right-click and choose "Update Field" '
                                                       'to build the table of contents.'),
                           ('end', None)):
            if kind == 'instr':
                el = OxmlElement('w:instrText')
                el.set(qn('xml:space'), 'preserve')
                el.text = text
                run._r.append(el)
            elif kind == 'text':
                run = p.add_run(text)
                run.italic = True
                run = p.add_run()
            else:
                el = OxmlElement('w:fldChar')
                el.set(qn('w:fldCharType'), kind)
                run._r.append(el)
        settings = self.word_doc.settings.element
        if settings.find(qn('w:updateFields')) is None:
            upd = OxmlElement('w:updateFields')
            upd.set(qn('w:val'), 'true')
            settings.append(upd)
        self.word_doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def _add_bibliography(self, inner: str) -> None:
        self.word_doc.add_heading('References', level=1)
        _, j = read_arg(inner, 0)  # the {99} width argument
        for idx, item in enumerate(split_top_level(inner[j:], '\\bibitem')[1:], start=1):
            _, k = read_optional(item, 0)
            _, k = read_arg(item, k)
            p = self.word_doc.add_paragraph(style='Bibliography' if 'Bibliography' in self.word_doc.styles else None)
            p.add_run(f'[{idx}] ')
            self._apply_inline(item[k:].strip(), p)
            self._trim_runs(p)

    # ------------------------------------------------------------------
    # Inline level
    # ------------------------------------------------------------------

    def _apply_inline(self, text: str, para_obj) -> None:
        # Kept for backwards compatibility: render LaTeX text into a paragraph
        self._inline(text, para_obj, {})

    def _inline(self, s: str, para, fmt: dict) -> None:
        # Renders inline LaTeX into runs on `para`. `fmt` is the current
        # formatting (bold/italic/...). Groups get their own copy so
        # {\bf x} only affects x.
        fmt = dict(fmt)
        buf: List[str] = []
        math = fmt.get('math', False)

        def flush():
            if buf:
                self._emit(para, ''.join(buf), fmt)
                buf.clear()

        i = 0
        n = len(s)
        while i < n:
            c = s[i]

            if c == '\\':
                m = re.match(r'\\([A-Za-z]+)\*?|\\(.)', s[i:], re.DOTALL)
                if m is None:
                    i += 1
                    continue
                name, sym = m.group(1), m.group(2)
                j = i + m.end()

                if sym is not None:
                    if sym == '\\':
                        # \\ line break (with optional [spacing])
                        flush()
                        para.add_run().add_break()
                        _, j = read_optional(s, j)
                    elif sym in '&%$#_{}':
                        buf.append(sym)
                    elif sym in ',;:! ':
                        buf.append(' ' if sym in '; ' else ('\u2009' if sym in ',:' else ''))
                    elif sym in '\'`^"~':
                        # accents like \'e - just keep the letter
                        arg, j = read_arg(s, j)
                        buf.append(self._accent(sym, arg or ''))
                    else:
                        buf.append(sym)
                    i = j
                    continue

                # Declarations and unknown commands eat the spaces after them
                j_space = skip_spaces(s, j)

                if name in _FORMAT_COMMANDS:
                    arg, j2 = read_arg(s, j)
                    flush()
                    self._inline(arg or '', para, {**fmt, **_FORMAT_COMMANDS[name]})
                    i = j2
                elif name in _DECLARATIONS:
                    flush()
                    fmt.update(_DECLARATIONS[name])
                    i = j_space
                elif name in ('href',):
                    url, j2 = read_arg(s, j)
                    label, j2 = read_arg(s, j2)
                    flush()
                    self._add_hyperlink(para, (url or '').replace('\\', ''), label or url or '', fmt)
                    i = j2
                elif name == 'url':
                    url, j2 = read_arg(s, j)
                    flush()
                    self._add_hyperlink(para, url or '', None, fmt, raw_text=url or '')
                    i = j2
                elif name == 'footnote':
                    _, j2 = read_optional(s, j)
                    arg, j2 = read_arg(s, j2)
                    flush()
                    self._inline(' (' + (arg or '').strip() + ')', para, fmt)
                    i = j2
                elif name in _CITE_COMMANDS:
                    note1, j2 = read_optional(s, j)
                    note2, j2 = read_optional(s, j2)
                    keys, j2 = read_arg(s, j2)
                    buf.append(self._cite_text(name, keys or '', note1, note2))
                    i = j2
                elif name in ('ref', 'eqref', 'autoref', 'cref', 'Cref', 'pageref', 'nameref'):
                    key, j2 = read_arg(s, j)
                    buf.append(self._ref_text(name, (key or '').strip()))
                    i = j2
                elif name == 'label':
                    key, j2 = read_arg(s, j)
                    self._set_label((key or '').strip(), self._ref_target)
                    i = j2
                elif name in ('SI', 'qty', 'si', 'unit', 'num', 'ang', 'SIrange', 'qtyrange'):
                    _, j2 = read_optional(s, j)
                    a1, j2 = read_arg(s, j2)
                    text = ''
                    if name in ('SI', 'qty'):
                        a2, j2 = read_arg(s, j2)
                        text = (a1 or '').strip() + '\u00a0' + si_unit_text(a2)
                    elif name in ('SIrange', 'qtyrange'):
                        a2, j2 = read_arg(s, j2)
                        a3, j2 = read_arg(s, j2)
                        unit = si_unit_text(a3)
                        text = f"{(a1 or '').strip()}\u00a0{unit} to {(a2 or '').strip()}\u00a0{unit}"
                    elif name in ('si', 'unit'):
                        text = si_unit_text(a1)
                    elif name == 'num':
                        text = (a1 or '').strip().replace('e', '×10^').replace('E', '×10^')
                        if '^' in text:
                            base, exp = text.split('^', 1)
                            text = base + exp.translate(_SUPERSCRIPTS)
                    else:
                        text = (a1 or '').strip() + '°'
                    buf.append(text)
                    i = j2
                elif name == 'includegraphics':
                    opts, j2 = read_optional(s, j)
                    path, j2 = read_arg(s, j2)
                    flush()
                    self._inline_image(para, (path or '').strip(), opts)
                    i = j2
                elif name in _DROP_WITH_ARG:
                    _, j2 = read_optional(s, j)
                    _, j2 = read_arg(s, j2)
                    i = j2
                elif name == 'caption':
                    # already handled by the float code
                    _, j2 = read_optional(s, j)
                    _, j2 = read_arg(s, j2)
                    i = j2
                elif name in ('textcolor',):
                    _, j2 = read_arg(s, j)
                    arg, j2 = read_arg(s, j2)
                    flush()
                    self._inline(arg or '', para, fmt)
                    i = j2
                elif math and name == 'frac':
                    num, j2 = read_arg(s, j)
                    den, j2 = read_arg(s, j2)
                    flush()
                    self._inline(self._wrap(num) + '/' + self._wrap(den), para, fmt)
                    i = j2
                elif math and name == 'sqrt':
                    _, j2 = read_optional(s, j)
                    arg, j2 = read_arg(s, j2)
                    buf.append('√')
                    flush()
                    self._inline(self._wrap(arg), para, fmt)
                    i = j2
                elif math and name in _MATH_SYMBOLS:
                    buf.append(_MATH_SYMBOLS[name])
                    i = j
                elif name in _TEXT_COMMANDS:
                    buf.append(_TEXT_COMMANDS[name])
                    # \LaTeX{} - swallow the empty group
                    i = j + 2 if s.startswith('{}', j) else j
                elif name in _MATH_SYMBOLS:
                    buf.append(_MATH_SYMBOLS[name])
                    i = j
                elif name in ('begin', 'end'):
                    # stray environment markers inside inline text
                    _, j2 = read_arg(s, j)
                    i = j2
                else:
                    # Unknown command: render its argument if it has one
                    _, j2 = read_optional(s, j)
                    if j2 < n and s[j2] == '{':
                        arg, j2 = read_group(s, j2)
                        flush()
                        self._inline(arg, para, fmt)
                        i = j2
                    else:
                        i = j_space
                continue

            if c == '{':
                arg, j = read_group(s, i)
                flush()
                self._inline(arg, para, fmt)
                i = j
                continue
            if c == '}':
                i += 1
                continue

            if c == '$' and not math:
                end = s.find('$', i + 1)
                end = n if end == -1 else end
                flush()
                self._inline(s[i + 1:end], para, {**fmt, 'italic': True, 'math': True})
                i = end + 1
                continue

            if math and c in '^_':
                arg, j = read_arg(s, i + 1)
                flush()
                key = 'superscript' if c == '^' else 'subscript'
                if arg is not None and arg.startswith('\\') and arg[1:] in _MATH_SYMBOLS:
                    arg = _MATH_SYMBOLS[arg[1:]]
                self._inline(arg or '', para, {**fmt, key: True})
                i = j
                continue

            if c == '~':
                buf.append('\u00a0')
            elif c == '-' and not math and s.startswith('---', i):
                buf.append('—')
                i += 3
                continue
            elif c == '-' and not math and s.startswith('--', i):
                buf.append('–')
                i += 2
                continue
            elif s.startswith('``', i):
                buf.append('“')
                i += 2
                continue
            elif s.startswith("''", i):
                buf.append('”')
                i += 2
                continue
            elif c == '`':
                buf.append('‘')
            elif c == "'" and not math:
                buf.append('’')
            elif c in ' \t\n':
                if math:
                    pass  # spaces don't matter in math
                elif not buf or buf[-1] != ' ':
                    buf.append(' ')
            elif c == '&' and math:
                pass
            else:
                buf.append(c)
            i += 1

        flush()

    @staticmethod
    def _wrap(arg: Optional[str]) -> str:
        # Adds brackets around multi-character fraction parts: (a+b)/c
        arg = (arg or '').strip()
        if len(arg) <= 1 or re.fullmatch(r'\\[A-Za-z]+|\d+', arg):
            return arg
        return '(' + arg + ')'

    @staticmethod
    def _accent(sym: str, letter: str) -> str:
        import unicodedata
        marks = {"'": '\u0301', '`': '\u0300', '^': '\u0302', '"': '\u0308', '~': '\u0303'}
        letter = letter.strip('{}')
        if len(letter) != 1:
            return letter
        return unicodedata.normalize('NFC', letter + marks[sym])

    @staticmethod
    def _style_run(run, fmt: dict) -> None:
        if fmt.get('bold'):
            run.bold = True
        if fmt.get('italic'):
            run.italic = True
        if fmt.get('underline'):
            run.underline = True
        if fmt.get('smallcaps'):
            run.font.small_caps = True
        if fmt.get('superscript'):
            run.font.superscript = True
        elif fmt.get('subscript'):
            run.font.subscript = True
        if fmt.get('mono'):
            run.font.name = 'Courier New'

    def _emit(self, para, text: str, fmt: dict) -> None:
        if not text:
            return
        run = para.add_run(text)
        self._style_run(run, fmt)

    def _add_hyperlink(self, para, url: str, label: Optional[str], fmt: dict,
                       raw_text: Optional[str] = None) -> None:
        # python-docx has no public API for creating links, so build the XML
        r_id = para.part.relate_to(url, RT.HYPERLINK, is_external=True)
        link = OxmlElement('w:hyperlink')
        link.set(qn('r:id'), r_id)

        # Render the label into a scratch paragraph, then move its runs into the link
        scratch = self.word_doc.add_paragraph()
        if raw_text is not None:
            scratch.add_run(raw_text)
        else:
            self._inline(label or url, scratch, fmt)
        for run in scratch.runs:
            run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
            run.underline = True
            link.append(run._r)
        scratch._p.getparent().remove(scratch._p)
        para._p.append(link)

    def _inline_image(self, para, path: str, opts: Optional[str]) -> None:
        resolved = self._resolve_image(path)
        if resolved is None:
            self._warn(f"Image not found: {path}")
            para.add_run(f"[Image file not found: {path}]")
            return
        try:
            para.add_run().add_picture(resolved, width=self._image_width(opts))
        except Exception as e:
            logger.warning(f"Could not insert image {resolved}: {e}")
            para.add_run(f"[Image found but error loading: {path}]")

    @staticmethod
    def _trim_runs(para) -> None:
        # Strips leading/trailing whitespace across the paragraph's runs
        # (a link at either end counts as text, so the space next to it stays)
        items = [el for el in para._p if el.tag in (qn('w:r'), qn('w:hyperlink'))
                 and (el.tag == qn('w:hyperlink') or ''.join(el.itertext()))]
        if not items:
            return
        if items[0].tag == qn('w:r'):
            first = Run(items[0], para)
            first.text = first.text.lstrip(' \u00a0')
        if items[-1].tag == qn('w:r'):
            last = Run(items[-1], para)
            last.text = last.text.rstrip(' \u00a0')
