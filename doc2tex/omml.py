# Word equations (OMML) -> LaTeX math
# Word saves equations as "Office Math Markup" XML: <m:f> is a fraction,
# <m:sSup> a superscript, <m:nary> a sum/integral and so on. Each element
# maps pretty directly to a LaTeX construct, so this is a small recursive walk.

from docx.oxml.ns import qn

from .symbols import MATH_COMMANDS, MATH_FUNCTIONS

M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

# Characters that mean something in LaTeX math and need escaping
_MATH_ESCAPES = {'{': r'\{', '}': r'\}', '%': r'\%', '#': r'\#', '$': r'\$',
                 '&': r'\&', '_': r'\_', '\\': r'\backslash ', '~': r'\sim '}

# Big operators (m:nary chr) -> LaTeX
_NARY = {'∑': r'\sum', '∏': r'\prod', '∐': r'\coprod', '∫': r'\int', '∬': r'\iint',
         '∭': r'\iiint', '∮': r'\oint', '⋃': r'\bigcup', '⋂': r'\bigcap',
         '⋁': r'\bigvee', '⋀': r'\bigwedge', '⨁': r'\bigoplus', '⨂': r'\bigotimes'}

# Accents (m:acc chr) -> LaTeX
_ACCENTS = {'̂': r'\hat', '^': r'\hat', '̃': r'\tilde', '~': r'\tilde',
            '̇': r'\dot', '̈': r'\ddot', '⃗': r'\vec', '→': r'\vec',
            '̄': r'\bar', '̅': r'\bar', '¯': r'\bar', '̌': r'\check',
            '́': r'\acute', '̀': r'\grave', '̆': r'\breve'}

# Delimiters (m:d begChr/endChr) -> what goes after \left / \right
_DELIMS = {'(': '(', ')': ')', '[': '[', ']': ']', '{': r'\{', '}': r'\}',
           '|': '|', '‖': r'\|', '⟨': r'\langle', '⟩': r'\rangle', '〈': r'\langle',
           '〉': r'\rangle', '⌊': r'\lfloor', '⌋': r'\rfloor', '⌈': r'\lceil',
           '⌉': r'\rceil', '': '.'}


def _m(tag: str) -> str:
    return f'{{{M_NS}}}{tag}'


def _prop(el, pr_tag: str, child: str, default=None):
    # Reads e.g. <m:dPr><m:begChr m:val="["/></m:dPr>
    pr = el.find(_m(pr_tag))
    if pr is None:
        return default
    c = pr.find(_m(child))
    if c is None:
        return default
    val = c.get(_m('val'))
    return default if val is None else val


def _child(el, tag: str):
    return el.find(_m(tag))


def _group(tex: str) -> str:
    # Wraps in braces unless it's a single character / command
    tex = tex.strip()
    if len(tex) == 1 or (tex.startswith('\\') and tex[1:].isalpha()):
        return tex
    return '{' + tex + '}'


def _text_to_latex(text: str, in_array: bool) -> str:
    out = []
    for ch in text:
        if ch == '&' and in_array:
            out.append('&')
        elif ch in _MATH_ESCAPES:
            out.append(_MATH_ESCAPES[ch])
        elif ch in MATH_COMMANDS:
            cmd = MATH_COMMANDS[ch]
            out.append(cmd + (' ' if cmd[-1].isalpha() else ''))
        else:
            out.append(ch)
    return ''.join(out)


class OmmlConverter:
    def __init__(self):
        self.in_array = False

    def convert(self, el) -> str:
        # Entry point: <m:oMath> or <m:oMathPara>
        tex = self._node(el)
        return ' '.join(tex.split())

    def _children(self, el) -> str:
        return ''.join(self._node(c) for c in el)

    def _node(self, el) -> str:
        tag = el.tag
        if not isinstance(tag, str) or tag.endswith('Pr'):
            return ''  # property elements (m:rPr, m:fPr, w:rPr ...) and comments
        if not tag.startswith('{' + M_NS):
            # A w:r or similar sneaking in: just take its text
            if tag != qn('w:r'):
                return ''
            return _text_to_latex(''.join(t.text or '' for t in el.iter(qn('w:t'))), self.in_array)
        name = tag.split('}')[1]
        handler = getattr(self, '_' + name, None)
        if handler is not None:
            return handler(el)
        return self._children(el)

    # --- leaf text --------------------------------------------------------

    def _r(self, el) -> str:
        text = ''.join(t.text or '' for t in el.iter(_m('t')))
        plain = _prop(el, 'rPr', 'sty') == 'p' or el.find(f"{_m('rPr')}/{_m('nor')}") is not None
        if plain and text.strip() in MATH_FUNCTIONS:
            return '\\' + text.strip() + ' '
        if plain and len(text.strip()) > 1 and text.strip().isalpha():
            return r'\mathrm{' + text.strip() + '}'
        return _text_to_latex(text, self.in_array)

    # --- structures -------------------------------------------------------

    def _f(self, el) -> str:
        num = self._node(_child(el, 'num')) if _child(el, 'num') is not None else ''
        den = self._node(_child(el, 'den')) if _child(el, 'den') is not None else ''
        kind = _prop(el, 'fPr', 'type', 'bar')
        if kind == 'lin':
            return f'{_group(num)}/{_group(den)}'
        if kind == 'noBar':
            return rf'\binom{{{num}}}{{{den}}}'
        return rf'\frac{{{num}}}{{{den}}}'

    def _part(self, el, tag: str) -> str:
        c = _child(el, tag)
        return self._node(c).strip() if c is not None else ''

    def _sSup(self, el) -> str:
        return f"{_group(self._part(el, 'e'))}^{{{self._part(el, 'sup')}}}"

    def _sSub(self, el) -> str:
        return f"{_group(self._part(el, 'e'))}_{{{self._part(el, 'sub')}}}"

    def _sSubSup(self, el) -> str:
        return (f"{_group(self._part(el, 'e'))}_{{{self._part(el, 'sub')}}}"
                f"^{{{self._part(el, 'sup')}}}")

    def _sPre(self, el) -> str:
        return (f"{{}}_{{{self._part(el, 'sub')}}}^{{{self._part(el, 'sup')}}}"
                f"{_group(self._part(el, 'e'))}")

    def _rad(self, el) -> str:
        deg = self._part(el, 'deg')
        e = self._part(el, 'e')
        if deg and _prop(el, 'radPr', 'degHide') not in ('1', 'on', 'true'):
            return rf'\sqrt[{deg}]{{{e}}}'
        return rf'\sqrt{{{e}}}'

    def _d(self, el) -> str:
        beg = _prop(el, 'dPr', 'begChr', '(')
        end = _prop(el, 'dPr', 'endChr', ')')
        sep = _prop(el, 'dPr', 'sepChr', '|')
        parts = [self._node(e).strip() for e in el.findall(_m('e'))]
        sep_tex = r' \mid ' if sep == '|' else ' ' + _text_to_latex(sep, False) + ' '
        inner = sep_tex.join(parts)
        return (r'\left' + _DELIMS.get(beg, _text_to_latex(beg, False)) + ' ' + inner +
                r' \right' + _DELIMS.get(end, _text_to_latex(end, False)))

    def _nary(self, el) -> str:
        chr_ = _prop(el, 'naryPr', 'chr', '∫')
        op = _NARY.get(chr_, _text_to_latex(chr_, False).strip())
        out = op
        if _prop(el, 'naryPr', 'subHide') not in ('1', 'on', 'true'):
            sub = self._part(el, 'sub')
            if sub:
                out += f'_{{{sub}}}'
        if _prop(el, 'naryPr', 'supHide') not in ('1', 'on', 'true'):
            sup = self._part(el, 'sup')
            if sup:
                out += f'^{{{sup}}}'
        return out + ' ' + self._part(el, 'e') + ' '

    def _func(self, el) -> str:
        name = self._part(el, 'fName')
        bare = name.strip()
        if bare in MATH_FUNCTIONS:  # typed as normal italic text
            name = '\\' + bare
        return name + ' ' + _group(self._part(el, 'e')) + ' '

    def _limLow(self, el) -> str:
        return f"{self._part(el, 'e')}_{{{self._part(el, 'lim')}}}"

    def _limUpp(self, el) -> str:
        return rf"\overset{{{self._part(el, 'lim')}}}{{{self._part(el, 'e')}}}"

    def _acc(self, el) -> str:
        chr_ = _prop(el, 'accPr', 'chr', '̂')
        cmd = _ACCENTS.get(chr_, r'\hat')
        return f"{cmd}{{{self._part(el, 'e')}}}"

    def _bar(self, el) -> str:
        cmd = r'\underline' if _prop(el, 'barPr', 'pos', 'bot') == 'bot' else r'\overline'
        return f"{cmd}{{{self._part(el, 'e')}}}"

    def _groupChr(self, el) -> str:
        chr_ = _prop(el, 'groupChrPr', 'chr', '⏟')
        e = self._part(el, 'e')
        if chr_ == '⏞':
            return rf'\overbrace{{{e}}}'
        if chr_ == '⏟':
            return rf'\underbrace{{{e}}}'
        return e

    def _borderBox(self, el) -> str:
        return rf"\boxed{{{self._part(el, 'e')}}}"

    def _eqArr(self, el) -> str:
        prev, self.in_array = self.in_array, True
        rows = [self._node(e).strip() for e in el.findall(_m('e'))]
        self.in_array = prev
        return r'\begin{aligned} ' + r' \\ '.join(rows) + r' \end{aligned}'

    def _m(self, el) -> str:
        rows = []
        for mr in el.findall(_m('mr')):
            rows.append(' & '.join(self._node(e).strip() for e in mr.findall(_m('e'))))
        return r'\begin{matrix} ' + r' \\ '.join(rows) + r' \end{matrix}'

    def _oMathPara(self, el) -> str:
        # Several equations stacked in one display block
        maths = [self._node(m).strip() for m in el.findall(_m('oMath'))]
        if len(maths) > 1:
            return r'\begin{gathered} ' + r' \\ '.join(maths) + r' \end{gathered}'
        return maths[0] if maths else ''


def omml_to_latex(el) -> str:
    # Converts an <m:oMath> / <m:oMathPara> element to a LaTeX math string
    # (without the surrounding $...$)
    return OmmlConverter().convert(el)
