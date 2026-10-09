# Unicode <-> LaTeX symbol tables, shared by both conversion directions.
# Word stores equations and symbols as real Unicode characters (α, ≤, ∑ ...),
# LaTeX wants commands (\alpha, \leq, \sum ...). This is the dictionary between them.

# Math symbols: unicode character -> LaTeX math command
MATH_COMMANDS = {
    'α': r'\alpha', 'β': r'\beta', 'γ': r'\gamma', 'δ': r'\delta', 'ε': r'\varepsilon',
    'ϵ': r'\epsilon', 'ζ': r'\zeta', 'η': r'\eta', 'θ': r'\theta', 'ϑ': r'\vartheta',
    'ι': r'\iota', 'κ': r'\kappa', 'λ': r'\lambda', 'μ': r'\mu', 'ν': r'\nu',
    'ξ': r'\xi', 'π': r'\pi', 'ρ': r'\rho', 'σ': r'\sigma', 'ς': r'\varsigma',
    'τ': r'\tau', 'υ': r'\upsilon', 'φ': r'\varphi', 'ϕ': r'\phi', 'χ': r'\chi',
    'ψ': r'\psi', 'ω': r'\omega',
    'Γ': r'\Gamma', 'Δ': r'\Delta', 'Θ': r'\Theta', 'Λ': r'\Lambda', 'Ξ': r'\Xi',
    'Π': r'\Pi', 'Σ': r'\Sigma', 'Υ': r'\Upsilon', 'Φ': r'\Phi', 'Ψ': r'\Psi',
    'Ω': r'\Omega',
    '×': r'\times', '·': r'\cdot', '⋅': r'\cdot', '÷': r'\div', '±': r'\pm', '∓': r'\mp',
    '≤': r'\leq', '≥': r'\geq', '≠': r'\neq', '≈': r'\approx', '≡': r'\equiv',
    '∼': r'\sim', '≃': r'\simeq', '≅': r'\cong', '∝': r'\propto', '≪': r'\ll', '≫': r'\gg',
    '∞': r'\infty', '∂': r'\partial', '∇': r'\nabla', '√': r'\surd',
    '∑': r'\sum', '∏': r'\prod', '∫': r'\int', '∬': r'\iint', '∭': r'\iiint', '∮': r'\oint',
    '→': r'\rightarrow', '←': r'\leftarrow', '↔': r'\leftrightarrow', '↦': r'\mapsto',
    '⇒': r'\Rightarrow', '⇐': r'\Leftarrow', '⇔': r'\Leftrightarrow',
    '↑': r'\uparrow', '↓': r'\downarrow',
    '∈': r'\in', '∉': r'\notin', '∋': r'\ni', '⊂': r'\subset', '⊃': r'\supset',
    '⊆': r'\subseteq', '⊇': r'\supseteq', '∪': r'\cup', '∩': r'\cap',
    '∀': r'\forall', '∃': r'\exists', '∅': r'\emptyset', '¬': r'\neg',
    '∧': r'\wedge', '∨': r'\vee', '⊕': r'\oplus', '⊗': r'\otimes',
    '∘': r'\circ', '°': r'^{\circ}', '∠': r'\angle', '⊥': r'\perp', '∥': r'\parallel',
    '…': r'\ldots', '⋯': r'\cdots', '⋮': r'\vdots', '⋱': r'\ddots',
    'ħ': r'\hbar', 'ℓ': r'\ell', 'ℜ': r'\Re', 'ℑ': r'\Im', 'ℵ': r'\aleph',
    'ℝ': r'\mathbb{R}', 'ℕ': r'\mathbb{N}', 'ℤ': r'\mathbb{Z}', 'ℚ': r'\mathbb{Q}', 'ℂ': r'\mathbb{C}',
    '⟨': r'\langle', '⟩': r'\rangle', '⌊': r'\lfloor', '⌋': r'\rfloor',
    '⌈': r'\lceil', '⌉': r'\rceil', '‖': r'\|',
    '′': "'", '″': "''", '−': '-', '∗': '*', '∣': '|',
}

# Characters plain pdflatex (utf8 + T1) can't typeset in normal text.
# They are wrapped in \ensuremath so they work in text and math alike.
TEXT_FALLBACKS = {
    ch: '\\ensuremath{' + cmd + '}'
    for ch, cmd in MATH_COMMANDS.items()
    if ch not in '·×÷±°…′″−∗'  # these already work in text mode (textcomp / T1)
}
TEXT_FALLBACKS.update({
    '°': r'\textdegree{}', '…': r'\ldots{}', '−': '-', '′': "'", '″': "''",
    ' ': '~', ' ': r'\,', '​': '',
})

# The other direction: LaTeX math command name -> readable unicode
# (used when turning LaTeX equations into Word text)
MATH_TO_UNICODE = {
    cmd[1:]: ch for ch, cmd in MATH_COMMANDS.items()
    if cmd.startswith('\\') and cmd[1:].isalpha()
}
MATH_TO_UNICODE.update({
    'epsilon': 'ε', 'varepsilon': 'ε', 'phi': 'φ', 'varphi': 'φ',
    'le': '≤', 'ge': '≥', 'ne': '≠', 'to': '→', 'gets': '←', 'land': '∧', 'lor': '∨',
    'sqrt': '√', 'degree': '°', 'implies': '⇒', 'iff': '⇔', 'mid': '|',
    'quad': ' ', 'qquad': '  ',
    'lbrace': '{', 'rbrace': '}', 'vert': '|', 'Vert': '‖',
})
for _fn in ('sin', 'cos', 'tan', 'sec', 'csc', 'cot', 'arcsin', 'arccos', 'arctan',
            'sinh', 'cosh', 'tanh', 'log', 'ln', 'exp', 'lim', 'max', 'min', 'sup',
            'inf', 'det', 'dim', 'ker', 'deg', 'gcd', 'arg', 'Pr', 'mod'):
    MATH_TO_UNICODE[_fn] = _fn

# Function names Word writes as plain text that LaTeX has commands for
MATH_FUNCTIONS = {
    'sin', 'cos', 'tan', 'sec', 'csc', 'cot', 'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh', 'coth', 'log', 'ln', 'lg', 'exp', 'lim', 'liminf',
    'limsup', 'max', 'min', 'sup', 'inf', 'det', 'dim', 'ker', 'deg', 'gcd',
    'arg', 'Pr',
}
