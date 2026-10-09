# Contributing to doc2tex

Thanks for wanting to help! doc2tex started as a small tool for our own lab reports, so any
fix, test document or idea is welcome - you don't need to be a LaTeX expert.

## Reporting a bug

Open an issue with the **Bug report** template. The most useful thing you can attach is a small
`.docx` or `.tex` file that shows the problem (cut it down to just the broken part if the
original is private). Also tell us:

- the command you ran (or the settings you used in the web UI)
- what you expected vs. what you got
- `doc2tex --version` and your Python version

## Setting up

```bash
git clone https://github.com/IOST-ASCOL/doc2tex.git
cd doc2tex
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -e ".[dev]"
```

Run the tests:

```bash
python -m pytest --cov=doc2tex --cov-report=term-missing
```

## Making changes

1. Fork the repo and create a branch from `main` (`fix-table-captions`, `omml-matrices`, ...).
2. Make your change. Have a look at the [technical overview](understand.md) first if you're
   touching the converters.
3. **Add a test.** Every line of the package is covered by the tests and CI fails if coverage
   drops below 100%, so new code needs a test that runs it. Building a small document in the
   test with `python-docx` (or a short LaTeX string) is usually enough - see
   `tests/test_features.py` for lots of examples.
4. Add a line to `CHANGELOG.md` under an "Unreleased" heading.
5. Open a pull request and fill in the template.

### Code style

- Keep it readable for students: plain Python, short functions, comments that explain *why*.
- Match the style of the file you're editing (4 spaces, single quotes, `# comments` above the code).
- Python 3.10 is the oldest version we support, so don't use newer syntax.

## Where to start

Issues labelled `good first issue` are a good starting point. Things we'd love help with:

- LaTeX math → editable Word equations (right now it becomes readable text)
- more `siunitx` / `biblatex` commands
- real Word footnotes for LaTeX → DOCX
- test documents from real reports that don't convert well

## Code of Conduct

Please be kind. By taking part in this project you agree to follow our
[Code of Conduct](CODE_OF_CONDUCT.md).
