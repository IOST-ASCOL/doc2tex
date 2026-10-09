# doc2tex - The main script that glues it all together
# You can use the CLI or the Web UI, but they both use this class eventually.

import os
import shutil
import subprocess
from pathlib import Path
from typing import Optional, List

from .options import ConversionOptions
from .latex import LatexGenerator
from .docx import DocxGenerator
from .utils import logger, is_valid_file, get_file_info
from .errors import ConversionError, InvalidFileFormatError, LatexCompilationError


class DocTeXConverter:
    """
    This is my main converter class.
    Basically, you give it a file, and it figures out if you want to go to LaTeX 
    or to Word based on the extension.
    """
    
    # I kept these here just to remember what we support
    ALLOWED_IN = ['docx', 'tex', 'latex']
    ALLOWED_OUT = ['docx', 'tex']
    DIRECTIONS = ('to_latex', 'to_docx')
    
    def __init__(self, settings: Optional[ConversionOptions] = None):
        # If the user didn't pass any settings, we just use defaults
        self.settings = settings or ConversionOptions()
        
        # Make sure settings are okay before we start
        self.settings.validate()
        self.last_generator = None  # the LatexGenerator/DocxGenerator from the last run
        self.warnings: List[str] = []  # things that didn't convert cleanly last time
        
        if self.settings.verbose:
            logger.setLevel('DEBUG')
            logger.debug("Okay, verbose mode is ON. Let's see what happens.")
    
    def convert(
        self,
        input_file: str,
        output_file: Optional[str] = None,
        direction: Optional[str] = None,
    ) -> str:
        # This is the function that does everything
        if not os.path.isfile(input_file):
            raise ConversionError(f"I can't find the file: {input_file}")

        # Just printing some info for the logs
        info = get_file_info(input_file)
        logger.info(f"Working on: {info['name']}")

        # 1. Figure out direction (to_latex or to_docx)
        direction = direction or self._detect_direction(input_file)
        if direction not in self.DIRECTIONS:
            raise ConversionError(f"Unknown direction: {direction}. Use one of {self.DIRECTIONS}")

        # 2. Pick a name for the output if we don't have one
        if output_file is None:
            output_file = self._calc_output_path(input_file, direction)

        # Make the folder if it's missing (I forgot this once and it crashed)
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        # 3. Call the generator
        self.warnings = []
        if direction == 'to_latex':
            result = self._run_latex_gen(input_file, output_file)
        else:
            result = self._run_docx_gen(input_file, output_file)
        self.warnings = list(getattr(self.last_generator, 'warnings', []))
        return result

    def _detect_direction(self, path: str) -> str:
        # Check extension and guess
        ext = Path(path).suffix.lower().lstrip('.')
        if ext == 'docx':
            return 'to_latex'
        elif ext in ['tex', 'latex']:
            return 'to_docx'
        else:
            raise InvalidFileFormatError(f"I don't know what to do with .{ext} files. Sorry!")

    def _calc_output_path(self, path: str, direction: str) -> str:
        # Swaps .docx for .tex or vice versa
        new_ext = '.tex' if direction == 'to_latex' else '.docx'
        return str(Path(path).with_suffix(new_ext))

    def _run_latex_gen(self, inp: str, out: str) -> str:
        # DOCX -> LaTeX
        # Need to make sure it's actually a docx file first
        if not is_valid_file(inp, ['docx']):
            raise InvalidFileFormatError("I need a .docx file to make LaTeX.")

        self.last_generator = LatexGenerator(self.settings)
        return self.last_generator.convert(inp, out)

    def _run_docx_gen(self, inp: str, out: str) -> str:
        # LaTeX -> DOCX
        if not is_valid_file(inp, ['tex', 'latex']):
            raise InvalidFileFormatError("I need a .tex file to make a Word doc.")

        self.last_generator = DocxGenerator(self.settings)
        return self.last_generator.convert(inp, out)

    @staticmethod
    def find_inputs(paths: List[str]) -> List[str]:
        # Expands folders into the .docx/.tex files inside them (not recursive).
        # Word's "~$file.docx" lock files are skipped.
        found: List[str] = []
        for p in paths:
            if os.path.isdir(p):
                for name in sorted(os.listdir(p)):
                    full = os.path.join(p, name)
                    if (os.path.isfile(full) and not name.startswith('~$')
                            and Path(name).suffix.lower() in ('.docx', '.tex', '.latex')):
                        found.append(full)
            else:
                found.append(p)
        return found

    @staticmethod
    def compile_pdf(tex_file: str, timeout: int = 300) -> str:
        # Runs latexmk (or pdflatex twice, for references) next to the .tex file
        tex_file = os.path.abspath(tex_file)
        workdir = os.path.dirname(tex_file)
        name = os.path.basename(tex_file)
        if shutil.which('latexmk'):
            commands = [['latexmk', '-pdf', '-interaction=nonstopmode', '-halt-on-error', name]]
        elif shutil.which('pdflatex'):
            cmd = ['pdflatex', '-interaction=nonstopmode', '-halt-on-error', name]
            commands = [cmd, cmd]
        else:
            raise LatexCompilationError(
                "No LaTeX installation found (need latexmk or pdflatex). "
                "Install TeX Live, MiKTeX or MacTeX, or upload the .tex to Overleaf.")
        for cmd in commands:
            try:
                proc = subprocess.run(cmd, cwd=workdir, capture_output=True, text=True,
                                      timeout=timeout, errors='replace')
            except subprocess.TimeoutExpired:
                raise LatexCompilationError(f"LaTeX took longer than {timeout}s, gave up")
            if proc.returncode != 0:
                errors = [ln for ln in proc.stdout.splitlines() if ln.startswith('!')]
                detail = errors[0] if errors else proc.stdout[-500:]
                raise LatexCompilationError(f"LaTeX failed on {name}: {detail}")
        pdf = str(Path(tex_file).with_suffix('.pdf'))
        logger.info(f"PDF written to {pdf}")
        return pdf

    def output_path_in(self, input_file: str, output_dir: str,
                       direction: Optional[str] = None) -> str:
        # Where a converted file goes inside output_dir
        d = direction or self._detect_direction(input_file)
        return os.path.join(output_dir, Path(self._calc_output_path(input_file, d)).name)

    def batch_convert(
        self,
        files: List[str],
        output_dir: Optional[str] = None,
        direction: Optional[str] = None,
    ) -> List[Optional[str]]:
        # This is useful if you have a whole folder of reports to convert.
        # Failed files show up as None in the results instead of stopping the batch.
        results: List[Optional[str]] = []
        for f in files:
            try:
                target = None
                if output_dir:
                    target = self.output_path_in(f, output_dir, direction)
                results.append(self.convert(f, target, direction=direction))
            except Exception as e:
                logger.warning(f"Skipping {f} because: {e}")
                results.append(None)
        return results

    # Old name, kept so existing scripts don't break
    batch = batch_convert
