# doc2tex - The main script that glues it all together
# You can use the CLI or the Web UI, but they both use this class eventually.

import os
from pathlib import Path
from typing import Optional, List

from .options import ConversionOptions
from .latex import LatexGenerator
from .docx import DocxGenerator
from .utils import logger, is_valid_file, cleanup_temp_dir, get_file_info
from .errors import ConversionError, InvalidFileFormatError


class DocTeXConverter:
    """
    This is my main converter class.
    Basically, you give it a file, and it figures out if you want to go to LaTeX 
    or to Word based on the extension.
    """
    
    # I kept these here just to remember what we support
    ALLOWED_IN = ['docx', 'tex', 'latex']
    ALLOWED_OUT = ['docx', 'tex']
    
    def __init__(self, settings: Optional[ConversionOptions] = None):
        # If the user didn't pass any settings, we just use defaults
        self.settings = settings or ConversionOptions()
        
        # Make sure settings are okay before we start
        self.settings.validate()
        
        if self.settings.verbose:
            logger.setLevel('DEBUG')
            logger.debug("Okay, verbose mode is ON. Let's see what happens.")
    
    def convert(
        self, 
        input_file: str, 
        output_file: Optional[str] = None,
        forced_direction: Optional[str] = None
    ) -> str:
        # This is the function that does everything
        if not os.path.isfile(input_file):
            raise ConversionError(f"I can't find the file: {input_file}")
        
        # Just printing some info for the logs
        info = get_file_info(input_file)
        logger.info(f"Working on: {info['name']}")
        
        # 1. Figure out direction (docx2latex or latex2docx)
        if forced_direction:
            direction = forced_direction
        else:
            direction = self._guess_direction(input_file)
            
        # 2. Pick a name for the output if we don't have one
        if output_file is None:
            output_file = self._calc_output_path(input_file, direction)
        
        # Make the folder if it's missing (I forgot this once and it crashed)
        out_dir = os.path.dirname(output_file)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
            
        # 3. Call the generator
        try:
            if direction == 'to_latex':
                return self._run_latex_gen(input_file, output_file)
            elif direction == 'to_docx':
                return self._run_docx_gen(input_file, output_file)
            else:
                raise ConversionError(f"Weird direction: {direction}. How did that happen?")
                
        except Exception as e:
            logger.error(f"Failed to convert {input_file}: {e}")
            raise
    
    def _guess_direction(self, path: str) -> str:
        # Checkextension and guess
        ext = Path(path).suffix.lower().lstrip('.')
        if ext == 'docx':
            return 'to_latex'
        elif ext in ['tex', 'latex']:
            return 'to_docx'
        else:
            raise InvalidFileFormatError(f"I don't know what to do with .{ext} files. Sorry!")
            
    def _calc_output_path(self, path: str, dir: str) -> str:
        # Swaps .docx for .tex or vice versa
        p = Path(path)
        new_ext = '.tex' if dir == 'to_latex' else '.docx'
        return str(p.with_suffix(new_ext))
        
    def _run_latex_gen(self, inp: str, out: str) -> str:
        # DOCX -> LaTeX
        # Need to make sure it's actually a docx file first
        if not is_valid_file(inp, ['docx']):
            raise InvalidFileFormatError("I need a .docx file to make LaTeX.")
            
        gen = LatexGenerator(self.settings)
        res = gen.convert(inp, out)
        
        # Clean up temp images if needed
        if self.settings.clean_temp_files and gen.temp_workspace:
            cleanup_temp_dir(gen.temp_workspace)
            
        return res
        
    def _run_docx_gen(self, inp: str, out: str) -> str:
        # LaTeX -> DOCX
        if not is_valid_file(inp, ['tex', 'latex']):
            raise InvalidFileFormatError("I need a .tex file to make a Word doc.")
            
        gen = DocxGenerator(self.settings)
        return gen.convert(inp, out)

    def batch(self, files: list, out_dir: Optional[str] = None) -> list:
        # This is useful if you have a whole folder of reports to convert
        results = []
        for f in files:
            try:
                # If they gave us a target folder, put it there
                if out_dir:
                    os.makedirs(out_dir, exist_ok=True)
                    name = Path(f).stem
                    # We have to guess the direction again for the extension
                    d = self._guess_direction(f)
                    ext = '.tex' if d == 'to_latex' else '.docx'
                    target = os.path.join(out_dir, name + ext)
                else:
                    target = None
                    
                results.append(self.convert(f, target))
            except Exception as e:
                logger.warning(f"Skipping {f} because: {e}")
                results.append(None)
        return results
