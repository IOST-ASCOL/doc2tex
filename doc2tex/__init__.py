# doc2tex package initialization
# Exports the main converter class and options for easy use

from typing import Optional

from .converter import DocTeXConverter
from .options import (
    ConversionOptions,
    DocumentType,
    FontSize,
    LineSpacing
)
from .errors import (
    DocTeXError,
    ConversionError,
    InvalidFileFormatError,
    LatexCompilationError,
)

__version__ = "2.0.0"

__all__ = [
    'convert', 'DocTeXConverter', 'ConversionOptions', 'DocumentType', 'FontSize',
    'LineSpacing', 'DocTeXError', 'ConversionError', 'InvalidFileFormatError',
    'LatexCompilationError', '__version__',
]


def convert(input_file: str, output_file: Optional[str] = None, **options) -> str:
    """One-liner: doc2tex.convert("report.docx", document_type="report").

    Keyword arguments are ConversionOptions fields; enum fields also take
    their string value ("report", "11pt", "onehalf").
    """
    return DocTeXConverter(ConversionOptions.from_dict(options)).convert(input_file, output_file)
