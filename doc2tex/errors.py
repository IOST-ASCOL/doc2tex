# Custom error classes for the converter
# Makes it easier to handle different types of errors

class DocTeXError(Exception):
    # Base error class
    pass


class ConversionError(DocTeXError):
    # When conversion fails
    pass


class FileNotFoundError(DocTeXError):
    # When input file doesn't exist
    pass


class InvalidFileFormatError(DocTeXError):
    # When file format is wrong
    pass


class InvalidOptionsError(DocTeXError):
    # When options are invalid
    pass


class ImageProcessingError(DocTeXError):
    # When image stuff fails
    pass


class LatexCompilationError(DocTeXError):
    # When LaTeX won't compile
    pass


class UnicodeHandlingError(DocTeXError):
    # When Unicode/encoding breaks
    pass
