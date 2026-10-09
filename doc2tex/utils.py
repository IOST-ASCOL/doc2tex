# Utility functions for the converter
# Has stuff like LaTeX escaping, file handling, image processing, etc.

import os
import re
import logging
import tempfile
import shutil
from pathlib import Path
from typing import Optional, List, Tuple
from PIL import Image
import hashlib

from .errors import ImageProcessingError, UnicodeHandlingError
from .symbols import TEXT_FALLBACKS


# Setup logging
def setup_logger(name: str = "doctex", verbose: bool = False) -> logging.Logger:
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    
    if not logger.handlers:
        handler = logging.StreamHandler()
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    
    return logger


logger = setup_logger()


# LaTeX special characters that need escaping
LATEX_SPECIAL_CHARS = {
    '&': r'\&',
    '%': r'\%',
    '$': r'\$',
    '#': r'\#',
    '_': r'\_',
    '{': r'\{',
    '}': r'\}',
    '~': r'\textasciitilde{}',
    '^': r'\textasciicircum{}',
    '\\': r'\textbackslash{}',
}


def escape_latex(text: str) -> str:
    # Escape special characters for LaTeX
    # Otherwise LaTeX will throw errors.
    # Done in a single pass so the output of one replacement
    # (like the braces in \textbackslash{}) never gets escaped again.
    # Symbols pdflatex can't print in text (α, ≤, →, ...) become math commands.
    if not text:
        return ""
    return ''.join(_ESCAPES.get(ch, ch) for ch in text)


_ESCAPES = {**TEXT_FALLBACKS, **LATEX_SPECIAL_CHARS}


def unescape_latex(text: str) -> str:
    # Reverse the escaping process (single pass, longest commands first)
    if not text:
        return ""
    return _UNESCAPE_RE.sub(lambda m: _UNESCAPE_MAP[m.group(0)], text)


_UNESCAPE_MAP = {escaped: char for char, escaped in LATEX_SPECIAL_CHARS.items()}
_UNESCAPE_MAP.update({
    r'\textasciitilde': '~',
    r'\textasciicircum': '^',
    r'\textbackslash': '\\',
})
_UNESCAPE_RE = re.compile(
    '|'.join(re.escape(k) for k in sorted(_UNESCAPE_MAP, key=len, reverse=True))
)


def sanitize_filename(filename: str) -> str:
    # Remove invalid characters from filename
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    filename = filename.strip('. ')
    
    # Limit length
    if len(filename) > 255:
        name, ext = os.path.splitext(filename)
        filename = name[:255-len(ext)] + ext
    
    return filename


def ensure_directory(path: str) -> Path:
    # Create directory if it doesn't exist
    dir_path = Path(path)
    dir_path.mkdir(parents=True, exist_ok=True)
    return dir_path


def get_temp_dir() -> str:
    # Create temporary directory for working files
    temp_dir = tempfile.mkdtemp(prefix='doctex_')
    logger.debug(f"Created temp directory: {temp_dir}")
    return temp_dir


def cleanup_temp_dir(temp_dir: str) -> None:
    # Delete temporary directory
    try:
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
            logger.debug(f"Cleaned up temp directory: {temp_dir}")
    except Exception as e:
        logger.warning(f"Failed to clean up temp directory: {e}")


def get_file_hash(filepath: str) -> str:
    # Calculate MD5 hash of file
    hash_md5 = hashlib.md5()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()


# Image processing functions
def optimize_image(
    image_path: str,
    output_path: str,
    max_width: Optional[int] = None,
    quality: int = 85
) -> Tuple[str, int, int]:
    # Resize and compress image
    try:
        with Image.open(image_path) as img:
            # Convert to RGB if needed
            if img.mode == 'RGBA':
                background = Image.new('RGB', img.size, (255, 255, 255))
                background.paste(img, mask=img.split()[3])
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Resize if too large
            if max_width and img.width > max_width:
                ratio = max_width / img.width
                new_height = int(img.height * ratio)
                img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
            
            # Save compressed image
            img.save(output_path, 'JPEG', quality=quality, optimize=True)
            
            logger.debug(f"Optimized image: {image_path} -> {output_path}")
            return output_path, img.width, img.height
            
    except Exception as e:
        raise ImageProcessingError(f"Failed to process image {image_path}: {e}")


def get_image_dimensions(image_path: str) -> Tuple[int, int]:
    # Get image width and height
    try:
        with Image.open(image_path) as img:
            return img.size
    except Exception as e:
        logger.warning(f"Failed to get image dimensions: {e}")
        return (0, 0)


def normalize_whitespace(text: str) -> str:
    # Clean up extra whitespace
    text = re.sub(r' +', ' ', text)
    text = re.sub(r'\n\n+', '\n\n', text)
    return text.strip()


def handle_unicode(text: str, encoding: str = 'utf-8') -> str:
    # Handle Unicode text safely
    try:
        if isinstance(text, bytes):
            text = text.decode(encoding, errors='replace')
        return text
    except Exception as e:
        raise UnicodeHandlingError(f"Failed to handle Unicode: {e}")


def extract_extension(filename: str) -> str:
    # Get file extension
    return os.path.splitext(filename)[1].lower().lstrip('.')


def is_valid_file(filepath: str, extensions: List[str]) -> bool:
    # Check if file exists and has correct extension
    if not os.path.isfile(filepath):
        return False
    
    ext = extract_extension(filepath)
    return ext in [e.lower() for e in extensions]


def format_file_size(size_bytes: int) -> str:
    # Convert bytes to human readable format
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} TB"


def get_file_info(filepath: str) -> dict:
    # Get file metadata
    stat = os.stat(filepath)
    return {
        'path': filepath,
        'name': os.path.basename(filepath),
        'size': stat.st_size,
        'size_formatted': format_file_size(stat.st_size),
        'extension': extract_extension(filepath),
        'modified': stat.st_mtime,
    }
