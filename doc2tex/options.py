# Configuration settings for the converter
# Just stores all the options like font size, margins, etc.

from dataclasses import dataclass, field
from typing import Optional, List
from enum import Enum


# Different types of LaTeX documents
class DocumentType(Enum):
    ARTICLE = "article"
    REPORT = "report"
    THESIS = "thesis"
    LETTER = "letter"
    BOOK = "book"


# Font sizes that LaTeX supports
class FontSize(Enum):
    PT_10 = "10pt"
    PT_11 = "11pt"
    PT_12 = "12pt"


# Line spacing options
class LineSpacing(Enum):
    SINGLE = "single"
    ONE_HALF = "onehalf"
    DOUBLE = "double"


@dataclass
class ConversionOptions:
    # Main class to store all conversion settings
    # Using dataclass because it's cleaner than writing __init__ manually
    
    # Basic document settings
    document_type: DocumentType = DocumentType.ARTICLE
    font_size: FontSize = FontSize.PT_12
    line_spacing: LineSpacing = LineSpacing.SINGLE
    page_margins: str = "top=2.5cm,bottom=2.5cm,left=2.5cm,right=2.5cm"
    
    # Bibliography stuff
    extract_bibliography: bool = False
    bibliography_style: str = "plain"
    
    # Language support
    unicode_support: bool = True
    language: str = "english"
    
    # Image handling
    preserve_images: bool = True
    image_max_width: str = "0.8\\textwidth"
    optimize_images: bool = False
    image_quality: int = 85
    
    # Extra packages and features
    custom_packages: List[str] = field(default_factory=list)
    preserve_comments: bool = False
    preserve_track_changes: bool = False
    
    # Technical settings
    output_encoding: str = "utf-8"
    clean_temp_files: bool = True
    verbose: bool = False
    
    # Whether to include full LaTeX document or just content
    include_preamble: bool = True
    standalone_document: bool = True
    
    # Convert to dictionary (useful for saving settings)
    def to_dict(self) -> dict:
        return {
            'document_type': self.document_type.value,
            'font_size': self.font_size.value,
            'line_spacing': self.line_spacing.value,
            'page_margins': self.page_margins,
            'extract_bibliography': self.extract_bibliography,
            'bibliography_style': self.bibliography_style,
            'unicode_support': self.unicode_support,
            'language': self.language,
            'preserve_images': self.preserve_images,
            'image_max_width': self.image_max_width,
            'optimize_images': self.optimize_images,
            'image_quality': self.image_quality,
            'custom_packages': self.custom_packages,
            'preserve_comments': self.preserve_comments,
            'preserve_track_changes': self.preserve_track_changes,
            'output_encoding': self.output_encoding,
            'clean_temp_files': self.clean_temp_files,
            'verbose': self.verbose,
            'include_preamble': self.include_preamble,
            'standalone_document': self.standalone_document,
        }
    
    # Create options from dictionary
    @classmethod
    def from_dict(cls, data: dict):
        # Need to convert string values back to enums
        if 'document_type' in data and isinstance(data['document_type'], str):
            data['document_type'] = DocumentType(data['document_type'])
        if 'font_size' in data and isinstance(data['font_size'], str):
            data['font_size'] = FontSize(data['font_size'])
        if 'line_spacing' in data and isinstance(data['line_spacing'], str):
            data['line_spacing'] = LineSpacing(data['line_spacing'])
        
        return cls(**data)
    
    # Get list of LaTeX packages needed based on options
    def get_latex_packages(self) -> List[str]:
        packages = [
            'graphicx',  # for images
            'hyperref',  # for links
            'amsmath',   # for equations
            'amssymb',   # math symbols
            'geometry',  # page margins
        ]
        
        # Add line spacing package if needed
        if self.line_spacing != LineSpacing.SINGLE:
            packages.append('setspace')
        
        # Unicode support packages
        if self.unicode_support:
            packages.extend(['fontenc', 'inputenc'])
        
        # Bibliography package
        if self.extract_bibliography:
            packages.append('natbib')
        
        # Table packages
        packages.extend(['booktabs', 'longtable', 'array'])
        
        # Add any custom packages user wants
        packages.extend(self.custom_packages)
        
        # Remove duplicates
        return list(set(packages))
    
    # Basic validation
    def validate(self) -> bool:
        if self.image_quality < 1 or self.image_quality > 100:
            raise ValueError("Image quality must be between 1 and 100")
        
        if not self.output_encoding:
            raise ValueError("Output encoding cannot be empty")
        
        return True
