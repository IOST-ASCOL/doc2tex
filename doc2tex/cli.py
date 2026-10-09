# Command line interface for doc2tex
# Used for batch processing or just quick converts from terminal

import argparse
import sys

from . import (
    __version__,
    DocTeXConverter,
    ConversionOptions,
    DocumentType,
    FontSize,
    LineSpacing,
    ConversionError,
    DocTeXError,
)
from .utils import logger, setup_logger


def create_parser() -> argparse.ArgumentParser:
    # Setting up all the flags for terminal usage
    parser = argparse.ArgumentParser(
        description='doc2tex - Simple DOCX ↔ LaTeX Converter',
        epilog="""
Example usage:
  doc2tex file.docx -o result.tex
  doc2tex file.tex
  doc2tex *.docx -d converted/
  doc2tex reports/ -d converted/          (every .docx/.tex in a folder)
  doc2tex thesis.docx --doc-type thesis --pdf
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    
    # Input files. You can pass multiple files here.
    parser.add_argument(
        'input',
        nargs='+',
        help='Files (or folders) to convert'
    )
    
    parser.add_argument(
        '-o', '--output',
        help='Where to save the result. If not set, it guesses name.'
    )
    
    parser.add_argument(
        '-d', '--output-dir',
        help='Folder to save results if batching'
    )
    
    parser.add_argument(
        '--direction',
        choices=['to_latex', 'to_docx'],
        help='Force it to go one way'
    )
    
    # Document settings
    doc_group = parser.add_argument_group('Doc Settings')
    doc_group.add_argument(
        '--doc-type',
        type=str,
        choices=['article', 'report', 'thesis', 'letter', 'book'],
        default='article'
    )
    doc_group.add_argument(
        '--font-size',
        type=str,
        choices=['10pt', '11pt', '12pt'],
        default='12pt'
    )
    doc_group.add_argument(
        '--spacing',
        type=str,
        choices=['single', 'onehalf', 'double'],
        default='single'
    )
    doc_group.add_argument(
        '--margins',
        type=str,
        default='top=2.5cm,bottom=2.5cm,left=2.5cm,right=2.5cm'
    )
    doc_group.add_argument('--language', default='english',
                           help='Document language for babel, e.g. german, french')
    doc_group.add_argument('--package', action='append', default=[], metavar='PKG',
                           help='Extra \\usepackage to add (can repeat)')

    # Word -> LaTeX
    tex_group = parser.add_argument_group('Word -> LaTeX')
    tex_group.add_argument('--extract-bib', action='store_true',
                           help='Turn the "References" section into a thebibliography list')
    tex_group.add_argument('--body-only', action='store_true',
                           help='Only write the content (no preamble), to \\input into your own template')
    tex_group.add_argument('--no-images', action='store_true', help='Dont handle images')
    tex_group.add_argument('--optimize-images', action='store_true', help='Shrink images (saves as JPEG)')
    tex_group.add_argument('--image-width', default='0.8\\textwidth',
                           help='Width for images whose size is unknown (default 0.8\\textwidth)')
    tex_group.add_argument('--pdf', action='store_true',
                           help='Also compile the .tex to PDF (needs latexmk or pdflatex)')

    # LaTeX -> Word
    docx_group = parser.add_argument_group('LaTeX -> Word')
    docx_group.add_argument('--no-numbering', action='store_true',
                            help="Don't number sections, figures and equations")
    
    # General stuff
    parser.add_argument('-v', '--verbose', action='store_true', help='Show more logs')
    parser.add_argument('--version', action='version', version=f'doc2tex {__version__}')
    
    return parser


def build_options(args: argparse.Namespace) -> ConversionOptions:
    # Map CLI args to our ConversionOptions class
    return ConversionOptions(
        document_type=DocumentType(args.doc_type),
        font_size=FontSize(args.font_size),
        line_spacing=LineSpacing(args.spacing),
        page_margins=args.margins,
        extract_bibliography=args.extract_bib,
        preserve_images=not args.no_images,
        optimize_images=args.optimize_images,
        image_max_width=args.image_width,
        language=args.language,
        custom_packages=args.package,
        include_preamble=not args.body_only,
        standalone_document=not args.body_only,
        number_sections=not args.no_numbering,
        verbose=args.verbose,
    )


def _finish(converter: DocTeXConverter, result: str, args: argparse.Namespace) -> None:
    # Shows warnings and compiles the PDF if asked to
    for w in converter.warnings:
        print(f"  warning: {w}")
    if args.pdf and result.endswith('.tex'):
        if args.body_only:
            raise ConversionError("--pdf needs a full document, it can't be used with --body-only")
        print(f"PDF: {converter.compile_pdf(result)}")


def main(argv=None) -> int:
    parser = create_parser()
    args = parser.parse_args(argv)
    
    # Setup the logger format
    setup_logger(verbose=args.verbose)

    inputs = DocTeXConverter.find_inputs(args.input)
    if not inputs:
        logger.error("No .docx or .tex files found.")
        return 1
    if args.output and (len(inputs) > 1 or args.output_dir):
        parser.error("-o only works with a single input file; use -d for several")

    try:
        options = build_options(args)
        converter = DocTeXConverter(options)
        
        # Batch conversion if more than one file
        if len(inputs) > 1 or args.output_dir:
            logger.info(f"Batch converting {len(inputs)} files...")
            failed = 0
            for f in inputs:
                try:
                    target = None
                    if args.output_dir:
                        target = converter.output_path_in(f, args.output_dir, args.direction)
                    result = converter.convert(f, target, direction=args.direction)
                    print(f"OK    {f} -> {result}")
                    _finish(converter, result, args)
                except DocTeXError as e:
                    failed += 1
                    print(f"FAIL  {f}: {e}")

            logger.info(f"Done! {len(inputs) - failed}/{len(inputs)} worked.")
            if failed:
                return 1
        
        # Single file conversion
        else:
            result = converter.convert(
                inputs[0],
                args.output,
                direction=args.direction
            )
            
            print(f"Success! Saved to: {result}")
            _finish(converter, result, args)

    except DocTeXError as e:
        logger.error(f"Error: {e}")
        return 1
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=args.verbose)
        return 1

    return 0


if __name__ == '__main__':
    sys.exit(main())
