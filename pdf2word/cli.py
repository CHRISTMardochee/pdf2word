"""
CLI entry point for pdf2word.
"""

import argparse
import logging
import sys

from .converter import PDFToWordConverter
from .docx_to_pdf import docx_to_pdf
from .config import save_api_key, remove_api_key


def main():
    parser = argparse.ArgumentParser(
        prog="pdf2word",
        description="Convert PDF to Word (.docx) and back, preserving formatting.",
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # --- convert command ---
    convert_parser = subparsers.add_parser(
        "convert", help="Convert a PDF to Word (.docx)"
    )
    convert_parser.add_argument("input", help="Input PDF file path")
    convert_parser.add_argument(
        "-o", "--output", help="Output DOCX file path (default: input with .docx extension)"
    )
    convert_parser.add_argument(
        "--ocr-engine",
        choices=["tesseract", "paddleocr"],
        default="tesseract",
        help="OCR engine for scanned PDFs (default: tesseract)",
    )
    convert_parser.add_argument(
        "--ocr-lang",
        default="fra+eng",
        help="OCR language(s) (default: fra+eng)",
    )
    convert_parser.add_argument(
        "--force-ocr",
        action="store_true",
        help="Force OCR even for text-based PDFs",
    )
    convert_parser.add_argument(
        "--no-enhance",
        action="store_true",
        help="Skip DOCX post-processing (paragraph merging, style fixes)",
    )
    convert_parser.add_argument(
        "--mode",
        choices=["auto", "text", "ocr", "hybrid", "combined", "smart", "libreoffice", "msword", "cloud", "docling"],
        default="auto",
        help="Conversion mode: auto, text, ocr, hybrid, combined, smart, msword (Windows), libreoffice (Linux), or cloud (API)",
    )
    convert_parser.add_argument(
        "--api-key",
        help="API key for cloud mode (ConvertAPI secret).",
    )
    convert_parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="Image resolution for hybrid mode (default: 300)",
    )
    convert_parser.add_argument(
        "--pages",
        help="Page numbers to convert (comma-separated, 0-indexed). E.g., 0,1,2",
    )
    convert_parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )

    # --- reconvert command ---
    reconvert_parser = subparsers.add_parser(
        "reconvert", help="Convert a Word (.docx) back to PDF"
    )
    reconvert_parser.add_argument("input", help="Input DOCX file path")
    reconvert_parser.add_argument(
        "-o", "--output-dir",
        help="Output directory for the PDF (default: same as input)",
    )
    reconvert_parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging",
    )

    # --- set-key command ---
    setkey_parser = subparsers.add_parser(
        "set-key", help="Save your ConvertAPI secret key globally for cloud conversions"
    )
    setkey_parser.add_argument("api_key", help="Your ConvertAPI secret key")

    # --- remove-key command ---
    removekey_parser = subparsers.add_parser(
        "remove-key", help="Remove the globally saved ConvertAPI secret key"
    )

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    # Set up logging
    is_verbose = getattr(args, "verbose", False)
    log_level = logging.DEBUG if is_verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    if args.command == "convert":
        _run_convert(args)
    elif args.command == "reconvert":
        _run_reconvert(args)
    elif args.command == "set-key":
        _run_set_key(args)
    elif args.command == "remove-key":
        _run_remove_key(args)


def _run_convert(args):
    """Execute the convert command."""
    input_pdf = args.input
    output_docx = args.output

    if output_docx is None:
        # Default: same name with .docx extension
        import os
        base = os.path.splitext(input_pdf)[0]
        output_docx = base + ".docx"

    pages = None
    if args.pages:
        pages = [int(p.strip()) for p in args.pages.split(",")]

    converter = PDFToWordConverter(
        ocr_engine=args.ocr_engine,
        ocr_lang=args.ocr_lang,
        enhance=not args.no_enhance,
        mode=args.mode,
        dpi=args.dpi,
        api_key=getattr(args, 'api_key', None),
    )

    result = converter.convert(
        input_pdf,
        output_docx,
        pages=pages,
        force_ocr=args.force_ocr,
    )

    print(f"\n[OK] Conversion complete!")
    print(f"   Output: {result['output_path']}")
    print(f"   Method: {result['method']}")
    print(f"   Pages:  {result['analysis']['page_count']}")
    print(f"   Enhanced: {result['enhanced']}")


def _run_reconvert(args):
    """Execute the reconvert command."""
    input_docx = args.input
    output_dir = args.output_dir

    pdf_path = docx_to_pdf(input_docx, output_dir)

    print(f"\n[OK] Reconversion complete!")
    print(f"   Output: {pdf_path}")


def _run_set_key(args):
    """Execute the set-key command."""
    try:
        save_api_key(args.api_key)
        print("\n[OK] API Key successfully saved!")
        print("You can now use '--mode cloud' without passing the key each time.")
    except Exception as e:
        print(f"\n[ERROR] Failed to save API keys: {e}")


def _run_remove_key(args):
    """Execute the remove-key command."""
    try:
        remove_api_key()
        print("\n[OK] API Key successfully removed.")
    except Exception as e:
        print(f"\n[ERROR] Failed to remove API keys: {e}")


if __name__ == "__main__":
    main()
