"""
Main pipeline orchestrator.
Routes PDFs through the appropriate conversion path and applies post-processing.
"""

import logging
import os

from .analyzer import PDFAnalyzer
from .text_converter import TextConverter
from .ocr_converter import OCRConverter
from .hybrid_converter import HybridConverter
from .combined_converter import CombinedConverter
from .smart_converter import SmartConverter
from .docx_enhancer import DocxEnhancer
from .cloud_converter import CloudAPIConverter

try:
    from .docling_converter import DoclingConverter
    HAS_DOCLING = True
except ImportError:
    HAS_DOCLING = False

try:
    from .msword_converter import NativeWordConverter
    HAS_MSWORD = True
except (ImportError, NotImplementedError):
    HAS_MSWORD = False

try:
    from .libreoffice_converter import LibreOfficeConverter
    HAS_LIBREOFFICE = True
except (ImportError, FileNotFoundError):
    HAS_LIBREOFFICE = False

logger = logging.getLogger(__name__)


class PDFToWordConverter:
    """
    Main PDF-to-Word conversion pipeline.
    
    Workflow:
    1. Analyze the PDF (text-based or scanned)
    2. Route to the appropriate converter
    3. Apply DOCX post-processing (enhancer) — except in hybrid mode
    """

    def __init__(self, ocr_engine: str = "tesseract", ocr_lang: str = "fra+eng",
                 enhance: bool = True, mode: str = "auto", dpi: int = 300, api_key: str | None = None):
        """
        Args:
            ocr_engine: OCR engine to use for scanned PDFs ('tesseract' or 'paddleocr').
            ocr_lang: Language(s) for OCR.
            enhance: Whether to apply DOCX post-processing (text/ocr modes only).
            mode: Conversion mode — 'auto' (default), 'text', 'ocr', 'hybrid', 'cloud', 'docling'.
            dpi: Image resolution for hybrid mode (default 300).
            api_key: ConvertAPI secret key for cloud mode.
        """
        self.ocr_engine = ocr_engine
        self.ocr_lang = ocr_lang
        self.enhance = enhance
        self.mode = mode
        self.dpi = dpi
        self.api_key = api_key

        self._text_converter = TextConverter()
        self._ocr_converter = OCRConverter(engine=ocr_engine, lang=ocr_lang)
        self._hybrid_converter = HybridConverter(dpi=dpi)
        self._combined_converter = CombinedConverter(dpi=dpi)
        self._smart_converter = SmartConverter()

        self._docling_converter = None
        if self.mode == "docling":
            if HAS_DOCLING:
                self._docling_converter = DoclingConverter()
            else:
                logger.warning("Docling not installed. Install with: pip install docling")

        self._cloud_converter = None
        if self.mode == "cloud":
            self._cloud_converter = CloudAPIConverter(api_key=self.api_key)
        # Graceful init: these may fail if the binary is not installed
        try:
            self._msword_converter = NativeWordConverter() if HAS_MSWORD else None
        except (NotImplementedError, RuntimeError, OSError) as e:
            logger.debug("MS Word converter unavailable: %s", e)
            self._msword_converter = None

        try:
            self._libreoffice_converter = LibreOfficeConverter() if HAS_LIBREOFFICE else None
        except (FileNotFoundError, RuntimeError, OSError) as e:
            logger.debug("LibreOffice converter unavailable: %s", e)
            self._libreoffice_converter = None
        self._enhancer = DocxEnhancer()

    def convert(self, input_pdf: str, output_docx: str,
                pages: list[int] | None = None,
                force_ocr: bool = False) -> dict:
        """
        Convert a PDF to DOCX.

        Args:
            input_pdf: Path to the input PDF file.
            output_docx: Path for the output DOCX file.
            pages: Optional list of 0-indexed page numbers to convert.
            force_ocr: If True, always use OCR even for text PDFs.

        Returns:
            dict with conversion results:
            - output_path: path to the generated DOCX
            - analysis: PDF analysis results
            - method: 'text', 'ocr', or 'hybrid'
            - enhanced: whether post-processing was applied
        """
        if not os.path.isfile(input_pdf):
            raise FileNotFoundError(f"PDF file not found: {input_pdf}")

        # Ensure output directory exists
        output_dir = os.path.dirname(os.path.abspath(output_docx))
        os.makedirs(output_dir, exist_ok=True)

        # Step 1: Analyze the PDF
        logger.info("=" * 60)
        logger.info("PDF to Word Conversion Pipeline")
        logger.info("Input:  %s", input_pdf)
        logger.info("Output: %s", output_docx)
        logger.info("Mode:   %s", self.mode)
        logger.info("=" * 60)

        analyzer = PDFAnalyzer(input_pdf)
        analysis = analyzer.analyze()

        if "error" in analysis:
            raise RuntimeError(f"PDF analysis failed: {analysis['error']}")

        logger.info("Analysis: %d pages, scanned=%s", analysis["page_count"], analysis["is_scanned"])

        # Step 2: Convert using appropriate method
        if self.mode == "docling":
            method = "docling"
            if self._docling_converter is None:
                logger.warning("Docling unavailable, falling back to 'smart' mode.")
                method = "smart"
                self._smart_converter.convert(input_pdf, output_docx, pages=pages)
            else:
                logger.info("Using Docling ML converter (IBM)")
                self._docling_converter.convert(input_pdf, output_docx, pages=pages)
        elif self.mode == "cloud":
            method = "cloud"
            logger.info("Using cloud API converter (ConvertAPI)")
            self._cloud_converter.convert(input_pdf, output_docx, pages=pages)
        elif self.mode == "msword":
            if not HAS_MSWORD or not self._msword_converter:
                logger.warning("Mode 'msword' requested but is unavailable (requires Windows and pywin32). Falling back to 'libreoffice' or 'smart'.")
                if HAS_LIBREOFFICE and self._libreoffice_converter:
                    method = "libreoffice"
                    logger.info("Falling back to LibreOffice converter")
                    self._libreoffice_converter.convert(input_pdf, output_docx, pages=pages)
                else:
                    method = "smart"
                    self._smart_converter.convert(input_pdf, output_docx, pages=pages)
            else:
                method = "msword"
                logger.info("Using native MS Word converter (pywin32)")
                self._msword_converter.convert(input_pdf, output_docx, pages=pages)
        elif self.mode == "libreoffice":
            if not HAS_LIBREOFFICE or not self._libreoffice_converter:
                logger.warning("Mode 'libreoffice' requested but LibreOffice is not installed. Falling back to 'smart' mode.")
                method = "smart"
                self._smart_converter.convert(input_pdf, output_docx, pages=pages)
            else:
                method = "libreoffice"
                logger.info("Using LibreOffice headless converter")
                self._libreoffice_converter.convert(input_pdf, output_docx, pages=pages)
        elif self.mode == "smart":
            # Smart mode: PyMuPDF-based extraction with reading order
            method = "smart"
            logger.info("Using smart converter (PyMuPDF)")
            self._smart_converter.convert(input_pdf, output_docx, pages=pages)
        elif self.mode == "combined":
            # Combined mode: background image + editable text blocks
            method = "combined"
            logger.info("Using combined converter (dpi=%d)", self.dpi)
            self._combined_converter.convert(input_pdf, output_docx, pages=pages)
        elif self.mode == "hybrid":
            # Hybrid mode: PDF pages as images + invisible text overlay
            method = "hybrid"
            logger.info("Using hybrid converter (dpi=%d)", self.dpi)
            self._hybrid_converter.convert(input_pdf, output_docx, pages=pages)
        elif force_ocr or self.mode == "ocr" or analysis["is_scanned"]:
            method = "ocr"
            logger.info("Using OCR converter (engine=%s)", self.ocr_engine)
            self._ocr_converter.convert(input_pdf, output_docx)
        else:
            method = "text"
            logger.info("Using text converter (pdf2docx)")
            self._text_converter.convert(input_pdf, output_docx, pages=pages)

        # Step 3: Post-process the DOCX (not needed for hybrid, combined, smart, or msword)
        enhanced = False
        if self.enhance and method not in ("hybrid", "combined", "smart", "msword", "libreoffice", "cloud"):
            logger.info("Applying DOCX enhancement...")
            self._enhancer.enhance(output_docx, source_pdf_path=input_pdf)
            enhanced = True

        logger.info("=" * 60)
        logger.info("Conversion complete: %s", output_docx)
        logger.info("=" * 60)

        return {
            "output_path": output_docx,
            "analysis": analysis,
            "method": method,
            "enhanced": enhanced,
        }
