"""
TextConverter module.
Converts text-based (native) PDFs to DOCX using pdf2docx as the primary engine.
"""

import logging
from pdf2docx import Converter as Pdf2DocxConverter

logger = logging.getLogger(__name__)


class TextConverter:
    """Convert text-based PDFs to DOCX using pdf2docx."""

    def __init__(self):
        pass

    def convert(self, pdf_path: str, docx_path: str, pages: list[int] | None = None) -> str:
        """
        Convert a text-based PDF to DOCX.

        Args:
            pdf_path: Path to the input PDF.
            docx_path: Path for the output DOCX.
            pages: Optional list of 0-indexed page numbers to convert.
                   If None, converts all pages.

        Returns:
            Path to the generated DOCX file.
        """
        logger.info("Converting text PDF: %s -> %s", pdf_path, docx_path)

        cv = Pdf2DocxConverter(pdf_path)
        # Tuning parameters to prevent text dropping on complex/overlapping layouts
        kwargs = {
            "connected_border_tolerance": 2.0,
            "line_overlap_threshold": 0.8,
            "line_margin_weight": 2.0,
            "word_margin_weight": 2.0,
            "clip_image_res_ratio": 2.0,
        }
        try:
            if pages is not None:
                # pdf2docx expects 0-indexed page numbers
                cv.convert(docx_path, pages=pages, **kwargs)
            else:
                cv.convert(docx_path, **kwargs)

            logger.info("Text conversion complete: %s", docx_path)
        except Exception as e:
            logger.error("Text conversion failed: %s", e)
            raise
        finally:
            cv.close()

        return docx_path
