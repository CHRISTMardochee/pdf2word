"""
PDFAnalyzer module.
Analyzes PDF files to determine their type (text-based vs scanned) and extract metadata.
"""

import fitz  # PyMuPDF
import logging

logger = logging.getLogger(__name__)


class PDFAnalyzer:
    """Analyze a PDF to determine its type and extract metadata."""

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self._doc = None

    def analyze(self) -> dict:
        """
        Analyze the PDF and return a dict with:
        - page_count
        - is_scanned: True if the PDF is mostly images (scanned)
        - text_ratio: ratio of text characters per page
        - has_images: True if images are found
        - metadata: PDF metadata dict
        - page_sizes: list of (width, height) tuples in points
        """
        try:
            self._doc = fitz.open(self.pdf_path)

            page_count = len(self._doc)
            is_scanned, text_ratio = self._check_if_scanned()
            has_images = self._check_has_images()
            metadata = dict(self._doc.metadata) if self._doc.metadata else {}
            page_sizes = [(page.rect.width, page.rect.height) for page in self._doc]

            result = {
                "page_count": page_count,
                "is_scanned": is_scanned,
                "text_ratio": text_ratio,
                "has_images": has_images,
                "metadata": metadata,
                "page_sizes": page_sizes,
            }

            logger.info(
                "PDF Analysis: %d pages, scanned=%s, text_ratio=%.2f, images=%s",
                page_count, is_scanned, text_ratio, has_images,
            )
            return result

        except Exception as e:
            logger.error("Failed to analyze PDF: %s", e)
            return {"error": str(e)}
        finally:
            if self._doc is not None:
                self._doc.close()

    def _check_if_scanned(self) -> tuple[bool, float]:
        """
        Check if PDF is scanned by measuring extractable text per page.
        Returns (is_scanned, avg_text_chars_per_page).
        """
        total_text = 0
        pages_to_check = min(5, len(self._doc))

        for i in range(pages_to_check):
            page = self._doc[i]
            text = page.get_text().strip()
            total_text += len(text)

        avg_chars = total_text / max(pages_to_check, 1)

        # If less than 50 chars average per page, it's likely scanned
        is_scanned = avg_chars < 50
        return is_scanned, avg_chars

    def _check_has_images(self) -> bool:
        """Check if PDF contains embedded images."""
        pages_to_check = min(3, len(self._doc))
        for i in range(pages_to_check):
            page = self._doc[i]
            images = page.get_images(full=True)
            if images:
                return True
        return False
