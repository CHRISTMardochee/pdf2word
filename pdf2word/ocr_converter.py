"""
OCRConverter module.
Converts scanned (image-based) PDFs to DOCX using OCR.
Supports Tesseract and PaddleOCR engines.
"""

import logging
import os
import tempfile

import fitz  # PyMuPDF
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

logger = logging.getLogger(__name__)


class OCRConverter:
    """Convert scanned PDFs to DOCX using OCR."""

    def __init__(self, engine: str = "tesseract", lang: str = "fra+eng"):
        """
        Args:
            engine: OCR engine to use ('tesseract' or 'paddleocr').
            lang: Language(s) for OCR. Default 'fra+eng' for French+English.
        """
        self.engine = engine.lower()
        self.lang = lang

    def convert(self, pdf_path: str, docx_path: str, dpi: int = 300) -> str:
        """
        Convert a scanned PDF to DOCX via OCR.

        Args:
            pdf_path: Path to the input PDF.
            docx_path: Path for the output DOCX.
            dpi: DPI for rendering PDF pages to images.

        Returns:
            Path to the generated DOCX file.
        """
        logger.info("OCR converting scanned PDF: %s -> %s (engine=%s)", pdf_path, docx_path, self.engine)

        doc = Document()
        pdf_doc = fitz.open(pdf_path)

        try:
            for page_num in range(len(pdf_doc)):
                logger.info("OCR processing page %d/%d", page_num + 1, len(pdf_doc))
                page = pdf_doc[page_num]

                # Render page to image
                mat = fitz.Matrix(dpi / 72, dpi / 72)
                pix = page.get_pixmap(matrix=mat)

                # Save as temp image
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                    tmp_path = tmp.name
                    pix.save(tmp_path)

                try:
                    # Run OCR
                    ocr_data = self._run_ocr(tmp_path)

                    # Build DOCX content from OCR results
                    self._build_page_content(doc, ocr_data, page.rect.width, page.rect.height)

                    # Add page break between pages (except last)
                    if page_num < len(pdf_doc) - 1:
                        doc.add_page_break()
                finally:
                    os.unlink(tmp_path)

            doc.save(docx_path)
            logger.info("OCR conversion complete: %s", docx_path)

        except Exception as e:
            logger.error("OCR conversion failed: %s", e)
            raise
        finally:
            pdf_doc.close()

        return docx_path

    def _run_ocr(self, image_path: str) -> list[dict]:
        """
        Run OCR on an image and return structured text data.

        Returns a list of dicts with keys:
        - text: recognized text
        - bbox: (x0, y0, x1, y1) bounding box
        - confidence: OCR confidence score
        """
        if self.engine == "tesseract":
            return self._run_tesseract(image_path)
        elif self.engine == "paddleocr":
            return self._run_paddleocr(image_path)
        else:
            raise ValueError(f"Unknown OCR engine: {self.engine}")

    def _run_tesseract(self, image_path: str) -> list[dict]:
        """Run Tesseract OCR on an image."""
        try:
            import pytesseract
        except ImportError:
            raise ImportError("pytesseract is required for Tesseract OCR. Install with: pip install pytesseract")

        img = Image.open(image_path)
        # Use image_to_data for detailed output with bounding boxes
        data = pytesseract.image_to_data(img, lang=self.lang, output_type=pytesseract.Output.DICT)

        results = []
        n_items = len(data["text"])

        current_line = []
        current_line_num = -1
        current_block_num = -1

        for i in range(n_items):
            text = data["text"][i].strip()
            conf = int(data["conf"][i])
            line_num = data["line_num"][i]
            block_num = data["block_num"][i]

            if conf < 0:
                continue

            # Detect line/block changes to group words into lines
            if line_num != current_line_num or block_num != current_block_num:
                if current_line:
                    results.append(self._merge_line(current_line))
                current_line = []
                current_line_num = line_num
                current_block_num = block_num

            if text:
                current_line.append({
                    "text": text,
                    "bbox": (data["left"][i], data["top"][i],
                             data["left"][i] + data["width"][i],
                             data["top"][i] + data["height"][i]),
                    "confidence": conf,
                    "block_num": block_num,
                    "line_num": line_num,
                })

        # Don't forget the last line
        if current_line:
            results.append(self._merge_line(current_line))

        return results

    def _run_paddleocr(self, image_path: str) -> list[dict]:
        """Run PaddleOCR on an image."""
        try:
            from paddleocr import PaddleOCR
        except ImportError:
            raise ImportError("paddleocr is required. Install with: pip install paddleocr")

        # Map language codes
        lang_map = {"fra": "fr", "eng": "en", "fra+eng": "fr"}
        paddle_lang = lang_map.get(self.lang, "fr")

        ocr = PaddleOCR(use_angle_cls=True, lang=paddle_lang, show_log=False)
        result = ocr.ocr(image_path, cls=True)

        results = []
        if result and result[0]:
            for line in result[0]:
                bbox_points = line[0]  # [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
                text = line[1][0]
                conf = line[1][1]

                x_coords = [p[0] for p in bbox_points]
                y_coords = [p[1] for p in bbox_points]

                results.append({
                    "text": text,
                    "bbox": (min(x_coords), min(y_coords), max(x_coords), max(y_coords)),
                    "confidence": conf * 100,
                })

        return results

    def _merge_line(self, words: list[dict]) -> dict:
        """Merge a list of word dicts into a single line dict."""
        text = " ".join(w["text"] for w in words)
        x0 = min(w["bbox"][0] for w in words)
        y0 = min(w["bbox"][1] for w in words)
        x1 = max(w["bbox"][2] for w in words)
        y1 = max(w["bbox"][3] for w in words)
        avg_conf = sum(w["confidence"] for w in words) / len(words)

        return {
            "text": text,
            "bbox": (x0, y0, x1, y1),
            "confidence": avg_conf,
            "block_num": words[0].get("block_num", 0),
        }

    def _build_page_content(self, doc: Document, ocr_data: list[dict],
                            page_width: float, page_height: float):
        """
        Build DOCX paragraphs from OCR data, grouping lines into paragraphs
        based on spatial proximity and block membership.
        """
        if not ocr_data:
            return

        # Group lines by block_num for paragraph reconstruction
        blocks = {}
        for item in ocr_data:
            block_id = item.get("block_num", 0)
            if block_id not in blocks:
                blocks[block_id] = []
            blocks[block_id].append(item)

        # Sort blocks by vertical position
        sorted_blocks = sorted(blocks.items(), key=lambda b: min(i["bbox"][1] for i in b[1]))

        for block_id, lines in sorted_blocks:
            # Sort lines within block by vertical position
            lines.sort(key=lambda l: l["bbox"][1])

            # Merge lines into a single paragraph text
            paragraph_text = " ".join(line["text"] for line in lines)

            if paragraph_text.strip():
                para = doc.add_paragraph(paragraph_text)
                # Set base font size
                for run in para.runs:
                    run.font.size = Pt(11)
