"""
PDF to Word (pdf2word) Conversion Pipeline.
Converts PDFs to Word docx format while preserving layout and reconstructing paragraphs.
"""

__version__ = "0.1.0"

from .converter import PDFToWordConverter

def convert(input_pdf: str, output_docx: str, ocr_engine: str = "tesseract", enhance: bool = True) -> str:
    """
    Convenience function to convert a PDF to Word.
    """
    converter = PDFToWordConverter(ocr_engine=ocr_engine, enhance=enhance)
    return converter.convert(input_pdf, output_docx)
