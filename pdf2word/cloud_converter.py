"""
Cloud API Converter module.
Uses the ConvertAPI service to securely and perfectly convert complex PDFs to DOCX.
Requires an internet connection and a valid API key (CONVERTAPI_SECRET).
"""

import os
import logging
import convertapi

from .config import load_api_key

logger = logging.getLogger(__name__)

class CloudAPIConverter:
    """
    Converts PDF to DOCX using ConvertAPI.
    """

    def __init__(self, api_key: str | None = None):
        # 1. Parameter, 2. Env Var, 3. Config file
        self.api_key = api_key or os.environ.get("CONVERTAPI_SECRET") or load_api_key()
        if not self.api_key:
            raise ValueError(
                "A ConvertAPI secret key is required for cloud mode.\n"
                "Provide it via --api-key, set the CONVERTAPI_SECRET env var,\n"
                "or save it globally using: python -m pdf2word set-key YOUR_KEY"
            )
        convertapi.api_credentials = self.api_key

    def convert(self, input_pdf: str, output_docx: str, pages: list[int] | None = None) -> str:
        """
        Convert a PDF to DOCX using ConvertAPI.

        Args:
            input_pdf: Path to the input PDF.
            output_docx: Path for the output DOCX.
            pages: Ignored; ConvertAPI processes the whole file.
            
        Returns:
            Path to the generated DOCX file.
        """
        logger.info("Cloud API converting: %s -> %s", input_pdf, output_docx)
        
        if pages is not None:
            logger.warning("The 'pages' argument is ignored in cloud mode. The entire PDF will be converted.")

        input_pdf_abs = os.path.abspath(input_pdf)
        output_docx_abs = os.path.abspath(output_docx)

        if not os.path.isfile(input_pdf_abs):
            raise FileNotFoundError(f"Input PDF not found: {input_pdf_abs}")

        logger.info("Uploading PDF to ConvertAPI servers for processing...")
        try:
            # We use pdf to docx endpoint
            result = convertapi.convert('docx', {
                'File': input_pdf_abs,
            }, from_format='pdf')
            
            logger.info("Conversion complete. Downloading DOCX...")
            result.save_files(output_docx_abs)
            
            logger.info("Cloud conversion successfully saved to: %s", output_docx_abs)
            return output_docx_abs
            
        except convertapi.ApiError as e:
            logger.error("ConvertAPI error: %s", e)
            raise RuntimeError(f"Cloud API conversion failed: {e}") from e
        except Exception as e:
            logger.error("Unexpected error during cloud conversion: %s", e)
            raise RuntimeError(f"Cloud conversion failed: {e}") from e
