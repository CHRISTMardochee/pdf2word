"""
MS Word Native Converter module.
Uses pywin32 COM interop to open a PDF in Microsoft Word and save it as a DOCX.
This provides the most accurate and natively editable reproduction of a PDF on Windows.
"""

import logging
import os
import sys

logger = logging.getLogger(__name__)


class NativeWordConverter:
    """
    Converts PDF to DOCX using Microsoft Word's native PDF Reflow engine via COM.
    Only available on Windows systems with MS Word installed.
    """

    def __init__(self):
        # Fail fast if not on Windows
        if sys.platform != "win32":
            raise NotImplementedError(
                "NativeWordConverter is only supported on Windows. "
                "Please use an alternative converter mode (e.g., 'auto', 'text')."
            )
        
        # Check if pywin32 is installed
        try:
            import win32com.client
            self.win32com_client = win32com.client
        except ImportError:
            raise ImportError(
                "The 'pywin32' package is required for native MS Word conversion. "
                "Install it with: pip install pywin32"
            )

    def convert(self, input_pdf: str, output_docx: str, pages: list[int] | None = None) -> str:
        """
        Convert a PDF to DOCX using MS Word.

        Args:
            input_pdf: Path to the input PDF.
            output_docx: Path for the output DOCX.
            pages: Ignored for this converter, as MS Word imports the whole document natively.

        Returns:
            Path to the generated DOCX file.
        """
        logger.info("Native MS Word converting: %s -> %s", input_pdf, output_docx)

        if pages is not None:
            logger.warning("The 'pages' argument is ignored by NativeWordConverter; the entire PDF will be converted.")

        input_pdf_abs = os.path.abspath(input_pdf)
        output_docx_abs = os.path.abspath(output_docx)

        if not os.path.isfile(input_pdf_abs):
            raise FileNotFoundError(f"Input PDF not found: {input_pdf_abs}")

        word = None
        try:
            # Dispatch Word Application
            # Using DispatchEx instead of Dispatch ensures a new instance is created,
            # avoiding interference with any Word windows the user currently has open.
            word = self.win32com_client.DispatchEx("Word.Application")
            
            # Hide the application window and disable alerts to keep it silent
            word.Visible = False
            word.DisplayAlerts = 0  # wdAlertsNone

            logger.info("Opening PDF in MS Word (this may take a moment)...")
            
            # Open the PDF. 
            # ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
            doc = word.Documents.Open(
                FileName=input_pdf_abs, 
                ConfirmConversions=False, 
                ReadOnly=True, 
                AddToRecentFiles=False
            )

            logger.info("Saving as DOCX...")
            
            # SaveAs2 with FileFormat=16 (wdFormatXMLDocument, which is .docx)
            # wdFormatXMLDocument = 16
            doc.SaveAs2(FileName=output_docx_abs, FileFormat=16)
            
            doc.Close(SaveChanges=0)  # wdDoNotSaveChanges = 0
            logger.info("Native MS Word conversion complete: %s", output_docx_abs)
            
            return output_docx_abs

        except Exception as e:
            logger.error("Native MS Word conversion failed: %s", e)
            
            # Catch common COM errors, like Word not being installed
            error_msg = str(e)
            if "Invalid class string" in error_msg or "-2147221005" in error_msg:
                raise RuntimeError(
                    "Microsoft Word does not appear to be installed on this system. "
                    "Native conversion requires MS Word."
                ) from e
                
            raise RuntimeError(f"MS Word conversion error: {e}") from e

        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception as quit_err:
                    logger.warning("Failed to quit MS Word application cleanly: %s", quit_err)
