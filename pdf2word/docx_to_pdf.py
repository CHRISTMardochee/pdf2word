"""
DOCX to PDF reconversion module.
Uses LibreOffice in headless mode to convert DOCX back to PDF.
"""

import logging
import os
import subprocess
import shutil

logger = logging.getLogger(__name__)


def find_libreoffice() -> str | None:
    """Find the LibreOffice executable on the system."""
    # Check if soffice is on PATH
    soffice = shutil.which("soffice")
    if soffice:
        return soffice

    # Common installation paths
    common_paths = [
        # Windows
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        # Linux
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/local/bin/soffice",
        # macOS
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]

    for path in common_paths:
        if os.path.isfile(path):
            return path

    return None


def docx_to_pdf(docx_path: str, output_dir: str | None = None) -> str:
    """
    Convert a DOCX file to PDF using LibreOffice headless.

    Args:
        docx_path: Path to the input DOCX file.
        output_dir: Directory for the output PDF. If None, uses the same
                    directory as the input file.

    Returns:
        Path to the generated PDF file.

    Raises:
        FileNotFoundError: If LibreOffice is not installed.
        RuntimeError: If the conversion fails.
    """
    soffice = find_libreoffice()
    if soffice is None:
        raise FileNotFoundError(
            "LibreOffice not found. Please install LibreOffice.\n"
            "  Windows: https://www.libreoffice.org/download/\n"
            "  Linux:   sudo apt install libreoffice\n"
            "  macOS:   brew install --cask libreoffice"
        )

    if output_dir is None:
        output_dir = os.path.dirname(os.path.abspath(docx_path))

    os.makedirs(output_dir, exist_ok=True)

    cmd = [
        soffice,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        os.path.abspath(docx_path),
    ]

    logger.info("Converting DOCX to PDF: %s", " ".join(cmd))

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=120,
        )

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed:\n"
                f"stdout: {result.stdout}\n"
                f"stderr: {result.stderr}"
            )

        # Determine output PDF path
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        pdf_path = os.path.join(output_dir, base_name + ".pdf")

        if not os.path.isfile(pdf_path):
            raise RuntimeError(f"Expected output PDF not found: {pdf_path}")

        logger.info("DOCX to PDF conversion complete: %s", pdf_path)
        return pdf_path

    except subprocess.TimeoutExpired:
        raise RuntimeError("LibreOffice conversion timed out (120s)")
