"""
LibreOffice Headless Converter module.
Uses LibreOffice in headless mode to convert PDF to DOCX.
This is the best open-source alternative to MS Word's PDF Reflow engine,
suitable for Linux server deployments.
"""

import logging
import os
import shutil
import subprocess
import tempfile
import uuid

logger = logging.getLogger(__name__)

# Common LibreOffice binary names/paths by platform
_LIBRE_OFFICE_PATHS = [
    "libreoffice",
    "soffice",
    # Linux package managers
    "/usr/bin/libreoffice",
    "/usr/bin/soffice",
    # macOS
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    # Windows (if installed)
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


def _find_libreoffice() -> str | None:
    """Find the LibreOffice binary on the system."""
    for path in _LIBRE_OFFICE_PATHS:
        if shutil.which(path):
            return path
    return None


class LibreOfficeConverter:
    """
    Converts PDF to DOCX using LibreOffice's headless mode.
    Works on Linux, macOS, and Windows (if LibreOffice is installed).
    
    Quality notes:
    - Uses the writer_pdf_import filter for best text extraction
    - Quality is good for text-heavy PDFs
    - Complex layouts (multi-column, heavy graphics) may have some differences
    - Installing Microsoft fonts (ttf-mscorefonts-installer) improves fidelity
    """

    def __init__(self, soffice_path: str | None = None, timeout: int = 300):
        """
        Args:
            soffice_path: Optional explicit path to soffice/libreoffice binary.
                          If None, auto-detected.
            timeout: Maximum seconds to wait for conversion (default: 300 for large PDFs).
        """
        self.soffice_path = soffice_path or _find_libreoffice()
        self.timeout = timeout

        if not self.soffice_path:
            raise FileNotFoundError(
                "LibreOffice not found on this system. "
                "Install it with:\n"
                "  Ubuntu/Debian: sudo apt install libreoffice-writer\n"
                "  macOS: brew install --cask libreoffice\n"
                "  Windows: https://www.libreoffice.org/download/"
            )

        logger.info("LibreOffice binary: %s", self.soffice_path)

    def convert(self, input_pdf: str, output_docx: str,
                pages: list[int] | None = None) -> str:
        """
        Convert a PDF to DOCX using LibreOffice headless.

        Args:
            input_pdf: Path to the input PDF.
            output_docx: Path for the output DOCX.
            pages: Ignored — LibreOffice converts the entire document.

        Returns:
            Path to the generated DOCX file.
        """
        logger.info("LibreOffice converting: %s -> %s", input_pdf, output_docx)

        if pages is not None:
            logger.warning(
                "The 'pages' argument is ignored by LibreOfficeConverter; "
                "the entire PDF will be converted."
            )

        input_pdf_abs = os.path.abspath(input_pdf)
        output_docx_abs = os.path.abspath(output_docx)

        if not os.path.isfile(input_pdf_abs):
            raise FileNotFoundError(f"Input PDF not found: {input_pdf_abs}")

        # Use a unique temporary directory for output to avoid conflicts
        # and support concurrent conversions
        unique_id = uuid.uuid4().hex[:8]
        tmp_outdir = tempfile.mkdtemp(prefix=f"lo_convert_{unique_id}_")
        
        # Unique user profile to allow concurrent LibreOffice instances
        tmp_profile = tempfile.mkdtemp(prefix=f"lo_profile_{unique_id}_")

        try:
            # Build the LibreOffice command
            cmd = [
                self.soffice_path,
                "--headless",
                "--norestore",
                "--nofirststartwizard",
                f"-env:UserInstallation=file:///{tmp_profile.replace(os.sep, '/')}",
                '--infilter=writer_pdf_import',
                "--convert-to", "docx",
                "--outdir", tmp_outdir,
                input_pdf_abs,
            ]

            logger.info("Running: %s", " ".join(cmd))

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=self.timeout,
                cwd=tmp_outdir,
            )

            logger.debug("LibreOffice stdout: %s", result.stdout)
            if result.stderr:
                logger.debug("LibreOffice stderr: %s", result.stderr)

            if result.returncode != 0:
                raise RuntimeError(
                    f"LibreOffice conversion failed (exit code {result.returncode}):\n"
                    f"stdout: {result.stdout}\n"
                    f"stderr: {result.stderr}"
                )

            # LibreOffice outputs to tmp_outdir with the same basename but .docx extension
            pdf_basename = os.path.splitext(os.path.basename(input_pdf_abs))[0]
            tmp_docx = os.path.join(tmp_outdir, pdf_basename + ".docx")

            if not os.path.isfile(tmp_docx):
                # Sometimes the output name differs; find any .docx in the output dir
                docx_files = [f for f in os.listdir(tmp_outdir) if f.endswith(".docx")]
                if docx_files:
                    tmp_docx = os.path.join(tmp_outdir, docx_files[0])
                else:
                    raise RuntimeError(
                        f"LibreOffice did not produce a DOCX file. "
                        f"Output dir contents: {os.listdir(tmp_outdir)}\n"
                        f"stdout: {result.stdout}\n"
                        f"stderr: {result.stderr}"
                    )

            # Move the result to the desired output path
            output_dir = os.path.dirname(output_docx_abs)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            shutil.move(tmp_docx, output_docx_abs)

            logger.info("LibreOffice conversion complete: %s", output_docx_abs)
            return output_docx_abs

        except subprocess.TimeoutExpired:
            raise RuntimeError(
                f"LibreOffice conversion timed out after {self.timeout}s. "
                f"Try increasing the timeout for large PDFs."
            )

        finally:
            # Clean up temporary directories
            shutil.rmtree(tmp_outdir, ignore_errors=True)
            shutil.rmtree(tmp_profile, ignore_errors=True)
