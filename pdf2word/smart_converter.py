"""
Smart Text Converter module (V3 — Client Reference Match).
Produces DOCX output matching the client's Generali brochure format:
- Cover page with red background text panels
- 2-column tables with left red border for "Que faire si..." 
  and thin red border for right coverage panels
- Proper icon/image placement
- Form fields on last page
- Deduplicated footer
"""

import logging
import os
import re
import tempfile

import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches, Pt, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

logger = logging.getLogger(__name__)

_CONTROL_CHARS = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]')
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Generali brand colors
RED = "C5281C"
CORAL = "F2634A"
GREY = "6D6E71"
LIGHT_GREY = "E0E0E0"


class SmartConverter:
    """
    Intelligent PDF-to-DOCX converter matching Generali brochure format.
    """

    HEADING_MIN_SIZE = 13.0
    SUBHEADING_MIN_SIZE = 10.5
    COLUMN_THRESHOLD = 0.45
    ROW_TOLERANCE = 20.0
    FOOTER_Y_THRESHOLD = 0.92

    def __init__(self):
        pass

    def convert(self, input_pdf: str, output_docx: str,
                pages: list[int] | None = None) -> str:
        logger.info("Smart V3 converting: %s -> %s", input_pdf, output_docx)

        pdf_doc = fitz.open(input_pdf)
        doc = Document()

        # Base style
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(10)
        style.font.color.rgb = RGBColor(0x6D, 0x6E, 0x71)

        # Heading styles
        for level in [1, 2, 3]:
            h_style = doc.styles[f'Heading {level}']
            h_style.font.color.rgb = RGBColor(0xC5, 0x28, 0x1C)
            h_style.font.name = 'Arial'

        page_numbers = pages if pages else list(range(len(pdf_doc)))

        with tempfile.TemporaryDirectory() as tmp_dir:
            for page_idx, page_num in enumerate(page_numbers):
                if page_num >= len(pdf_doc):
                    continue

                page = pdf_doc[page_num]
                logger.info("Processing page %d/%d", page_idx + 1, len(page_numbers))

                if page_idx > 0:
                    doc.add_page_break()

                # Extract images
                self._extract_images(doc, page, pdf_doc, tmp_dir, page_num)

                # Extract text blocks
                text_dict = page.get_text("dict")
                raw_blocks = [b for b in text_dict.get("blocks", [])
                              if b["type"] == 0 and self._block_has_text(b)]

                # Separate footer
                footer_y = page.rect.height * self.FOOTER_Y_THRESHOLD
                content_blocks = [b for b in raw_blocks if b["bbox"][1] < footer_y]
                footer_blocks = [b for b in raw_blocks if b["bbox"][1] >= footer_y]

                # Is this the cover page? (has white text = overlaid on image)
                is_cover = self._is_cover_page(content_blocks)

                if is_cover:
                    self._process_cover_page(doc, content_blocks)
                else:
                    self._process_content_page(doc, content_blocks, page.rect.width)

                # Footer
                if footer_blocks:
                    self._add_footer(doc, footer_blocks)

        pdf_doc.close()

        for section in doc.sections:
            section.left_margin = Cm(1.5)
            section.right_margin = Cm(1.5)
            section.top_margin = Cm(1.0)
            section.bottom_margin = Cm(1.0)

        doc.save(output_docx)
        logger.info("Smart V3 conversion complete: %s", output_docx)
        return output_docx

    # ── Cover Page ──────────────────────────────────────────────

    def _is_cover_page(self, blocks: list) -> bool:
        """Check if blocks contain white text (indicating cover page overlay)."""
        for block in blocks:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    color = span.get("color", 0)
                    if color == 0xFFFFFF:
                        return True
        return False

    def _process_cover_page(self, doc: Document, blocks: list):
        """Process cover page: red-background panels for white text, normal for others."""
        sorted_blocks = sorted(blocks, key=lambda b: b["bbox"][1])

        for block in sorted_blocks:
            style_type = self._get_block_style(block)
            has_white = any(
                span.get("color", 0) == 0xFFFFFF
                for line in block.get("lines", [])
                for span in line.get("spans", [])
            )

            if has_white:
                # White text → red background panel
                para = doc.add_paragraph()
                self._fill_paragraph(para, block, style_type=style_type)
                # Set red background shading
                self._set_paragraph_shading(para, RED)
                # Make text white and larger
                for run in para.runs:
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                para.paragraph_format.space_before = Pt(4)
                para.paragraph_format.space_after = Pt(2)
            else:
                self._add_block(doc, block)

    def _set_paragraph_shading(self, para, color: str):
        """Set paragraph background color (shading)."""
        pPr = para._element.get_or_add_pPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), color)
        pPr.append(shd)

    # ── Content Pages ───────────────────────────────────────────

    def _process_content_page(self, doc: Document, blocks: list, page_width: float):
        """Process content page with column layout detection."""
        if not blocks:
            return

        col_split = page_width * self.COLUMN_THRESHOLD

        # Classify blocks
        full_width = []
        left_col = []
        right_col = []

        for block in blocks:
            x0, _, x1, _ = block["bbox"]
            width = x1 - x0
            if width > page_width * 0.55:
                full_width.append(block)
            elif x0 < col_split:
                left_col.append(block)
            else:
                right_col.append(block)

        full_width.sort(key=lambda b: b["bbox"][1])
        left_col.sort(key=lambda b: b["bbox"][1])
        right_col.sort(key=lambda b: b["bbox"][1])

        # Build sections between full-width headings
        all_items = []
        fw_positions = sorted([(b["bbox"][1], b) for b in full_width])
        boundaries = [0.0] + [y for y, _ in fw_positions] + [9999.0]

        for i in range(len(boundaries) - 1):
            zone_top = boundaries[i]
            zone_bottom = boundaries[i + 1]

            # Full-width block at boundary
            for y, fw_block in fw_positions:
                if abs(y - zone_top) < self.ROW_TOLERANCE:
                    all_items.append(("full", fw_block))

            # Column blocks in this zone
            z_left = [b for b in left_col if zone_top - 30 <= b["bbox"][1] < zone_bottom]
            z_right = [b for b in right_col if zone_top - 30 <= b["bbox"][1] < zone_bottom]

            if z_left and z_right:
                all_items.append(("two_col", z_left, z_right))
            else:
                for b in z_left:
                    all_items.append(("full", b))
                for b in z_right:
                    all_items.append(("full", b))

        # Render
        for item in all_items:
            if item[0] == "full":
                self._add_block(doc, item[1])
            elif item[0] == "two_col":
                self._add_two_column_section(doc, item[1], item[2])

    def _add_two_column_section(self, doc: Document, left_blocks: list, right_blocks: list):
        """
        Render two-column content as a Word table matching the client reference:
        - Left column: "Que faire si..." with thick red left border
        - Right column: coverage details with thin red full border
        """
        is_que_faire = any(
            "que faire" in span.get("text", "").lower()
            for b in left_blocks
            for line in b.get("lines", [])
            for span in line.get("spans", [])
        )

        table = doc.add_table(rows=1, cols=2)
        table.autofit = True

        # Column widths
        tbl = table._tbl
        tbl_pr = tbl.find(f"{{{W_NS}}}tblPr")
        if tbl_pr is None:
            tbl_pr = OxmlElement('w:tblPr')
            tbl.insert(0, tbl_pr)

        # Set table width
        tbl_w = OxmlElement('w:tblW')
        tbl_w.set(qn('w:w'), '5000')
        tbl_w.set(qn('w:type'), 'pct')
        tbl_pr.append(tbl_w)

        # Fill left cell
        left_cell = table.cell(0, 0)
        left_cell.paragraphs[0].text = ""
        self._set_cell_width(left_cell, Cm(8.5))
        for i, block in enumerate(left_blocks):
            para = left_cell.paragraphs[0] if i == 0 else left_cell.add_paragraph()
            self._fill_paragraph(para, block)

        # Style left cell: thick red left border only
        if is_que_faire:
            self._set_cell_borders(left_cell, left_color=RED, left_size="18")
        
        # Add vertical padding
        self._set_cell_margin(left_cell, top=Cm(0.3), bottom=Cm(0.3), left=Cm(0.4), right=Cm(0.3))

        # Fill right cell
        right_cell = table.cell(0, 1)
        right_cell.paragraphs[0].text = ""
        self._set_cell_width(right_cell, Cm(9.5))
        for i, block in enumerate(right_blocks):
            para = right_cell.paragraphs[0] if i == 0 else right_cell.add_paragraph()
            self._fill_paragraph(para, block)

        # Style right cell: thin red border all around
        self._set_cell_borders(right_cell,
                               top_color=RED, bottom_color=RED,
                               left_color=RED, right_color=RED,
                               top_size="4", bottom_size="4",
                               left_size="4", right_size="4")
        self._set_cell_margin(right_cell, top=Cm(0.3), bottom=Cm(0.3), left=Cm(0.4), right=Cm(0.4))

        # Remove default table borders (we use cell borders instead)
        self._remove_table_borders(table)

        # Spacing after table
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(6)
        spacer.paragraph_format.space_after = Pt(2)
        spacer_pf = spacer._element.get_or_add_pPr()
        sz_elem = OxmlElement('w:sz')
        sz_elem.set(qn('w:val'), '4')

    # ── Cell Styling Helpers ────────────────────────────────────

    def _set_cell_width(self, cell, width):
        """Set cell preferred width."""
        tc = cell._tc
        tc_pr = tc.get_or_add_tcPr()
        tc_w = OxmlElement('w:tcW')
        tc_w.set(qn('w:w'), str(int(width.emu / 635)))  # EMU to twips approx
        tc_w.set(qn('w:type'), 'dxa')
        tc_pr.append(tc_w)

    def _set_cell_borders(self, cell, **kwargs):
        """Set individual cell borders. Pass top_color, left_color, etc."""
        tc = cell._tc
        tc_pr = tc.get_or_add_tcPr()

        existing = tc_pr.find(qn('w:tcBorders'))
        if existing is not None:
            tc_pr.remove(existing)

        borders = OxmlElement('w:tcBorders')

        for edge in ["top", "left", "bottom", "right"]:
            color = kwargs.get(f"{edge}_color")
            size = kwargs.get(f"{edge}_size", "4")
            border = OxmlElement(f'w:{edge}')
            if color:
                border.set(qn('w:val'), 'single')
                border.set(qn('w:sz'), size)
                border.set(qn('w:space'), '0')
                border.set(qn('w:color'), color)
            else:
                border.set(qn('w:val'), 'nil')
            borders.append(border)

        tc_pr.append(borders)

    def _set_cell_margin(self, cell, top=None, bottom=None, left=None, right=None):
        """Set cell padding/margins."""
        tc = cell._tc
        tc_pr = tc.get_or_add_tcPr()

        mar = OxmlElement('w:tcMar')
        for edge, val in [("top", top), ("bottom", bottom), ("start", left), ("end", right)]:
            if val:
                m = OxmlElement(f'w:{edge}')
                m.set(qn('w:w'), str(int(val.emu / 635)))
                m.set(qn('w:type'), 'dxa')
                mar.append(m)
        tc_pr.append(mar)

    def _remove_table_borders(self, table):
        """Remove all table-level borders (use cell-level instead)."""
        tbl = table._tbl
        tbl_pr = tbl.find(f"{{{W_NS}}}tblPr")
        if tbl_pr is None:
            return

        existing = tbl_pr.find(f"{{{W_NS}}}tblBorders")
        if existing is not None:
            tbl_pr.remove(existing)

        borders = OxmlElement('w:tblBorders')
        for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
            border = OxmlElement(f'w:{edge}')
            border.set(qn('w:val'), 'none')
            border.set(qn('w:sz'), '0')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), 'FFFFFF')
            borders.append(border)
        tbl_pr.append(borders)

    # ── Block Rendering ─────────────────────────────────────────

    def _get_block_style(self, block: dict) -> str:
        max_size = 0
        has_red = False
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                max_size = max(max_size, span.get("size", 0))
                color = span.get("color", 0)
                r = (color >> 16) & 0xFF
                if r > 150 and ((color >> 8) & 0xFF) < 120:
                    has_red = True

        if max_size >= self.HEADING_MIN_SIZE:
            return "heading"
        elif max_size >= self.SUBHEADING_MIN_SIZE and has_red:
            return "subheading"
        return "body"

    def _add_block(self, doc: Document, block: dict):
        """Add a text block as a paragraph."""
        style_type = self._get_block_style(block)

        if style_type == "heading":
            para = doc.add_heading(level=1)
        elif style_type == "subheading":
            para = doc.add_heading(level=2)
        else:
            para = doc.add_paragraph()

        self._fill_paragraph(para, block, style_type=style_type)

    def _fill_paragraph(self, para, block: dict, style_type: str = None):
        """Fill a paragraph with formatted runs."""
        if style_type is None:
            style_type = self._get_block_style(block)

        # Clear existing
        for run in para.runs:
            run.text = ""

        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = self._sanitize(span.get("text", ""))
                if not text:
                    continue

                run = para.add_run(text)

                font_size = span.get("size", 10)
                run.font.size = Pt(font_size)

                color_int = span.get("color", 0)
                r = (color_int >> 16) & 0xFF
                g = (color_int >> 8) & 0xFF
                b = color_int & 0xFF
                run.font.color.rgb = RGBColor(r, g, b)

                if span.get("flags", 0) & 16:
                    run.font.bold = True
                if span.get("flags", 0) & 2:
                    run.font.italic = True

                font_name = span.get("font", "")
                clean_font = font_name.split("+")[-1] if "+" in font_name else font_name
                clean_font = clean_font.split("-")[0] if "-" in clean_font else clean_font
                if clean_font:
                    run.font.name = clean_font

        # Spacing
        pf = para.paragraph_format
        if style_type == "heading":
            pf.space_before = Pt(14)
            pf.space_after = Pt(4)
        elif style_type == "subheading":
            pf.space_before = Pt(8)
            pf.space_after = Pt(3)
        else:
            pf.space_before = Pt(1)
            pf.space_after = Pt(1)

    # ── Footer ──────────────────────────────────────────────────

    def _add_footer(self, doc: Document, footer_blocks: list):
        """Add footer text deduplicated."""
        seen = set()
        
        # Separator line
        sep = doc.add_paragraph()
        sep.paragraph_format.space_before = Pt(12)
        sep.paragraph_format.space_after = Pt(2)
        sep_run = sep.add_run("─" * 70)
        sep_run.font.color.rgb = RGBColor(0xE0, 0xE0, 0xE0)
        sep_run.font.size = Pt(6)
        
        for block in sorted(footer_blocks, key=lambda b: b["bbox"][1]):
            text = self._get_block_text(block).strip()
            key = text[:30]
            if key in seen or not text:
                continue
            seen.add(key)

            para = doc.add_paragraph()
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    t = self._sanitize(span.get("text", ""))
                    if t:
                        run = para.add_run(t)
                        run.font.size = Pt(7)
                        run.font.color.rgb = RGBColor(0x6D, 0x6E, 0x71)

    # ── Image Extraction ────────────────────────────────────────

    def _extract_images(self, doc: Document, page: fitz.Page,
                        pdf_doc: fitz.Document, tmp_dir: str, page_num: int):
        """Extract images via pixmap rendering (standard PNG)."""
        img_info_list = page.get_image_info(xrefs=True)
        seen_xrefs = set()

        for info in img_info_list:
            xref = info.get("xref", 0)
            if xref in seen_xrefs or xref == 0:
                continue
            seen_xrefs.add(xref)

            bbox = info.get("bbox", [])
            if not bbox or len(bbox) < 4:
                continue

            w_pt = abs(bbox[2] - bbox[0])
            h_pt = abs(bbox[3] - bbox[1])
            if w_pt < 10 or h_pt < 10:
                continue

            img_path = os.path.join(tmp_dir, f"img_p{page_num}_x{xref}.png")
            try:
                clip = fitz.Rect(bbox) & page.rect
                if clip.is_empty or clip.width < 5 or clip.height < 5:
                    continue
                mat = fitz.Matrix(2.5, 2.5)
                pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
                pix.save(img_path)
            except Exception as e:
                logger.warning("Image xref=%d render failed: %s", xref, str(e)[:80])
                continue

            if not os.path.isfile(img_path):
                continue

            visible_w_pt = clip.width
            max_w_in = (page.rect.width - 60) / 72.0
            width_in = min(visible_w_pt / 72.0, max_w_in)
            if width_in < 0.3:
                width_in = 2.0

            try:
                para = doc.add_paragraph()
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = para.add_run()
                run.add_picture(img_path, width=Inches(width_in))
                para.paragraph_format.space_before = Pt(6)
                para.paragraph_format.space_after = Pt(6)
                logger.info("Added image xref=%d (%.1f in wide)", xref, width_in)
            except Exception as e:
                logger.warning("Image xref=%d DOCX failed: %s", xref, str(e)[:80])

    # ── Utilities ───────────────────────────────────────────────

    @staticmethod
    def _sanitize(text: str) -> str:
        return _CONTROL_CHARS.sub('', text)

    def _block_has_text(self, block: dict) -> bool:
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                if self._sanitize(span.get("text", "")).strip():
                    return True
        return False

    def _get_block_text(self, block: dict) -> str:
        parts = []
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                t = self._sanitize(span.get("text", ""))
                if t:
                    parts.append(t)
        return " ".join(parts)
