import os
import fitz  # PyMuPDF for PDF + images
import docx
from openpyxl import load_workbook
from PIL import Image
from PIL import ImageDraw
import tempfile


class AuditRAMHighlighter:

    def __init__(self, file_path, search_text):
        self.file_path = file_path
        self.search_text = search_text.lower().strip()
        self.ext = os.path.splitext(file_path)[1].lower()

    # ------------------------------------------------------
    # --------- MAIN EXECUTION FUNCTION --------------------
    # ------------------------------------------------------
    def run(self, output_path):
        if self.ext == ".pdf":
            self._process_pdf(output_path)

        elif self.ext in [".png", ".jpg", ".jpeg"]:
            self._process_image(output_path)

        elif self.ext == ".docx":
            self._process_word(output_path)

        elif self.ext == ".xlsx":
            self._process_excel(output_path)

        else:
            raise ValueError("Unsupported file format.")

    # ------------------------------------------------------
    # --------- PDF PROCESSING (with PyMuPDF) --------------
    # ------------------------------------------------------
    def _process_pdf(self, output_file):
        pdf = fitz.open(self.file_path)

        for page in pdf:
            text_instances = page.search_for(self.search_text, hit_max=5000)
            for inst in text_instances:
                # Draw a red rectangle with transparent fill
                page.add_rect_annot(inst).set_colors(stroke=(1, 0, 0))
                page.add_rect_annot(inst).update()

        pdf.save(output_file, deflate=True)
        pdf.close()

    # ------------------------------------------------------
    # --------- IMAGE PROCESSING (PIL) ----------------------
    # ------------------------------------------------------
    def _process_image(self, output_file):
        image = Image.open(self.file_path)
        draw = ImageDraw.Draw(image)

        # Convert image → temp PDF for text extraction
        temp_pdf = tempfile.mktemp(".pdf")
        image.save(temp_pdf, "PDF")

        pdf = fitz.open(temp_pdf)
        page = pdf[0]

        text_instances = page.search_for(self.search_text)
        for inst in text_instances:
            x0, y0, x1, y1 = inst
            draw.rectangle((x0, y0, x1, y1), outline="red", width=3)

        image.save(output_file)

    # ------------------------------------------------------
    # --------- WORD DOCUMENT PROCESSING (.docx) ------------
    # ------------------------------------------------------
    def _process_word(self, output_file):
        doc = docx.Document(self.file_path)

        # Convert Word → temp PDF for measurement & annotation
        temp_pdf = tempfile.mktemp(".pdf")

        # Save as PDF (MS Word COM automation on Windows only)
        try:
            import comtypes.client
            word = comtypes.client.CreateObject('Word.Application')
            doc_obj = word.Documents.Open(self.file_path)
            doc_obj.SaveAs(temp_pdf, FileFormat=17)
            doc_obj.Close()
            word.Quit()
        except:
            raise RuntimeError("Word-to-PDF conversion requires MS Word on Windows.")

        # Now annotate PDF and save final
        self.file_path = temp_pdf
        self._process_pdf(output_file)

    # ------------------------------------------------------
    # --------- EXCEL PROCESSING (.xlsx) --------------------
    # ------------------------------------------------------
    def _process_excel(self, output_file):

        wb = load_workbook(self.file_path)
        result_pdf = tempfile.mktemp(".pdf")

        # Convert Excel → PDF (Windows + MS Excel only)
        try:
            import comtypes.client
            excel = comtypes.client.CreateObject("Excel.Application")
            wb_obj = excel.Workbooks.Open(self.file_path)
            wb_obj.ExportAsFixedFormat(0, result_pdf)
            wb_obj.Close()
            excel.Quit()

        except:
            raise RuntimeError("Excel-to-PDF conversion requires MS Excel on Windows.")

        # Annotate generated PDF
        self.file_path = result_pdf
        self._process_pdf(output_file)
