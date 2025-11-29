# AuditRAM – Automated Text Search & Bounding Box Highlighter

AuditRAM is a Python-based auditing tool that automatically finds and marks text inside documents by drawing **red, non-filled bounding boxes** around identified text elements — without altering or modifying the original document.

This project fulfills the requirements of the SCIT Pune **AuditRAM Assignment**.

---

#  Features

- Accepts multiple file formats:
  - PDF
  - Image files (JPG, PNG)
  - Word files (.docx)
  - Excel files (.xlsx)
- Draws a **red bounding box** over all matched text
- Supports **OCR using Tesseract** for scanned images
- Exports to:
  - PDF
  - PNG/JPG (for images)
- GUI application (Tkinter)
- Command-line interface (CLI)
- Converts DOCX/XLSX to PDF internally
- Does NOT modify the original file

---

#  Project Structure

