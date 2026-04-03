# PDF to Excel Converter

Convert PDF bank statements to Excel spreadsheets with OCR support.

## Features

- **EasyOCR** for PDFs with broken/embedded fonts (auto-detection)
- **Drag & drop** upload with live PDF preview
- **Exclusion zones** — draw rectangles on the preview to ignore headers, logos, footers
- **Page range extension** — apply zones to following or previous pages
- **Cancel** long conversions
- **FR/EN** interface with language switcher
- **Auto column detection** for French bank statements (Caisse d'Epargne, etc.)

## Quick Start

```bash
pip install flask openpyxl pymupdf Pillow easyocr numpy
python pdfxlsx.py
# Open http://localhost:5000
```

## Requirements

- Python 3.9+
- Flask, openpyxl, PyMuPDF, Pillow
- EasyOCR + NumPy (for OCR mode)

## How it works

1. Upload a PDF
2. Preview pages and optionally mark zones to exclude
3. Click Convert — EasyOCR reads each page
4. Download the Excel file with columns: Date, Label, Value date, Debit, Credit
