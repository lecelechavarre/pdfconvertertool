# PDF Converter Tool

A simple desktop application built with Python and Tkinter that converts:

- PDF → Word (.docx)
- Word (.docx) → PDF (.pdf)

The application includes a built-in preview system and preserves document formatting as much as possible during conversion.

---

## Features

- Convert PDF files to editable Word documents
- Convert Word documents to PDF
- Preview converted files before downloading
- Preserves:
  - Paragraph spacing
  - Alignment (left, center, right, justify)
  - Indentation
  - Bold, italic, and underline formatting
- Runs conversion in the background to prevent UI freezing
- Automatically installs required dependencies

---

## Requirements

- Python 3.8 or higher

The application will automatically install required packages if they are missing:

- PyMuPDF
- python-docx
- ReportLab
- pdf2docx
- Pillow

---

## How to Run

1. Clone the repository:

```bash
git clone https://github.com/lecelechavarre/pdfconvertertool.git
cd pdfconvertertool
```

2. Run the script:

```bash
python pfdconverter.py
```

## How It Works

### PDF → Word

- Uses `pdf2docx` to convert PDF files into editable Word documents.
- Generates a preview before allowing download.

### Word → PDF

- Reads the Word document using `python-docx`.
- Rebuilds the document layout using `ReportLab`.
- Preserves formatting like spacing, alignment, and inline styles.

## Project Structure
```
pdfconvertertool/  
│
├── pdfconverter.py
├── README.md
└── requirements.txt
```
