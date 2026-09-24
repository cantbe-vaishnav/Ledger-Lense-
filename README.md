# LedgerLens — Annual Report Analyzer

A privacy-conscious, browser-based annual-report workspace. Source documents are read in the browser and are not uploaded by this project.

## Run it

Open `index.html` in a modern browser. For the PDF, spreadsheet, and chart libraries to load, the browser needs an internet connection on first use.

## Supported source files

- PDF annual reports with selectable text
- Excel workbooks (`.xlsx`, `.xls`)
- CSV and text files

## What it does

- Reads source text and looks for revenue, EBITDA, net income, and operating-cash-flow values
- Derives fiscal years from source content or file order
- Compares extracted measures year over year
- Highlights narrative excerpts related to growth, risk, and operational priorities
- Exports a lightweight text summary

Scanned/image-only PDFs need OCR before their text can be analyzed. Financial values are presented in the units reported by the source; the sample workspace uses INR crores.
