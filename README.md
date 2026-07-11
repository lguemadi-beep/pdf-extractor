# PDF Data Extractor

Extracts data from good-quality (text-based, not scanned) PDF files and
produces a professionally formatted Excel (.xlsx) report. Works as a
desktop app, a one-off command-line tool, or an always-on background
service that watches a folder for new PDFs.

## What it extracts, per PDF

**"Facture Proforma" invoices (LAHLAH KARIM template)** are recognized
automatically and parsed field-by-field: invoice number, date, demandeur,
client name/NIF/address, company info, every line item (date, désignation,
qté, P.U HT, montant HT), sous-total HT, TVA %, montant TVA, and total TTC.

**Any other PDF** falls back to a generic extraction: every table found
on every page, "Label: Value" style fields, and any dates/amounts found
in the text — so the tool still works on other documents, just less
precisely tailored.

## Excel output

One timestamped file, e.g. `PDF_Extraction_Report_20260711_194421.xlsx`.

If one or more "Facture Proforma" invoices are found, a single `Factures`
sheet lists one row per prestation line, in this exact column order:
**Fournisseur, Demandeur, Numéro facture, Date de facture, Description,
Montant HT** — with a grand-total row at the bottom.

Any PDFs that *don't* match the invoice template (if mixed into the same
folder) still get a generic report alongside it:

| Sheet | Contents |
|---|---|
| `Summary` | One row per PDF: pages, #tables, #fields, status |
| `Fields` | Every key/value pair found, across all PDFs |
| `Tables` | Every table, stacked long-format, with source file/page |
| `<file>_p#_t#` | A clean, ready-to-use copy of each individual table |

Formatting includes a styled title block, colored header row, frozen
header, banded (zebra-striped) rows, autofilter, auto-sized columns, and
Excel number formatting on all monetary columns.

## Adapting to a different invoice template

`core/facture_parser.py` is written specifically for the LAHLAH KARIM
"Facture Proforma" layout (company block on the left, N°/Date/Demandeur
on the right, a `DÉSIGNATION DES PRESTATIONS` line-item table, then
Sous-total HT / TVA / Total TTC). If you get invoices from a different
supplier with a different layout, that file's regular expressions are
the place to adjust — happy to tailor it further if you share more
samples or a different template.

## Install as a standalone Windows app (no Python needed on the target PC)

PyInstaller can only build for the OS it runs on, so the `.exe` must be
built once on a Windows machine that *does* have Python — after that,
the resulting file runs anywhere, Python or not.

**Step 1 — on one Windows PC with Python 3.10+ installed:**
1. Copy this whole `pdf_extractor_app` folder onto that PC.
2. Double-click `build_exe.bat` (or run it from a terminal).
   It installs the build tools and produces:
   - `dist\PDF_Extractor.exe` — the desktop app (double-click to run)
   - `dist\PDF_Extractor_CLI.exe` — headless/command-line version, for services

**Step 2 — deploy anywhere:**
Copy `dist\PDF_Extractor.exe` (a single file) to any other Windows PC —
no install, no Python, just double-click it. Same for
`PDF_Extractor_CLI.exe` if you want the command-line/service version.

To run it as a real always-on Windows service (auto-processes new PDFs
even after a reboot, no window open), use
`deploy\install_windows_service.bat` — it wires up
`dist\PDF_Extractor_CLI.exe` with NSSM. See that file's comments.

Don't have a Windows PC with Python handy? Install Python temporarily
just for this step from https://www.python.org/downloads/ (check "Add
python.exe to PATH"), run `build_exe.bat` once, then uninstall Python —
the `.exe` it produced no longer needs it.

## Install — Python required (all platforms)

```bash
cd pdf_extractor_app
pip install -r requirements.txt
```

Requires Python 3.10+. Tkinter (for the GUI) ships with standard Python
installers on Windows/Mac; on Linux install it with e.g.
`sudo apt install python3-tk` if it's missing.

## Run — Desktop app (recommended for everyday use)

```bash
python main.py
```

1. Click **Browse...** to select the folder containing your PDFs
   (this is mandatory — the buttons stay disabled until you pick one).
2. Optionally choose a different output folder for the Excel report.
3. Click **Extract now** for a one-off run, or **Start background
   service** to keep watching the folder and auto-extract any new PDF
   that gets added.

## Run — Command line / headless

One-off extraction:
```bash
python cli.py --folder "/path/to/pdfs"
```

Custom output folder, including sub-folders:
```bash
python cli.py --folder "/path/to/pdfs" --output "/path/to/reports" --recursive
```

Run continuously as a watch-folder service (blocks, Ctrl+C to stop):
```bash
python cli.py --folder "/path/to/pdfs" --watch
```

## Run as a real background service (Python-based, alternative to the .exe method above)

**Linux (systemd):**
1. Copy the app to e.g. `/opt/pdf_extractor_app`.
2. Edit `deploy/pdf_extractor.service` (set the real folder path and user).
3. `sudo cp deploy/pdf_extractor.service /etc/systemd/system/`
4. `sudo systemctl enable --now pdf_extractor`

**Windows (NSSM):**
1. Download [NSSM](https://nssm.cc/download).
2. Edit `PDF_FOLDER` in `deploy/install_windows_service.bat`.
3. Right-click the `.bat` → *Run as administrator*.

## Notes & limits

- Designed for **text-based / good quality PDFs**. Scanned image PDFs
  (photos of documents) contain no extractable text layer, so tables/
  fields won't be found for those — an OCR step (e.g. Tesseract) could
  be added if you need that.
- Key/value detection looks for short `Label: Value` style lines. If
  your documents use a different layout, the regex in
  `core/extractor.py` (`KEY_VALUE_PATTERN`) can be tuned.
- Logs are written to `logs/pdf_extractor.log` (rotating, 3 x 1MB).
- Existing report files are never overwritten — each run creates a new
  timestamped file.

## Project structure

```
pdf_extractor_app/
├── main.py                 # Desktop GUI (Tkinter)
├── cli.py                  # Command-line / headless entry point
├── build_exe.bat           # Builds standalone .exe (run on a Windows PC with Python)
├── core/
│   ├── extractor.py        # Generic PDF -> tables / fields / dates / amounts
│   ├── facture_parser.py   # Specialized parser for "Facture Proforma" invoices
│   ├── excel_writer.py     # Professional .xlsx report generation
│   ├── watcher.py          # Folder-watching (background service)
│   └── pipeline.py         # Ties extraction + Excel writing together
├── deploy/
│   ├── pdf_extractor.service        # systemd unit (Linux)
│   └── install_windows_service.bat  # NSSM installer (Windows, uses the .exe)
├── logs/                    # Rotating log files (created at runtime)
├── requirements.txt
├── requirements-build.txt   # requirements.txt + PyInstaller, for build_exe.bat
└── README.md
```
