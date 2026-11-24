# file-converter

Converts PDF files into formats you can edit. Runs entirely on your computer.

| Format | Output | What you get |
|---|---|---|
| Word | `.docx` | Editable document with layout, tables and images; fonts mapped to installed ones |
| Excel | `.xlsx` | One sheet per page, text and tables in page order; merged cells, cell colours and bold kept; currency, percentages and numbers become real numbers |
| PowerPoint | `.pptx` | One slide per page: editable text boxes with the original fonts and colours, native shapes and lines, images |
| Plain text | `.txt` | Text only |
| Markdown | `.md` | Headings, bold, tables |
| HTML | `.html` | Web page, editable in the browser |
| CSV | `.csv` | Tables as rows |
| Images | `.zip` | One PNG per page |

## Run it

### Web app

Double-click `run.bat`, or run it from a terminal in this folder:

```
.\run.bat
```

It starts the server and opens http://127.0.0.1:5000 in your browser. Drop in PDFs, pick a format, and click Convert. Press Ctrl+C in the terminal window to stop it.

You can also start the server directly:

```
.venv\Scripts\python app.py
```

### Command line

```
.venv\Scripts\python cli.py report.pdf --to docx xlsx
.venv\Scripts\python cli.py report.pdf --to pptx --pages 1-3,5 --out converted
.venv\Scripts\python cli.py secret.pdf --to docx --password hunter2
```

| Option | Meaning |
|---|---|
| `--to` | One or more formats: `docx`, `xlsx`, `pptx`, `txt`, `md`, `html`, `csv`, `png`. Default is `docx`. |
| `--pages` | Page range such as `1-3,5`. Default is all pages. |
| `--out` | Output folder. Default is next to the PDF. |
| `--password` | Password for a protected PDF. |

### First-time setup

`run.bat` does this for you: if `.venv` is missing, it creates it and installs `requirements.txt`. This needs Python 3 on your PATH.

To set up by hand:

```
py -3 -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
```

## Limits

- **Scanned PDFs** (photos of pages) have no text to extract. The tool warns you; run OCR on them first.
- Complex layouts will not match the PDF pixel for pixel, and fonts are substituted if not installed.
- Excel works best on tables with ruled lines; otherwise columns are guessed from spacing.
- PowerPoint: pages with gradients or more than 600 shapes get their line art as one background picture (text and images stay editable).

## Files

- `converter.py` — all conversion logic (`convert(pdf, fmt, out, pages, password)`)
- `app.py` + `templates/index.html` — web UI
- `cli.py` — command line
