# PDF Table Extractor & Converter

A powerful, user-friendly Python tool to extract structured tables from PDF files (stored locally or fetched via URL) and export them into **CSV**, **Excel (.xlsx)**, **JSON**, **HTML**, and **Markdown**.

Supports both an **Interactive Mode** (with native Windows file/folder pickers) and a **CLI Mode** for script automation and scheduled workflows.

---

## Features

- **Multiple Input Sources**:
  - **Browse PC**: Opens a native Windows file selection dialog.
  - **Manual Path / Drag & Drop**: Paste or drop file paths directly into the terminal (handles quotes and spaces automatically).
  - **Download via URL**: Downloads PDFs from web links (`http://` or `https://`) with download progress reporting, browser user-agent headers, and `%PDF` validation.

- **Flexible Output Destinations**:
  - **Browse PC**: Opens a native Windows directory picker.
  - **Customizable**: Choose custom folder paths and custom output filename prefixes.
  - **Auto-Open**: Automatically reveal the exported files in Windows File Explorer upon completion.

- **Multi-Format Export**:
  - **CSV (`.csv`)**: Individual CSV per table, plus an optional merged CSV.
  - **Excel (`.xlsx`)**: Clean multi-sheet workbook where each table gets its own tab (e.g., `Table_1_p1`), plus an optional `All_Tables_Merged` sheet.
  - **JSON (`.json`)**: Structured record arrays.
  - **HTML (`.html`)**: Clean, responsive styled web page for quick viewing.
  - **Markdown (`.md`)**: GitHub-flavored markdown tables with parsing metrics.

- **Smart Table Extraction**:
  - **Auto-Fallback Engine**: Tries Camelot's `lattice` (bordered grid lines) first. If no tables are detected, it automatically attempts `stream` (borderless/whitespace-aligned tables).
  - **Page Selection**: Extract all pages (`all`), specific pages (`1`), or ranges (`1-5`, `1,3,7`).
  - **Header Promotion**: Automatically trims cell whitespace and promotes the first row to clean column headers.
  - **Table Merging**: Option to combine multiple tables across pages into a unified dataset.
  - **Password Support**: Handles encrypted or password-protected PDFs.
  - **Accuracy & Quality Preview**: Prints row/column dimensions, Camelot accuracy percentage, whitespace percentage, and a table head preview before saving.

---

## Requirements & Prerequisites

### 1. Python Dependencies

Install the required Python packages:

```bash
pip install camelot-py[cv] pandas openpyxl requests tabulate
```

> **Note**: `tkinter` comes pre-installed with standard Windows Python installations.

### 2. Ghostscript (Required by Camelot)

Camelot relies on **Ghostscript** for reading PDF raster data:
1. Download Ghostscript for Windows from [Ghostscript Official Downloads](https://www.ghostscript.com/download/gsdnld.html) (e.g., AGPL release 64-bit installer).
2. During installation, ensure the Ghostscript `bin` and `lib` folders are added to your system `PATH` (e.g., `C:\Program Files\gs\gs10.xx.x\bin`).

---

## Usage

### 1. Interactive Mode (Default)

Simply run the script with no arguments:

```bash
python pdf-csv.py
```

You will be guided through a series of prompts:
1. **Choose PDF source**: Browse PC dialog, enter path / drag-and-drop, or download via URL.
2. **Choose Output folder**: Browse folder dialog, use `./output`, or enter custom path.
3. **Set filename prefix**: Default is the original PDF filename.
4. **Select pages**: `all` or specific pages like `1-3`.
5. **Select flavor**: Auto-detect (`auto`), `lattice`, or `stream`.
6. **Password protection**: Optional password if encrypted.
7. **Select formats**: Choose one or multiple (`csv`, `excel`, `json`, `html`, `markdown`, or `all`).
8. **Header promotion & merging**: Choose whether to promote first row to column headers and whether to produce a combined table.
9. **Review & Explorer**: Preview summary and accuracy metrics, then optionally open the output folder in Windows File Explorer.

---

### 2. Command-Line (CLI) Mode

For automated pipelines, background jobs, or batch processing, supply CLI flags:

#### Basic Example (Local PDF to CSV)
```bash
python pdf-csv.py -i document.pdf -o ./output -f csv
```

#### Export to Multiple Formats (CSV + Excel + Markdown)
```bash
python pdf-csv.py -i document.pdf -o ./results -f csv,excel,markdown
```

#### Download from URL & Export to a Multi-Sheet Excel Workbook
```bash
python pdf-csv.py -u "https://example.com/financial_report.pdf" -o ./downloads -f excel --merge
```

#### Extract Borderless Tables from Specific Pages
```bash
python pdf-csv.py -i invoice.pdf --pages 1-3 --flavor stream -f json,excel
```

---

## CLI Options Reference

| Option | Flag | Description | Default |
|---|---|---|---|
| `--input` | `-i` | Path to local PDF file | None |
| `--url` | `-u` | URL of PDF to download and process | None |
| `--output` | `-o` | Destination folder path | `./output` |
| `--prefix` | `-p` | Filename prefix for exported files | PDF filename stem |
| `--format` | `-f` | Comma-separated formats: `csv`, `excel`, `json`, `html`, `markdown` | `csv` |
| `--pages` | | Pages to extract: `'all'`, `'1'`, `'1,2'`, `'1-5'` | `'all'` |
| `--flavor` | | Extraction flavor: `auto`, `lattice`, or `stream` | `auto` |
| `--password` | | Password for encrypted PDF | None |
| `--merge` | | Merge all extracted tables into a combined file/sheet | `False` |
| `--no-headers` | | Do not promote the first row to column headers | `False` |
| `--no-preview` | | Suppress table preview output in terminal | `False` |
| `--clean-download` | | Delete downloaded temporary PDF after processing | `False` |

---

## Troubleshooting

- **`ValueError: Please make sure that Ghostscript is installed`**:
  Install Ghostscript from [Ghostscript.com](https://www.ghostscript.com/download/gsdnld.html) and add its `bin` directory to your Windows system `PATH`.
- **0 tables found**:
  By default, `--flavor auto` checks `lattice` (tables with borders) and falls back to `stream` (tables with whitespace columns). If tables are still not detected, try explicitly passing `--flavor stream` or adjust page ranges.
