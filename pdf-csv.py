"""
PDF Table Extractor & Converter
Extract tables from PDF files (local or URL) and export them to CSV, Excel, JSON, HTML, or Markdown.
Supports interactive mode with file/folder dialogs and CLI automation.
"""

import os
import sys
import argparse
import tempfile
import urllib.parse
from pathlib import Path
from typing import List, Optional, Tuple

# Ensure stdout/stderr handle UTF-8 safely on Windows terminals
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import pandas as pd
import requests
import camelot


def get_file_dialog() -> Optional[str]:
    """Open a native Windows file selection dialog to choose a PDF file."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        file_path = filedialog.askopenfilename(
            title="Select PDF File",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        root.destroy()
        return file_path if file_path else None
    except Exception as e:
        print(f"[!] Warning: Graphical file dialog unavailable ({e}). Using console input.")
        return None


def get_folder_dialog(default_dir: str = "") -> Optional[str]:
    """Open a native Windows folder selection dialog to choose an output folder."""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        folder_path = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=default_dir or os.getcwd(),
        )
        root.destroy()
        return folder_path if folder_path else None
    except Exception as e:
        print(f"[!] Warning: Graphical folder dialog unavailable ({e}). Using console input.")
        return None


def sanitize_input_path(path_str: str) -> str:
    """Clean user input path by stripping wrapping quotes and whitespace."""
    if not path_str:
        return ""
    cleaned = path_str.strip().strip("'\"")
    return os.path.expanduser(cleaned)


def download_pdf_from_url(url: str, target_dir: Path) -> Path:
    """
    Download a PDF file from a given URL to the target directory.
    Uses browser-like headers to avoid common 403 Forbidden errors.
    """
    parsed = urllib.parse.urlparse(url)
    if not parsed.scheme or not parsed.netloc:
        raise ValueError(f"Invalid URL: '{url}'. Please include http:// or https://")

    # Extract filename from URL or fallback
    url_filename = os.path.basename(parsed.path)
    if not url_filename or not url_filename.lower().endswith(".pdf"):
        url_filename = f"downloaded_{int(pd.Timestamp.now().timestamp())}.pdf"

    target_path = target_dir / url_filename

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        )
    }

    print(f"\n[+] Downloading PDF from: {url}")
    response = requests.get(url, headers=headers, stream=True, timeout=30)
    response.raise_for_status()

    total_size = int(response.headers.get("content-length", 0))
    downloaded = 0
    chunk_size = 8192

    with open(target_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=chunk_size):
            if chunk:
                f.write(chunk)
                downloaded += len(chunk)
                if total_size > 0:
                    percent = (downloaded / total_size) * 100
                    print(f"\r    Downloading: {downloaded / 1024:.1f} KB / {total_size / 1024:.1f} KB ({percent:.1f}%)", end="")
                else:
                    print(f"\r    Downloaded: {downloaded / 1024:.1f} KB", end="")
    print()

    # Validate PDF magic header
    with open(target_path, "rb") as f:
        header = f.read(5)
        if not header.startswith(b"%PDF-"):
            print("[!] Warning: Downloaded file may not be a valid PDF document (missing %PDF- header).")

    print(f"[OK] Saved downloaded file to: {target_path}")
    return target_path


def clean_dataframe(df: pd.DataFrame, promote_header: bool = True) -> pd.DataFrame:
    """
    Clean DataFrame extracted by Camelot:
    - Strips whitespace from cell strings
    - Optionally promotes row 0 to table columns if appropriate
    """
    df_cleaned = df.copy()
    # Strip whitespace across all string columns
    df_cleaned = df_cleaned.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

    if promote_header and len(df_cleaned) > 1:
        # Use first row as header
        new_header = df_cleaned.iloc[0]
        df_cleaned = df_cleaned.iloc[1:].copy()
        
        # Ensure column headers are unique and non-empty
        cols = []
        counts = {}
        for i, val in enumerate(new_header):
            name = str(val).strip() if pd.notna(val) and str(val).strip() else f"Column_{i+1}"
            if name in counts:
                counts[name] += 1
                cols.append(f"{name}_{counts[name]}")
            else:
                counts[name] = 1
                cols.append(name)
        df_cleaned.columns = cols
        df_cleaned.reset_index(drop=True, inplace=True)

    return df_cleaned


def extract_tables(
    pdf_path: str,
    pages: str = "all",
    flavor: str = "auto",
    password: Optional[str] = None,
) -> Tuple[List[camelot.core.Table], str]:
    """
    Extract tables from a PDF using Camelot with automatic flavor fallback.
    Flavors:
      - 'lattice': For tables with grid lines/borders.
      - 'stream': For tables separated by whitespace (borderless).
      - 'auto': Attempts 'lattice' first; if 0 tables found, falls back to 'stream'.
    """
    pdf_path = str(pdf_path)
    kwargs = {"pages": pages}
    if password:
        kwargs["password"] = password

    if flavor in ("lattice", "stream"):
        print(f"\n[+] Extracting tables using '{flavor}' flavor (pages: {pages})...")
        tables = camelot.read_pdf(pdf_path, flavor=flavor, **kwargs)
        used_flavor = flavor
    else:  # 'auto'
        print(f"\n[+] Extracting tables using 'lattice' flavor (pages: {pages})...")
        tables = camelot.read_pdf(pdf_path, flavor="lattice", **kwargs)
        used_flavor = "lattice"
        if len(tables) == 0:
            print("[INFO] No tables found with 'lattice' (grid lines). Attempting 'stream' flavor...")
            tables = camelot.read_pdf(pdf_path, flavor="stream", **kwargs)
            used_flavor = "stream"

    return tables, used_flavor


def preview_tables(tables: List[camelot.core.Table], promote_header: bool = True, max_rows: int = 4):
    """Print an informative summary and preview of extracted tables."""
    print(f"\n{'='*70}")
    print(f" EXTRACTED TABLES SUMMARY: Found {len(tables)} table(s)")
    print(f"{'='*70}")

    for idx, table in enumerate(tables, start=1):
        report = table.parsing_report
        accuracy = report.get("accuracy", "N/A")
        whitespace = report.get("whitespace", "N/A")
        page = report.get("page", table.page)
        shape = table.df.shape

        print(f"\n--- [Table {idx}] Page: {page} | Dimensions: {shape[0]} rows x {shape[1]} cols | Accuracy: {accuracy}% | Whitespace: {whitespace}% ---")
        
        df_preview = clean_dataframe(table.df, promote_header=promote_header)
        preview_text = df_preview.head(max_rows).to_string(index=False)
        print(preview_text)
        if len(df_preview) > max_rows:
            print(f"... and {len(df_preview) - max_rows} more row(s)")
    print(f"{'='*70}\n")


def export_tables(
    tables: List[camelot.core.Table],
    output_dir: Path,
    base_name: str,
    formats: List[str],
    promote_header: bool = True,
    merge_tables: bool = False,
) -> List[Path]:
    """
    Export extracted tables to requested formats (csv, excel, json, html, markdown).
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    exported_files: List[Path] = []
    
    cleaned_dfs = [clean_dataframe(table.df, promote_header=promote_header) for table in tables]
    
    # 1. Export Excel (.xlsx) - Single workbook with sheets per table + optional merged sheet
    if "excel" in formats or "xlsx" in formats:
        excel_path = output_dir / f"{base_name}_tables.xlsx"
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            for idx, (table, df) in enumerate(zip(tables, cleaned_dfs), start=1):
                sheet_title = f"Table_{idx}_p{table.page}"[:31]
                df.to_excel(writer, sheet_name=sheet_title, index=False)
            
            if merge_tables and len(cleaned_dfs) > 1:
                try:
                    merged_df = pd.concat(cleaned_dfs, ignore_index=True)
                    merged_df.to_excel(writer, sheet_name="All_Tables_Merged"[:31], index=False)
                except Exception as e:
                    print(f"[!] Warning: Could not merge tables into Excel sheet ({e})")
        
        exported_files.append(excel_path)
        print(f"[OK] Created Excel workbook: {excel_path} ({len(tables)} sheets)")

    # 2. Export CSV (.csv)
    if "csv" in formats:
        for idx, (table, df) in enumerate(zip(tables, cleaned_dfs), start=1):
            csv_path = output_dir / f"{base_name}_table_{idx}_page_{table.page}.csv"
            df.to_csv(csv_path, index=False, encoding="utf-8-sig")
            exported_files.append(csv_path)
            print(f"[OK] Created CSV: {csv_path}")
        
        if merge_tables and len(cleaned_dfs) > 1:
            try:
                merged_df = pd.concat(cleaned_dfs, ignore_index=True)
                merged_csv_path = output_dir / f"{base_name}_tables_merged.csv"
                merged_df.to_csv(merged_csv_path, index=False, encoding="utf-8-sig")
                exported_files.append(merged_csv_path)
                print(f"[OK] Created Merged CSV: {merged_csv_path}")
            except Exception as e:
                print(f"[!] Warning: Could not create merged CSV ({e})")

    # 3. Export JSON (.json)
    if "json" in formats:
        for idx, (table, df) in enumerate(zip(tables, cleaned_dfs), start=1):
            json_path = output_dir / f"{base_name}_table_{idx}_page_{table.page}.json"
            df.to_json(json_path, orient="records", indent=2, force_ascii=False)
            exported_files.append(json_path)
            print(f"[OK] Created JSON: {json_path}")

        if merge_tables and len(cleaned_dfs) > 1:
            try:
                merged_df = pd.concat(cleaned_dfs, ignore_index=True)
                merged_json_path = output_dir / f"{base_name}_tables_merged.json"
                merged_df.to_json(merged_json_path, orient="records", indent=2, force_ascii=False)
                exported_files.append(merged_json_path)
                print(f"[OK] Created Merged JSON: {merged_json_path}")
            except Exception as e:
                print(f"[!] Warning: Could not create merged JSON ({e})")

    # 4. Export HTML (.html)
    if "html" in formats:
        html_path = output_dir / f"{base_name}_tables.html"
        css_style = """
        <style>
            body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; margin: 30px; background: #f8fafc; color: #1e293b; }
            h1 { color: #0f172a; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; }
            h2 { color: #334155; margin-top: 30px; }
            table { border-collapse: collapse; width: 100%; margin: 15px 0 30px 0; background: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
            th { background-color: #3b82f6; color: white; text-align: left; padding: 12px 16px; font-weight: 600; }
            td { padding: 10px 16px; border-bottom: 1px solid #f1f5f9; }
            tr:nth-child(even) { background-color: #f8fafc; }
            tr:hover { background-color: #f1f5f9; }
            .meta { font-size: 0.875rem; color: #64748b; margin-bottom: 8px; }
        </style>
        """
        html_content = [f"<!DOCTYPE html><html><head><meta charset='utf-8'><title>{base_name} Tables</title>{css_style}</head><body>"]
        html_content.append(f"<h1>Extracted Tables - {base_name}</h1>")
        
        for idx, (table, df) in enumerate(zip(tables, cleaned_dfs), start=1):
            html_content.append(f"<h2>Table {idx} (Page {table.page})</h2>")
            html_content.append(f"<div class='meta'>Dimensions: {df.shape[0]} rows x {df.shape[1]} columns | Accuracy: {table.parsing_report.get('accuracy', 'N/A')}%</div>")
            html_content.append(df.to_html(index=False, classes="table"))

        html_content.append("</body></html>")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write("\n".join(html_content))
        exported_files.append(html_path)
        print(f"[OK] Created HTML file: {html_path}")

    # 5. Export Markdown (.md)
    if "markdown" in formats or "md" in formats:
        md_path = output_dir / f"{base_name}_tables.md"
        md_lines = [f"# Extracted Tables: {base_name}\n"]
        for idx, (table, df) in enumerate(zip(tables, cleaned_dfs), start=1):
            md_lines.append(f"## Table {idx} (Page {table.page})")
            md_lines.append(f"*Accuracy: {table.parsing_report.get('accuracy', 'N/A')}% | Whitespace: {table.parsing_report.get('whitespace', 'N/A')}%*\n")
            try:
                md_lines.append(df.to_markdown(index=False) + "\n\n")
            except Exception:
                md_lines.append(df.to_string(index=False) + "\n\n")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write("\n".join(md_lines))
        exported_files.append(md_path)
        print(f"[OK] Created Markdown file: {md_path}")

    return exported_files


def run_interactive():
    """Interactive console prompt with dialogs and rich configuration."""
    print("=" * 65)
    print("        PDF Table Extractor & Converter (Interactive Mode)")
    print("=" * 65)

    # 1. Source Selection
    pdf_path = None
    downloaded_temp_path = None

    while not pdf_path:
        print("\nChoose PDF source:")
        print("  [1] Browse PC (Open Windows File Picker)")
        print("  [2] Enter local file path manually (or drag & drop)")
        print("  [3] Download from URL (http/https)")
        print("  [q] Quit")
        choice = input("Enter choice (1/2/3/q) [1]: ").strip().lower()

        if choice in ("", "1"):
            print("Opening file dialog...")
            selected = get_file_dialog()
            if selected and os.path.isfile(selected):
                pdf_path = selected
            else:
                print("[!] No file selected or file dialog cancelled.")
        elif choice == "2":
            manual_input = sanitize_input_path(input("Enter or drag & drop PDF file path: "))
            if manual_input and os.path.isfile(manual_input):
                pdf_path = manual_input
            else:
                print(f"[!] File not found: '{manual_input}'. Please try again.")
        elif choice == "3":
            url_input = input("Enter PDF URL: ").strip()
            if url_input:
                try:
                    temp_dir = Path(tempfile.gettempdir()) / "pdf_tables_download"
                    temp_dir.mkdir(parents=True, exist_ok=True)
                    pdf_path = str(download_pdf_from_url(url_input, temp_dir))
                    downloaded_temp_path = pdf_path
                except Exception as e:
                    print(f"[!] Failed to download PDF: {e}")
        elif choice == "q":
            print("Exiting.")
            sys.exit(0)
        else:
            print("Invalid option. Please enter 1, 2, 3, or q.")

    base_default_name = Path(pdf_path).stem

    # 2. Output Directory Selection
    print("\nChoose Output Folder:")
    default_out = str(Path.cwd() / "output")
    print(f"  [1] Browse PC (Open Folder Picker)")
    print(f"  [2] Use default folder: '{default_out}'")
    print("  [3] Enter custom folder path manually")
    out_choice = input("Enter choice (1/2/3) [1]: ").strip()

    output_dir_str = None
    if out_choice in ("", "1"):
        print("Opening folder dialog...")
        selected_folder = get_folder_dialog(default_out)
        if selected_folder:
            output_dir_str = selected_folder
        else:
            print(f"No folder selected. Falling back to default: '{default_out}'")
            output_dir_str = default_out
    elif out_choice == "2":
        output_dir_str = default_out
    elif out_choice == "3":
        custom_out = sanitize_input_path(input("Enter output folder path: "))
        output_dir_str = custom_out if custom_out else default_out
    else:
        output_dir_str = default_out

    output_dir = Path(output_dir_str)

    # 3. Base Filename Prefix
    custom_prefix = input(f"\nEnter output filename prefix [{base_default_name}]: ").strip()
    base_name = custom_prefix if custom_prefix else base_default_name

    # 4. Extraction Pages
    pages_input = input("\nEnter pages to extract (e.g. 'all', '1', '1,2', '1-5') [all]: ").strip()
    pages = pages_input if pages_input else "all"

    # 5. Extraction Flavor
    print("\nChoose extraction flavor:")
    print("  [1] Auto-detect (lattice first, then stream fallback) [Recommended]")
    print("  [2] Lattice (best for tables with clear grid lines / borders)")
    print("  [3] Stream (best for tables without borders, aligned by whitespace)")
    flavor_choice = input("Enter choice (1/2/3) [1]: ").strip()
    flavor_map = {"1": "auto", "2": "lattice", "3": "stream", "": "auto"}
    flavor = flavor_map.get(flavor_choice, "auto")

    # 6. Password Protected?
    password = None
    is_encrypted = input("\nIs the PDF password protected? (y/N) [n]: ").strip().lower()
    if is_encrypted == "y":
        password = input("Enter PDF password: ").strip()

    # 7. Output Formats
    print("\nSelect output formats (separate multiple with commas):")
    print("  [1] CSV (.csv) [Default]")
    print("  [2] Excel (.xlsx - workbook with separate sheet per table)")
    print("  [3] JSON (.json)")
    print("  [4] HTML (.html - clean responsive web page)")
    print("  [5] Markdown (.md)")
    print("  [a] All of the above")
    format_choice = input("Enter choice (e.g., 1, 2 or 'a') [1]: ").strip().lower()
    
    fmt_dict = {"1": "csv", "2": "excel", "3": "json", "4": "html", "5": "markdown"}
    selected_formats = []
    if format_choice in ("a", "all"):
        selected_formats = ["csv", "excel", "json", "html", "markdown"]
    elif not format_choice:
        selected_formats = ["csv"]
    else:
        for token in format_choice.split(","):
            token = token.strip()
            if token in fmt_dict:
                selected_formats.append(fmt_dict[token])
            elif token in ("csv", "excel", "json", "html", "markdown", "xlsx", "md"):
                selected_formats.append(token)
    if not selected_formats:
        selected_formats = ["csv"]

    # 8. Promote first row to header?
    promote_choice = input("\nTreat the first row of tables as column headers? (Y/n) [Y]: ").strip().lower()
    promote_header = promote_choice != "n"

    # 9. Merge tables option
    merge_choice = input("Also generate a merged/combined file for all extracted tables? (y/N) [N]: ").strip().lower()
    merge_tables = merge_choice == "y"

    # Extraction Execution
    try:
        tables, used_flavor = extract_tables(
            pdf_path=pdf_path,
            pages=pages,
            flavor=flavor,
            password=password,
        )

        if len(tables) == 0:
            print("\n[!] No tables were detected in the specified pages.")
            print("Tip: If the PDF has borderless tables, try running again with 'stream' flavor.")
            return

        # Preview
        preview_tables(tables, promote_header=promote_header)

        # Export
        print(f"\n[+] Exporting tables to {output_dir}...")
        exported = export_tables(
            tables=tables,
            output_dir=output_dir,
            base_name=base_name,
            formats=selected_formats,
            promote_header=promote_header,
            merge_tables=merge_tables,
        )

        print("\n" + "=" * 65)
        print(f" [OK] Extraction successfully completed! ({len(tables)} tables extracted)")
        print(f" Output directory: {output_dir.resolve()}")
        print("=" * 65)

        # 10. Open output folder in Windows Explorer?
        open_folder = input("\nOpen output folder in Windows Explorer? (Y/n) [Y]: ").strip().lower()
        if open_folder != "n":
            try:
                os.startfile(output_dir)
            except Exception as e:
                print(f"[!] Could not open folder: {e}")

    except Exception as e:
        print(f"\n[!] Error during extraction: {e}")
        import traceback
        traceback.print_exc()

    finally:
        # Cleanup temporary downloaded file if requested
        if downloaded_temp_path and os.path.exists(downloaded_temp_path):
            cleanup = input(f"\nDelete temporary downloaded PDF ({os.path.basename(downloaded_temp_path)})? (y/N) [N]: ").strip().lower()
            if cleanup == "y":
                try:
                    os.remove(downloaded_temp_path)
                    print("[OK] Temporary download removed.")
                except Exception as e:
                    print(f"[!] Could not delete temp file: {e}")


def run_cli(args: argparse.Namespace):
    """Run in non-interactive CLI mode using parsed command-line arguments."""
    pdf_path = None
    downloaded_temp = None

    if args.url:
        temp_dir = Path(tempfile.gettempdir()) / "pdf_tables_download"
        temp_dir.mkdir(parents=True, exist_ok=True)
        pdf_path = str(download_pdf_from_url(args.url, temp_dir))
        downloaded_temp = pdf_path
    elif args.input:
        cleaned = sanitize_input_path(args.input)
        if not os.path.isfile(cleaned):
            print(f"[!] Error: Input file '{cleaned}' does not exist.")
            sys.exit(1)
        pdf_path = cleaned
    else:
        print("[!] Error: Either --input <path> or --url <url> is required in CLI mode.")
        sys.exit(1)

    output_dir = Path(args.output) if args.output else Path.cwd() / "output"
    base_name = args.prefix if args.prefix else Path(pdf_path).stem
    formats = [f.strip().lower() for f in args.format.split(",")]

    try:
        tables, used_flavor = extract_tables(
            pdf_path=pdf_path,
            pages=args.pages,
            flavor=args.flavor,
            password=args.password,
        )

        if len(tables) == 0:
            print("[!] No tables detected.")
            sys.exit(0)

        if not args.no_preview:
            preview_tables(tables, promote_header=not args.no_headers)

        export_tables(
            tables=tables,
            output_dir=output_dir,
            base_name=base_name,
            formats=formats,
            promote_header=not args.no_headers,
            merge_tables=args.merge,
        )
        print(f"\n[OK] Done. Tables saved to: {output_dir.resolve()}")

    finally:
        if downloaded_temp and args.clean_download:
            try:
                os.remove(downloaded_temp)
            except Exception:
                pass


def main():
    parser = argparse.ArgumentParser(
        description="PDF Table Extractor & Converter - Extract tables from PDF (file or URL) to CSV, Excel, JSON, HTML, Markdown.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Interactive mode (dialogs + prompts):
  python pdf-csv.py

  # Extract local PDF to CSV and Excel:
  python pdf-csv.py -i document.pdf -o ./results -f csv,excel

  # Download from URL and export all pages using 'stream' flavor:
  python pdf-csv.py -u https://example.com/data.pdf --flavor stream --pages all -f excel --merge
        """
    )

    parser.add_argument("-i", "--input", help="Path to local PDF file")
    parser.add_argument("-u", "--url", help="URL of PDF file to download and extract")
    parser.add_argument("-o", "--output", help="Directory where extracted tables will be saved (default: ./output)")
    parser.add_argument("-p", "--prefix", help="Output filename prefix (default: PDF filename stem)")
    parser.add_argument("-f", "--format", default="csv", help="Comma-separated export formats: csv, excel, json, html, markdown (default: csv)")
    parser.add_argument("--pages", default="all", help="Pages to extract: 'all', '1', '1,2', '1-5' (default: all)")
    parser.add_argument("--flavor", default="auto", choices=["auto", "lattice", "stream"], help="Extraction flavor (default: auto)")
    parser.add_argument("--password", help="PDF password if protected")
    parser.add_argument("--merge", action="store_true", help="Merge all extracted tables into a combined file")
    parser.add_argument("--no-headers", action="store_true", help="Do not promote first row to column headers")
    parser.add_argument("--no-preview", action="store_true", help="Skip printing table preview in terminal")
    parser.add_argument("--clean-download", action="store_true", help="Remove downloaded PDF after processing")

    # If no command-line arguments are provided, launch interactive mode
    if len(sys.argv) == 1:
        run_interactive()
    else:
        args = parser.parse_args()
        run_cli(args)


if __name__ == "__main__":
    main()
