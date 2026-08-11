import re
import os
import math
import sys
import pandas as pd
import tempfile 
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed
from pypdf import PdfWriter

def _render_chunk(pages_chunk, month, year, template_dir, template_name, out_path):
    """Runs in a separate process: render a subset of pages into its own standalone PDF."""
    env = Environment(loader=FileSystemLoader(template_dir))
    template = env.get_template(template_name)
    html_string = template.render(pages=pages_chunk, month=month, year=year)
    HTML(string=html_string, base_url=template_dir).write_pdf(out_path)
    return out_path

BULAN = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]

# TEMPLATE_NAME = "report_3x2.html"
# PAGE_SIZE = 6  # records per page

LAYOUTS = {
    "PORTRAIT": {"template": "report_2x2.html", "page_size": 4},
    "LANDSCAPE": {"template": "report_3x2.html", "page_size": 6},
}
DEFAULT_LAYOUT = "LANDSCAPE"

def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

TEMPLATE_DIR = os.path.join(get_base_path(), "templates")


def read_table(file_path: str):
    """Read a CSV or Excel file into a DataFrame, and derive the list of
    value fields (every column except 'name') from the file itself."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".csv", ".txt"):
        df = pd.read_csv(file_path, sep=None, engine="python", encoding="utf-8-sig")
    elif ext in (".xlsx", ".xls"):
        df = pd.read_excel(file_path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    # normalize column headers: lowercase + spaces -> underscores
    df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]

    # everything except 'name' is a value field, in file order
    fields = [col for col in df.columns if col != "name"]

    return df, fields


def fmt(value) -> str:
    """Format a numeric value for display, showing '-' for missing/zero."""
    if pd.isna(value) or value == 0:
        return "-"
    try:
        s = f"{value:,.0f}"
        return s.replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(value)


def parse_currency_value(value) -> float:
    """Normalize a currency-formatted value (string or number) into a float.
    Handles both Indonesian format (1.000,00) and English format (1,000.00)."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    value = str(value).strip()
    value = re.sub(r'[^\d.,]', '', value)
    if not value:
        return 0.0

    if re.search(r'\.\d{3},', value):
        value = value.replace('.', '').replace(',', '.')
    else:
        value = value.replace(',', '')

    try:
        return float(value)
    except ValueError:
        return 0.0

LABEL_OVERRIDES = {
    "total": "GAJI YANG DITRANSFER",
}

def field_label(field: str) -> str:
    """Derive a display label from a field key, e.g. 'gaji_bruto' -> 'GAJI BRUTO'."""
    if field in LABEL_OVERRIDES:
        return LABEL_OVERRIDES[field]
    return field.replace("_", " ").upper()

DEDUCTION = {"bpjs_tk", "pph_21", "uang_makan", "seragam"} #HARDCODE, UPDATE THIS LATER

def normalize_record(row: pd.Series, fields: list) -> dict:
    """Turn a raw row into a dict with every expected field present."""
    record = {"name": row.get("name", "Unknown"), "rows": []}

    for field in fields:
        raw_value = parse_currency_value(row[field]) if field in row else 0.0
        if raw_value > 0:
            record["rows"].append({
                "key": field,
                "label": field_label(field),
                "value": fmt(raw_value),
                "negative": field in DEDUCTION
            })
        # record[field] = fmt(raw_value)

    return record

def preview_data(input_path: str):
    """Read a file and return (columns, rows, fields) for a UI preview table,
    without generating the PDF."""
    df, fields = read_table(input_path)

    if "name" not in df.columns:
        raise ValueError(f"Missing required column: 'name'. Found columns: {', '.join(df.columns)}")

    columns = ["Name"] + [field_label(f) for f in fields]

    rows = []
    for _, row in df.iterrows():
        r = [row.get("name", "Unknown")]
        for field in fields:
            raw_value = parse_currency_value(row[field]) if field in row else 0.0
            r.append(fmt(raw_value))
        rows.append(r)

    return columns, rows, fields

def build_pages(df: pd.DataFrame, fields: list, page_size: int) -> list:
    """Group normalized records into chunks of PAGE_SIZE for pagination."""
    records = [normalize_record(row, fields) for _, row in df.iterrows()]
    pages = [records[i:i + page_size] for i in range(0, len(records), page_size)]
    return pages


def render_html(pages: list, template_name: str) -> str:
    """Render the Jinja2 template with the paginated records."""
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
    template = env.get_template(template_name)

    now = datetime.now()
    month = BULAN[now.month - 1]
    year = now.year

    # field_defs = [{"key": f, "label": field_label(f)} for f in fields]

    return template.render(pages=pages, month=month, year=year)

"""Full pipeline: read file -> normalize -> render -> save PDF."""
def generate_pdf(input_path: str, output_path: str, layout: str = DEFAULT_LAYOUT) -> str:
    if layout not in LAYOUTS:
        raise ValueError(f"Unknown layout: {layout}. Choose from {list(LAYOUTS)}.")

    config = LAYOUTS[layout]
    template_name = config["template"]
    page_size = config["page_size"]

    df, fields = read_table(input_path)

    if "name" not in df.columns:
        raise ValueError(f"Missing required column: 'name'. Found columns: {', '.join(df.columns)}")

    pages = build_pages(df, fields, page_size)
    html_string = render_html(pages, template_name)
    HTML(string=html_string, base_url=TEMPLATE_DIR).write_pdf(output_path)
    return output_path

def generate_pdf_parallel(input_path: str, output_path: str, layout: str = DEFAULT_LAYOUT,
                           workers = None,  progress_callback = None) -> str:

    if layout not in LAYOUTS:
        raise ValueError(f"Unknown layout: {layout}. Choose from {list(LAYOUTS)}.")
    config = LAYOUTS[layout]
    template_name = config["template"]
    page_size = config["page_size"]

    df, fields = read_table(input_path)

    if "name" not in df.columns:
        raise ValueError(f"Missing required column: 'name'. Found columns: {', '.join(df.columns)}")

    pages = build_pages(df, fields, page_size)
    if not pages:
        raise ValueError("No records found to generate.")

    workers = workers or min(os.cpu_count() or 1, len(pages))

    # not worth the process-spawn overhead for a handful of pages
    if workers <= 1 or len(pages) < workers * 2:
        html_string = render_html(pages, template_name)
        HTML(string=html_string, base_url=TEMPLATE_DIR).write_pdf(output_path)
        return output_path

    now = datetime.now()
    month = BULAN[now.month - 1]
    year = now.year

    target_chunks = min(len(pages), workers * 4)
    chunk_size = max(1, math.ceil(len(pages) / target_chunks))
    chunks = [pages[i:i + chunk_size] for i in range(0, len(pages), chunk_size)]

    with tempfile.TemporaryDirectory() as tmpdir:
        chunk_paths = [os.path.join(tmpdir, f"chunk_{i}.pdf") for i in range(len(chunks))]

        with ProcessPoolExecutor(max_workers=workers) as executor:
            to_index = {
                executor.submit(_render_chunk, chunk, month, year, TEMPLATE_DIR, template_name, path): i
                for i, (chunk, path) in enumerate(zip(chunks, chunk_paths))
            }
            done_count = 0
            total = len(to_index)
            for t in as_completed(to_index):
                t.result()  # raises here if that chunk's render failed
                done_count += 1
                if progress_callback:
                    percent = int((done_count / total) * 100)
                    progress_callback(percent)

        merger = PdfWriter()
        for path in chunk_paths:
            merger.append(path)
        merger.write(output_path)
        merger.close()

    return output_path

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    sample = os.path.join(os.path.dirname(__file__), "sample_data.csv")
    if os.path.exists(sample):
        out = generate_pdf_parallel(sample, "test_report.pdf")
        print(f"Generated: {out}")
    else:
        print("No sample_data.csv found for manual testing.")