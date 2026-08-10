import re
import os
import sys
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from datetime import datetime

BULAN = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]

TEMPLATE_NAME = "report.html"
PAGE_SIZE = 4  # records per page

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

def build_pages(df: pd.DataFrame, fields: list) -> list:
    """Group normalized records into chunks of PAGE_SIZE for pagination."""
    records = [normalize_record(row, fields) for _, row in df.iterrows()]
    pages = [records[i:i + PAGE_SIZE] for i in range(0, len(records), PAGE_SIZE)]
    return pages


def render_html(pages: list) -> str:
    """Render the Jinja2 template with the paginated records."""
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
    template = env.get_template(TEMPLATE_NAME)

    now = datetime.now()
    month = BULAN[now.month - 1]
    year = now.year

    # field_defs = [{"key": f, "label": field_label(f)} for f in fields]

    return template.render(pages=pages, month=month, year=year)


def generate_pdf(input_path: str, output_path: str) -> str:
    """Full pipeline: read file -> normalize -> render -> save PDF."""
    df, fields = read_table(input_path)

    if "name" not in df.columns:
        raise ValueError(f"Missing required column: 'name'. Found columns: {', '.join(df.columns)}")

    pages = build_pages(df, fields)
    html_string = render_html(pages)
    HTML(string=html_string, base_url=TEMPLATE_DIR).write_pdf(output_path)
    return output_path


if __name__ == "__main__":
    sample = os.path.join(os.path.dirname(__file__), "sample_data.csv")
    if os.path.exists(sample):
        out = generate_pdf(sample, "test_report.pdf")
        print(f"Generated: {out}")
    else:
        print("No sample_data.csv found for manual testing.")