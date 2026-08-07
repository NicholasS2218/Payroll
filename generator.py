import re
import os
import sys
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
from datetime import datetime

BULAN = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]
FIELDS = ["gaji_bruto", "bpjs_tk", "pph_21", "uang_makan", "seragam"]

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
TEMPLATE_NAME = "report.html"

PAGE_SIZE = 4  # records per page

def get_base_path():
    if getattr(sys, 'frozen', False):
        # running as a PyInstaller bundle
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

TEMPLATE_DIR = os.path.join(get_base_path(), "templates")

def read_table(file_path: str) -> pd.DataFrame:
    """Read a CSV or Excel file into a DataFrame."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".csv", ".txt"):
        df = pd.read_csv(file_path, sep=None, engine="python",  encoding="utf-8-sig")
        print(df.columns.tolist())
    elif ext in (".xlsx", ".xls"):
        df = pd.read_excel(file_path)
        print(df.columns.tolist())
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    # normalize column headers: lowercase + spaces -> underscores
    df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]

    return df


def fmt(value) -> str:
    """Format a numeric value for display, showing '-' for missing/zero."""
    if pd.isna(value) or value == 0:
        return "-"
    try:
        return f"{value:,.0f}"
    except (ValueError, TypeError):
        return str(value)

def parse_currency_value(value) -> float:
    """Normalize a currency-formatted value (string or number) into a float.
    Handles both Indonesian format (1.000,00) and English format (1,000.00)."""
    if pd.isna(value):
        return 0.0

    # if it's already numeric (pandas parsed it fine), just use it directly
    if isinstance(value, (int, float)):
        return float(value)

    value = str(value).strip()
    value = re.sub(r'[^\d.,]', '', value)  # strip currency symbols, spaces, etc.

    if not value:
        return 0.0

    if re.search(r'\.\d{3},', value):
        # Indonesian format: 1.000,00
        value = value.replace('.', '')
        value = value.replace(',', '.')
    else:
        # English format: 1,000.00
        value = value.replace(',', '')

    try:
        return float(value)
    except ValueError:
        return 0.0


def normalize_record(row: pd.Series) -> dict:
    """Turn a raw row into a dict with every expected field present,
    and compute the total from the raw numeric values."""
    raw = {
        field: parse_currency_value(row[field]) if field in row else 0.0
        for field in FIELDS
    }

    total = raw["gaji_bruto"] - raw["bpjs_tk"] - raw["pph_21"] - raw["uang_makan"] - raw["seragam"]

    record = {"name": row.get("name", "Unknown")}
    for field in FIELDS:
        record[field] = fmt(raw[field])
    record["total"] = fmt(total)

    return record


def build_pages(df: pd.DataFrame) -> list:
    """Group normalized records into chunks of PAGE_SIZE for pagination."""
    records = [normalize_record(row) for _, row in df.iterrows()]
    pages = [records[i:i + PAGE_SIZE] for i in range(0, len(records), PAGE_SIZE)]
    return pages


def render_html(pages: list) -> str:
    """Render the Jinja2 template with the paginated records."""
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
    template = env.get_template(TEMPLATE_NAME)

    now = datetime.now()
    month = BULAN[now.month - 1]
    year = now.year

    return template.render(pages = pages, month = month, year = year)


def generate_pdf(input_path: str, output_path: str) -> str:
    """Full pipeline: read file -> normalize -> render -> save PDF.
    Returns the output_path on success."""
    df = read_table(input_path)

    required_cols = {"name"} | set(FIELDS)
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(
            f"Missing expected column(s): {', '.join(missing)}. "
            f"Found columns: {', '.join(df.columns)}"
        )

    pages = build_pages(df)
    html_string = render_html(pages)
    HTML(string=html_string, base_url=TEMPLATE_DIR).write_pdf(output_path)
    return output_path


if __name__ == "__main__":
    # Quick manual test with a sample CSV, if present.
    sample = os.path.join(os.path.dirname(__file__), "sample_data.csv")
    if os.path.exists(sample):
        out = generate_pdf(sample, "test_report.pdf")
        print(f"Generated: {out}")
    else:
        print("No sample_data.csv found for manual testing.")