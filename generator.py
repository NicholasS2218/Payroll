import re
import os
import math
import sys
import pandas as pd
import tempfile 
from jinja2 import Environment, FileSystemLoader
from datetime import datetime

BULAN = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]

LAYOUTS = {
    "A4": {"template": "report_detail.html", "page_size": 2},
    # "A6": {"template": "report_detail_A6.html", "page_size": 1},
}
DEFAULT_LAYOUT = "A4"

def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

TEMPLATE_DIR = os.path.join(get_base_path(), "templates")


def read_table(file_path: str):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".csv", ".txt"):
        raw = pd.read_csv(file_path, sep=None, engine="python", encoding="utf-8-sig", header=None, dtype=str)
    elif ext in (".xlsx", ".xls"):
        raw = pd.read_excel(file_path, header=None, dtype=str)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    meta = {"company": "", "bulan": ""}
    header_idx = None
    for i in range(len(raw)):
        row = raw.iloc[i]
        first = str(row.iloc[0]).strip().lower()
        if first == "perusahaan":
            meta["company"] = first_value(row.iloc[1:])
        elif first == "bulan":
            val = first_value(row.iloc[1:])
            try:
                d = pd.to_datetime(val)
                meta["bulan"] = f"{BULAN[d.month - 1]} {d.year}"
            except (ValueError, TypeError):
                meta["bulan"] = val.upper()
        elif first == "nama":
            header_idx = i
            break
    if header_idx is None:
        raise ValueError("Could not find column 'Nama'")

    headers = ["" if pd.isna(h) else " ".join(str(h).split()) for h in raw.iloc[header_idx]]
    body = raw.iloc[header_idx + 1:].reset_index(drop=True)
    body.columns = range(len(headers))  # positional, so duplicate headers are fine
    body = body[body[0].notna() & (body[0].str.strip() != "")].reset_index(drop=True)

    if not meta["bulan"]:
        now = datetime.now()
        meta["bulan"] = f"{BULAN[now.month - 1]} {now.year}"

    # # normalize column headers: lowercase + spaces -> underscores
    # df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]

    # # everything except 'name' is a value field, in file order
    # fields = [col for col in df.columns if col != "name"]

    return headers, body, meta


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

UPPER_LABELS = {"total penghasilan", "total pengurang", "gaji bersih"}

def field_label(field: str) -> str:
    # return field.replace("_", " ").upper()
    label = field.replace("_", " ")
    return label.upper() if label.lower() in UPPER_LABELS else label

def first_value(cells) -> str:
    for c in cells:
        if pd.notna(c) and str(c).strip():
            return str(c).strip()
    return ""

def terbilang_value(n: int) -> str:
    satuan = ["", "satu", "dua", "tiga", "empat", "lima", "enam",
              "tujuh", "delapan", "sembilan", "sepuluh", "sebelas"]
    if n < 12:
        return satuan[n]
    if n < 20:
        return f"{terbilang_value(n - 10)} belas"
    if n < 100:
        return f"{terbilang_value(n // 10)} puluh {terbilang_value(n % 10)}".strip()
    if n < 200:
        return f"seratus {terbilang_value(n - 100)}".strip()
    if n < 1000:
        return f"{terbilang_value(n // 100)} ratus {terbilang_value(n % 100)}".strip()
    if n < 2000:
        return f"seribu {terbilang_value(n - 1000)}".strip()
    if n < 1_000_000:
        return f"{terbilang_value(n // 1000)} ribu {terbilang_value(n % 1000)}".strip()
    if n < 1_000_000_000:
        return f"{terbilang_value(n // 1_000_000)} juta {terbilang_value(n % 1_000_000)}".strip()
    if n < 1_000_000_000_000:
        return f"{terbilang_value(n // 1_000_000_000)} miliar {terbilang_value(n % 1_000_000_000)}".strip()
    return f"{terbilang_value(n // 1_000_000_000_000)} triliun {terbilang_value(n % 1_000_000_000_000)}".strip()


def terbilang(value) -> str:
    n = int(round(value))
    words = terbilang_value(n) if n > 0 else "nol"
    return f"{words.capitalize()} rupiah"

def get_sections(headers: list) -> dict:
    """Split columns into earnings / deductions / summary using marker headers."""
    low = [h.lower() for h in headers]

    def pos(name):
        if name not in low:
            raise ValueError(f"Missing required column: '{name}'. "
                             f"Found columns: {', '.join(h for h in headers if h)}")
        return low.index(name)

    gaji, tp, tpg = pos("gaji"), pos("total penghasilan"), pos("total pengurang")
    return {
        "earn": range(gaji, tp + 1),
        "deduct": range(tp + 1, tpg + 1),
        "summary": range(tpg + 1, len(headers)),
    }


def build_rows(row, headers, idxs): # , negative_after=None):
    out = []
    for i in idxs:
        label = headers[i]
        if not label:
            continue
        value = parse_currency_value(row[i])
        is_total = label.lower().startswith("total")
        # if value == 0 and not is_total:
        #     continue
        out.append({
            "label": field_label(label),
            "value": fmt(value),
            "total": is_total,
            "negative": False,
            # "negative": negative_after is not None and i > negative_after,
        })
    return out


def normalize_record(row: pd.Series, headers: list, sections: dict, meta: dict) -> dict:
    low = [h.lower() for h in headers]

    def info(name):
        if name not in low:
            return ""
        v = row[low.index(name)]
        return "" if pd.isna(v) else str(v).strip()

    summary_idxs = list(sections["summary"])

    # base = "gaji bersih" if present, otherwise the first summary column
    base = next((i for i in summary_idxs if low[i] == "gaji bersih"),
                summary_idxs[0] if summary_idxs else None)

    summary = build_rows(row, headers, summary_idxs)
    # summary = build_rows(row, headers, summary_idxs, negative_after=base)

    if base is None:
        transfer = 0.0
    else:
        deductions = sum(parse_currency_value(row[i])
                         for i in summary_idxs if i > base)
        transfer = parse_currency_value(row[base]) - deductions

    summary.append({
        "label": "Nilai Dibayar",
        "value": fmt(transfer),
        "total": True,
        "negative": False,
    })

    # every column before "Gaji" is an info field, in file order
    info_items = [
        {"label": headers[i],
         "value": "" if pd.isna(row[i]) else str(row[i]).strip()}
        for i in range(sections["earn"][0]) if headers[i]
    ]
    half = (len(info_items) + 2) // 2
    info_left = info_items[:half]
    info_right = [{"label": "Bulan", "value": meta["bulan"]}] + info_items[half:]

    return {
        "nama": info("nama") or "Unknown",
        "info_left": info_left,
        "info_right": info_right,
        "earnings": build_rows(row, headers, sections["earn"]),
        "deductions": build_rows(row, headers, sections["deduct"]),
        "summary": summary,
        "terbilang": terbilang(transfer),
    }

def preview_data(input_path: str):
    headers, body, meta = read_table(input_path)
    sections = get_sections(headers)
    first_money = sections["earn"][0]

    last = len(headers) - 1
    idxs = [i for i in range(len(headers)) if headers[i] or i == last]

    # names that exist in both the earnings and deductions sections
    earn_names = {field_label(headers[i]).lower() for i in sections["earn"] if headers[i]}
    deduct_names = {field_label(headers[i]).lower() for i in sections["deduct"] if headers[i]}
    both = earn_names & deduct_names

    seen, columns = {}, []
    for i in idxs:
        label = field_label(headers[i] or "Total")
        if label.lower() in both:
            if i in sections["earn"]:
                label += " (+)"
            elif i in sections["deduct"]:
                label += " (-)"
        seen[label] = seen.get(label, 0) + 1
        columns.append(label if seen[label] == 1 else f"{label} ({seen[label]})")

    def cell(row, i):
        if i < first_money:
            v = row[i]
            return "" if pd.isna(v) else str(v).strip()
        return fmt(parse_currency_value(row[i]))

    rows = [[cell(row, i) for i in idxs] for _, row in body.iterrows()]

    return columns, rows, meta

def build_pages(body, headers, meta, page_size, only=None):
    if only is not None:
        body = body.iloc[sorted(only)]

    sections = get_sections(headers)
    records = [normalize_record(row, headers, sections, meta) for _, row in body.iterrows()]
    return [records[i:i + page_size] for i in range(0, len(records), page_size)]


def render_html(pages: list, template_name: str, meta: dict) -> str:
    env = Environment(loader=FileSystemLoader(TEMPLATE_DIR))
    template = env.get_template(template_name)

    # now = datetime.now()
    # month = BULAN[now.month - 1]
    # year = now.year

    # field_defs = [{"key": f, "label": field_label(f)} for f in fields]

    return template.render(pages=pages, **meta)

PAPER_SIZES = {                 
    "A4": ("A4 portrait", 210), 
    "A6": ("A6 landscape", 148),
}

PRINT_SNIPPET = """
<style>
@page { size: __PAPER__; }
@media screen {
    body { background: #888; }
    .page { background: #fff; width: __WIDTH__mm; margin: 10px auto; }
}
</style>
<script>
window.addEventListener('load', () => setTimeout(() => window.print(), 300));
</script>
"""

def generate_html(input_path: str, layout: str = DEFAULT_LAYOUT, only=None) -> str:
    """Render the slips as one HTML page that opens the print dialog when loaded."""
    if layout not in LAYOUTS:
        raise ValueError(f"Unknown layout: {layout}. Choose from {list(LAYOUTS)}.")
    config = LAYOUTS[layout]

    headers, body, meta = read_table(input_path)
    pages = build_pages(body, headers, meta, config["page_size"], only)
    if not pages:
        raise ValueError("No records found to print.")

    paper, width = PAPER_SIZES.get(layout, PAPER_SIZES[DEFAULT_LAYOUT])
    snippet = PRINT_SNIPPET.replace("__PAPER__", paper).replace("__WIDTH__", str(width))

    html = render_html(pages, config["template"], meta)
    return html.replace("</body>", snippet + "</body>")

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    sample = os.path.join(os.path.dirname(__file__), "sample_data.csv")
    if os.path.exists(sample):
        out = generate_pdf_parallel(sample, "test_report.pdf")
        print(f"Generated: {out}")
    else:
        print("No sample_data.csv found for manual testing.")