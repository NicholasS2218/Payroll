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

def render_chunk(pages_chunk, meta, template_dir, template_name, out_path):
    env = Environment(loader=FileSystemLoader(template_dir))
    template = env.get_template(template_name)
    html_string = template.render(pages=pages_chunk, **meta)
    HTML(string=html_string, base_url=template_dir).write_pdf(out_path)
    return out_path

BULAN = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]

LAYOUTS = {
    # "PORTRAIT 2x2": {"template": "report_2x2.html", "page_size": 4},
    # "LANDSCAPE 3x2": {"template": "report_3x2.html", "page_size": 6},
    "A4": {"template": "report_detail.html", "page_size": 2},
    "A6": {"template": "report_detail_A6.html", "page_size": 1},
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

    return {
        "nama": info("nama") or "Unknown",
        "no_nik": info("no nik"),
        "jabatan": info("jabatan"),
        "department": info("department"),
        "status": info("status"),
        "bank": info("bank"),
        "no_rekening": info("no rekening"),
        "bulan": meta["bulan"],
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

    seen, columns = {}, []
    for i in idxs:
        label = field_label(headers[i] or "Total")
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

"""Full pipeline: read file -> normalize -> render -> save PDF."""
def generate_pdf(input_path: str, output_path: str, layout: str = DEFAULT_LAYOUT) -> str:
    if layout not in LAYOUTS:
        raise ValueError(f"Unknown layout: {layout}. Choose from {list(LAYOUTS)}.")

    config = LAYOUTS[layout]
    template_name = config["template"]
    page_size = config["page_size"]

    headers, body, meta = read_table(input_path)
    pages = build_pages(body, headers, meta, page_size)
    html_string = render_html(pages, template_name, meta)
    HTML(string=html_string, base_url=TEMPLATE_DIR).write_pdf(output_path)
    return output_path

def generate_pdf_parallel(input_path: str, output_path: str, layout: str= DEFAULT_LAYOUT,
                           workers= None,  progress_callback= None, only= None) -> str:

    if layout not in LAYOUTS:
        raise ValueError(f"Unknown layout: {layout}. Choose from {list(LAYOUTS)}.")
    config = LAYOUTS[layout]
    template_name = config["template"]
    page_size = config["page_size"]

    headers, body, meta = read_table(input_path)
    pages = build_pages(body, headers, meta, page_size, only)
    if not pages:
        raise ValueError("No records found to generate.")

    workers = workers or min(os.cpu_count() or 1, len(pages))

    # not worth the process-spawn overhead for a handful of pages
    if workers <= 1 or len(pages) < workers * 2:
        html_string = render_html(pages, template_name, meta)
        HTML(string=html_string, base_url=TEMPLATE_DIR).write_pdf(output_path)
        return output_path

    target_chunks = min(len(pages), workers * 4)
    chunk_size = max(1, math.ceil(len(pages) / target_chunks))
    chunks = [pages[i:i + chunk_size] for i in range(0, len(pages), chunk_size)]

    with tempfile.TemporaryDirectory() as tmpdir:
        chunk_paths = [os.path.join(tmpdir, f"chunk_{i}.pdf") for i in range(len(chunks))]

        with ProcessPoolExecutor(max_workers=workers) as executor:
            to_index = {
                executor.submit(render_chunk, chunk, meta, TEMPLATE_DIR, template_name, path): i
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