# Payroll Report Generator

Generates an A4 PDF payroll report (4 records per page, 2x2 card layout) from a CSV or Excel file. Runs fully offline as a desktop app.

---

## For Users (just running the app)

1. Install the **GTK3 runtime** (required once, needed by the PDF engine):
   https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases
   - Download the latest `gtk3-runtime-x.x.x-x-x-ts-win64.exe`
   - Run it, keep default options, finish install
2. Run `app.exe`
3. Click **Select File** → choose your CSV/Excel file
4. Click **Generate Report** → choose where to save → PDF opens automatically

No Python, no internet connection, and no other installs needed.

---

## For Developers (building the .exe from source)

### 1. Set up the environment
```bash
python -m venv venv
venv\Scripts\activate          # Windows
pip install -r requirements.txt
```
Make sure GTK3 runtime (link above) is also installed on your dev machine, or WeasyPrint won't even run locally.

### 2. Install PyInstaller
```bash
pip install pyinstaller
```

### 3. Generate a spec file
```bash
pyi-makespec --onefile --windowed app.py
```
This creates `app.spec`.

### 4. Edit `app.spec`
Inside the `Analysis(...)` block, update:
```python
datas=[('templates', 'templates')],
hiddenimports=['weasyprint', 'PIL._tkinter_finder'],
```

### 5. Build the executable
```bash
pyinstaller app.spec
```
Output will be at `dist/app.exe`.

### 6. Test before sharing
Run `dist/app.exe` directly (not from the VSCode terminal/venv) to confirm it works standalone. Ideally test on a machine other than your dev machine.

---

## Project Structure
```
payroll-report-generator/
├── app.py                    # Tkinter UI
├── generator.py              # CSV/Excel reading, normalization, PDF generation
├── templates/
│   └── report.html           # A4 2x2 card layout (Jinja2)
├── requirements.txt
├── app.spec                  # PyInstaller build config
└── README.md
```

## Notes
- Input files must follow these columns: `name`, `gaji bruto`, `bpjs tk`, `pph 21`, `uang makan`, `seragam` (case-insensitive, spaces are normalized to underscores).
- Currency values with `,`/`.` formatting (e.g. `1.000,00` or `1,000.00`) are auto-normalized.
- CSV delimiter (`,` or `;`) is auto-detected.
- Change it in the FIELDS in generator.py