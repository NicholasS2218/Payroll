import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

from generator import generate_pdf

class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Payroll Report Generator")
        self.root.geometry("480x260")
        self.root.resizable(False, False)

        self.input_path = tk.StringVar(value="No file selected")

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 20, "pady": 10}

        title = tk.Label(self.root, text="Payroll Report Generator", font=("Arial", 16, "bold"))
        title.pack(pady=(20, 5))

        subtitle = tk.Label(
            self.root,
            text="Select a CSV or Excel file to generate an A4 PDF report.",
            font=("Arial", 10),
            fg="#555",
        )
        subtitle.pack(pady=(0, 15))

        select_btn = tk.Button(self.root, text="Select File", width=20, command=self.select_file)
        select_btn.pack(**pad)

        self.file_label = tk.Label(self.root, textvariable=self.input_path, font=("Arial", 9), fg="#333")
        self.file_label.pack()

        self.generate_btn = tk.Button(
            self.root, text="Generate Report", width=20, command=self.generate_report, state=tk.DISABLED
        )
        self.generate_btn.pack(pady=15)

        self.status_label = tk.Label(self.root, text="", font=("Arial", 9), fg="#007700")
        self.status_label.pack()

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select CSV or Excel file",
            filetypes=[("Spreadsheet files", "*.csv *.xlsx *.xls"), ("All files", "*.*")],
        )
        if path:
            self.input_path.set(path)
            self.generate_btn.config(state=tk.NORMAL)
            self.status_label.config(text="")

    def generate_report(self):
        input_path = self.input_path.get()
        if not os.path.isfile(input_path):
            messagebox.showerror("Error", "Please select a valid file first.")
            return

        default_name = os.path.splitext(os.path.basename(input_path))[0] + "_report.pdf"
        output_path = filedialog.asksaveasfilename(
            title="Save report as",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output_path:
            return

        self.generate_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Generating report...", fg="#555")
        self.root.update_idletasks()

        # Run generation in a background thread so the UI doesn't freeze.
        thread = threading.Thread(target=self._run_generation, args=(input_path, output_path))
        thread.start()

    def _run_generation(self, input_path, output_path):
        try:
            generate_pdf(input_path, output_path)
            self.root.after(0, self._on_success, output_path)
        except Exception as e:
            self.root.after(0, self._on_error, str(e))

    def _on_success(self, output_path):
        self.status_label.config(text="Report generated successfully!", fg="#007700")
        self.generate_btn.config(state=tk.NORMAL)
        if messagebox.askyesno("Success", f"Report saved to:\n{output_path}\n\nOpen it now?"):
            self._open_file(output_path)

    def _on_error(self, error_message):
        self.status_label.config(text="Failed to generate report.", fg="#cc0000")
        self.generate_btn.config(state=tk.NORMAL)
        messagebox.showerror("Error", f"Something went wrong:\n\n{error_message}")

    def _open_file(self, path):
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')


if __name__ == "__main__":
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()