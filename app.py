import os
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk

from generator import generate_pdf_parallel, generate_pdf, preview_data

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BRAND_COLOR = "#f4640d"
BRAND_COLOR_HOVER = "#d3560b"
TEXT_COLOR = "#FFFFFF"

class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Payroll Report Generator")
        self.root.geometry("760x560")
        self.root.minsize(640, 480)
        # self.root.resizable(False, False)

        self.input_path = None # tk.StringVar(value="No file selected")
        self.all_columns = []
        self.all_rows = []

        self._build_ui()

    def _build_ui(self):
        # scrollable container - everything below lives inside this instead of self.root
        # self.scroll_frame = ctk.CTkScrollableFrame(self.root, fg_color="transparent")
        # self.scroll_frame.pack(fill="both", expand=True)
 
        # container = self.scroll_frame

        # uncomment the above if u want scroll
        container = self.root
        container.grid_rowconfigure(3, weight=1)  # table row expands to fill remaining space
        container.grid_columnconfigure(0, weight=1)

        # header
        header = ctk.CTkFrame(container, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=24, pady=(24, 8))
        # header.pack(fill="x", padx=24, pady=(24, 8))
        # pad = {"padx": 20, "pady": 10}

        title = ctk.CTkLabel(header, text="Payroll Report Generator", font=ctk.CTkFont(size=32, weight="bold"))
        title.pack()
        #title.pack(pady=(20, 5))

        subtitle = ctk.CTkLabel(
            header,
            text="Select a CSV or Excel file",
            font=ctk.CTkFont(size=16),
            text_color="#666666",
        )
        subtitle.pack()

        file_row = ctk.CTkFrame(container, fg_color="transparent")
        file_row.grid(row=1, column=0, sticky="ew", padx=24, pady=(8, 4))
        # file_row.pack(fill="x", padx=24, pady=(8, 4))

        self.select_btn = ctk.CTkButton(
            file_row, text="Select File", width=120, command=self.select_file, 
            fg_color=BRAND_COLOR, hover_color=BRAND_COLOR_HOVER, text_color=TEXT_COLOR
        )
        self.select_btn.pack(side="left")

        self.file_label = ctk.CTkLabel(file_row, text="No file selected", font=ctk.CTkFont(size=14), text_color="#555555")
        self.file_label.pack(side="left", padx=(12,0))

        # search bar
        search_row = ctk.CTkFrame(container, fg_color="transparent")
        search_row.grid(row=2, column=0, sticky="ew", padx=24, pady=(4, 0))
        # search_row.pack(fill="x", padx=24, pady=(4, 0))

        search_subtitle = ctk.CTkLabel(
            search_row, text="Search Employee", font=ctk.CTkFont(size=12), text_color="#555555", anchor="w"
        )
        search_subtitle.pack(fill="x", pady=(2, 0))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_changed)

        self.search_entry = ctk.CTkEntry(
            search_row, placeholder_text="Search by name...", textvariable=self.search_var,
        )
        self.search_entry.pack(fill="x")

        # preview table
        table_frame = ctk.CTkFrame(container)
        table_frame.grid(row=3, column=0, sticky="nsew", padx=24, pady=(12, 8))
        # table_frame.pack(fill="both", expand=True, padx=24, pady=(12, 8))

        self._build_table(table_frame)

        # status + generate
        footer = ctk.CTkFrame(container, fg_color="transparent")
        footer.grid(row=4, column=0, sticky="ew", padx=24, pady=(4, 24))
        # footer.pack(fill="x", padx=24, pady=(4, 24))

        self.status_label = ctk.CTkLabel(footer, text="", font=ctk.CTkFont(size=14), text_color="#007700")
        self.status_label.pack(side="left")

        self.generate_btn = ctk.CTkButton(
            footer, text="Generate Report", width=120, command=self.generate_report, state="disabled",
            fg_color=BRAND_COLOR, hover_color=BRAND_COLOR_HOVER, text_color=TEXT_COLOR
        )
        self.generate_btn.pack(side="right")

        # self.status_label = tk.Label(self.root, text="", font=("Arial", 9), fg="#007700")
        # self.status_label.pack()

    def _build_table(self, parent):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Preview.Treeview", rowheight=26, font=("Arial", 10), background="white", fieldbackground="white")
        style.configure("Preview.Treeview.Heading", font=("Arial", 10, "bold"))

        tree_container = tk.Frame(parent, bg="white")
        tree_container.pack(fill="both", expand=True, padx=1, pady=1)

        self.tree = ttk.Treeview(tree_container, style="Preview.Treeview", show="headings")

        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        self._set_placeholder()

    def _set_placeholder(self):
        self.tree["columns"] = ["info"]
        self.tree.column("info", anchor="center")
        self.tree.heading("info", text="")
        self.tree.delete(*self.tree.get_children())
        self.tree.insert("", "end", values=("Select a file to preview its data here.",))

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select CSV or Excel file",
            filetypes=[("Spreadsheet files", "*.csv *.xlsx *.xls"), ("All files", "*.*")],
        )

        if not path:
            return

        try:
            columns, rows, _ = preview_data(path)
        except Exception as e:
            messagebox.showerror("Error reading file", str(e))
            return

        self.input_path = path
        self.file_label.configure(text=os.path.basename(path))

        self.all_columns = columns
        self.all_rows = rows

        self.search_var.set("") 

        self._populate_table(columns, rows)
        self.generate_btn.configure(state="normal")
        self.status_label.configure(text=f"{len(rows)} record(s) loaded.", text_color="#007700")
        
        # if path:
        #     self.input_path.set(path)
        #     self.generate_btn.config(state=tk.NORMAL)
        #     self.status_label.config(text="")

    def _on_search_changed(self, *args):
        if not self.all_rows:
            return

        query = self.search_var.get().strip().lower()
        if not query:
            filtered = self.all_rows
        else:
            # column 0 is always "Name"
            filtered = [row for row in self.all_rows if query in str(row[0]).lower()]

        self._populate_table(self.all_columns, filtered)
        self.status_label.configure(text=f"{len(filtered)} of {len(self.all_rows)} record(s) shown.", text_color="#555555")

    def _populate_table(self, columns, rows):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns

        font = tkfont.Font(font=("Arial", 10))
        header_font = tkfont.Font(font=("Arial", 10, "bold"))

        for i, col in enumerate(columns):
            self.tree.heading(col, text=col)
            # width = longest of header or any value in that column, plus padding
            content_width = max(
                [header_font.measure(col)] + [font.measure(str(row[i])) for row in rows],
                default=80,
            )
            anchor = "w" if i == 0 else "center"
            self.tree.column(col, anchor=anchor, width=content_width + 24, minwidth=60, stretch=True)

        for row in rows:
            self.tree.insert("", "end", values=row)

    def generate_report(self):
        if not self.input_path or not os.path.isfile(self.input_path):
            messagebox.showerror("Error", "Please select a valid file first.")
            return

        default_name = os.path.splitext(os.path.basename(self.input_path))[0] + "_report.pdf"
        output_path = filedialog.asksaveasfilename(
            title="Save report as",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF files", "*.pdf")],
        )
        if not output_path:
            return

        self.generate_btn.configure(state="disabled")
        self.select_btn.configure(state="disabled")
        self.status_label.configure(text="Generating report...", text_color="#555555")
        # self.root.update_idletasks()

        # Run generation in a background thread so the UI doesn't freeze.
        thread = threading.Thread(target=self._run_generation, args=(self.input_path, output_path))
        thread.start()

    def _run_generation(self, input_path, output_path):
        try:
            def on_progress(percent):
                self.root.after(0, lambda: self.status_label.configure(
                    text=f"Progress ... {percent}%", text_color="#555555"
                ))
            generate_pdf_parallel(input_path, output_path, progress_callback=on_progress)
            self.root.after(0, self._on_success, output_path)
        except Exception as e:
            self.root.after(0, self._on_error, str(e))

    def _on_success(self, output_path):
        self.status_label.configure(text="Report generated successfully!", text_color="#007700")
        self.generate_btn.configure(state="normal")
        self.select_btn.configure(state="normal")
        if messagebox.askyesno("Success", f"Report saved to:\n{output_path}\n\nOpen it now?"):
            self._open_file(output_path)

    def _on_error(self, error_message):
        self.status_label.configure(text="Failed to generate report.", text_color="#cc0000")
        self.generate_btn.configure(state="normal")
        self.select_btn.configure(state="normal")
        messagebox.showerror("Error", f"Something went wrong:\n\n{error_message}")

    def _open_file(self, path):
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()

    root = ctk.CTk()
    app = ReportApp(root)
    root.mainloop()