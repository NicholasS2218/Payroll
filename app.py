try:
    import pyi_splash
    pyi_splash.update_text("Loading...")
except ImportError:
    pyi_splash = None

import os
import sys
import threading
import tkinter as tk
import tkinter.font as tkfont
import customtkinter as ctk
import traceback
import shutil

from concurrent.futures.process import BrokenProcessPool
from tkinter import filedialog, messagebox, ttk

import tempfile
from pathlib import Path
import webbrowser
from generator import generate_pdf_parallel, generate_pdf, preview_data, generate_html, TEMPLATE_DIR

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

BRAND_COLOR = "#f4640d"
BRAND_COLOR_HOVER = "#d3560b"
TEXT_COLOR = "#FFFFFF"
LOADING_THRESHOLD = 500 

TEMPLATE_FILE = "template.xlsx"

LAYOUT_LABELS = {
    "A4 (Portrait)": "A4",
    "A6 (Landscape)": "A6",
}

class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Payroll Report Generator v.1.0")
        self.root.geometry("760x560")
        self.root.after(0, lambda: self.root.state("zoomed"))
        self.root.minsize(640, 480)
        # self.root.resizable(False, False)

        self.input_path = None # tk.StringVar(value="No file selected")
        self.all_columns = []
        self.all_rows = []

        self.selected = set()
        self.jabatan_idx = None
        self.dept_idx = None

        self._insert_job = None
        self._search_job = None

        self.build_ui()

    def build_ui(self):
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

        self.template_btn = ctk.CTkButton(
            file_row, text="Download Template", width=140, command=self.download_template,
            fg_color=BRAND_COLOR, hover_color=BRAND_COLOR_HOVER, text_color=TEXT_COLOR
        )
        self.template_btn.pack(side="right")

        # search bar
        search_row = ctk.CTkFrame(container, fg_color="transparent")
        search_row.grid(row=2, column=0, sticky="ew", padx=24, pady=(4, 0))
        # search_row.pack(fill="x", padx=24, pady=(4, 0))

        self.meta_label = ctk.CTkLabel(
            search_row, text="", font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#333333", anchor="w"
        )
        self.meta_label.pack(fill="x", pady=(0, 4))

        search_subtitle = ctk.CTkLabel(
            search_row, text="Search Employee", font=ctk.CTkFont(size=14), text_color="#555555", anchor="w"
        )
        search_subtitle.pack(fill="x", pady=(2, 0))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self.on_search_typed)

        self.search_entry = ctk.CTkEntry(
            search_row, placeholder_text="Search by name...", textvariable=self.search_var,
        )
        self.search_entry.pack(fill="x")

        filter_row = ctk.CTkFrame(search_row, fg_color="transparent")
        filter_row.pack(fill="x", pady=(8, 0))

        self.jabatan_var = tk.StringVar(value="All")
        self.dept_var = tk.StringVar(value="All")

        menu_style = dict(
            width = 200, values = ["All"], 
            fg_color = BRAND_COLOR, button_color = BRAND_COLOR_HOVER,
            button_hover_color = BRAND_COLOR_HOVER, text_color = TEXT_COLOR,
            command = lambda _: self.on_search_changed(),
        )

        ctk.CTkLabel(filter_row, text="Jabatan:", font=ctk.CTkFont(size=14),
                     text_color="#555555").pack(side="left")
        self.jabatan_menu = ctk.CTkOptionMenu(filter_row, variable=self.jabatan_var, **menu_style)
        self.jabatan_menu.pack(side="left", padx=(6, 20))

        ctk.CTkLabel(filter_row, text="Department:", font=ctk.CTkFont(size=14),
                     text_color="#555555").pack(side="left")
        self.dept_menu = ctk.CTkOptionMenu(filter_row, variable=self.dept_var, **menu_style)
        self.dept_menu.pack(side="left", padx=(6, 0))

        # preview table
        table_frame = ctk.CTkFrame(container)
        table_frame.grid(row=3, column=0, sticky="nsew", padx=24, pady=(12, 8))
        # table_frame.pack(fill="both", expand=True, padx=24, pady=(12, 8))

        self.build_table(table_frame)
        self.loading_frame = ctk.CTkFrame(table_frame, fg_color="#ffffff")
        self.loading_label = ctk.CTkLabel(self.loading_frame, text="Loading...",
                                          font=ctk.CTkFont(size=16), text_color="#555555")
        self.loading_label.place(relx=0.5, rely=0.45, anchor="center")
        self.loading_bar = ctk.CTkProgressBar(self.loading_frame, width=320, mode="indeterminate",
                                              progress_color=BRAND_COLOR)
        self.loading_bar.place(relx=0.5, rely=0.55, anchor="center")

        layout_row = ctk.CTkFrame(container, fg_color="transparent")
        layout_row.grid(row=4, column=0, sticky="ew", padx=24, pady=(8, 4))

        layout_label = ctk.CTkLabel(layout_row, text="Layout:", font=ctk.CTkFont(size=14), text_color="#555555")
        layout_label.pack(side="left", padx=(0, 8))

        self.layout_var = tk.StringVar(value="A4 (Portrait)")
        self.layout_selector = ctk.CTkSegmentedButton(
            layout_row,
            values                  = list(LAYOUT_LABELS),
            variable                = self.layout_var,
            fg_color                = "#eeeeee",
            selected_color          = BRAND_COLOR,
            selected_hover_color    = BRAND_COLOR_HOVER,
            text_color              = TEXT_COLOR
        )
        self.layout_selector.pack(side="left")
        
        # status + generate
        footer = ctk.CTkFrame(container, fg_color="transparent")
        footer.grid(row=5, column=0, sticky="ew", padx=24, pady=(4, 24))
        # footer.pack(fill="x", padx=24, pady=(4, 24))

        self.status_label = ctk.CTkLabel(footer, text="", font=ctk.CTkFont(size=14), text_color="#007700")
        self.status_label.pack(side="left")

        # self.generate_btn = ctk.CTkButton(
        #     footer, text="Generate Report", width=120, command=self.generate_report, state="disabled",
        #     fg_color=BRAND_COLOR, hover_color=BRAND_COLOR_HOVER, text_color=TEXT_COLOR
        # )
        # self.generate_btn.pack(side="right")

        self.print_btn = ctk.CTkButton(
            footer, text="Print Payroll", width=120, command=self.preview_print, state="disabled",
            fg_color=BRAND_COLOR, hover_color=BRAND_COLOR_HOVER, text_color=TEXT_COLOR
        )
        self.print_btn.pack(side="right", padx=(0, 8))

        # self.status_label = tk.Label(self.root, text="", font=("Arial", 9), fg="#007700")
        # self.status_label.pack()

    def build_table(self, parent):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Preview.Treeview", rowheight=26, font=("Arial", 10), background="white", fieldbackground="white")
        style.configure("Preview.Treeview.Heading", font=("Arial", 10, "bold"))

        tree_container = tk.Frame(parent, bg="white")
        tree_container.pack(fill="both", expand=True, padx=1, pady=1)

        self.tree = ttk.Treeview(tree_container, style="Preview.Treeview",
                                show="tree headings", selectmode="extended")
        self.tree.bind("<Button-1>", self.on_tree_click)

        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        # Shift + mouse wheel scrolls sideways
        self.tree.bind("<Shift-MouseWheel>",
                       lambda e: self.tree.xview_scroll(-1 * (e.delta // 120), "units"))

        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        self.set_placeholder()

    def show_loading(self, text="Loading..."):
        self.loading_label.configure(text=text)
        self.loading_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.loading_frame.lift()
        self.loading_bar.start()
        self.select_btn.configure(state="disabled")
        self.print_btn.configure(state="disabled")
        self.root.update_idletasks()

    def hide_loading(self):
        self.loading_bar.stop()
        self.loading_frame.place_forget()
        self.select_btn.configure(state="normal")

    def on_search_typed(self, *args):
        """Debounce: wait 250 ms after the last keystroke before filtering."""
        if self._search_job:
            self.root.after_cancel(self._search_job)
        self._search_job = self.root.after(250, self.on_search_changed)

    def load_file(self, path):
        """Runs in a background thread: only reads the file, never touches the UI."""
        try:
            result = preview_data(path)
        except Exception as e:
            self.root.after(0, self.load_failed, e)
            return
        self.root.after(0, self.load_done, path, result)

    def load_failed(self, e):
        self.hide_loading()
        self.refresh_generate_btn()
        messagebox.showerror("Error reading file", str(e))

    def on_tree_click(self, event):
        if not self.all_rows:
            return
        # only react to clicks inside the checkbox column; anywhere else just highlights the row
        if self.tree.identify_column(event.x) != "#0":
            return
        if self.tree.identify_region(event.x, event.y) not in ("tree", "cell"):
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        idx = int(iid)
        if idx in self.selected:
            self.selected.discard(idx)
            self.tree.item(iid, text="☐")
        else:
            self.selected.add(idx)
            self.tree.item(iid, text="☑")
        self.update_selection_status()
        return "break"   # don't also highlight the row when toggling the checkbox
    
    def set_visible_selection(self, state: bool):
        """Select / clear every row currently shown (respects the search filter)."""
        if not self.all_rows:
            return
        for iid in self.tree.get_children():
            idx = int(iid)
            if state:
                self.selected.add(idx)
            else:
                self.selected.discard(idx)
            self.tree.item(iid, text="☑" if state else "☐")
        self.update_selection_status()

    def refresh_generate_btn(self):
        state = "normal" if self.print_indices() else "disabled"
        # self.generate_btn.configure(state=state)
        self.print_btn.configure(state=state)

    def update_selection_status(self, shown=None):
        self.update_header()
        shown = len(self.tree.get_children()) if shown is None else shown
        self.status_label.configure(
            text=f"{shown} of {len(self.all_rows)} shown, {len(self.selected)} selected for print.",
            text_color="#555555",
        )
        self.refresh_generate_btn()

    def all_visible_selected(self) -> bool:
        iids = self.tree.get_children()
        return bool(iids) and all(int(i) in self.selected for i in iids)

    def toggle_all_visible(self):
        """Header checkbox: select ONLY the rows shown; if that's already the selection, clear it."""
        if not self.all_rows:
            return
        select = not self.all_visible_selected()
        for iid in self.tree.get_children():
            idx = int(iid)
            if select:
                self.selected.add(idx)
            else:
                self.selected.discard(idx)
            self.tree.item(iid, text="☑" if select else "☐")
        self.update_selection_status()

    def print_indices(self):
        """Rows that will be printed: ticked AND currently shown."""
        # visible = {int(i) for i in self.tree.get_children()}
        # return sorted(self.selected & visible)
        return sorted(self.selected)

    def update_header(self):
        self.tree.heading("#0", text="☑" if self.all_visible_selected() else "☐")

    def setup_filters(self):
        """Find the Jabatan / Department columns and fill the dropdowns with their values."""
        cols = [c.lower() for c in self.all_columns]
        self.jabatan_idx = cols.index("jabatan") if "jabatan" in cols else None
        self.dept_idx = cols.index("department") if "department" in cols else None

        for idx, menu, var in (
            (self.jabatan_idx, self.jabatan_menu, self.jabatan_var),
            (self.dept_idx, self.dept_menu, self.dept_var),
        ):
            values = sorted({row[idx] for row in self.all_rows if row[idx]}) if idx is not None else []
            menu.configure(values=["All"] + values, state="normal" if idx is not None else "disabled")
            var.set("All")

    def set_placeholder(self):
        self.tree["columns"] = ["info"]
        self.tree.column("#0", width=0, minwidth=0, stretch=False)
        self.tree.heading("#0", text="")
        self.tree.delete(*self.tree.get_children())
        self.tree.insert("", "end", values=("Select a file to preview its data here.",))

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select CSV or Excel file",
            filetypes=[("Spreadsheet files", "*.csv *.xlsx *.xls"), ("All files", "*.*")],
        )
        if not path:
            return

        self.show_loading("Reading file...")
        threading.Thread(target=self.load_file, args=(path,), daemon=True).start()

    def load_done(self, path, result):
        columns, rows, meta = result

        self.input_path = path
        self.file_label.configure(text=os.path.basename(path))

        left = " ".join(p for p in ["Payroll", meta.get("company")] if p)
        right = meta.get("bulan")
        self.meta_label.configure(text="  |  ".join(p for p in [left, right] if p))

        self.all_columns = columns
        self.all_rows = rows

        self.setup_filters()
        self.search_var.set("")
        if self._search_job:                      # clearing the search must not trigger a second rebuild
            self.root.after_cancel(self._search_job)
            self._search_job = None

        self.selected = set(range(len(rows)))     # everything selected by default
        self.populate_table(columns, rows, list(range(len(rows))),
                            on_done=self.update_selection_status)
   
    def download_template(self):
        src = os.path.join(TEMPLATE_DIR, TEMPLATE_FILE)
        if not os.path.isfile(src):
            messagebox.showerror("Template not found", f"The template file is missing:\n{src}")
            return

        dest = filedialog.asksaveasfilename(
            title="Save template as",
            defaultextension=".xlsx",
            initialfile=TEMPLATE_FILE,
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not dest:
            return

        try:
            shutil.copyfile(src, dest)
        except PermissionError:
            messagebox.showerror("Cannot save", "That file is open in another program. Close it or choose a different name.")
            return
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        if messagebox.askyesno("Template saved", f"Saved to:\n{dest}\n\nOpen it now?"):
            self.open_file(dest)

    def on_search_changed(self, *args):
        self._search_job = None
        if not self.all_rows:
            return

        query = self.search_var.get().strip().lower()
        jab = self.jabatan_var.get()
        dept = self.dept_var.get()

        def matches(row):
            if query and query not in str(row[0]).lower():
                return False
            if jab != "All" and self.jabatan_idx is not None and row[self.jabatan_idx] != jab:
                return False
            if dept != "All" and self.dept_idx is not None and row[self.dept_idx] != dept:
                return False
            return True

        pairs = [(i, row) for i, row in enumerate(self.all_rows) if matches(row)]

        self.populate_table(self.all_columns, [r for _, r in pairs], [i for i, _ in pairs],
                            on_done=lambda: self.update_selection_status(shown=len(pairs)))

    def populate_table(self, columns, rows, indices, on_done=None):
        # stop any insert still running from a previous search/filter
        if self._insert_job:
            self.root.after_cancel(self._insert_job)
            self._insert_job = None

        items = list(zip(indices, rows))
        if len(items) > LOADING_THRESHOLD:
            self.show_loading(f"Loading rows... 0/{len(items)}")

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns

        self.tree.column("#0", width=44, minwidth=44, stretch=False, anchor="center")
        self.tree.heading("#0", text="☐", command=self.toggle_all_visible)

        font = tkfont.Font(font=("Arial", 10))
        header_font = tkfont.Font(font=("Arial", 10, "bold"))

        for i, col in enumerate(columns):
            self.tree.heading(col, text=col)
            # measure only the longest value per column instead of every cell
            longest = max((str(r[i]) for r in rows), key=len, default="")
            content_width = max(header_font.measure(col), font.measure(longest))
            anchor = "w" if i == 0 else "center"
            w = content_width + 24
            self.tree.column(col, anchor=anchor, width=w, minwidth=w, stretch=False)

        CHUNK = 300

        def step(start=0):
            for idx, row in items[start:start + CHUNK]:
                mark = "☑" if idx in self.selected else "☐"
                self.tree.insert("", "end", iid=str(idx), text=mark, values=row)
            nxt = start + CHUNK
            if nxt < len(items):
                self.loading_label.configure(text=f"Loading rows... {nxt}/{len(items)}")
                self._insert_job = self.root.after(1, step, nxt)
            else:
                self._insert_job = None
                self.hide_loading()
                if on_done:
                    on_done()

        step()
    
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

        try:
            with open(output_path, "ab"):
                pass
        except PermissionError:
            messagebox.showerror("Cannot save", "That file is open in another program. Close it or choose a different name.")
            return

        # self.generate_btn.configure(state="disabled")
        self.select_btn.configure(state="disabled")
        self.status_label.configure(text="Generating report...", text_color="#555555")
        # self.root.update_idletasks()

        # Run generation in a background thread so the UI doesn't freeze.
        thread = threading.Thread(
            target=self.run_generation, 
            args=(self.input_path, output_path, LAYOUT_LABELS[self.layout_var.get()], self.print_indices())
        )
        thread.start()

    def run_generation(self, input_path, output_path, layout, only):
        try:
            def on_progress(percent):
                self.root.after(0, lambda: self.status_label.configure(
                    text=f"Generating... {percent}%", text_color="#555555"
                ))
            generate_pdf_parallel(input_path, output_path, layout=layout, progress_callback=on_progress, only=only)
            self.root.after(0, self.on_success, output_path)
        except Exception as e:
            traceback.print_exc()
            self.root.after(0, self.on_error, self.describe_error(e))

    def on_success(self, output_path):
        self.status_label.configure(text="Report generated successfully!", text_color="#007700")
        self.refresh_generate_btn()
        self.select_btn.configure(state="normal")
        if messagebox.askyesno("Success", f"Report saved to:\n{output_path}\n\nOpen it now?"):
            self.open_file(output_path)

    def on_error(self, error_message):
        self.status_label.configure(text="Failed to generate report.", text_color="#cc0000")
        self.refresh_generate_btn()
        self.select_btn.configure(state="normal")
        messagebox.showerror("Error", f"Something went wrong:\n\n{error_message}")

    def describe_error(self, e: Exception) -> str:
        if isinstance(e, PermissionError):
            return ("Cannot write to the output file.\n\n"
                    "It is most likely still open in a PDF viewer or another program, "
                    "or the folder is protected. Close it, or save under a different name, "
                    f"then try again.\n\nDetails: {e}")
        if isinstance(e, FileNotFoundError):
            return f"A file or folder could not be found.\n\nDetails: {e}"
        if isinstance(e, BrokenProcessPool):
            return ("A background worker crashed while rendering. "
                    "Try again, or check that the template files are present.")
        if isinstance(e, ValueError):
            return str(e)  # your own messages, e.g. missing column
        return f"{type(e).__name__}: {e}"

    def open_file(self, path):
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform.startswith("darwin"):
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')

    def preview_print(self):
        if not self.input_path or not os.path.isfile(self.input_path):
            messagebox.showerror("Error", "Please select a valid file first.")
            return
        try:
            html = generate_html(self.input_path, LAYOUT_LABELS[self.layout_var.get()], self.print_indices())
            fd, path = tempfile.mkstemp(suffix=".html", prefix="slip_")
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(html)
            webbrowser.open(Path(path).as_uri())
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", self.describe_error(e))


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()

    root = ctk.CTk()
    app = ReportApp(root)

    if pyi_splash:
        pyi_splash.close()
        
    root.mainloop()