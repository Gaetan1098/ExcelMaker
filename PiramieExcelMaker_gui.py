"""
PiramieExcelMaker_gui.py
---------------------------------
Drag-and-drop (or file-picker fallback) GUI with two flows:

  A) Run Core Only:
     - Requires MASTER (.xlsm) only
     - Calls process_excel_file(MASTER)

  B) Ingest + Core:
     - Requires MASTER (.xlsm) + MONTH report (.xlsx/.xls)
     - Calls ingest_month_into_apr_bundle(MASTER, MONTH)  [creates backup]
     - Then calls process_excel_file(MASTER)

Dependencies:
  pip install pandas openpyxl tkinterdnd2  (tkinterdnd2 is optional)
"""

import os
import sys
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Optional drag & drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

# 1) Monthly ingest (appends/normalizes APR Bundle + backup)
from PiramieExcelMaker_append import ingest_month_into_apr_bundle
# 2) Your existing core (builds Molo Molo Auto etc.)
from PiramieExcelMaker_core import process_excel_file


class DropZone(ttk.Frame):
    def __init__(self, master, text, on_file_selected, filetypes=()):
        super().__init__(master, padding=10, borderwidth=2, relief="groove")
        self.on_file_selected = on_file_selected
        self.filetypes = filetypes

        self.label = ttk.Label(self, text=text, anchor="center", wraplength=280, justify="center")
        self.label.pack(fill="both", expand=True)

        self.btn = ttk.Button(self, text="Browse...", command=self.browse_file)
        self.btn.pack(pady=(8, 0))

        if DND_AVAILABLE:
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select file",
            filetypes=self.filetypes or [("All files", "*.*")],
        )
        if path:
            self.on_file_selected(path)
            self.label.configure(text=os.path.basename(path))

    def _on_drop(self, event):
        raw = event.data
        # handle braces from Windows drag payloads and multiple paths; take first
        if raw.startswith("{") and raw.endswith("}"):
            raw = raw[1:-1]
        path = raw.split()[0]
        if os.path.isfile(path):
            self.on_file_selected(path)
            self.label.configure(text=os.path.basename(path))


class App(TkinterDnD.Tk if DND_AVAILABLE else tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("APR Bundle Ingest & Core Processor")
        self.geometry("820x560")
        self.minsize(780, 520)

        # State
        self.master_path = None       # .xlsm
        self.month_report_path = None # .xlsx/.xls

        self._build_ui()

    def _build_ui(self):
        # Title
        title = ttk.Label(self, text="APR Bundle Pipeline", font=("Segoe UI", 16, "bold"))
        title.pack(pady=(12, 4))

        tips = ttk.Label(
            self,
            text="1) Select MASTER (.xlsm)   2) (Optional) Select MONTH (.xlsx/.xls)   3) Choose an action below",
            foreground="#555",
        )
        tips.pack(pady=(0, 10))

        zones = ttk.Frame(self)
        zones.pack(fill="both", expand=True, padx=16)

        self.zone_master = DropZone(
            zones,
            "Drop/select MASTER workbook (.xlsm)",
            self._set_master,
            filetypes=[("Excel Macro-Enabled", "*.xlsm")],
        )
        self.zone_master.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=8)

        self.zone_month = DropZone(
            zones,
            "Drop/select MONTH purchases report (.xlsx/.xls) — optional for 'Run Core Only'",
            self._set_month,
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003", "*.xls")],
        )
        self.zone_month.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=8)

        zones.grid_columnconfigure(0, weight=1)
        zones.grid_columnconfigure(1, weight=1)
        zones.grid_rowconfigure(0, weight=1)

        # Action bar (two buttons)
        action_bar = ttk.Frame(self)
        action_bar.pack(fill="x", padx=16, pady=(0, 10))

        self.btn_core_only = ttk.Button(action_bar, text="Run Core Only", command=self._run_core_only)
        self.btn_core_only.pack(side="left", padx=(0, 8))

        self.btn_ingest_core = ttk.Button(action_bar, text="Ingest + Core", command=self._run_ingest_then_core)
        self.btn_ingest_core.pack(side="left")

        self.btn_quit = ttk.Button(action_bar, text="Quit", command=self.destroy)
        self.btn_quit.pack(side="right")

        # Log box
        self.text = tk.Text(self, height=16, wrap="word", state="disabled")
        self.text.pack(fill="both", expand=False, padx=16, pady=(0, 16))

    def _log(self, msg: str):
        self.text.configure(state="normal")
        self.text.insert("end", msg + "\n")
        self.text.see("end")
        self.text.configure(state="disabled")
        self.update_idletasks()

    # ---------- file setters ----------
    def _set_master(self, path: str):
        if not path.lower().endswith(".xlsm"):
            messagebox.showwarning("Invalid file", "MASTER must be a .xlsm file.")
            return
        self.master_path = path
        self._log(f"MASTER selected: {path}")

    def _set_month(self, path: str):
        if not (path.lower().endswith(".xlsx") or path.lower().endswith(".xls")):
            messagebox.showwarning("Invalid file", "MONTH report must be .xlsx or .xls.")
            return
        self.month_report_path = path
        self._log(f"MONTH report selected: {path}")

    # ---------- actions ----------
    def _run_core_only(self):
        if not self.master_path or not os.path.isfile(self.master_path):
            messagebox.showerror("Missing MASTER", "Please select a valid MASTER (.xlsm) file.")
            return
        try:
            self._log("Running Core Only...")
            process_excel_file(self.master_path)
            self._log("Core processing complete.")
            messagebox.showinfo("Success", "Core run finished successfully.")
        except Exception as e:
            tb = traceback.format_exc()
            self._log("ERROR (Core Only):\n" + tb)
            messagebox.showerror("Run failed", str(e))

    def _run_ingest_then_core(self):
        if not self.master_path or not os.path.isfile(self.master_path):
            messagebox.showerror("Missing MASTER", "Please select a valid MASTER (.xlsm) file.")
            return
        if not self.month_report_path or not os.path.isfile(self.month_report_path):
            messagebox.showerror("Missing MONTH", "Please select a valid MONTH (.xlsx/.xls) file.")
            return

        try:
            self._log("Step 1/2: Ingesting month into APR Bundle (backup first)...")

            # call the column-by-column appender
            summary = ingest_month_into_apr_bundle(self.master_path, self.month_report_path)

            # log what that function actually returns
            self._log(f"  Backup created: {summary.get('master_backup', '(n/a)')}")
            rows_appended = summary.get("rows_appended", 0)
            fr = summary.get("first_row_written")
            lr = summary.get("last_row_written")
            range_str = f" (rows {fr}–{lr})" if fr and lr else ""
            self._log(f"  Rows appended: {rows_appended}{range_str}")
            self._log(f"  Sheet updated: {summary.get('sheet', 'APR Bundle')}")

            self._log("Step 2/2: Running core processor on MASTER...")
            process_excel_file(self.master_path)
            self._log("Core processing complete.")

        except Exception as e:
            tb = traceback.format_exc()
            self._log("ERROR (Ingest + Core):\n" + tb)
            messagebox.showerror("Run failed", str(e))


if __name__ == "__main__":
    app = App()
    app.mainloop()
