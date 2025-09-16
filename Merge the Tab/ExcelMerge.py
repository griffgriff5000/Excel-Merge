#!/usr/bin/env python3
# ExcelMergeGUI.py
import re
import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

# ---------- Core merge logic ----------

def read_all_sheets(
    xls_path: Path,
    exclude_patterns,
    include_hidden: bool,
    header_row: int | None,
    auto_header: bool,
    keep_empty: bool,
    log=print,
):
    """Read all sheets from an Excel file into {name: DataFrame}."""
    from openpyxl import load_workbook

    log(f"Opening workbook: {xls_path}")
    wb = load_workbook(filename=xls_path, read_only=True, data_only=True)
    sheet_names = []
    for ws in wb.worksheets:
        if not include_hidden and ws.sheet_state != "visible":
            continue
        sheet_names.append(ws.title)

    # Apply exclude regex filters
    if exclude_patterns:
        def excluded(name: str) -> bool:
            return any(re.search(rx, name, flags=re.I) for rx in exclude_patterns)
        before = len(sheet_names)
        sheet_names = [n for n in sheet_names if not excluded(n)]
        log(f"Excluding {before - len(sheet_names)} sheet(s) by pattern.")

    if not sheet_names:
        log("No sheets selected to read.")
        return {}

    # Auto header detection helper
    def detect_header(df: pd.DataFrame) -> int:
        scan_rows = min(len(df), 30)
        best_row, best_count = 0, -1
        for r in range(scan_rows):
            cnt = df.iloc[r].notna().sum()
            if cnt > best_count:
                best_row, best_count = r, cnt
        return best_row

    sheets = {}
    for name in sheet_names:
        log(f"Reading sheet: {name}")
        if auto_header:
            df = pd.read_excel(xls_path, sheet_name=name, header=None, engine="openpyxl")
            if df.empty:
                sheets[name] = df
                continue
            h = detect_header(df)
            df.columns = df.iloc[h].astype(str).str.strip()
            df = df.iloc[h + 1:].reset_index(drop=True)
        else:
            df = pd.read_excel(xls_path, sheet_name=name, header=header_row, engine="openpyxl")

        if not keep_empty:
            df = df.dropna(how="all")
        sheets[name] = df

    return sheets


def unify_columns_union(dfs: list[pd.DataFrame]) -> list[pd.DataFrame]:
    """Make all DataFrames share the union of all columns (order = discovered order)."""
    cols = []
    seen = set()
    for df in dfs:
        for c in map(str, df.columns):
            if c not in seen:
                seen.add(c)
                cols.append(c)
    out = []
    for df in dfs:
        df = df.copy()
        df.columns = list(map(str, df.columns))
        out.append(df.reindex(columns=cols))
    return out


def merge_workbook(
    in_path: Path,
    out_path: Path,
    add_source: bool,
    include_hidden: bool,
    auto_header: bool,
    header_row: int | None,
    exclude_patterns: list[str],
    keep_empty: bool,
    out_sheet_name: str,
    log=print,
):
    sheets = read_all_sheets(
        in_path,
        exclude_patterns=exclude_patterns,
        include_hidden=include_hidden,
        header_row=(header_row if header_row is not None else 0),
        auto_header=auto_header,
        keep_empty=keep_empty,
        log=log,
    )

    if not sheets:
        raise RuntimeError("No data read from any sheet (after filters).")

    # Optionally add SourceSheet column before union
    dfs = []
    for name, df in sheets.items():
        if df.empty:
            continue
        df2 = df.copy()
        if add_source:
            df2["SourceSheet"] = name
        dfs.append(df2)

    if not dfs:
        merged = pd.DataFrame()
    else:
        aligned = unify_columns_union(dfs)
        merged = pd.concat(aligned, axis=0, ignore_index=True)

    # Write output
    if out_path.suffix.lower() == ".csv":
        merged.to_csv(out_path, index=False)
        log(f"Written CSV: {out_path}")
    else:
        with pd.ExcelWriter(out_path, engine="openpyxl") as xlw:
            merged.to_excel(xlw, index=False, sheet_name=out_sheet_name)
        log(f"Written Excel: {out_path} (sheet: {out_sheet_name})")

    return merged.shape  # (rows, cols)


# ---------- GUI ----------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Sheet Merger")
        self.geometry("720x520")
        self.minsize(720, 520)

        # State
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.add_source = tk.BooleanVar(value=True)
        self.include_hidden = tk.BooleanVar(value=False)
        self.keep_empty = tk.BooleanVar(value=False)
        self.auto_header = tk.BooleanVar(value=True)
        self.header_row = tk.IntVar(value=0)
        self.exclude_regex = tk.StringVar(value="")   # e.g. ^Summary$|Archive
        self.out_format = tk.StringVar(value="xlsx")  # xlsx or csv
        self.out_sheet_name = tk.StringVar(value="Consolidated")

        self._build()

    def _build(self):
        pad = {"padx": 8, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True)

        # Input file
        row = 0
        ttk.Label(frm, text="Input workbook (.xlsx/.xlsm):", font=("Segoe UI", 10, "bold")).grid(row=row, column=0, sticky="w", **pad)
        row += 1
        in_row = ttk.Frame(frm)
        in_row.grid(row=row, column=0, sticky="ew", **pad)
        in_row.columnconfigure(0, weight=1)
        ttk.Entry(in_row, textvariable=self.input_path).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(in_row, text="Browse…", command=self.select_input).grid(row=0, column=1)
        row += 1

        # Output file + format
        ttk.Label(frm, text="Output:", font=("Segoe UI", 10, "bold")).grid(row=row, column=0, sticky="w", **pad)
        row += 1

        out_row = ttk.Frame(frm)
        out_row.grid(row=row, column=0, sticky="ew", **pad)
        out_row.columnconfigure(0, weight=1)
        ttk.Entry(out_row, textvariable=self.output_path).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(out_row, text="Choose…", command=self.select_output).grid(row=0, column=1)

        fmt_row = ttk.Frame(frm)
        fmt_row.grid(row=row+1, column=0, sticky="w", **pad)
        ttk.Radiobutton(fmt_row, text="Excel (.xlsx)", value="xlsx", variable=self.out_format, command=self._toggle_sheet_name).pack(side="left", padx=(0, 12))
        ttk.Radiobutton(fmt_row, text="CSV (.csv)", value="csv", variable=self.out_format, command=self._toggle_sheet_name).pack(side="left")
        row += 2

        # Options
        ttk.Label(frm, text="Options:", font=("Segoe UI", 10, "bold")).grid(row=row, column=0, sticky="w", **pad)
        row += 1

        opts = ttk.Frame(frm)
        opts.grid(row=row, column=0, sticky="ew", **pad)
        for i in range(3):
            opts.columnconfigure(i, weight=1)

        ttk.Checkbutton(opts, text="Add SourceSheet column", variable=self.add_source).grid(row=0, column=0, sticky="w", pady=2)
        ttk.Checkbutton(opts, text="Include hidden sheets", variable=self.include_hidden).grid(row=0, column=1, sticky="w", pady=2)
        ttk.Checkbutton(opts, text="Keep completely empty rows", variable=self.keep_empty).grid(row=0, column=2, sticky="w", pady=2)

        # Headers
        hdr = ttk.Frame(frm)
        hdr.grid(row=row+1, column=0, sticky="ew", **pad)
        ttk.Checkbutton(hdr, text="Auto-detect header row", variable=self.auto_header, command=self._toggle_header_ctrls).grid(row=0, column=0, sticky="w")
        ttk.Label(hdr, text="or header row index (0-based):").grid(row=0, column=1, sticky="e", padx=(16, 6))
        self.header_spin = ttk.Spinbox(hdr, from_=0, to=100, textvariable=self.header_row, width=6, state="disabled")
        self.header_spin.grid(row=0, column=2, sticky="w")

        # Exclude
        ex = ttk.Frame(frm)
        ex.grid(row=row+2, column=0, sticky="ew", **pad)
        ex.columnconfigure(1, weight=1)
        ttk.Label(ex, text="Exclude sheet names (regex, | separated):").grid(row=0, column=0, sticky="w", padx=(0, 6))
        ttk.Entry(ex, textvariable=self.exclude_regex).grid(row=0, column=1, sticky="ew")

        # Sheet name for Excel output
        sh = ttk.Frame(frm)
        sh.grid(row=row+3, column=0, sticky="ew", **pad)
        ttk.Label(sh, text="Output sheet name (for .xlsx):").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.sheet_entry = ttk.Entry(sh, textvariable=self.out_sheet_name, width=24)
        self.sheet_entry.grid(row=0, column=1, sticky="w")

        row += 4

        # Action buttons
        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, sticky="ew", **pad)
        btns.columnconfigure(0, weight=1)
        ttk.Button(btns, text="Merge", command=self.run_merge).pack(side="right")
        ttk.Button(btns, text="Quit", command=self.destroy).pack(side="right", padx=(0, 8))

        # Log
        ttk.Label(frm, text="Log:", font=("Segoe UI", 10, "bold")).grid(row=row+1, column=0, sticky="w", **pad)
        self.log = tk.Text(frm, height=10, wrap="word")
        self.log.grid(row=row+2, column=0, sticky="nsew", padx=8, pady=(0, 8))
        frm.rowconfigure(row+2, weight=1)

        # Initial state
        self._toggle_header_ctrls()
        self._toggle_sheet_name()

    # ----- helpers -----

    def _println(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.update_idletasks()

    def _toggle_header_ctrls(self):
        if self.auto_header.get():
            self.header_spin.configure(state="disabled")
        else:
            self.header_spin.configure(state="normal")

    def _toggle_sheet_name(self):
        if self.out_format.get() == "xlsx":
            self.sheet_entry.configure(state="normal")
        else:
            self.sheet_entry.configure(state="disabled")

    def select_input(self):
        path = filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if path:
            self.input_path.set(path)
            # suggest default output
            p = Path(path)
            default = p.with_name(p.stem + "_merged.xlsx")
            self.output_path.set(str(default))

    def select_output(self):
        if not self.input_path.get():
            initial = str(Path.home())
        else:
            initial = str(Path(self.input_path.get()).parent)
        if self.out_format.get() == "csv":
            path = filedialog.asksaveasfilename(
                title="Save merged file as CSV",
                defaultextension=".csv",
                initialdir=initial,
                initialfile="merged.csv",
                filetypes=[("CSV", "*.csv")]
            )
        else:
            path = filedialog.asksaveasfilename(
                title="Save merged file as Excel",
                defaultextension=".xlsx",
                initialdir=initial,
                initialfile="merged.xlsx",
                filetypes=[("Excel", "*.xlsx")]
            )
        if path:
            self.output_path.set(path)

    def run_merge(self):
        try:
            in_path = self.input_path.get().strip()
            out_path = self.output_path.get().strip()
            if not in_path:
                messagebox.showwarning("Missing input", "Please choose an input workbook.")
                return
            in_p = Path(in_path)
            if not in_p.exists():
                messagebox.showerror("Not found", f"Input file not found:\n{in_p}")
                return

            if not out_path:
                # default based on input + format
                out_p = in_p.with_name(in_p.stem + ("_merged.csv" if self.out_format.get()=="csv" else "_merged.xlsx"))
                self.output_path.set(str(out_p))
            out_p = Path(self.output_path.get())

            exclude_patterns = [s for s in re.split(r"\|", self.exclude_regex.get().strip()) if s] if self.exclude_regex.get().strip() else []

            self._println("---- Starting merge ----")
            self._println(f"Input:  {in_p}")
            self._println(f"Output: {out_p}")
            if exclude_patterns:
                self._println(f"Exclude regex: {exclude_patterns}")
            self._println(f"Auto header: {self.auto_header.get()}   Header row: {self.header_row.get() if not self.auto_header.get() else '(auto)'}")
            self._println(f"Include hidden: {self.include_hidden.get()}   Add SourceSheet: {self.add_source.get()}   Keep empty rows: {self.keep_empty.get()}")
            if self.out_format.get() == "xlsx":
                self._println(f"Output sheet name: {self.out_sheet_name.get()}")

            rows, cols = merge_workbook(
                in_path=in_p,
                out_path=out_p,
                add_source=self.add_source.get(),
                include_hidden=self.include_hidden.get(),
                auto_header=self.auto_header.get(),
                header_row=None if self.auto_header.get() else self.header_row.get(),
                exclude_patterns=exclude_patterns,
                keep_empty=self.keep_empty.get(),
                out_sheet_name=self.out_sheet_name.get(),
                log=self._println,
            )

            self._println(f"SUCCESS: Wrote {rows} rows × {cols} columns.")
            messagebox.showinfo("Done", f"Merged successfully.\n\nRows: {rows}\nCols: {cols}\n\nSaved to:\n{out_p}")
        except Exception as e:
            self._println(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))


def main():
    # Better DPI scaling on Windows (optional)
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = App()
    app.mainloop()


if __name__ == "__main__":
    # Ensure pandas shows full tracebacks if needed
    pd.set_option("display.max_colwidth", 200)
    try:
        main()
    except Exception as exc:
            # Last-ditch error surface
            messagebox.showerror("Fatal error", str(exc))
            sys.exit(1)
