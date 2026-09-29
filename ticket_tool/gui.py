"""Tkinter GUI: notes .txt -> preview -> Excel. Works on Windows, macOS, Linux."""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tkinter as tk
from datetime import date
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional

from . import __version__
from . import excel_io
from .parser import ParseResult, TicketRow, parse_file

CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".ticket_notes_tool.json")
WARN_BG = "#fff1c2"

# key, heading, width, anchor, editable
COLUMNS = [
    ("no", "#", 40, "center", False),
    ("ticket_id", "Ticket ID", 130, "w", True),
    ("priority", "Priority", 70, "center", True),
    ("status", "Status", 95, "center", True),
    ("title", "Title", 250, "w", True),
    ("description", "Description", 340, "w", True),
    ("time_taken", "Time Taken", 80, "center", True),
]


def open_file(path: str) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)                                # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title(f"Ticket Notes to Excel  v{__version__}")
        root.geometry("1150x780")
        root.minsize(900, 620)

        self.cfg = self._load_cfg()
        self.rows: List[TicketRow] = []
        self.notes_path = ""
        self.dirty = False

        self.strip_var = tk.BooleanVar(value=True)
        self.backup_var = tk.BooleanVar(value=True)
        self.fit_var = tk.BooleanVar(value=True)
        self.notes_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.parse_info = tk.StringVar(value="No notes file loaded.")
        self.excel_info = tk.StringVar(value="No Excel file selected.")

        self._build_ui()
        last = self.cfg.get("last_excel")
        if last and os.path.exists(last):
            self._set_excel(last)

    # ------------------------------------------------------------ config
    @staticmethod
    def _load_cfg() -> dict:
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return {}

    def _save_cfg(self) -> None:
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as fh:
                json.dump(self.cfg, fh)
        except Exception:
            pass

    # ------------------------------------------------------------ UI
    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 5}
        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        # Step 1 ------------------------------------------------------
        f1 = ttk.LabelFrame(main, text="Step 1 - Select your notes file (.txt)")
        f1.grid(row=0, column=0, sticky="ew", **pad)
        f1.columnconfigure(0, weight=1)
        ttk.Entry(f1, textvariable=self.notes_var, state="readonly").grid(
            row=0, column=0, sticky="ew", padx=8, pady=6)
        ttk.Button(f1, text="Browse...", command=self._browse_notes).grid(
            row=0, column=1, padx=8)
        ttk.Checkbutton(
            f1, variable=self.strip_var, command=self._reparse,
            text="Strip ticket-system boilerplate ('Recent Comments' heading and "
                 "timestamp/author lines)").grid(row=1, column=0, columnspan=2,
                                                 sticky="w", padx=8)
        ttk.Label(f1, textvariable=self.parse_info, foreground="#555").grid(
            row=2, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 6))

        # Step 2 ------------------------------------------------------
        f2 = ttk.LabelFrame(main, text="Step 2 - Preview (double-click a cell to "
                                        "edit; rows in yellow are missing a priority)")
        f2.grid(row=1, column=0, sticky="nsew", **pad)
        f2.columnconfigure(0, weight=1)
        f2.rowconfigure(0, weight=3)
        f2.rowconfigure(2, weight=1)

        holder = ttk.Frame(f2)
        holder.grid(row=0, column=0, sticky="nsew", padx=8, pady=6)
        holder.columnconfigure(0, weight=1)
        holder.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(holder, columns=[c[0] for c in COLUMNS],
                                 show="headings", selectmode="extended")
        for key, head, width, anchor, _ in COLUMNS:
            self.tree.heading(key, text=head)
            self.tree.column(key, width=width, anchor=anchor,
                             stretch=key in ("title", "description"))
        self.tree.tag_configure("warn", background=WARN_BG)
        ys = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(holder, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Delete>", lambda e: self._delete_selected())
        self.tree.bind("<BackSpace>", lambda e: self._delete_selected())

        bar = ttk.Frame(f2)
        bar.grid(row=1, column=0, sticky="ew", padx=8)
        ttk.Button(bar, text="Remove selected rows",
                   command=self._delete_selected).pack(side="left")
        ttk.Button(bar, text="Re-read notes file",
                   command=lambda: self._reparse(force=True)).pack(side="left", padx=6)
        ttk.Label(bar, text="Repeated Ticket IDs are kept - only rows you remove "
                            "here are left out.", foreground="#555").pack(side="left", padx=10)

        self.detail = tk.Text(f2, height=7, wrap="word", state="disabled",
                              relief="solid", borderwidth=1)
        self.detail.grid(row=2, column=0, sticky="nsew", padx=8, pady=6)

        # Step 3 ------------------------------------------------------
        f3 = ttk.LabelFrame(main, text="Step 3 - Select your daily Excel file")
        f3.grid(row=2, column=0, sticky="ew", **pad)
        f3.columnconfigure(0, weight=1)
        ttk.Entry(f3, textvariable=self.excel_var, state="readonly").grid(
            row=0, column=0, sticky="ew", padx=8, pady=6)
        ttk.Button(f3, text="Browse...", command=self._browse_excel).grid(row=0, column=1, padx=4)
        ttk.Button(f3, text="Create new...", command=self._create_excel).grid(row=0, column=2, padx=8)
        opts = ttk.Frame(f3)
        opts.grid(row=1, column=0, columnspan=3, sticky="w", padx=8)
        ttk.Label(opts, text="Sheet:").pack(side="left")
        self.sheet_box = ttk.Combobox(opts, textvariable=self.sheet_var, state="readonly", width=24)
        self.sheet_box.pack(side="left", padx=6)
        self.sheet_box.bind("<<ComboboxSelected>>", lambda e: self._refresh_excel_info())
        ttk.Checkbutton(opts, variable=self.backup_var,
                        text="Back up the file first (backups/ folder)").pack(side="left", padx=12)
        ttk.Checkbutton(opts, variable=self.fit_var,
                        text="Grow row height for long text").pack(side="left")
        ttk.Label(f3, textvariable=self.excel_info, foreground="#555").grid(
            row=2, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 6))

        # Step 4 ------------------------------------------------------
        f4 = ttk.LabelFrame(main, text="Step 4 - Write to Excel")
        f4.grid(row=3, column=0, sticky="ew", **pad)
        self.write_btn = ttk.Button(f4, text="Add rows to Excel", command=self._write)
        self.write_btn.pack(side="left", padx=8, pady=8)
        self.open_btn = ttk.Button(f4, text="Open Excel file", command=self._open_excel)
        self.open_btn.pack(side="left")
        self.result_var = tk.StringVar()
        ttk.Label(f4, textvariable=self.result_var, foreground="#1a6b1a").pack(side="left", padx=12)

    # ------------------------------------------------------------ notes
    def _browse_notes(self) -> None:
        path = filedialog.askopenfilename(
            title="Select notes file",
            initialdir=self.cfg.get("last_notes_dir") or os.path.expanduser("~"),
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if path:
            self.cfg["last_notes_dir"] = os.path.dirname(path)
            self._save_cfg()
            self._load_notes(path)

    def _load_notes(self, path: str) -> None:
        try:
            result: ParseResult = parse_file(path, self.strip_var.get())
        except Exception as exc:
            messagebox.showerror("Could not read file", str(exc))
            return
        self.notes_path = path
        self.notes_var.set(path)
        self.rows = result.rows
        self.dirty = False
        self.result_var.set("")
        msg = f"Detected {len(self.rows)} rows."
        if result.ignored_lines:
            msg += f"  {len(result.ignored_lines)} stray line(s) outside any ticket were ignored."
        if not self.rows:
            msg = "No tickets detected. Expected lines like: 1. T20260929.0043 - Title | P3 | Assigned"
        self.parse_info.set(msg)
        self._refresh_tree()

    def _reparse(self, force: bool = False) -> None:
        if not self.notes_path:
            return
        if self.dirty and not messagebox.askyesno(
                "Discard edits?", "Re-reading the notes file will discard your edits "
                                  "and removed rows in the preview. Continue?"):
            if not force:                      # undo the checkbox toggle
                self.strip_var.set(not self.strip_var.get())
            return
        self._load_notes(self.notes_path)

    # ------------------------------------------------------------ preview
    @staticmethod
    def _flat(text: str, limit: int = 220) -> str:
        s = " / ".join(x.strip() for x in text.splitlines() if x.strip())
        return s if len(s) <= limit else s[:limit - 1] + "..."

    def _refresh_tree(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self.rows):
            tags = ("warn",) if not r.priority or not r.ticket_id else ()
            self.tree.insert("", "end", iid=str(i), tags=tags, values=(
                i + 1, r.ticket_id, r.priority, r.status, r.title,
                self._flat(r.description), r.time_taken))
        self._show_detail("")

    def _show_detail(self, text: str) -> None:
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", text)
        self.detail.configure(state="disabled")

    def _on_select(self, _e=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        r = self.rows[int(sel[0])]
        self._show_detail(f"{r.ticket_id}  |  {r.priority or '-'}  |  {r.status or '-'}\n"
                          f"{r.title}\n\n{r.description}")

    def _delete_selected(self) -> None:
        sel = sorted((int(i) for i in self.tree.selection()), reverse=True)
        if not sel:
            return
        for i in sel:
            del self.rows[i]
        self.dirty = True
        self.parse_info.set(f"{len(self.rows)} rows in preview ({len(sel)} removed).")
        self._refresh_tree()

    def _on_double_click(self, event) -> None:
        if self.tree.identify_region(event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        col = int(self.tree.identify_column(event.x).lstrip("#")) - 1
        key, head, _, _, editable = COLUMNS[col]
        if not iid or not editable:
            return
        row = self.rows[int(iid)]
        attr = key
        new = self._edit_dialog(f"Edit {head}", getattr(row, attr), multiline=(key in ("description", "title")))
        if new is None or new == getattr(row, attr):
            return
        setattr(row, attr, new.strip("\n") if key == "description" else new.strip())
        self.dirty = True
        self._refresh_tree()
        self.tree.selection_set(iid)

    def _edit_dialog(self, title: str, value: str, multiline: bool) -> Optional[str]:
        win = tk.Toplevel(self.root)
        win.title(title)
        win.transient(self.root)
        txt = tk.Text(win, width=90, height=18 if multiline else 2, wrap="word")
        txt.pack(fill="both", expand=True, padx=10, pady=10)
        txt.insert("1.0", value)
        txt.focus_set()
        out = {"v": None}

        def ok():
            v = txt.get("1.0", "end-1c")
            out["v"] = v if multiline else " ".join(v.split())
            win.destroy()

        btns = ttk.Frame(win)
        btns.pack(pady=(0, 10))
        ttk.Button(btns, text="OK", command=ok).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="left", padx=6)
        win.grab_set()
        self.root.wait_window(win)
        return out["v"]

    # ------------------------------------------------------------ excel
    def _browse_excel(self) -> None:
        path = filedialog.askopenfilename(
            title="Select your Excel file",
            initialdir=os.path.dirname(self.cfg.get("last_excel", "")) or os.path.expanduser("~"),
            filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if path:
            self._set_excel(path)

    def _create_excel(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Create new Excel file", defaultextension=".xlsx",
            initialfile=f"Tickets_{date.today():%Y-%m-%d}.xlsx",
            initialdir=os.path.dirname(self.cfg.get("last_excel", "")) or os.path.expanduser("~"),
            filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            excel_io.create_new_workbook(path)
        except PermissionError:
            messagebox.showerror("Cannot create file", "That file is open in Excel. Close it first.")
            return
        except Exception as exc:
            messagebox.showerror("Cannot create file", str(exc))
            return
        self._set_excel(path)

    def _set_excel(self, path: str) -> None:
        try:
            sheets = excel_io.list_sheets(path)
        except excel_io.ExcelWriteError as exc:
            messagebox.showerror("Excel file", str(exc))
            return
        self.excel_var.set(path)
        self.cfg["last_excel"] = path
        self._save_cfg()
        self.sheet_box["values"] = sheets
        self.sheet_var.set(excel_io.DEFAULT_SHEET if excel_io.DEFAULT_SHEET in sheets else sheets[0])
        self._refresh_excel_info()

    def _refresh_excel_info(self) -> None:
        path = self.excel_var.get()
        if not path:
            return
        try:
            info = excel_io.inspect_sheet(path, self.sheet_var.get())
        except excel_io.ExcelWriteError as exc:
            self.excel_info.set(str(exc))
            return
        note = "" if info["headers_ok"] else \
            "   WARNING: row 1 headers differ from the template - columns A-F will be used."
        self.excel_info.set(f"'{info['sheet']}': {info['existing_rows']} existing rows. "
                            f"New rows will start at row {info['next_row']}.{note}")

    def _open_excel(self) -> None:
        path = self.excel_var.get()
        if path and os.path.exists(path):
            open_file(path)

    # ------------------------------------------------------------ write
    def _write(self) -> None:
        if not self.rows:
            messagebox.showwarning("Nothing to write", "Load a notes file first (Step 1).")
            return
        path = self.excel_var.get()
        if not path:
            messagebox.showwarning("No Excel file", "Select or create an Excel file (Step 3).")
            return
        sheet = self.sheet_var.get()
        try:
            missing = sum(1 for r in self.rows if not r.priority)
            if missing and not messagebox.askyesno(
                    "Missing priority",
                    f"{missing} row(s) have no priority (highlighted yellow). "
                    "Add them with a blank priority?"):
                return
            already = excel_io.count_existing_matches(path, self.rows, sheet)
            if already and not messagebox.askyesno(
                    "Possible repeat",
                    f"{already} of these {len(self.rows)} rows already exist in the sheet, "
                    "identical in every column - you may have run this notes file already.\n\n"
                    "Add all rows again anyway?"):
                return
            res = excel_io.append_rows(path, self.rows, sheet, self.backup_var.get(),
                                       self.fit_var.get())
        except excel_io.ExcelWriteError as exc:
            messagebox.showerror("Could not write", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Unexpected error", f"{type(exc).__name__}: {exc}")
            return
        self.result_var.set(f"Added {res['count']} rows to '{res['sheet']}' "
                            f"(rows {res['first_row']}-{res['last_row']}).")
        self._refresh_excel_info()
        if res["backup"]:
            self.result_var.set(self.result_var.get() + "  Backup saved.")


def main() -> None:
    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)   # crisp text on Windows
        except Exception:
            pass
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
