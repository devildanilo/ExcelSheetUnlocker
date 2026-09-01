# Copyright (c) 2026 Danilo Di Pietrantonio
# All rights reserved.
#
# Use and free, non-commercial distribution permitted.
# Commercial use, sale, or other financial gain from the
# software requires express written authorization from
# the copyright holder.
# Modification requires express written authorization
# from the copyright holder.
# See LICENSE.txt for full terms.

import zipfile
import shutil
import os
import sys
import re
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from xml.etree import ElementTree as ET

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False


# ── Resource helper (works in dev and in PyInstaller .exe) ─────────────────
def _resource(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(__file__), relative_path)


# ── Palette ────────────────────────────────────────────────────────────────
BG       = "#0f1117"
BG2      = "#1a1d27"
BG3      = "#22263a"
BG4      = "#181c2a"
ACCENT   = "#4ade80"
ACCENT2  = "#22c55e"
TEXT     = "#e2e8f0"
TEXT_DIM = "#64748b"
DANGER   = "#f87171"
WARN     = "#fbbf24"
BORDER   = "#2d3148"
INDIGO   = "#818cf8"

FONT_MAIN  = ("Segoe UI", 10)
FONT_BOLD  = ("Segoe UI", 10, "bold")
FONT_MONO  = ("Consolas", 9)
FONT_SM    = ("Segoe UI", 8)
FONT_SHEET = ("Segoe UI", 9)


# ── Scan: find protected sheets and their names ────────────────────────────
def scan_xlsx(file_path):
    sheets = []
    with zipfile.ZipFile(file_path, 'r') as z:
        names = z.namelist()
        REQUIRED_ENTRIES = {"xl/workbook.xml", "[Content_Types].xml"}
        if not REQUIRED_ENTRIES.issubset(set(names)):
            raise ValueError("Not a valid .xlsx file")

        wb_path = next((n for n in names if n.endswith("xl/workbook.xml")
                        or n == "xl/workbook.xml"), None)
        if not wb_path:
            raise FileNotFoundError("xl/workbook.xml not found")

        wb_xml  = z.read(wb_path)
        wb_root = ET.fromstring(wb_xml)
        ns_main = {
            'ms': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }

        rid_to_name = {}
        for sh in wb_root.findall('.//ms:sheet', ns_main):
            rid  = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', '')
            name = sh.get('name', '')
            rid_to_name[rid] = name

        rels_path = next(
            (n for n in names if n.endswith("xl/_rels/workbook.xml.rels")), None
        )
        rid_to_xml = {}
        if rels_path:
            rels_xml  = z.read(rels_path)
            rels_root = ET.fromstring(rels_xml)
            for rel in rels_root:
                rid      = rel.get('Id', '')
                target   = rel.get('Target', '')
                xml_name = os.path.basename(target)
                rid_to_xml[rid] = xml_name

        xml_to_display = {}
        for rid, sheet_name in rid_to_name.items():
            xml_name = rid_to_xml.get(rid, '')
            if xml_name:
                xml_to_display[xml_name] = sheet_name

        ws_prefix = "xl/worksheets/"
        ws_files  = sorted(n for n in names
                           if n.startswith(ws_prefix) and n.endswith(".xml"))
        for ws_path in ws_files:
            xml_name  = os.path.basename(ws_path)
            content   = z.read(ws_path).decode("utf-8", errors="ignore")
            protected = bool(re.search(r'<sheetProtection', content))
            display   = xml_to_display.get(xml_name, xml_name)
            sheets.append({
                'xml_name':   xml_name,
                'sheet_name': display,
                'protected':  protected,
            })
    return sheets


# ── Core: unlock only selected sheet xml files ────────────────────────────
def remove_protection_selective(file_path, selected_xml_names,
                                create_copy, log_callback):
    base, ext   = os.path.splitext(file_path)
    output_file = (base + "_unlocked" + ext) if create_copy else file_path

    tmp_dir        = tempfile.mkdtemp()
    temp_zip       = os.path.join(tmp_dir, "temp.zip")
    extract_folder = os.path.join(tmp_dir, "extracted")
    result_zip     = os.path.join(tmp_dir, "result.zip")

    try:
        shutil.copy(file_path, temp_zip)
        with zipfile.ZipFile(temp_zip, 'r') as z:
            for member in z.infolist():
                member_path = os.path.realpath(
                    os.path.join(extract_folder, member.filename)
                )
                if not member_path.startswith(os.path.realpath(extract_folder) + os.sep):
                    raise ValueError(f"Zip Slip detected: {member.filename}")
                z.extract(member, extract_folder)

        sheets_path = os.path.join(extract_folder, "xl", "worksheets")
        if not os.path.isdir(sheets_path):
            raise FileNotFoundError("No worksheets folder — is this a valid .xlsx?")

        modified = 0
        for fname in os.listdir(sheets_path):
            if fname not in selected_xml_names:
                continue
            full = os.path.join(sheets_path, fname)
            with open(full, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
            new_content = re.sub(
                r'<sheetProtection[^>]*/?>(?:</sheetProtection>)?',
                '', content
            )
            if new_content != content:
                modified += 1
                log_callback(f"   ↳ unlocked: {fname}", "info")
            with open(full, "w", encoding="utf-8") as f:
                f.write(new_content)

        if modified == 0:
            log_callback("⚠️  Tags not found in selected sheets — saved anyway.", "warn")
        else:
            log_callback(f"🔓  Unlocked {modified} sheet(s).", "ok")

        with zipfile.ZipFile(result_zip, 'w', zipfile.ZIP_DEFLATED) as z:
            for root_dir, _, files in os.walk(extract_folder):
                for fn in files:
                    fp2 = os.path.join(root_dir, fn)
                    z.write(fp2, os.path.relpath(fp2, extract_folder))

        shutil.move(result_zip, output_file)
        action = "Saved copy" if create_copy else "Overwritten"
        log_callback(f"✅  {action}: {os.path.basename(output_file)}", "ok")
        return True, output_file

    except Exception as e:
        log_callback(f"❌  Error: {e}", "error")
        return False, str(e)

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ── SheetRow ──────────────────────────────────────────────────────────────
class SheetRow:
    def __init__(self, parent, sheet_info):
        self.xml_name   = sheet_info['xml_name']
        self.sheet_name = sheet_info['sheet_name']
        self.protected  = sheet_info['protected']
        self.enabled    = tk.BooleanVar(value=self.protected)

        self.frame = tk.Frame(parent, bg=BG4)
        self.frame.pack(fill="x", padx=0, pady=0)

        tk.Label(self.frame, text="", width=3, bg=BG4).pack(side="left")

        icon       = "🔒" if self.protected else "🔓"
        icon_color = WARN if self.protected else TEXT_DIM
        tk.Label(self.frame, text=icon, font=("Segoe UI", 9),
                 bg=BG4, fg=icon_color).pack(side="left")

        self.chk = tk.Checkbutton(
            self.frame,
            variable=self.enabled,
            bg=BG4, fg=TEXT,
            activebackground=BG4,
            selectcolor=BG2,
            bd=0, highlightthickness=0,
            state="normal" if self.protected else "disabled",
            cursor="hand2" if self.protected else "arrow"
        )
        self.chk.pack(side="left", padx=(4, 0))

        color = TEXT if self.protected else TEXT_DIM
        tk.Label(self.frame, text=self.sheet_name,
                 font=FONT_SHEET, bg=BG4, fg=color,
                 anchor="w").pack(side="left", fill="x", expand=True, padx=4)

        badge_text  = "protected" if self.protected else "no lock"
        badge_color = WARN if self.protected else TEXT_DIM
        tk.Label(self.frame, text=badge_text,
                 font=FONT_SM, bg=BG4, fg=badge_color,
                 width=10, anchor="e").pack(side="right", padx=(0, 8))

    def destroy(self):
        self.frame.destroy()


# ── FileRow ───────────────────────────────────────────────────────────────
class FileRow:
    def __init__(self, parent, path, remove_cb, log_cb, canvas_refresh_cb):
        self.path              = path
        self.remove_cb         = remove_cb
        self.log_cb            = log_cb
        self.canvas_refresh_cb = canvas_refresh_cb
        self.enabled           = tk.BooleanVar(value=True)
        self.sheet_rows        = []
        self.expanded          = False
        self.scan_done         = False
        self.success           = False

        self.outer = tk.Frame(parent, bg=BORDER, pady=1)
        self.outer.pack(fill="x", padx=0, pady=2)

        self.header = tk.Frame(self.outer, bg=BG3)
        self.header.pack(fill="x", padx=1, pady=(1, 0))

        self.chk = tk.Checkbutton(
            self.header,
            variable=self.enabled,
            bg=BG3, fg=TEXT,
            activebackground=BG3,
            selectcolor=BG2,
            bd=0, highlightthickness=0,
            cursor="hand2"
        )
        self.chk.pack(side="left", padx=(6, 0))

        self.toggle_btn = tk.Button(
            self.header, text="▶",
            font=("Segoe UI", 8),
            bg=BG3, fg=TEXT_DIM,
            activebackground=BG3, activeforeground=ACCENT,
            bd=0, highlightthickness=0,
            cursor="hand2",
            command=self._toggle
        )
        self.toggle_btn.pack(side="left", padx=(4, 0))

        self.name_lbl = tk.Label(
            self.header,
            text=os.path.basename(path),
            font=FONT_MONO, bg=BG3, fg=TEXT, anchor="w"
        )
        self.name_lbl.pack(side="left", fill="x", expand=True, padx=6)

        self.status_lbl = tk.Label(
            self.header, text="scanning…",
            font=FONT_SM, bg=BG3, fg=TEXT_DIM,
            width=14, anchor="e"
        )
        self.status_lbl.pack(side="right", padx=(0, 8))

        for w in (self.header, self.name_lbl):
            w.bind("<Button-3>", lambda e: self.remove_cb(self))

        self.sheet_panel = tk.Frame(self.outer, bg=BG4)

        threading.Thread(target=self._scan, daemon=True).start()

    def _scan(self):
        try:
            sheets  = scan_xlsx(self.path)
            n_prot  = sum(1 for s in sheets if s['protected'])
            total   = len(sheets)

            def update():
                if n_prot == 0:
                    self.status_lbl.config(text="no protection", fg=TEXT_DIM)
                    self.toggle_btn.config(fg=TEXT_DIM, state="disabled")
                else:
                    self.status_lbl.config(
                        text=f"🔒 {n_prot}/{total} sheets", fg=WARN)

                for s in sheets:
                    row = SheetRow(self.sheet_panel, s)
                    self.sheet_rows.append(row)

                self.scan_done = True
                self.canvas_refresh_cb()

            self.header.after(0, update)

        except Exception as e:
            def show_err():
                self.status_lbl.config(text="scan error", fg=DANGER)
                self.log_cb(
                    f"⚠️  Could not scan {os.path.basename(self.path)}: {e}", "warn")
            self.header.after(0, show_err)

    def _toggle(self):
        if not self.scan_done:
            return
        self.expanded = not self.expanded
        if self.expanded:
            self.sheet_panel.pack(fill="x", padx=1, pady=(0, 1))
            self.toggle_btn.config(text="▼", fg=ACCENT)
        else:
            self.sheet_panel.pack_forget()
            self.toggle_btn.config(text="▶", fg=TEXT_DIM)
        self.canvas_refresh_cb()

    def selected_xml_names(self):
        return [r.xml_name for r in self.sheet_rows
                if r.protected and r.enabled.get()]

    def select_all_sheets(self):
        for r in self.sheet_rows:
            if r.protected:
                r.enabled.set(True)

    def deselect_all_sheets(self):
        for r in self.sheet_rows:
            r.enabled.set(False)

    def set_status(self, text, color):
        # May be called from a worker thread during batch processing —
        # marshal the actual widget mutation onto the Tk main loop.
        self.status_lbl.after(0, self._set_status_direct, text, color)

    def _set_status_direct(self, text, color):
        self.status_lbl.config(text=text, fg=color)
        self.name_lbl.config(fg=color)

    def destroy(self):
        self.outer.destroy()


# ── Main App ───────────────────────────────────────────────────────────────
class ExcelSheetUnlockerApp:
    def __init__(self, root):
        self.root    = root
        self.rows    = []
        self._running = False
        self._setup_window()
        self._build_ui()

    def _setup_window(self):
        self.root.title("Excel Sheet Unlocker")
        self.root.configure(bg=BG)
        try:
            self.root.iconbitmap(_resource("icon.ico"))
        except Exception:
            pass
        self.root.resizable(True, True)
        w, h = 700, 700
        self.root.geometry(f"{w}x{h}")
        self.root.minsize(520, 560)
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        PAD = 24

        # Header
        hdr = tk.Frame(self.root, bg=BG)
        hdr.pack(fill="x", padx=PAD, pady=(20, 10))
        tk.Label(hdr, text="🔓", font=("Segoe UI", 22),
                 bg=BG, fg=ACCENT).pack(side="left")
        tk.Label(hdr, text="  Excel Sheet Unlocker",
                 font=("Segoe UI", 18, "bold"),
                 bg=BG, fg=TEXT).pack(side="left")
        tk.Label(hdr, text="  —  remove sheet protection",
                 font=("Segoe UI", 10), bg=BG, fg=TEXT_DIM
                 ).pack(side="left", pady=(5, 0))

        # Drop zone
        dz_border = tk.Frame(self.root, bg=BORDER)
        dz_border.pack(fill="x", padx=PAD)
        self.drop_frame = tk.Frame(dz_border, bg=BG2, pady=18)
        self.drop_frame.pack(fill="both", padx=1, pady=1)
        self.dz_icon = tk.Label(self.drop_frame, text="📂",
                                font=("Segoe UI", 24), bg=BG2, fg=ACCENT)
        self.dz_icon.pack()
        hint = ("Drop .xlsx files here  •  click to browse  •  multiple files OK"
                if DND_AVAILABLE else
                "Click to browse .xlsx files  •  multiple files OK")
        self.dz_label = tk.Label(self.drop_frame, text=hint,
                                 font=("Segoe UI", 10), bg=BG2, fg=TEXT_DIM)
        self.dz_label.pack(pady=(4, 0))
        for w in (self.drop_frame, self.dz_icon, self.dz_label):
            w.bind("<Button-1>", lambda e: self._browse())
            w.bind("<Enter>",    lambda e: self._hover(True))
            w.bind("<Leave>",    lambda e: self._hover(False))
        if DND_AVAILABLE:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)

        # Options row
        opt = tk.Frame(self.root, bg=BG)
        opt.pack(fill="x", padx=PAD, pady=(10, 0))
        self.create_copy_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            opt, text="  Create copy  (_unlocked.xlsx next to original)",
            variable=self.create_copy_var,
            font=FONT_MAIN, bg=BG, fg=TEXT,
            activebackground=BG, activeforeground=TEXT,
            selectcolor=BG2, bd=0, highlightthickness=0,
            cursor="hand2", command=self._update_copy_hint
        ).pack(side="left")
        self.copy_hint = tk.Label(opt, text="", font=FONT_SM, bg=BG, fg=WARN)
        self.copy_hint.pack(side="left", padx=(8, 0), pady=(2, 0))
        tk.Label(opt, text="▶ click arrow to expand sheets  •  right-click file to remove",
                 font=FONT_SM, bg=BG, fg=TEXT_DIM).pack(side="right")
        self._update_copy_hint()

        # List header
        lh = tk.Frame(self.root, bg=BG)
        lh.pack(fill="x", padx=PAD, pady=(10, 4))
        self.count_lbl = tk.Label(lh, text="No files added",
                                  font=FONT_BOLD, bg=BG, fg=TEXT_DIM)
        self.count_lbl.pack(side="left")
        for label, cmd, color in (
            ("Select all sheets",   self._select_all,   ACCENT),
            ("Deselect all sheets", self._deselect_all, INDIGO),
            ("Clear all",           self._clear_files,  DANGER),
        ):
            tk.Button(lh, text=label, font=FONT_SM,
                      bg=BG, fg=color,
                      activebackground=BG, activeforeground=color,
                      bd=0, cursor="hand2", command=cmd
                      ).pack(side="right", padx=(0, 10))

        # Scrollable file list
        lb_border = tk.Frame(self.root, bg=BORDER)
        lb_border.pack(fill="both", expand=True, padx=PAD)
        lb_outer = tk.Frame(lb_border, bg=BG3)
        lb_outer.pack(fill="both", expand=True, padx=1, pady=1)

        self.canvas = tk.Canvas(lb_outer, bg=BG3, bd=0, highlightthickness=0, height=180)
        self.canvas.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(lb_outer, orient="vertical",
                          bg=BG3, troughcolor=BG3, bd=0,
                          command=self.canvas.yview)
        sb.pack(side="right", fill="y")
        self.canvas.config(yscrollcommand=sb.set)

        self.list_inner = tk.Frame(self.canvas, bg=BG3)
        self._cw = self.canvas.create_window(
            (0, 0), window=self.list_inner, anchor="nw")
        self.list_inner.bind("<Configure>", self._on_list_cfg)
        self.canvas.bind("<Configure>",     self._on_canvas_cfg)
        self.canvas.bind("<MouseWheel>",    self._on_scroll)

        # Log
        log_border = tk.Frame(self.root, bg=BORDER)
        log_border.pack(fill="x", padx=PAD, pady=(8, 0))
        log_inner = tk.Frame(log_border, bg=BG2)
        log_inner.pack(fill="both", padx=1, pady=1)
        self.log = tk.Text(log_inner, height=4,
                           bg=BG2, fg=TEXT_DIM, font=FONT_MONO,
                           bd=0, highlightthickness=0,
                           state="disabled", wrap="word")
        self.log.pack(fill="both", padx=8, pady=6)
        self.log.tag_config("ok",    foreground=ACCENT)
        self.log.tag_config("info",  foreground=TEXT_DIM)
        self.log.tag_config("warn",  foreground=WARN)
        self.log.tag_config("error", foreground=DANGER)

        # Progress + button + footer — declared bottom-up
        self._style_progress()

        footer = tk.Frame(self.root, bg=BG)
        footer.pack(side="bottom", fill="x", padx=PAD, pady=(0, 8))
        tk.Label(
            footer,
            text="© 2026 Danilo Di Pietrantonio  •  Modification requires authorization  •  tangina.software@gmail.com",
            font=FONT_SM, bg=BG, fg=TEXT_DIM,
            wraplength=480, justify="center"
        ).pack()

        self.run_btn = tk.Button(
            self.root, text="🔓   Unlock Selected Sheets",
            font=("Segoe UI", 12, "bold"),
            bg=ACCENT, fg="#0f1117",
            activebackground=ACCENT2, activeforeground="#0f1117",
            bd=0, pady=12, cursor="hand2",
            command=self._run
        )
        self.run_btn.pack(side="bottom", fill="x", padx=PAD, pady=(4, 0))

        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(side="bottom", fill="x", padx=PAD, pady=(10, 0))

        self._log("Ready — add .xlsx files, expand to pick sheets, then click Unlock.", "info")

    # ── Style ─────────────────────────────────────────────────────────────
    def _style_progress(self):
        s = ttk.Style()
        s.theme_use("default")
        s.configure("TProgressbar",
                     troughcolor=BG3, background=ACCENT,
                     bordercolor=BG3, lightcolor=ACCENT,
                     darkcolor=ACCENT, thickness=6)

    # ── Canvas helpers ────────────────────────────────────────────────────
    def _on_list_cfg(self, e):
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_cfg(self, e):
        self.canvas.itemconfig(self._cw, width=e.width)

    def _on_scroll(self, e):
        self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    def _refresh_canvas(self):
        self.list_inner.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    # ── Misc ──────────────────────────────────────────────────────────────
    def _update_copy_hint(self):
        self.copy_hint.config(
            text="" if self.create_copy_var.get()
            else "⚠ original will be overwritten"
        )

    def _hover(self, on):
        c = BG3 if on else BG2
        self.drop_frame.config(bg=c)
        self.dz_icon.config(bg=c)
        self.dz_label.config(bg=c)

    def _log(self, msg, tag="info"):
        # May be called from a worker thread (batch processing) or the
        # main thread (UI actions) — always marshal onto the Tk main loop.
        self.root.after(0, self._log_direct, msg, tag)

    def _log_direct(self, msg, tag):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    # ── File management ───────────────────────────────────────────────────
    def _browse(self):
        paths = filedialog.askopenfilenames(
            title="Select Excel files",
            filetypes=[("Excel files", "*.xlsx *.xlsm")]
        )
        for p in paths:
            self._add_file(p)

    def _on_drop(self, event):
        for p in self.root.tk.splitlist(event.data):
            if p.lower().endswith((".xlsx", ".xlsm")):
                self._add_file(p)
            else:
                self._log(f"⚠️  Skipped (not .xlsx/.xlsm): {os.path.basename(p)}", "warn")

    def _add_file(self, path):
        if any(r.path == path for r in self.rows):
            return
        row = FileRow(self.list_inner, path,
                      self._remove_row, self._log, self._refresh_canvas)
        self.rows.append(row)
        self._update_count()

    def _remove_row(self, row):
        row.destroy()
        self.rows.remove(row)
        self._update_count()
        self._refresh_canvas()

    def _clear_files(self):
        for r in self.rows:
            r.destroy()
        self.rows.clear()
        self._update_count()
        self._refresh_canvas()

    def _select_all(self):
        for r in self.rows:
            r.enabled.set(True)
            r.select_all_sheets()

    def _deselect_all(self):
        for r in self.rows:
            r.enabled.set(False)
            r.deselect_all_sheets()

    def _update_count(self):
        n = len(self.rows)
        self.count_lbl.config(
            text="No files added" if n == 0
            else f"{n} file{'s' if n > 1 else ''} loaded  ✓",
            fg=TEXT_DIM if n == 0 else ACCENT
        )

    # ── Processing ────────────────────────────────────────────────────────
    def _run(self):
        if self._running:
            return
        active = [r for r in self.rows
                  if r.enabled.get() and r.selected_xml_names()]
        if not active:
            messagebox.showwarning(
                "Nothing to process",
                "No protected sheets are selected.\n"
                "Expand a file (▶) and check the sheets you want to unlock."
            )
            return
        self._running = True
        self.run_btn.config(state="disabled", bg=BG3, fg=TEXT_DIM,
                            text="⏳   Processing...")
        threading.Thread(
            target=self._process_all,
            args=(active, self.create_copy_var.get()),
            daemon=True
        ).start()

    def _process_all(self, active_rows, create_copy):
        total = len(active_rows)
        ok    = 0
        self.root.after(0, self.progress.configure, {"maximum": total, "value": 0})

        for i, row in enumerate(active_rows, 1):
            selected = row.selected_xml_names()
            self._log(
                f"\n── [{i}/{total}] {os.path.basename(row.path)}"
                f"  ({len(selected)} sheet(s) selected)", "info"
            )
            row.set_status("working…", WARN)

            success, _ = remove_protection_selective(
                row.path, selected, create_copy, self._log
            )

            row.success = success
            if success:
                ok += 1
                row.set_status("✓ done", ACCENT)
            else:
                row.set_status("✗ error", DANGER)

            self.root.after(0, self.progress.configure, {"value": i})

        tag = "ok" if ok == total else "warn"
        self._log(f"\n🏁  Done — {ok}/{total} file(s) processed.", tag)

        def autoclear():
            for row in list(self.rows):
                if row.success:
                    row.destroy()
                    self.rows.remove(row)
            self._update_count()
            self._refresh_canvas()
            self.progress["value"] = 0
            self._running = False

        self.root.after(1200, autoclear)
        self.root.after(0, self.run_btn.config, {
            "state": "normal", "bg": ACCENT, "fg": "#0f1117",
            "text": "🔓   Unlock Selected Sheets"
        })


# ── Entry point ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    ExcelSheetUnlockerApp(root)
    root.mainloop()