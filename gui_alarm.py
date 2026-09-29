#!/usr/bin/env python3
"""
GUI Excel -> *.alm (Free Alarm Clock) — hanya jalur generate .alm instan.
Jalan di Mac / Windows / Linux, tanpa install tambahan selain openpyxl.

  pip install openpyxl
  python3 gui_alarm.py

Alur: Pilih Excel -> Muat -> Export .alm -> Restore di FreeAlarmClock.
"""
import sys
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

sys.path.insert(0, str(Path(__file__).parent))
from excel_to_freealarm import baca_excel

APP_TITLE = "Excel → Free Alarm Clock (.alm)"

class AlarmGUI:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("900x560")
        self.excel_path = tk.StringVar(value=str(Path(__file__).parent / "contoh_ews_rj.xlsx"))
        self.alarms = []

        # --- Bar atas: file ---
        top = ttk.Frame(root, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="File Excel:").pack(side="left")
        ttk.Entry(top, textvariable=self.excel_path, width=55).pack(side="left", padx=6)
        ttk.Button(top, text="Pilih…", command=self.pilih_file).pack(side="left")
        ttk.Button(top, text="Muat", command=self.muat).pack(side="left", padx=4)

        # --- Toolbar: hanya generate .alm ---
        bar = ttk.Frame(root, padding=(8, 0))
        bar.pack(fill="x")
        ttk.Button(bar, text="1. Muat / Validasi", command=self.muat).pack(side="left")
        ttk.Button(bar, text="2. Export .alm instan", command=self.export_alm).pack(side="left", padx=4)
        self.status = tk.StringVar(value="Siap. Pilih file Excel lalu klik Muat.")
        ttk.Label(root, textvariable=self.status, padding=(8, 4)).pack(fill="x")

        # --- Progress 0-100% ---
        prow = ttk.Frame(root, padding=(8, 0))
        prow.pack(fill="x")
        self.prog = ttk.Progressbar(prow, mode="determinate", maximum=100)
        self.prog.pack(side="left", fill="x", expand=True)
        self.prog_label = tk.StringVar(value="0% (0/0)")
        ttk.Label(prow, textvariable=self.prog_label, width=16).pack(side="left", padx=6)

        # --- Tabel ---
        cols = ("Waktu", "Label", "Tipe", "Hari", "Suara", "Pesan")
        self.tree = ttk.Treeview(root, columns=cols, show="headings", height=14)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=130 if c != "Label" else 200)
        self.tree.pack(fill="both", expand=True, padx=8)
        sb = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)

        # --- Log ---
        self.log = tk.Text(root, height=8)
        self.log.pack(fill="both", expand=False, padx=8, pady=(8, 8))
        self.log_msg("GUI siap. Pilih file Excel lalu klik Muat.\n")

    def log_msg(self, s):
        self.log.insert("end", s)
        self.log.see("end")

    def set_progress(self, done, total):
        pct = int(done * 100 / total) if total else 0
        self.prog["value"] = pct
        self.prog_label.set(f"{pct}% ({done}/{total})")
        if total:
            self.status.set(f"Proses {done}/{total} ({pct}%)"
                            + (" — SELESAI 100%" if done >= total else ""))

    def reset_progress(self):
        self.prog["value"] = 0
        self.prog_label.set("0% (0/0)")

    def pilih_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p:
            self.excel_path.set(p)
            self.muat()

    def muat(self):
        # Reset progress setiap load Excel baru
        self.reset_progress()
        p = self.excel_path.get()
        if not Path(p).exists():
            messagebox.showerror("Error", f"File tidak ketemu:\n{p}")
            return
        try:
            self.alarms = baca_excel(p)
        except Exception as e:
            messagebox.showerror("Error baca Excel", str(e))
            return
        for i in self.tree.get_children():
            self.tree.delete(i)
        for a in self.alarms:
            self.tree.insert("", "end", values=(
                a["waktu"], a["label"], a["tipe"],
                ",".join(a["weekdays_en"]), a["suara"], a["pesan"]))
        self.status.set(f"{len(self.alarms)} alarm dimuat dari {Path(p).name} — siap export.")
        self.log_msg(f"Dimuat {len(self.alarms)} alarm (progress di-reset 0%):\n")
        for a in self.alarms:
            self.log_msg(f"  {a['waktu']} | {a['tipe']} | {','.join(a['weekdays_en'])} | {a['label']}\n")

    def export_alm(self):
        if not self.alarms:
            self.muat()
        if not self.alarms:
            return
        out = filedialog.asksaveasfilename(defaultextension=".alm",
            filetypes=[("Alarm backup", "*.alm")], initialfile="hasil.alm")
        if not out:
            return
        try:
            from excel_to_alm import alarm_ke_baris
            lines = []
            for i, a in enumerate(self.alarms, 1):
                lines.append(alarm_ke_baris(a, datetime.now()))
                self.set_progress(i, len(self.alarms))
            with open(out, "w", encoding="utf-8-sig", newline="") as f:
                f.write("\r\n".join(lines) + "\r\n")
            self.log_msg(f"OK: {len(lines)} alarm -> {out} (100%)\n")
            self.log_msg("Lanjut di FreeAlarmClock: hapus list lama, File > Restore > pilih file itu.\n")
            messagebox.showinfo("Selesai 100%",
                f"{len(lines)} alarm tersimpan ke:\n{out}\n\nDi FreeAlarmClock: hapus list lama, File > Restore > pilih file itu.")
        except Exception as e:
            messagebox.showerror("Gagal export", str(e))

def main():
    root = tk.Tk()
    try:
        root.tk.call("tk", "windowingsystem")
    except Exception:
        pass
    AlarmGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
