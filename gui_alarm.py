#!/usr/bin/env python3
"""
GUI untuk import Excel -> Free Alarm Clock + Alarm standalone.
Jalan di Mac / Windows / Linux, tanpa install tambahan selain openpyxl.

  pip install openpyxl
  python3 gui_alarm.py
"""
import sys
import time
import threading
import subprocess
import platform
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

sys.path.insert(0, str(Path(__file__).parent))
from excel_to_freealarm import baca_excel

APP_TITLE = "Excel → Free Alarm Clock"

class AlarmGUI:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("900x620")
        self.excel_path = tk.StringVar(value=str(Path(__file__).parent / "contoh_jadwal.xlsx"))
        self.alarms = []
        self.running = False
        self.run_thread = None

        # --- Bar atas: file ---
        top = ttk.Frame(root, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="File Excel:").pack(side="left")
        ttk.Entry(top, textvariable=self.excel_path, width=55).pack(side="left", padx=6)
        ttk.Button(top, text="Pilih…", command=self.pilih_file).pack(side="left")
        ttk.Button(top, text="Muat", command=self.muat).pack(side="left", padx=4)
        ttk.Button(top, text="Template", command=self.buat_template).pack(side="left")

        # --- Toolbar ---
        bar = ttk.Frame(root, padding=(8, 0))
        bar.pack(fill="x")
        ttk.Button(bar, text="1. Cek Validasi", command=self.muat).pack(side="left")
        ttk.Button(bar, text="2. Jalankan Alarm (Mac/Win)", command=self.start_alarm).pack(side="left", padx=4)
        ttk.Button(bar, text="Stop", command=self.stop_alarm).pack(side="left")
        ttk.Button(bar, text="3. Export .alm instan", command=self.export_alm).pack(side="left", padx=4)
        self.btn_win = ttk.Button(bar, text="4. Kirim via klik Add (legacy, jangan dipakai)", command=self.kirim_windows)
        self.btn_win.pack(side="left", padx=4)
        if sys.platform != "win32":
            self.btn_win.state(["disabled"])
            ttk.Label(bar, text="(khusus Windows)").pack(side="left")
        self.status = tk.StringVar(value="Siap. Klik Muat.")
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
        self.tree = ttk.Treeview(root, columns=cols, show="headings", height=12)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=130 if c != "Label" else 200)
        self.tree.pack(fill="both", expand=True, padx=8)
        sb = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)

        # --- Form tambah manual ---
        f = ttk.LabelFrame(root, text="Tambah manual → simpan ke Excel", padding=8)
        f.pack(fill="x", padx=8, pady=6)
        self.e_waktu = self._field(f, "Waktu HH:MM", 0, "07:00")
        self.e_label = self._field(f, "Label", 1, "Bel Masuk", w=25)
        self.e_hari = self._field(f, "Hari", 2, "Senin,Selasa,Rabu,Kamis,Jumat", w=30)
        self.e_tipe = self._field(f, "Tipe", 3, "Weekly")
        self.e_suara = self._field(f, "Suara", 4, "School.mp3")
        ttk.Button(f, text="+ Tambah", command=self.tambah_manual).grid(row=1, column=5, padx=6)

        # --- Log ---
        self.log = tk.Text(root, height=8)
        self.log.pack(fill="both", expand=False, padx=8, pady=(0, 8))
        self.log_msg("GUI siap. Pilih file Excel lalu klik Muat.\n")

    def _field(self, parent, label, col, default="", w=15):
        ttk.Label(parent, text=label).grid(row=0, column=col, sticky="w")
        v = tk.StringVar(value=default)
        ttk.Entry(parent, textvariable=v, width=w).grid(row=1, column=col, padx=4)
        return v

    def log_msg(self, s):
        self.log.insert("end", s)
        self.log.see("end")

    def set_progress(self, done, total):
        pct = int(done * 100 / total) if total else 0
        self.prog["value"] = pct
        self.prog_label.set(f"{pct}% ({done}/{total})")
        self.status.set(f"Proses {done}/{total} ({pct}%)" + (" — SELESAI 100%" if done >= total and total else ""))

    def pilih_file(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if p:
            self.excel_path.set(p)
            self.muat()

    def buat_template(self):
        import buat_template  # noqa
        self.log_msg("Template dibuat: contoh_jadwal.xlsx\n")
        self.muat()

    def muat(self):
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
        self.status.set(f"{len(self.alarms)} alarm dimuat dari {Path(p).name}")
        self.log_msg(f"Dimuat {len(self.alarms)} alarm:\n")
        for a in self.alarms:
            self.log_msg(f"  {a['waktu']} | {a['tipe']} | {','.join(a['weekdays_en'])} | {a['label']}\n")

    def tambah_manual(self):
        from openpyxl import load_workbook
        p = self.excel_path.get()
        try:
            wb = load_workbook(p)
            ws = wb["Jadwal"] if "Jadwal" in wb.sheetnames else wb.active
            ws.append([self.e_waktu.get(), self.e_label.get(), self.e_hari.get(),
                       self.e_tipe.get(), self.e_suara.get(), self.e_label.get()])
            wb.save(p)
            self.log_msg(f"Ditambah: {self.e_waktu.get()} - {self.e_label.get()}\n")
            self.muat()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # --- Alarm standalone (thread) ---
    def cocok_hari_ini(self, a):
        if a["tipe"] == "Daily":
            return True
        if a["tipe"] == "Once":
            if not a["weekdays_en"]:
                return True
            return datetime.now().strftime("%Y-%m-%d") in a["weekdays_en"]
        if "Every day" in a["weekdays_en"]:
            return True
        return datetime.now().strftime("%A") in a["weekdays_en"]

    def bunyi(self, a):
        self.log_msg(f"\n🔔 {datetime.now():%H:%M:%S} — {a['label']} ({a['waktu']})\n")
        try:
            if Path(a["suara"]).exists():
                if platform.system() == "Darwin":
                    subprocess.Popen(["afplay", a["suara"]])
                    return
                elif platform.system() == "Windows":
                    import winsound
                    winsound.PlaySound(a["suara"], winsound.SND_FILENAME | winsound.SND_ASYNC)
                    return
        except Exception as e:
            self.log_msg(f"(gagal putar suara: {e})\n")
        try:
            if platform.system() == "Windows":
                import winsound
                winsound.Beep(880, 600)
            else:
                subprocess.run(["afplay", "/System/Library/Sounds/Glass.aiff"], timeout=5)
        except Exception:
            self.root.bell()

    def start_alarm(self):
        if not self.alarms:
            self.muat()
        if not self.alarms:
            return
        if self.running:
            messagebox.showinfo("Info", "Alarm sudah berjalan.")
            return
        self.running = True
        self.status.set("Alarm berjalan… (Stop untuk berhenti)")
        self.log_msg(f"Alarm berjalan. Hari ini: {datetime.now():%A %Y-%m-%d}\n")
        self.run_thread = threading.Thread(target=self._loop, daemon=True)
        self.run_thread.start()

    def _loop(self):
        sudah = set()
        while self.running:
            now = datetime.now().strftime("%H:%M")
            today = datetime.now().strftime("%Y-%m-%d")
            for a in self.alarms:
                if not self.cocok_hari_ini(a):
                    continue
                key = (today, a["waktu"], a["label"])
                if a["waktu"] == now and key not in sudah:
                    sudah.add(key)
                    self.root.after(0, self.bunyi, a)
                    self.root.after(0, lambda m=f"✔ Berbunyi: {a['waktu']} {a['label']}\n": self.log_msg(m))
            time.sleep(10)

    def stop_alarm(self):
        self.running = False
        self.status.set("Berhenti.")
        self.log_msg("Alarm dihentikan.\n")

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
            Path(out).write_text("\ufeff" + "\r\n".join(lines) + "\r\n", encoding="utf-8")
            self.log_msg(f"OK: {len(lines)} alarm -> {out} (100%)\n")
            self.log_msg("Lanjut di FreeAlarmClock: File > Restore > pilih file itu.\n")
            messagebox.showinfo("Selesai 100%",
                f"{len(lines)} alarm tersimpan ke:\n{out}\n\nBuka FreeAlarmClock > File > Restore > pilih file itu.")
        except Exception as e:
            messagebox.showerror("Gagal export", str(e))

    def kirim_windows(self):
        if not messagebox.askokcancel("Jalur lama",
            "Jalur klik-Add ini TIDAK mengisi jam/label otomatis (hanya buka dialog).\n"
            "Yang 09:59 (None) itu dialog kosong yang ter-OK.\n\nLanjut pakai jalur lama?\n"
            "Disarankan BATAL lalu pakai 'Export .alm instan' + File > Restore."):
            return
        try:
            from excel_to_freealarm import run_otomasi
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return
        if not self.alarms:
            self.muat()
        self.log_msg("Menghubungi Free Alarm Clock… jangan sentuh mouse.\n")
        t = threading.Thread(target=self._kirim_thread, daemon=True)
        t.start()

    def _kirim_thread(self):
        total = len(self.alarms)
        self.root.after(0, self.set_progress, 0, total)
        self.root.after(0, lambda: self.log_msg("Menghubungi Free Alarm Clock (timeout 10 dtk)...\n"))
        def cb(done, total_, alarm, ok):
            self.root.after(0, self.set_progress, done, total_)
            self.root.after(0, self.log_msg,
                f"[{'OK' if ok else 'GAGAL'} {done}/{total_} {done*100//total_}%] {alarm['waktu']} {alarm['label']}\n"
                + ("  -> Tekan OK di dialog untuk simpan, lanjut otomatis.\n" if ok else ""))
        def log_cb(m):
            self.root.after(0, self.log_msg, m)
        try:
            from excel_to_freealarm import run_otomasi
            ok_count = run_otomasi(self.alarms, delay=1.0, on_progress=cb,
                                   interactive=False, log=log_cb)
            self.root.after(0, self.set_progress, total, total)
            self.root.after(0, lambda: self.log_msg(
                f"\n===== SELESAI 100%: {ok_count}/{total} dialog Add terbuka =====\n"
                f"Verifikasi 100%: 1) jumlah baris di FreeAlarmClock = {total}, "
                f"2) cek 'Next 3 Alarms' di panel bawah, 3) File > Backup.\n"))
            self.root.after(0, lambda: messagebox.showinfo(
                "Selesai 100%", f"{ok_count}/{total} alarm diproses.\nCek list di FreeAlarmClock + Next 3 Alarms."))
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Gagal otomasi", str(e) + "\nPastikan FreeAlarmClock terbuka di Windows."))

def main():
    root = tk.Tk()
    # tema lebih modern bila tersedia
    try:
        root.tk.call("tk", "windowingsystem")
    except Exception:
        pass
    AlarmGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
