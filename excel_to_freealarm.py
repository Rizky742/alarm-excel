#!/usr/bin/env python3
"""
Import jadwal dari Excel ke Free Alarm Clock (https://freealarmapp.com/)
Format Excel DINAMIS: urutan kolom bebas, nama header fleksibel ID/EN.

Header yang dikenali (case-insensitive, boleh acak, boleh sebagian):
  Waktu: Jam/Waktu/Time/Pukul | Label: Nama/Label/Kegiatan/Acara
  Hari: Hari/Day | Tanggal: Tanggal/Date/Tgl
  Tipe: Tipe/Jenis/Occurs (opsional, auto bila kosong)
  Suara: Suara/Sound/MP3 (opsional) | Pesan: Pesan/Note (opsional)

Contoh semua valid:
  Jam | Kegiatan | Hari                -> Tipe auto=Weekly, Suara default
  Time | Title | Date                  -> Tipe auto=Once
  Pukul | Nama | Senin | Selasa | ...  -> wide-format checklist v/x/1
  Waktu | Label | Hari | Tipe | Suara | Pesan  (format lengkap lama)

  Tipe: Once / Daily / Weekly (v5.3 juga dukung Monthly, Yearly)
  Hari untuk Weekly: Senin,Selasa,Rabu,Kamis,Jumat,Sabtu,Minggu
    atau: Weekdays (=Senin-Jumat), Weekends (=Sabtu-Minggu), Setiap Hari / Every day
  Hari untuk Once: tanggal YYYY-MM-DD (opsional, kosong = hari ini/besok ikut jam)
  Hari untuk Daily: kosongkan saja.

Cara pakai:
  # 1. Test dulu (bisa di Mac):
  python excel_to_freealarm.py contoh_jadwal.xlsx --dry-run

  # 2. Di Windows, install:
  pip install openpyxl pywinauto

  # 3. Buka FreeAlarmClock.exe, lalu:
  python excel_to_freealarm.py contoh_jadwal.xlsx --inspect   (lihat nama tombol, sekali saja)
  python excel_to_freealarm.py contoh_jadwal.xlsx             (otomasi tambah alarm)

Kenapa otomasi GUI, bukan tulis Data.ini langsung?
  FreeAlarmClock menyimpan setting di Registry (HKCU\\SOFTWARE\\ComfortSoftware\\FreeAC)
  atau Data.ini (versi Portable) dalam format biner proprietary + backup *.alm
  yang tidak didokumentasikan. Menulis langsung berisiko merusak setting.
  Cara aman yang didukung resmi: lewat jendela "Add" / "-addalarm".
"""

import sys
import re
import argparse
from datetime import datetime, time as dtime, date as ddate
from pathlib import Path

HARI_MAP = {
    "senin": "Monday", "monday": "Monday", "mon": "Monday",
    "selasa": "Tuesday", "tuesday": "Tuesday", "tue": "Tuesday", "tues": "Tuesday",
    "rabu": "Wednesday", "wednesday": "Wednesday", "wed": "Wednesday",
    "kamis": "Thursday", "thursday": "Thursday", "thu": "Thursday", "thur": "Thursday", "thurs": "Thursday",
    "jumat": "Friday", "jumat ": "Friday", "friday": "Friday", "fri": "Friday", "jum'at": "Friday",
    "sabtu": "Saturday", "saturday": "Saturday", "sat": "Saturday",
    "minggu": "Sunday", "ahad": "Sunday", "sunday": "Sunday", "sun": "Sunday", "mgg": "Sunday",
}

# Header dinamis: canonical -> daftar alias (sudah dinormalisasi)
HEADER_ALIASES = {
    "waktu": ["waktu", "jam", "time", "pukul", "pukuljam", "hour", "hhmm", "waktuhhmm"],
    "label": ["label", "nama", "name", "kegiatan", "acara", "keterangan", "judul", "title", "aktivitas", "activity", "bel", "alarm"],
    "hari": ["hari", "day", "days", "haribel", "jadwalhari", "weekday"],
    "tanggal": ["tanggal", "date", "tgl"],
    "tipe": ["tipe", "type", "jenis", "occurs", "repeat", "pengulangan", "mode"],
    "suara": ["suara", "sound", "musik", "music", "mp3", "file", "nada", "bunyi", "audio"],
    "pesan": ["pesan", "message", "msg", "catatan", "note", "deskripsi", "description", "reminder"],
}
# Kolom hari wide-format: nama kolom -> English day
WIDE_DAY_COLS = {
    "senin": "Monday", "monday": "Monday", "mon": "Monday",
    "selasa": "Tuesday", "tuesday": "Tuesday", "sel": "Tuesday",
    "rabu": "Wednesday", "wednesday": "Wednesday", "rab": "Wednesday",
    "kamis": "Thursday", "thursday": "Thursday", "kam": "Thursday",
    "jumat": "Friday", "friday": "Friday", "jum": "Friday", "jumat'": "Friday",
    "sabtu": "Saturday", "saturday": "Saturday", "sab": "Saturday",
    "minggu": "Sunday", "sunday": "Sunday", "mgg": "Sunday", "min": "Sunday",
}

def parse_args():
    p = argparse.ArgumentParser(description="Import Excel -> Free Alarm Clock")
    p.add_argument("excel", help="File .xlsx jadwal")
    p.add_argument("--sheet", default=None, help="Nama sheet (default: sheet aktif / 'Jadwal')")
    p.add_argument("--dry-run", action="store_true", help="Hanya baca Excel + tampilkan, tanpa otomasi")
    p.add_argument("--inspect", action="store_true", help="List kontrol GUI FreeAlarmClock (Windows saja)")
    p.add_argument("--delay", type=float, default=1.0, help="Jeda antar langkah GUI (detik)")
    p.add_argument("--skip-first-row-if-header", action="store_true", default=True)
    return p.parse_args()

def _norm_header(s):
    if s is None:
        return ""
    s = str(s).strip().lower()
    # buang isi kurung: "Waktu (HH:MM)" -> "waktu"
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"[^a-z0-9]", "", s)
    return s

def _deteksi_kolom(headers):
    """Map index kolom -> nama kanonis. Kembalikan (mapping, wide_day_idx)."""
    mapping = {}
    wide_day_idx = {}  # idx -> English day
    for i, h in enumerate(headers):
        n = _norm_header(h)
        if not n:
            continue
        found = None
        for canon, aliases in HEADER_ALIASES.items():
            if n in aliases:
                found = canon
                break
        if found:
            # jangan timpa kalau sudah ada (ambil pertama)
            if found not in mapping.values():
                mapping[i] = found
            continue
        # cek wide day columns: "Senin", "Selasa", ...
        if n in WIDE_DAY_COLS and i not in mapping:
            wide_day_idx[i] = WIDE_DAY_COLS[n]
    return mapping, wide_day_idx

def _is_truthy(v):
    if v is None:
        return False
    s = str(v).strip().lower()
    return s in ("1", "v", "x", "✓", "ya", "y", "yes", "true", "ok", "•", "*")

def baca_excel(path, sheet=None):
    """Parser DINAMIS: urutan kolom bebas, nama header fleksibel (ID/EN), opsional.

    Didukung:
    - Header alias: Jam/Waktu/Time/Pukul | Nama/Label/Kegiatan | Hari/Day |
      Tanggal/Date | Tipe/Jenis/Occurs | Suara/Sound/MP3 | Pesan/Note
    - Tanpa header (posisi legacy): Waktu | Label | Hari | Tipe | Suara | Pesan
    - Wide-format: kolom Senin..Minggu terpisah berisi checklist (v/x/1/Ya)
    - Sel Excel bertipe Time/Date/datetime/float otomatis terbaca
    - Kolom opsional boleh hilang (default: Tipe=auto, Suara=School.mp3, Pesan=Label)
    - Semua sheet bisa dipilih via --sheet, default: sheet 'Jadwal' atau aktif
    """
    try:
        import openpyxl
    except ImportError:
        print("Install dulu: pip install openpyxl")
        sys.exit(1)
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else (wb["Jadwal"] if "Jadwal" in wb.sheetnames else wb.active)
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []

    # Cari baris header: baris yang mengandung kolom waktu / label
    header_idx = None
    mapping, wide_day_idx = {}, {}
    for ri in range(min(5, len(rows))):
        m, w = _deteksi_kolom(rows[ri] if rows[ri] else [])
        if "waktu" in m.values() or ("label" in m.values() and "hari" in m.values()) or w:
            header_idx = ri
            mapping, wide_day_idx = m, w
            break
    if header_idx is None:
        # Legacy tanpa header: posisi tetap
        start = 0
        # cek baris pertama apakah header teks "waktu..."
        if rows[0] and rows[0][0] is not None and _norm_header(rows[0][0]) in HEADER_ALIASES["waktu"]:
            start = 1
        return _baca_posisional(rows, start)

    print(f"[INFO] Header terdeteksi di baris {header_idx+1}: {mapping}"
          + (f" + kolom hari wide: {wide_day_idx}" if wide_day_idx else ""))
    hasil = []
    idx_of = {v: k for k, v in mapping.items()}  # canon -> idx
    for i, r in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if r is None or all(c is None or str(c).strip() == "" for c in r):
            continue
        get = lambda canon: (r[idx_of[canon]] if canon in idx_of and idx_of[canon] < len(r) else None)

        waktu_raw = get("waktu")
        if waktu_raw is None or (isinstance(waktu_raw, str) and waktu_raw.strip() == ""):
            continue
        waktu = normalisasi_waktu(waktu_raw)
        if not waktu:
            print(f"[BARIS {i}] SKIP: Waktu tidak valid: {waktu_raw!r} (pakai HH:MM / format Time Excel)")
            continue

        label = get("label")
        label = str(label).strip() if label not in (None, "") else f"Alarm {waktu}"
        hari_raw = get("hari")
        tanggal_raw = get("tanggal")
        tipe_raw = get("tipe")
        suara = get("suara")
        pesan = get("pesan")
        hari = str(hari_raw).strip() if hari_raw not in (None, "") else ""
        tanggal = str(tanggal_raw).strip() if tanggal_raw not in (None, "") else ""
        tipe = str(tipe_raw).strip().capitalize() if tipe_raw not in (None, "") else ""
        suara = str(suara).strip() if suara not in (None, "") else "School.mp3"
        pesan = str(pesan).strip() if pesan not in (None, "") else label

        # Wide-format: gabung checklist Senin..Minggu
        if wide_day_idx:
            checked = [en for idx, en in wide_day_idx.items()
                       if idx < len(r) and _is_truthy(r[idx])]
            if checked:
                hari = ",".join(checked)
                if not tipe:
                    tipe = "Weekly"

        # Gabung tanggal + hari untuk Once
        if tanggal and not hari:
            hari = tanggal
        elif tanggal and tanggal not in hari:
            hari = f"{hari} {tanggal}".strip()

        # Auto-infer tipe bila kosong
        if not tipe:
            tipe = infer_tipe(hari)
        if tipe not in ("Once", "Daily", "Weekly", "Monthly", "Yearly"):
            # normalisasi varian: "one time", "harian", "mingguan", ...
            tipe = normalisasi_tipe(tipe)

        hasil.append({
            "baris": i, "waktu": waktu, "label": label,
            "hari": hari, "tipe": tipe, "suara": suara, "pesan": pesan,
            "weekdays_en": parse_hari(hari, tipe),
        })
    return hasil

def _baca_posisional(rows, start):
    hasil = []
    for i, r in enumerate(rows[start:], start=start + 1):
        if not r or r[0] is None or (isinstance(r[0], str) and r[0].strip() == ""):
            continue
        waktu = normalisasi_waktu(r[0])
        if not waktu:
            print(f"[BARIS {i}] SKIP: Waktu tidak valid: {r[0]!r}")
            continue
        label = str(r[1]).strip() if len(r) > 1 and r[1] not in (None, "") else f"Alarm {waktu}"
        hari = str(r[2]).strip() if len(r) > 2 and r[2] not in (None, "") else ""
        tipe = str(r[3]).strip().capitalize() if len(r) > 3 and r[3] not in (None, "") else ""
        suara = str(r[4]).strip() if len(r) > 4 and r[4] not in (None, "") else "School.mp3"
        pesan = str(r[5]).strip() if len(r) > 5 and r[5] not in (None, "") else label
        if not tipe:
            tipe = infer_tipe(hari)
        tipe = normalisasi_tipe(tipe)
        hasil.append({
            "baris": i, "waktu": waktu, "label": label,
            "hari": hari, "tipe": tipe, "suara": suara, "pesan": pesan,
            "weekdays_en": parse_hari(hari, tipe),
        })
    return hasil

def infer_tipe(hari):
    if not hari or hari.strip().lower() in ("setiap hari", "every day", "tiap hari", "harian", "daily"):
        return "Daily"
    if re.search(r"\d{4}-\d{2}-\d{2}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}", hari):
        return "Once"
    return "Weekly"

def normalisasi_tipe(t):
    t = str(t).strip().lower().replace(" ", "")
    m = {
        "once": "Once", "onetime": "Once", "sekali": "Once", "sekalisaja": "Once",
        "daily": "Daily", "harian": "Daily", "setiaphari": "Daily", "tiaphari": "Daily", "everyday": "Daily",
        "weekly": "Weekly", "mingguan": "Weekly",
        "monthly": "Monthly", "bulanan": "Monthly",
        "yearly": "Yearly", "tahunan": "Yearly",
    }
    return m.get(t, "Weekly")

def normalisasi_waktu(s):
    # Dukung tipe asli Excel: datetime.time, datetime.datetime, float (fraksi hari), int
    if isinstance(s, dtime):
        return f"{s.hour:02d}:{s.minute:02d}"
    if isinstance(s, datetime):
        return s.strftime("%H:%M")
    if isinstance(s, ddate):
        return None  # tanggal saja tanpa jam -> bukan waktu valid
    if isinstance(s, (int, float)) and not isinstance(s, bool):
        # Excel menyimpan jam sebagai fraksi: 0.5 = 12:00
        if 0 <= s < 1:
            total = round(s * 24 * 60)
            return f"{total // 60:02d}:{total % 60:02d}"
        # angka seperti 700 -> 07:00, 1330 -> 13:30
        iv = int(s)
        if 0 <= iv <= 2359:
            h, m = iv // 100, iv % 100
            if 0 <= h <= 23 and 0 <= m <= 59:
                return f"{h:02d}:{m:02d}"
        return None
    s = str(s).strip()
    # dukung datetime dari openpyxl, "07:00", "7:00", "07.00", "07:00:00"
    if re.match(r"^\d{1,2}[:.]\d{2}(:\d{2})?$", s):
        s = s.replace(".", ":")
        parts = s.split(":")
        h, m = int(parts[0]), int(parts[1])
        if 0 <= h <= 23 and 0 <= m <= 59:
            return f"{h:02d}:{m:02d}"
        return None
    # coba parse tanggal+jam
    for fmt in ("%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M", "%H:%M"):
        try:
            return datetime.strptime(s, fmt).strftime("%H:%M")
        except ValueError:
            pass
    return None

def parse_hari(hari_str, tipe):
    """Kembalikan list hari English untuk Weekly, atau tanggal untuk Once."""
    if tipe == "Daily":
        return ["Every day"]
    hari_str = str(hari_str)
    if tipe == "Once":
        # format YYYY-MM-DD atau DD/MM/YYYY
        m = re.search(r"\d{4}-\d{2}-\d{2}", hari_str)
        if m:
            return [m.group(0)]
        m2 = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", hari_str)
        if m2:
            try:
                d, mo, y = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
                if y < 100:
                    y += 2000
                return [f"{y:04d}-{mo:02d}-{d:02d}"]
            except ValueError:
                pass
        return []
    h = hari_str.strip().lower()
    if h in ("", "setiap hari", "every day", "everyday", "tiap hari"):
        return ["Every day"]
    if h in ("weekdays", "hari kerja", "senin-jumat", "senin-jumat"):
        return ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    if h in ("weekends", "akhir pekan", "sabtu-minggu"):
        return ["Saturday", "Sunday"]
    out = []
    for part in re.split(r"[,;/|]+", hari_str):
        k = part.strip().lower()
        if k in HARI_MAP:
            out.append(HARI_MAP[k])
    return out or ["Every day"]

# ---------- Bagian otomasi Windows (pywinauto) ----------

def cari_main_window(app=None):
    """Cari jendela UTAMA FreeAlarmClock (bukan dialog). Atasi error '2 elements match'.

    Strategi: ambil semua jendela visible yang judulnya cocok, pilih yang punya
    tombol Add + menu File/Alarm (ciri main window)."""
    from pywinauto import Desktop, Application
    cands = Desktop(backend="uia").windows(title_re=".*Free Alarm Clock.*", visible_only=True)
    if not cands:
        # fallback: semua termasuk hidden
        cands = Desktop(backend="uia").windows(title_re=".*Free Alarm Clock.*")
    if not cands:
        raise RuntimeError("Jendela 'Free Alarm Clock' tidak ditemukan.")
    if len(cands) == 1:
        return cands[0]
    # Beberapa cocok (main + dialog) -> pilih yang punya tombol Add
    for w in cands:
        try:
            if w.child_window(title_re="^Add.*", control_type="Button").exists(timeout=1):
                return w
        except Exception:
            continue
    # fallback: jendela paling besar / pertama
    try:
        return max(cands, key=lambda w: w.rectangle().width() * w.rectangle().height())
    except Exception:
        return cands[0]

def inspect_gui():
    try:
        from pywinauto import Application
    except ImportError:
        print("pip install pywinauto  (Windows saja)")
        return
    print("Mencari jendela Free Alarm Clock...")
    app = Application(backend="uia").connect(title_re=".*Free Alarm Clock.*")
    win = app.window(title_re=".*Free Alarm Clock.*")
    print("=== MAIN WINDOW ===")
    win.print_control_identifiers()
    print("\nTIP: Buka dialog Add manual, lalu jalankan lagi dengan --inspect-dialog")
    print("untuk melihat kontrol dialog 'Alarm settings'.")

def _dialog_terbuka(main_win):
    """True bila ada top-level window baru selain main (kemungkinan dialog Add)."""
    try:
        from pywinauto import Desktop
        wins = Desktop(backend="uia").windows(visible_only=True)
        mh = None
        try:
            mh = main_win.handle
        except Exception:
            pass
        for w in wins:
            try:
                if mh and w.handle == mh:
                    continue
                t = w.window_text()
            except Exception:
                continue
            if "Free Alarm Clock" in t or t.strip() == "":
                return True
        # fallback: window aktif bukan main
        try:
            from pywinauto import Application
            app = Application(backend="uia").connect(title_re=".*Free Alarm Clock.*", timeout=3)
            act = app.active()
            return act.handle != mh
        except Exception:
            return False
    except Exception:
        return False

def tambah_satu_alarm_gui(main_win, alarm, delay=1.0, log=None):
    """Otomasi 1 alarm. Kembalikan True jika dialog Add berhasil dibuka.
    Disesuaikan dengan FreeAlarmClock v5.x English UI.
    """
    import time
    from pywinauto.keyboard import send_keys
    def say(m):
        print(m)
        if log:
            try:
                log(m + "\n")
            except Exception:
                pass
    opened = False
    err = ""
    # 1. Klik Add — coba beberapa varian (toolbar FreeAlarmClock beda-beda)
    for title_re, ctype in [("^Add.*", "Button"), (".*Add.*", "Button"),
                            (".*Add.*", "SplitButton"), (".*Add.*", None)]:
        try:
            kw = {"title_re": title_re}
            if ctype:
                kw["control_type"] = ctype
            btn = main_win.child_window(**kw)
            btn.wait("enabled visible ready", timeout=3)
            try:
                btn.click_input()
            except Exception:
                btn.click()
            opened = True
            say(f"  [GUI] Klik Add OK ({title_re}/{ctype})")
            break
        except Exception as e:
            err = str(e)[:160]
    if not opened:
        # dump tombol yang ada untuk diagnosa
        try:
            btns = [b.window_text() for b in main_win.descendants(control_type="Button")]
            say(f"  [GUI] Tombol Add tidak ketemu ({err}). Tombol yang ada: {btns[:12]}")
        except Exception as e2:
            say(f"  [GUI] Tombol Add tidak ketemu ({err}). Dump gagal: {e2}")
        # fallback keyboard: Insert (shortcut Add di FreeAlarmClock) lalu Alt+A (menu Alarm>Add)
        for keys in ("{INSERT}", "%a"):
            try:
                main_win.set_focus()
                time.sleep(0.3)
                send_keys(keys)
                say(f"  [GUI] Coba keyboard {keys}")
                time.sleep(delay)
                # cek apakah dialog muncul
                if _dialog_terbuka(main_win):
                    opened = True
                    say(f"  [GUI] Dialog terbuka via keyboard {keys}")
                    break
            except Exception as e3:
                say(f"  [GUI] Keyboard {keys} gagal: {e3}")
    time.sleep(delay)

    try:
        from pywinauto import Application
        app = Application(backend="uia").connect(title_re=".*Free Alarm Clock.*")
        # dialog biasanya title kosong / "Free Alarm Clock" - cari dialog teratas selain main
        dlgs = app.windows()
        dlg = None
        for d in dlgs:
            try:
                t = d.window_text()
            except Exception:
                continue
            # dialog Add biasanya punya tombol OK/Cancel + teks "Occurs" / "Label"
            if d != main_win.wrapper_object() if hasattr(main_win, "wrapper_object") else True:
                pass
            if "Alarm" in t or t.strip() == "":
                # heuristik: yang punya Edit "Label" / ComboBox "Occurs"
                try:
                    if d.child_window(title_re=".*Occurs.*|.*Label.*|.*Time.*", control_type=".*").exists(timeout=1):
                        dlg = d
                        break
                except Exception:
                    pass
        if dlg is None:
            # fallback: ambil window aktif
            dlg = app.active()
        print(f"  [GUI] Dialog: {dlg.window_text()!r}")
        dlg.set_focus()
        time.sleep(0.5)

        # --- Isi Occurs (One Time / Daily / Weekly / Monthly / Yearly) ---
        # mapping UI: "One Time" untuk Once
        occurs_label = {"Once": "One Time", "Daily": "Daily", "Weekly": "Weekly",
                        "Monthly": "Monthly", "Yearly": "Yearly"}[alarm["tipe"]]
        try:
            combo = dlg.child_window(control_type="ComboBox").wrapper_object()
            combo.select(occurs_label)
            print(f"  [GUI] Occurs -> {occurs_label}")
        except Exception as e:
            print(f"  [GUI] ComboBox Occurs tidak auto ({e}), pilih manual: {occurs_label}")
            send_keys("{TAB}")

        time.sleep(0.4)
        # --- Isi Time (HH:MM) ---
        try:
            # cari edit time: biasanya 2 spin / 1 edit
            edits = dlg.descendants(control_type="Edit")
            # heuristik: isi time lalu label
            print(f"  [GUI] Ditemukan {len(edits)} Edit box, isi manual Time={alarm['waktu']} Label={alarm['label']}")
        except Exception:
            pass

        # Panduan manual yang selalu tampil agar user bisa lanjut walau selector beda versi:
        print(f"  >>> ISI MANUAL di dialog: Time={alarm['waktu']} | Occurs={occurs_label} | "
              f"Hari={','.join(alarm['weekdays_en'])} | Label={alarm['label']} | Sound={alarm['suara']}")
        print("  >>> Tekan Enter/OK untuk simpan, script lanjut 2 detik...")
        time.sleep(2)
        # Jangan auto-OK agar tidak salah simpan — user tekan OK, atau uncomment baris bawah:
        # send_keys("{ENTER}")
        return opened
    except Exception as e:
        print(f"  [GUI ERROR] {e}")
        return False

def hitung_alarm_di_app():
    """Coba baca jumlah alarm yang tampil di list utama (verifikasi). Kembalikan int atau None."""
    try:
        main = cari_main_window()
        # List alarm biasanya berupa ListItem / DataItem per baris
        items = main.descendants(control_type="ListItem") or main.descendants(control_type="DataItem")
        # fallback: cari teks jam HH:MM di descendants
        if not items:
            return None
        return len(items)
    except Exception:
        return None

def run_otomasi(alarms, delay, on_progress=None, interactive=True, log=None):
    """on_progress(done, total, alarm, ok) dipanggil tiap 1 alarm selesai (untuk progress bar GUI/CLI).
    interactive=False untuk dipanggil dari GUI (tanpa input ENTER yang bikin stuck di .exe windowed).
    log(msg) opsional untuk kirim pesan ke GUI."""
    def say(m):
        print(m)
        if log:
            try:
                log(m + "\n")
            except Exception:
                pass
    try:
        from pywinauto import Application
        import time
    except ImportError:
        say("Butuh Windows + pip install pywinauto")
        if not interactive:
            raise
        sys.exit(1)
    say("Menghubungi Free Alarm Clock (timeout 10 detik)...")
    main = None
    last_err = None
    for _ in range(2):
        try:
            from pywinauto import Application
            # connect ke proses dulu (tidak ambigu), baru cari main window
            try:
                app = Application(backend="uia").connect(path="FreeAlarmClock.exe", timeout=10)
                main = cari_main_window(app)
            except Exception:
                app = Application(backend="uia").connect(title_re=".*Free Alarm Clock.*", timeout=10)
                main = cari_main_window(app)
            break
        except Exception as e:
            last_err = e
            s = str(e)
            if "elements that match" in s or "Ambiguous" in s:
                try:
                    main = cari_main_window()
                    break
                except Exception as e2:
                    last_err = e2
            import time as _t
            _t.sleep(1)
    if main is None:
        say(f"GAGAL terhubung: {last_err}")
        say(" Checklist: 1) FreeAlarmClock.exe sudah dibuka? 2) Tutup dialog Add yang terbuka, sisakan 1 main window.")
        say(" 3) Bahasa UI English? 4) Jalankan AlarmExcel sebagai Administrator bila perlu.")
        if not interactive:
            raise RuntimeError(str(last_err))
        sys.exit(1)
    try:
        main.set_focus()
    except Exception as e:
        say(f"Gagal fokus jendela: {e}")
    total = len(alarms)
    say(f"Terhubung: {main.window_text()!r} — {total} alarm akan ditambahkan.")
    say("PENTING: dialog Add akan dibuka satu-satu. Tekan OK di tiap dialog untuk simpan.")
    if interactive:
        try:
            input("Tekan ENTER untuk mulai...")
        except (EOFError, OSError):
            pass
    ok_count, fail = 0, []
    for i, a in enumerate(alarms, 1):
        say(f"\n[{i}/{total}] ({i*100//total}%) {a['waktu']} - {a['label']} ({a['tipe']} {a['hari']})")
        ok = tambah_satu_alarm_gui(main, a, delay=delay, log=say)
        ok_count += 1 if ok else 0
        if not ok:
            fail.append(a)
        if on_progress:
            try:
                on_progress(i, total, a, ok)
            except Exception:
                pass
    print(f"\n===== SELESAI: {ok_count}/{total} dialog Add terbuka ({ok_count*100//max(total,1)}%) =====")
    if fail:
        print(f"Gagal buka dialog: {len(fail)} alarm, cek manual:")
        for a in fail:
            print(f"  - {a['waktu']} {a['label']}")
    # Verifikasi jumlah di aplikasi
    n = hitung_alarm_di_app()
    if n is not None:
        print(f"Verifikasi: FreeAlarmClock menampilkan ~{n} baris alarm.")
        print(f"Cocokkan dengan Excel ({total}). Cek juga 'Next 3 Alarms' di bawah + File > Backup.")
    else:
        print("Verifikasi manual: cocokkan jumlah baris di FreeAlarmClock dengan Excel,")
        print("cek 'Next 3 Alarms' di panel bawah, lalu File > Backup untuk simpan *.alm.")
        print("100% = semua baris Excel sudah dilewati + muncul di list aplikasi.")
    return ok_count

def main():
    args = parse_args()
    p = Path(args.excel)
    if not p.exists():
        print(f"File tidak ketemu: {p}")
        sys.exit(1)
    alarms = baca_excel(str(p), sheet=args.sheet)
    if not alarms:
        print("Tidak ada data valid di Excel.")
        sys.exit(1)
    print(f"Dibaca {len(alarms)} alarm dari {p.name}:")
    for a in alarms:
        print(f"  {a['waktu']:5s} | {a['tipe']:6s} | {','.join(a['weekdays_en']):30s} | {a['label']} | {a['suara']}")
    if args.dry_run:
        print("\n--dry-run: berhenti di sini (tidak menyentuh FreeAlarmClock).")
        return
    if args.inspect:
        inspect_gui()
        return
    if sys.platform != "win32":
        print("\nOtomasi GUI hanya bisa di Windows. Di Mac pakai --dry-run untuk cek,")
        print("atau pakai alarm_standalone.py untuk bunyi langsung tanpa FreeAlarmClock.")
        return
    run_otomasi(alarms, delay=args.delay)

if __name__ == "__main__":
    main()
