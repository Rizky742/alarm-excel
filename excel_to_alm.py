#!/usr/bin/env python3
"""
Excel -> *.alm (Free Alarm Clock backup) — INSTAN, tanpa klik Add satu-satu.
Format hasil reverse-engineering dari file backup asli v5.3:

  <BOM>YYYYMMDDHHMMSS<FLAGS> <CODE><AlarmLabel>...</AlarmLabel><AlarmSound>...</AlarmSound>\r\n

  - YYYYMMDDHHMMSS: trigger berikutnya. Once/Daily = tanggal nyata;
    Weekly = 00000000 + jam (contoh 00000000100000 = 10:00)
  - FLAGS Once/Daily  = 111011111011152  (dari test1/test2)
  - FLAGS Weekly Senin = 111010000011152 (dari test3, Senin saja)
  - CODE: 10 = One Time, 11 = Daily, 12 = Weekly (13/14 = Monthly/Yearly, belum diverifikasi)
  - Sound disimpan lowercase: twinkle, cuckoo, piano, alarm, school...

Cara pakai:
  python3 excel_to_alm.py contoh_ews_rj.xlsx hasil.alm --tanggal 2026-09-30
  # lalu di FreeAlarmClock: File > Restore > pilih hasil.alm

 tanggal default = hari ini (Once lewat jam -> otomatis besok).
"""
import sys
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from xml.sax.saxutils import escape

sys.path.insert(0, str(Path(__file__).parent))
from excel_to_freealarm import baca_excel

FLAGS_ONCE_DAILY = "111011111011152"
FLAGS_WEEKLY_SENIN = "111010000011152"
CODE = {"Once": "10", "Daily": "11", "Weekly": "12", "Monthly": "13", "Yearly": "14"}

# sound name -> nama file lowercase yang dipakai .alm
def norm_sound(s):
    s = (s or "school").strip().lower()
    s = s.replace(".mp3", "").replace(".wav", "")
    return s or "school"

def next_datetime(waktu_hhmm, tipe, ref_date=None):
    """Hitung YYYYMMDDHHMMSS untuk .alm."""
    h, m = int(waktu_hhmm[:2]), int(waktu_hhmm[3:5])
    if tipe == "Weekly":
        return f"00000000{h:02d}{m:02d}00"
    today = ref_date or datetime.now()
    dt = today.replace(hour=h, minute=m, second=0, microsecond=0)
    if tipe == "Once" and dt <= today:
        dt += timedelta(days=1)
    # Daily: simpan trigger berikutnya juga (seperti file asli test2)
    if tipe == "Daily" and dt <= today:
        dt += timedelta(days=1)
    return dt.strftime("%Y%m%d%H%M%S")

def alarm_ke_baris(a, ref_date=None):
    tipe = a["tipe"] if a["tipe"] in CODE else "Weekly"
    # Weekly selain Senin-only: fallback ke Daily (aman, bunyi tiap hari)
    # karena bitmask hari lain belum 100% dipetakan.
    flags = FLAGS_ONCE_DAILY
    if tipe == "Weekly":
        days = set(a.get("weekdays_en", []))
        if days == {"Monday"}:
            flags = FLAGS_WEEKLY_SENIN
        elif days == {"Every day"} or not days:
            tipe = "Daily"  # fallback aman
            flags = FLAGS_ONCE_DAILY
        else:
            # fallback aman: Daily agar tetap bunyi (catat di log)
            print(f"  [NOTE] {a['waktu']} {a['label']}: Weekly {sorted(days)} belum dipetakan bitmask -> pakai Daily")
            tipe = "Daily"
            flags = FLAGS_ONCE_DAILY
    dt = next_datetime(a["waktu"], tipe, ref_date)
    code = CODE[tipe]
    # Samakan file asli: hanya & dan < yang di-escape, > dibiarkan mentah
    label = a["label"].replace("&", "&amp;").replace("<", "&lt;")
    sound = norm_sound(a["suara"]).replace("&", "&amp;").replace("<", "&lt;")
    return f"{dt}{flags} {code}<AlarmLabel>{label}</AlarmLabel><AlarmSound>{sound}</AlarmSound>"

def main():
    ap = argparse.ArgumentParser(description="Excel -> .alm instan")
    ap.add_argument("excel", help="File .xlsx")
    ap.add_argument("output", help="File .alm hasil")
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--tanggal", default=None, help="Tanggal acuan YYYY-MM-DD (default: hari ini)")
    args = ap.parse_args()

    ref = datetime.strptime(args.tanggal, "%Y-%m-%d") if args.tanggal else datetime.now()
    alarms = baca_excel(args.excel, sheet=args.sheet)
    if not alarms:
        print("Tidak ada data valid.")
        sys.exit(1)
    lines = [alarm_ke_baris(a, ref) for a in alarms]
    out = Path(args.output)
    # newline="" agar \r\n tidak digandakan jadi \r\r\n di Windows
    with open(out, "w", encoding="utf-8-sig", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")
    print(f"OK: {len(lines)} alarm -> {out} ({len(lines)*100//len(lines)}% selesai)")
    for l in lines:
        print(" ", l[:60] + "...")

if __name__ == "__main__":
    main()
