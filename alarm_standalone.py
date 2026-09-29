#!/usr/bin/env python3
"""
Alarm standalone dari Excel — jalan di Mac / Windows / Linux TANPA FreeAlarmClock.
Cocok untuk testing format Excel sebelum import ke FreeAlarmClock di Windows,
atau sebagai pengganti bel sekolah sederhana.

  python3 alarm_standalone.py contoh_jadwal.xlsx --check   # cek jadwal hari ini
  python3 alarm_standalone.py contoh_jadwal.xlsx --run     # jalan terus, bunyi tiap jadwal

Butuh: pip install openpyxl
Suara: pakai file MP3 jika ada, else beep sistem (afplay di Mac / winsound di Windows).
"""
import sys, time, subprocess, platform
from datetime import datetime
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from excel_to_freealarm import baca_excel

HARI_INDO = ["Senin","Selasa","Rabu","Kamis","Jumat","Sabtu","Minggu"]

def hari_ini_indo():
    # Monday=0
    return HARI_INDO[datetime.now().weekday()]

def cocok_hari_ini(alarm):
    tipe = alarm["tipe"]
    if tipe == "Daily":
        return True
    if tipe == "Once":
        # weekdays_en berisi [YYYY-MM-DD] atau []
        if not alarm["weekdays_en"]:
            return True
        return datetime.now().strftime("%Y-%m-%d") in alarm["weekdays_en"]
    # Weekly
    en_today = datetime.now().strftime("%A")  # Monday...
    if "Every day" in alarm["weekdays_en"]:
        return True
    return en_today in alarm["weekdays_en"]

def bunyi(alarm):
    print(f"\n🔔 {datetime.now().strftime('%H:%M:%S')} — {alarm['label']} ({alarm['waktu']}) — {alarm['pesan']}")
    suara = alarm["suara"]
    try:
        if Path(suara).exists():
            if platform.system() == "Darwin":
                subprocess.Popen(["afplay", suara])
            elif platform.system() == "Windows":
                import winsound
                winsound.PlaySound(suara, winsound.SND_FILENAME | winsound.SND_ASYNC)
            else:
                subprocess.Popen(["xdg-open", suara])
            return
    except Exception as e:
        print(f"(gagal putar {suara}: {e})")
    # fallback beep
    try:
        if platform.system() == "Windows":
            import winsound
            for _ in range(3):
                winsound.Beep(880, 500); time.sleep(0.2)
        else:
            print("\a")
            subprocess.run(["afplay", "/System/Library/Sounds/Glass.aiff"], timeout=5)
    except Exception:
        print("\a BEEP!")

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    excel = sys.argv[1]
    mode = sys.argv[2] if len(sys.argv) > 2 else "--check"
    alarms = baca_excel(excel)
    print(f"Hari ini: {hari_ini_indo()} {datetime.now().strftime('%Y-%m-%d')}")
    aktif = [a for a in alarms if cocok_hari_ini(a)]
    aktif.sort(key=lambda x: x["waktu"])
    print(f"Jadwal aktif hari ini ({len(aktif)}):")
    for a in aktif:
        print(f"  {a['waktu']} - {a['label']}")
    if mode == "--check":
        return
    print("\nBerjalan... Ctrl+C untuk berhenti.")
    sudah = set()
    while True:
        now = datetime.now().strftime("%H:%M")
        today = datetime.now().strftime("%Y-%m-%d")
        for a in aktif:
            key = (today, a["waktu"], a["label"])
            if a["waktu"] == now and key not in sudah:
                bunyi(a)
                sudah.add(key)
        time.sleep(10)

if __name__ == "__main__":
    main()
