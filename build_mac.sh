#!/bin/bash
# Build aplikasi desktop Mac (.app + binary) — jalankan di Mac
# Hasil: dist/AlarmExcel.app dan dist/AlarmExcel
set -e
cd "$(dirname "$0")"
pip3 install -r requirements.txt pyinstaller
# --windowed = tanpa terminal, --onedir = folder app (lebih stabil untuk tkinter)
python3 -m PyInstaller \
  --noconfirm --clean \
  --name "AlarmExcel" \
  --windowed --onedir \
  --add-data "contoh_jadwal.xlsx:." \
  gui_alarm.py
echo ""
echo "SELESAI. Coba double-click: dist/AlarmExcel.app"
echo "Atau via terminal: ./dist/AlarmExcel/AlarmExcel"
