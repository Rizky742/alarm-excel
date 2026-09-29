@echo off
REM Build aplikasi desktop Windows (.exe) — jalankan di PC Windows client
REM Hasil: dist\AlarmExcel.exe (satu file, tinggal double-click)
pip install -r requirements.txt pyinstaller
python -m PyInstaller ^
  --noconfirm --clean ^
  --name "AlarmExcel" ^
  --windowed --onefile ^
  --add-data "contoh_jadwal.xlsx;." ^
  gui_alarm.py
echo.
echo SELESAI. File: dist\AlarmExcel.exe
pause
