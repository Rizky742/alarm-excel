@echo off
REM IMPORT sekali klik — copy hasil.alm lalu buka FreeAlarmClock untuk Restore
REM Letakkan file ini SATU FOLDER dengan FreeAlarmClock.exe + hasil.alm
copy /Y hasil.alm "%~dp0hasil.alm" >nul
echo.
echo 1. FreeAlarmClock akan dibuka...
echo 2. Klik File ^> Restore ^> pilih hasil.alm ^> Open
echo 3. Cek list + Next 3 Alarms
echo.
start "" "%~dp0FreeAlarmClock.exe"
pause
