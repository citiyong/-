@echo off
chcp 65001 >nul
echo Run the cleanup using the latest list on your Desktop? Deleted files go to the Recycle Bin.
set /p ok=Type Y and press Enter: 
if /i not "%ok%"=="Y" goto :eof
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0pc-cleanup.ps1" -Apply
pause
