@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Starte ImmoVerwaltung 2.0 Professional...
py -3 main.py
if errorlevel 1 (
    echo.
    echo Python konnte nicht gestartet werden. Bitte Python/PySide6 pruefen.
    pause
)
