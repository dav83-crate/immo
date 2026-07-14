@echo off
chcp 65001 >nul
title ImmoVerwaltung 5.2.1 EXE erstellen
cd /d "%~dp0"

set "PYTHON_EXE=%CD%\.venv\Scripts\python.exe"

if exist "%PYTHON_EXE%" (
    echo Verwende virtuelle Umgebung:
    echo "%PYTHON_EXE%"
) else (
    echo Keine .venv gefunden. Verwende py Launcher.
    set "PYTHON_EXE=py"
)

echo.
echo Pruefe Python...
%PYTHON_EXE% --version
if errorlevel 1 (
    echo FEHLER: Python wurde nicht gefunden.
    pause
    exit /b 1
)

echo.
echo Installiere Requirements und PyInstaller...
%PYTHON_EXE% -m pip install -r requirements.txt
%PYTHON_EXE% -m pip install pyinstaller pillow

if not exist "assets\dbs_logo.png" (
    echo FEHLER: assets\dbs_logo.png fehlt.
    pause
    exit /b 1
)

if not exist "assets\dbs_logo.ico" (
    echo FEHLER: assets\dbs_logo.ico fehlt.
    pause
    exit /b 1
)

echo.
echo Alte Build-Dateien loeschen...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo.
echo EXE wird mit Logo und Fenstericon erstellt...
%PYTHON_EXE% -m PyInstaller ^
 --onefile ^
 --windowed ^
 --clean ^
 --noconfirm ^
 --name "ImmoVerwaltung_5_2_1" ^
 --icon "assets\dbs_logo.ico" ^
 --add-data "assets;assets" ^
 --distpath "%USERPROFILE%\Desktop" ^
 "main.py"

if errorlevel 1 (
    echo FEHLER: EXE konnte nicht erstellt werden.
    pause
    exit /b 1
)

echo.
echo Fertig:
echo "%USERPROFILE%\Desktop\ImmoVerwaltung_5_2_1.exe"
echo.
echo Hinweis:
echo Windows kann alte Symbole zwischenspeichern. Bei Bedarf EXE umbenennen
echo oder den Windows-Explorer neu starten.
pause
