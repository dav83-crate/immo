@echo off
chcp 65001 >nul
title ImmoVerwaltung 5.2.1 EXE via Spec
cd /d "%~dp0"

set "PYTHON_EXE=%CD%\.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" set "PYTHON_EXE=py"

%PYTHON_EXE% -m pip install -r requirements.txt
%PYTHON_EXE% -m pip install pyinstaller pillow

if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

%PYTHON_EXE% -m PyInstaller --clean --noconfirm "ImmoVerwaltung_5_2_1.spec"

if errorlevel 1 (
    echo FEHLER: Build fehlgeschlagen.
    pause
    exit /b 1
)

echo Fertig: dist\ImmoVerwaltung_5_2_1.exe
pause
