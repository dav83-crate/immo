@echo off
title ImmoVerwaltung Backup erstellen
set "PY=%CD%\.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
%PY% scripts\backup_erstellen.py
pause
