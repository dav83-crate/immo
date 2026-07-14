@echo off
title ImmoVerwaltung Feedback eintragen
set "PY=%CD%\.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"
%PY% scripts\feedback_eintragen.py
pause
