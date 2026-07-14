# Version 5.2.1 – EXE-Logo- und Ressourcen-Fix

## Behoben
- Logo wird in einer PyInstaller-Onefile-EXE über sys._MEIPASS gefunden.
- Daten und config.json bleiben neben der EXE und werden nicht im Temp-Ordner gespeichert.
- Fenstericon wird für QApplication und MainWindow gesetzt.
- dbs_logo.ico mit mehreren Windows-Größen ergänzt.
- PyInstaller bindet assets vollständig ein.
- EXE-Dateisymbol wird über --icon gesetzt.
- CONFIG_DIR-Fehler beseitigt; user_ui_state.json liegt im Datenordner.

## EXE erstellen
1. EXE_ERSTELLEN.bat starten.
2. Die fertige EXE liegt auf dem Desktop.
3. Alternativ EXE_ERSTELLEN_MIT_SPEC.bat verwenden.

## Windows-Symbolcache
Windows kann das alte EXE-Symbol zwischenspeichern. In diesem Fall die neue EXE anders benennen oder den Explorer neu starten.

Die Excelstruktur wurde nicht verändert.
