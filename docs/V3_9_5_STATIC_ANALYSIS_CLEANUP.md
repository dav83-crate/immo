# Version 3.9.5 – Static Analysis Cleanup

Bereinigt:
- ungenutzte Signalparameter (`column` -> `_column`)
- statische Hilfsmethoden mit `@staticmethod`
- sichere Behandlung von `QLayoutItem | None`
- Initialisierung von `dashboard_treffer` in `__init__`
- sichere Excelwert-Konvertierung für date/time/datetime/timedelta und Formelfelder
- MergedCell-kompatible Spaltenbreiten über `get_column_letter`
- breite Exception-Klauseln auf konkrete Standardfehler eingegrenzt
- ungenutzte Variable `rect` entfernt
- doppeltes Escape im Datums-RegEx beseitigt
- Pfadparameter akzeptieren `str | Path`
- ungenutzten Import `QTimer` entfernt
- Schattenvariablen bereinigt

Die XLSX-Beta-Struktur wurde nicht verändert.
