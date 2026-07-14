# Version 2.0.13 – XLSX-Struktur bereinigt

## Ziel
Überflüssige Excel-Dateien wurden entfernt, ohne die Arbeitsweise des Programms einzuschränken.

## Entfernt
- Export-Excel-Dateien
- Analyse-Exportdateien
- temporäre Excel-Dateien außerhalb des Datenordners

## Behalten
Alle Excel-Dateien, die in `DATA_FILES` aktiv vom Programm genutzt werden.

## Wichtig
- Aktive Stammdaten-/Moduldateien im Ordner `daten/` bleiben erhalten.
- Die Programmfunktion bleibt erhalten.
- Keine aktiven Daten wurden absichtlich entfernt.
- Seitenbutton-Schriftfarbe bleibt `#f00000`.
