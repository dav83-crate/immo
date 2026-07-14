# Version 2.0.5 – Eingabemaske Reihenfolge-Fix

## Behoben
Die gruppierte Eingabemaske hat Werte beim Speichern nicht mehr exakt in der ursprünglichen SCHEMA-/Excel-Spaltenreihenfolge zurückgegeben.

Dadurch konnten Felder wie Objektordner oder Wohnungsordner falsch befüllt werden.

## Änderung
- Optische Gruppierung bleibt erhalten.
- Jedes Eingabefeld wird intern mit seinem Original-Spaltenindex gespeichert.
- Beim Speichern werden Werte wieder exakt in Tabellenreihenfolge zurückgegeben.
- Keine Excel-Strukturänderung.
