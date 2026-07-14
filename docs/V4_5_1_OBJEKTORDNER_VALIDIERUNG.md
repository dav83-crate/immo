# Version 4.5.1 – Objektordner-Validierung

## Behoben
- Werte wie m3, m², kWh, reine Zahlen und Einheiten werden nie als Objektordner verwendet.
- Objektordner werden nur aus den stabilen Stammdaten Objekte, Wohnungen und Mieter geladen.
- Alle Module verwenden dieselbe zentrale Objektordner-Funktion.
- Nicht eindeutig zuordenbare Datensätze bleiben leer statt falsch zugeordnet zu werden.
- Neuer Bereich `Objektordner-Prüfung` mit Bericht und Excel-Export.

## Wichtig
- Keine bestehende Exceldatei verändert.
- Keine Spalte ergänzt, gelöscht oder umsortiert.
- Der Prüfbereich arbeitet ausschließlich lesend.
- Seitenbuttonfarbe bleibt #f00000.
