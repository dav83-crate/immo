# Version 2.0.11 – Mieter Suche Objektordner Fix

## Behoben
Bei Mieter konnte die Suche/Anzeige durch fehlende oder falsch erkannte Objektordner/Wohnungsordner leer bleiben.

## Änderung
- Mieter-Suche nutzt jetzt alle normalen Felder.
- Zusätzlich werden Objektordner/Wohnungsordner und deren Ordnernamen berücksichtigt.
- Falls Ort im Format `Objekt / Wohnung` gepflegt ist, wird auch darüber gesucht.
- Nach `Neu anlegen` wird der Suchfilter geleert, damit der neue Mieter sofort sichtbar ist.
- `self.anzeige` wird wieder korrekt auf die gefilterten Zeilen gesetzt.

## Wichtig
- Keine Excel-Strukturänderung.
- Seitenbutton-Schriftfarbe bleibt #f00000.
