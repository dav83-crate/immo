# Version 2.0.7 – Buchhaltung Eigenbestand / HV-Rechnungen

## Neu
Die Buchhaltung trennt nun zwischen:
- HV-Rechnungen als Einnahmen
- Versorger-/Betriebskostenrechnungen nur bei Eigenbestand

## Additive Erweiterungen
### Objekte
Neue Spalte am Ende:
- Buchhaltung

Werte:
- Eigenbestand
- Fremdverwaltung

### Neue Tabelle
- HV-Rechnungen (`hv_rechnungen.xlsx`)

## Buchhaltungslogik
### Einnahmen
- aus `HV-Rechnungen`

### Ausgaben
- Rechnungen nur für Objekte mit `Buchhaltung = Eigenbestand`
- Betriebskosten/Versorgerabschläge nur für Objekte mit `Buchhaltung = Eigenbestand`

## Wichtig
Bestehende Spalten wurden nicht verschoben.
Bestehende Beta-Daten bleiben erhalten.
