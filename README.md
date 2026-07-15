# ImmoVerwaltung

**Aktueller Entwicklungsstand: Version 6.0 Enterprise Foundation mit dazugehöriger Mitarbeiterapp für Zeiterfassung und Schädensmeldung**

Windows-Desktopanwendung für Immobilien-, Mieter-, Dokumenten-, Aufgaben-, Fristen- und Finanzverwaltung.

## Schwerpunkte

- Objekt-, Wohnungs- und Mieterverwaltung
- Rechnungen, Betriebskosten und Buchhaltung
- PDF-Import für Kontoauszüge und Belege
- Objektordner als zentraler Beziehungsschlüssel
- Mitarbeiter-Login, Rollen und Rechte
- Kalender, Aufgaben, Fristen und Workflow-Center
- Smart Objektakte, Objektgalerie und Verwaltungsleitstand
- Berichte, Backups und Systemprüfung
- Offline-Betrieb mit bestehender Beta-Excelstruktur

## Voraussetzungen

- Windows 10 oder Windows 11
- Python 3.11 oder neuer

## Installation

```powershell
py -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

Alternativ kann unter Windows `STARTEN.bat` verwendet werden.

## EXE erstellen

```powershell
EXE_ERSTELLEN.bat
```

Das Logo und das Windows-Icon werden aus `assets/` eingebettet. Eine alternative Erstellung über die Spec-Datei ist mit `EXE_ERSTELLEN_MIT_SPEC.bat` möglich.

## Daten und Datenschutz

Produktive Excel-Dateien im Ordner `daten/`, lokale Konfigurationen, Backups und Exporte sind über `.gitignore` ausgeschlossen. Diese Dateien dürfen nicht in ein öffentliches Repository eingecheckt werden.

## Zentrale Entwicklungsregeln

- Die bestehende Beta-Excelstruktur bleibt unverändert.
- Objektordner bleiben der zentrale Beziehungsschlüssel.
- Neue Versionen müssen abwärtskompatibel bleiben.
- Die Navigationsschriftfarbe `#f00000` wird beibehalten.

## Dokumentation

Weitere Hinweise befinden sich in `docs/`, im Benutzerhandbuch und im `CHANGELOG.md`.

## Lizenz

Proprietär. Alle Rechte vorbehalten. Siehe `LICENSE`.


## Version 6.0

Enterprise Objekt 360°, Dokumenten-Center 2.0 und Workflow-Regeln PRO.

⚠️ Copyright Notice

Dieses Repository dient ausschließlich der Dokumentation und Entwicklung.

Das Kopieren, Verwenden oder Weiterentwickeln des Quellcodes außerhalb einer ausdrücklich schriftlich erteilten Genehmigung ist nicht gestattet.

Alle Rechte verbleiben bei David Dornheim.
