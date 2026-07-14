# V50.1 Anzeige-Fix

## Behoben
Beim Wechsel über die Seitenleiste konnten vorherige Inhalte/Tabs sichtbar bleiben und erst beim Bewegen des Mauszeigers korrekt neu gezeichnet werden.

## Ursache
Die Fade-Animation wurde auf den gesamten QStackedWidget-Inhaltsbereich gelegt.
Bei komplexen Seiten mit Tabs und Tabellen kann das in Qt zu verzögertem Repaint führen.

## Änderung
- Fade-Animation des Inhaltsbereichs entfernt.
- Seitenwechsel erfolgt sofort und stabil.
- Aktiver Menüpunkt bleibt optisch markiert.
- Aktiver Menüpunkt wird leicht höher dargestellt.

## Daten
Keine Änderung an Excel-Struktur oder Daten.
