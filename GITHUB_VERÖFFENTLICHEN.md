# Repository auf GitHub veröffentlichen

## Variante über die GitHub-Webseite

1. Auf GitHub ein neues leeres Repository anlegen.
2. Keine zusätzliche README, `.gitignore` oder Lizenz erzeugen lassen.
3. Die von ChatGPT gelieferte Repository-ZIP entpacken.
4. Im Projektordner ein Terminal öffnen.
5. Die von GitHub angezeigten Befehle für ein bestehendes Repository ausführen.

## Beispiel

```powershell
git remote add origin https://github.com/DEIN-NAME/immoverwaltung.git
git branch -M main
git push -u origin main
```

Für ein privates Projekt sollte das Repository bei GitHub als **Private** angelegt werden.
