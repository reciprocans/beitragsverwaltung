# Vorstands-Dokumente – Version 5

Kleine passwortgeschützte Dokumentenablage für Docker/Portainer.

## Funktionen
- Dark Mode
- Upload über den `+`-Button oben links
- PDF, Word (DOC/DOCX) und Excel (XLS/XLSX)
- Browser-Vorschau als PDF
- Titel, Rubrik, Kommentar/Notiz und **Datum des Schriftstücks**
- Rubriken: Vorstandssitzungen, Veranstaltungen, Mitgliederlisten, Briefe, Sonstiges
- automatische Sortierung nach Dokumentdatum, neuestes zuerst
- standardmäßig wird nur das aktuellste vorhandene Jahr angezeigt
- ältere Jahre lassen sich per Jahres-Schalter aufrufen
- Archiv für den Admin
- komplettes ZIP-Backup für den Admin (Datenbank, Originale, Vorschauen)
- drei Rollen: Admin, Redakteur, Leser
- Passwörter im Admin-Bereich änderbar
- persistente SQLite-Datenbank und Dateien unter `/data`

## Rollen
### Admin
Kann lesen, herunterladen, hochladen, bearbeiten, löschen, archivieren, archivierte Dokumente wiederherstellen, Backups herunterladen und die Passwörter aller drei Konten ändern.

### Redakteur
Kann lesen, herunterladen, hochladen, bearbeiten und löschen. Kein Zugriff auf Archiv, Backup oder Passwortverwaltung.

### Leser
Kann Dokumente nur ansehen und herunterladen.

## Erstanmeldung
Die Start-Zugangsdaten stehen in `docker-compose.yml`:

- Benutzer `admin` / Passwort `BITTE-AENDERN`
- Benutzer `redakteur` / Passwort `REDAKTEUR-AENDERN`
- Benutzer `leser` / Passwort `LESER-AENDERN`

Diese Passwörter gelten nur beim **ersten Anlegen** der jeweiligen Konten. Danach werden Änderungen in der Datenbank gespeichert und bei Container-Neustarts nicht überschrieben.

## Update einer bestehenden Installation
Das vorhandene Docker-Volume unbedingt weiterverwenden. Die Datenbank wird beim Start automatisch erweitert. Bestehende Dokumente bekommen zunächst ihr bisheriges Upload-Datum als Schriftstückdatum; dieses kann ein Admin oder Redakteur anschließend korrigieren.

Falls vorher nur das gemeinsame `APP_PASSWORD` benutzt wurde, wird dieses beim ersten Start der neuen Version – sofern kein Admin-Konto existiert – als Admin-Passwort übernommen. Alternativ wird `ADMIN_PASSWORD` verwendet.

## Backup
Im Admin-Bereich erzeugt **Backup herunterladen** eine ZIP-Datei mit:
- `app.db`
- allen Originaldateien unter `uploads/`
- allen PDF-Vorschauen unter `previews/`

## Ports
- Host: `8188`
- Container: `8787`
