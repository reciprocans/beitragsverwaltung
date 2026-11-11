# Kolping-inspirierte Vereinsverwaltung V15 – Mitglieder + Vorstandsprotokolle

Die bisherige Vereinsverwaltung V14 und die hochgeladene **Vorstandsprotokolle V5** werden in einem Docker-Container unter **Port 2312** bereitgestellt. Zwei Bereiche, eine Anmeldung, gemeinsame Benutzerrollen, aber **getrennte SQLite-Datenbanken** im selben Volume.

- **Vereinsverwaltung:** `http://SERVER-IP:2312/`
- **Vorstandsprotokolle:** `http://SERVER-IP:2312/protokolle/`
- **Benutzerverwaltung:** `/users` (nur Admin)

## Benutzer und Rollen

| Konto | Rechte |
| --- | --- |
| `admin` | Vollzugriff, Zahlungsverwaltung, Beitragssätze, Import/Export, Backups, Löschen, Benutzerpasswörter |
| `redakteur` | Mitglieder und Familien anlegen/bearbeiten, Mandate und Anträge hochladen, Protokolle hochladen/bearbeiten/löschen; **keine** Beitragsbuchungen, Importe, Admin-Einstellungen oder Benutzerverwaltung |
| `leser` | Mitglieder, Familien, Statistiken, Beitragsstände und Protokolle **ansehen**; Dokumente herunterladen; **keine** Änderungen |

Die Rechte werden auf dem Server geprüft und nicht nur über versteckte Buttons gesteuert. Protokolle werden bei beiden Benutzerrollen mit derselben Sitzung geöffnet. Die vorhandenen Protokollfunktionen (Dokumentdatum, Rubriken, Word-/Excel-/PDF-Vorschau, Archiv und Protokoll-Backup) bleiben vorhanden.

## Installation / Update über Portainer

1. **Vor dem Update** das bisherige Mitglieder-Docker-Volume vollständig sichern, inklusive der hochgeladenen Dokumente.
2. Archiv entpacken, den Ordner `vereinsbeitraege` in dein Git-Repository hochladen. In Portainer den Stack mit der `docker-compose.yml` aus dem Ordner neu bauen (**Rebuild / Pull and redeploy**).
3. Portainer-Umgebungsvariablen setzen: `ADMIN_PASSWORD`, `EDITOR_PASSWORD`, `READER_PASSWORD` (je **mindestens 10 Zeichen**) und einen langen zufälligen `SECRET_KEY`. Bestehenden `SECRET_KEY` für die Sitzung beibehalten. `ADMIN_PASSWORD` für den Admin der bisherigen Installation wiederverwenden.
4. Das **bisherige Volume** `vereinsbeitraege_data` weiterverwenden. Ein neues Volume würde wie eine leere Installation aussehen.
5. Über `http://SERVER-IP:2312/login` anmelden mit `admin`, `redakteur` oder `leser` und dem jeweiligen Passwort. Passwörter können anschließend unter **Benutzer** vom Admin geändert werden.

**Wichtig:** Passwörter aus den Umgebungsvariablen werden nur beim ersten Erstellen des jeweiligen Kontos übernommen; danach liegen Hashes in `verein.db`. Wer bereits ein Admin-Konto mit einem älteren Passwort anlegt, muss dieses Konto über „Benutzer“ ändern – nicht allein über die Umgebungsvariable.

## Vorstandsprotokolle – Datenübernahme

- Die hochgeladene Vorstandsprotokolle-V5-ZIP enthält die **Software, keine vorhandene `app.db` oder Dokumentdateien**. Beim ersten Start entsteht eine leere Protokollablage unter `/data/app.db` mit `/data/uploads` und `/data/previews`.
- Sollten Vorstandsprotokolle bereits separat mit Daten betrieben worden sein, **vor dem Umzug deren Volume sichern** und `app.db`, `uploads/` und `previews/` (mit den bestehenden Dateinamen) in das Mitglieder-Volume `/data` übernehmen, während die Container gestoppt sind. Dateien nicht überschreiben, falls dort bereits Protokolle liegen. Hierfür ist eine einmalige kontrollierte Migration nötig.
- Die Vereinsdatenbank `verein.db` und `/data/documents` bleiben getrennt und unangetastet. Die Protokoll-Nutzerkonten aus einer alten `app.db` werden **nicht** verwendet: Die gemeinsame Anmeldung liegt ausschließlich in `verein.db`.

## Hinweise

- Mit dem Docker-Volume werden sowohl `verein.db`, Dokumente, Protokoll-DB als auch deren Originale und PDF-Vorschauen gespeichert. **Sichere das ganze Volume**, nicht nur eine einzelne Datenbank.
- Bei Zugriff übers Internet HTTPS (Reverse Proxy) verwenden und `COOKIE_SECURE=1` setzen.
- Für Word-/Excel-Protokollvorschauen enthält der Container LibreOffice. Der erste Docker-Build ist daher größer als in V14.
- Der Dienst benutzt den externen Port **2312** und den internen Port **8000**.
- Ein vollständiger Docker-Laufzeittest ist in der Erstellungsumgebung nicht möglich; Python-Syntax und die Paketstruktur wurden geprüft.
