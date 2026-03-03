# AutoBericht Schnellstart (DE)

Diese Anleitung ist für neue Anwender:innen mit technischem Grundverständnis. UI-Begriffe sind **englisch**, wie in der App.

## 1) Repo und Projektordner anlegen
1. Lege einen Repo-Ordner an, z. B. `C:\AutoBericht`.
2. Lege einen leeren Projektordner an, z. B. `C:\Projekt_X`.
3. Kopiere das komplette Repo nach `C:\AutoBericht`.
4. Starte `start-autobericht.cmd` im Repo-Root.

## 2) Projektordner wählen
- Im Startdialog auf **Open Project Folder** klicken.

![Startdialog](screenshots/annotated/01-first-run-modal.png)

## 3) Projekt initialisieren (Sprache/Metadaten)
1. Auf der linken Seite **Project** öffnen.
2. Im Feld **Locale** Sprache wählen (`DE`, `FR`, `IT`).
3. Moderator/Firma eintragen.

Rote Markierungen: Locale, Import Self-Assessment, Word/PPT Export.

![Project Seite](screenshots/annotated/02-project-page-overview.png)

## 4) Foto-Workflow vorbereiten
1. Originalfotos in `photos/raw/pm1`, `photos/raw/pm2`, `photos/raw/pm3` ablegen.
2. In der App auf **PhotoSorter** wechseln.
3. Oben: **Import / Export Photos**.

Rote Markierungen: Import/Export, Edit tags, Save sidecar, Show Unsorted, Clear Filters, `U`.

![PhotoSorter Übersicht](screenshots/annotated/06-photosorter-main.png)

## 5) Fotos importieren/exportieren
- **Import photos**: Import + Resize nach `photos/resized` (lange Seite max. 1920 px).
- **Rescan photos**: nur neu einlesen.
- **Export tagged folders**: exportiert in `photos/export` je Tag/Kategorie.

![Import/Export Dialog](screenshots/annotated/08-photosorter-import-export-modal.png)

## 6) Filter und Unsorted nutzen
- **Show Unsorted**: nur ungetaggte Bilder.
- Filter-Icon unten rechts am Bild: aktive Filter zurücksetzen.

![Filter aktiv](screenshots/annotated/07-photosorter-filter-active.png)

## 7) Beobachtungs-Tags pflegen
- **Edit tags** öffnen.
- Tags bei Bedarf ergänzen/entfernen.
- **Save debug log** für Supportfälle.

![Edit Tags](screenshots/annotated/09-photosorter-edit-tags-modal.png)

## 8) Bericht im AutoBericht bearbeiten
1. Kapitel links wählen.
2. Texte in **Finding** / **Recommendation** bearbeiten.
3. Mit **Include**, **Done**, Priorität und Sidebar-Filtern arbeiten.
4. Pro Zeile **Library Off / Append / Replace** bewusst setzen:
   - **Off**: nicht in die Library schreiben.
   - **Append**: aktuellen Text an vorhandenen Library-Text anhängen.
   - **Replace**: vorhandenen Library-Text komplett ersetzen.

![Kapitelbearbeitung 4.8](screenshots/annotated/04-chapter-48-editing.png)

## 9) Kapitel 4.8 ordnen
- In Kapitel 4.8 auf **Organize 4.8**.
- Reihenfolge/Filter prüfen, dann **Apply**.

![Organize 4.8](screenshots/annotated/05-organize-48-modal.png)

## 10) Word, PowerPoint und Library
- Auf der **Project**-Seite:
  - **Word Export**
  - **PowerPoint Export (Report)** / **PowerPoint Export (Training)**
  - **Generate / Update Library**
  - **Export Library Excel**

## 11) Datenlogik (wichtig)
- `project_sidecar.json`: Projektzustand (Antworten, Auswahl, Reihenfolge, Fototags, editierte Texte und gewählte Library-Aktion je Zeile).
- `library_*.json`: wiederverwendbare Inhalte (Textbausteine/Tag-Wissen über Projekte hinweg).
- `Append`/`Replace` wird erst wirksam, wenn du auf **Generate / Update Library** klickst.
