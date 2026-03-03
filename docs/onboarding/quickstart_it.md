# AutoBericht Guida Rapida (IT)

Questa guida è pensata per utenti tecnici. Le etichette UI restano in **inglese** (come nell'app).

## 1) Preparare repo e cartella progetto
1. Crea una cartella repo, ad es. `C:\AutoBericht`.
2. Crea una cartella progetto vuota, ad es. `C:\Progetto_X`.
3. Copia l'intero repo in `C:\AutoBericht`.
4. Avvia `start-autobericht.cmd` nella root del repo.

## 2) Selezionare la cartella progetto
- Nel dialog iniziale, clicca **Open Project Folder**.

![Dialog iniziale](screenshots/annotated/01-first-run-modal.png)

## 3) Inizializzare il progetto (lingua/metadati)
1. Apri **Project** a sinistra.
2. In **Locale**, scegli la lingua (`DE`, `FR`, `IT`).
3. Inserisci moderatore e azienda.

Cerchi rossi: Locale, Import Self-Assessment, Word/PPT Export.

![Pagina Project](screenshots/annotated/02-project-page-overview.png)

## 4) Preparare il flusso foto
1. Metti le foto originali in `photos/raw/pm1`, `photos/raw/pm2`, `photos/raw/pm3`.
2. Passa a **PhotoSorter**.
3. In alto: **Import / Export Photos**.

Cerchi rossi: Import/Export, Edit tags, Save sidecar, Show Unsorted, Clear Filters, `U`.

![Vista PhotoSorter](screenshots/annotated/06-photosorter-main.png)

## 5) Importare/esportare foto
- **Import photos**: import + resize in `photos/resized` (lato lungo max. 1920 px).
- **Rescan photos**: rilettura semplice.
- **Export tagged folders**: esporta in `photos/export` per tag/categorie.

![Dialog Import/Export](screenshots/annotated/08-photosorter-import-export-modal.png)

## 6) Usare filtri e unsorted
- **Show Unsorted**: mostra solo foto non taggate.
- Icona filtro in basso a destra sull'immagine: pulisce i filtri attivi.

![Filtro attivo](screenshots/annotated/07-photosorter-filter-active.png)

## 7) Gestire i tag osservazioni
- Apri **Edit tags**.
- Aggiungi/rimuovi tag se necessario.
- **Save debug log** per supporto.

![Edit Tags](screenshots/annotated/09-photosorter-edit-tags-modal.png)

## 8) Modificare il report in AutoBericht
1. Scegli un capitolo a sinistra.
2. Modifica **Finding** / **Recommendation**.
3. Usa **Include**, **Done**, priorità e filtri sidebar.
4. Per ogni riga imposta **Library Off / Append / Replace**:
   - **Off**: non scrivere nella library.
   - **Append**: aggiunge il testo corrente alla voce library esistente.
   - **Replace**: sostituisce completamente la voce library esistente.

![Modifica capitolo 4.8](screenshots/annotated/04-chapter-48-editing.png)

## 9) Organizzare il capitolo 4.8
- In 4.8 clicca **Organize 4.8**.
- Controlla ordine/filtro e conferma con **Apply**.

![Organize 4.8](screenshots/annotated/05-organize-48-modal.png)

## 10) Export Word, PowerPoint e Library
- Nella pagina **Project**:
  - **Word Export**
  - **PowerPoint Export (Report)** / **PowerPoint Export (Training)**
  - **Generate / Update Library**
  - **Export Library Excel**

## 11) Logica dati (importante)
- `project_sidecar.json`: stato del progetto (risposte, selezioni, ordine, tag foto, testi modificati e azione library scelta per ogni riga).
- `library_*.json`: contenuti riutilizzabili (testi/tag) per i progetti successivi.
- `Append`/`Replace` viene applicato solo dopo **Generate / Update Library**.
