(() => {
  const uiText = {
    "markdown_hint_title": "Markdown:",
    "markdown_hint_bold": "**bold**",
    "markdown_hint_italic": "*italic*",
    "markdown_hint_link": "[text](url)",
    "markdown_hint_list": "- List item",
    "markdown_hint_paragraph": "Blank line = new paragraph",
    "markdown_hint_linebreak": "Line breaks kept",
    "project_tool_import_title": "Import Self-Assessment",
    "project_tool_import_hint": "Pick the self-assessment file. It will be imported and copied into the project inputs folder.",
    "project_export_card_title": "Word Export",
    "project_export_card_hint": "Use the DOCX template from the project templates folder and create a report in outputs using the current sidecar data.",
    "project_export_ppt_report": "PowerPoint Export (Report)",
    "project_export_ppt_report_idms": "PowerPoint Export (Report IDMS)",
    "project_export_ppt_training": "PowerPoint Export (Training)",
    "project_export_ppt_card_title": "PowerPoint Export",
    "project_export_ppt_card_hint": "Use the PPTX template from the project templates folder and create slides in outputs using the current sidecar data.",
    "project_export_ppt_missing": "PowerPoint exporter is not available.",
    "project_export_ppt_running_report": "Exporting report slides ...",
    "project_export_ppt_running_report_idms": "Exporting report slides for IDMS ...",
    "project_export_ppt_running_training": "Exporting training slides ...",
    "project_export_ppt_done": "PowerPoint export complete.",
    "project_tool_library_title": "Library Export",
    "project_tool_library_hint": "Generate or update the library from the current project content.",
    "project_tool_library_excel_title": "Library Excel Export",
    "project_tool_library_excel_hint": "Export the current library JSON into an Excel workbook (human-readable and machine-ingestible).",
    "project_tool_library_excel": "Export Library Excel",
    "project_tool_library_excel_missing": "Library Excel exporter is not available.",
    "project_tool_action_plan_title": "Action Plan Export",
    "project_tool_action_plan_hint": "Use the Excel template from the project templates folder and create the action-plan workbook in outputs from the current report items.",
    "project_tool_action_plan": "Export Action Plan Excel",
    "project_tool_action_plan_running": "Exporting action plan ...",
    "project_tool_action_plan_done": "Action plan export complete.",
    "project_tool_log_title": "Debug Log",
    "project_tool_log_hint": "Save the current debug log to share diagnostics when troubleshooting.",
    "project_bootstrap_hint": "This folder is empty. Select report language to load seed content.",
    "project_meta_locale_select": "Select language",
    "chapter0_front_matter": "Customer context (introductory text)",
    "chapter0_front_matter_hint": "This text appears before the A/B/C summary points. Separate paragraphs with a blank line.",
    "checklist_button_title": "Suva checklists",
    "checklist_button_label": "Checklists",
    "checklist_overlay_title": "Suva checklists",
    "checklist_overlay_hint": "Click Copy to copy “See also: … [link]”. Use the filters to narrow the list. Source: Suva checklist list 67000.",
    "checklist_overlay_fallback": "No checklist list for this language; showing German list.",
    "checklist_copy_title": "Copy checklist link",
    "checklist_copy_button": "Copy",
    "checklist_filter_all": "All",
    "checklist_filter_ekas": "EKAS",
    "checklist_filter_vital": "Vital rules",
    "checklist_header_title": "Title",
    "checklist_header_link": "Link",
    "checklist_header_copy": "Copy",
    "checklist_see_also": "See also",
    "checklist_industry_transport": "Transport & storage",
    "checklist_industry_transport_short": "Tr/St",
    "checklist_industry_metall": "Metal",
    "checklist_industry_metall_short": "Met",
    "checklist_industry_buero": "Office",
    "checklist_industry_buero_short": "Off",
    "checklist_industry_forst": "Forestry",
    "checklist_industry_forst_short": "For",
    "checklist_industry_holz": "Wood",
    "checklist_industry_holz_short": "Wood",
    "checklist_industry_bau": "Construction & installations",
    "checklist_industry_bau_short": "Constr.",
    "checklist_industry_uebrige": "Other",
    "checklist_industry_uebrige_short": "Other",
    "photosorter_report_description": "Chapters and subchapters (1.x / 1.2 etc.)",
    "photosorter_observations_description": "These tags are used as Chapter 4.8 in the report.",
    "photosorter_training_description": "Seminar/training categories.",
    "project_logo_hint": "Place logo.png/logo.jpg in inputs, then select the logo manually. The app writes outputs/logo-large.png and outputs/logo-small.png.",
    "project_spider_hint": "Adjust chapter overrides and preview the spider chart live.",
    "chapter_preview_button": "Preview chapter",
    "chapter_preview_empty": "No included findings in this chapter.",
    "chapter_preview_title": "Word Export Preview",
    "obs_drag_disabled_title": "Enable Manual + All and clear the search to reorder",
    "obs_drag_title": "Drag to reorder",
    "obs_filter_all": "All",
    "obs_filter_no_photos": "No Photos",
    "obs_filter_with_photos": "With Photos",
    "obs_no_pending": "No pending changes",
    "obs_organize_button": "Organize 4.8",
    "obs_pending": "Pending changes",
    "obs_reorder_limited": "Reorder requires Manual + All and an empty search",
    "obs_reorder_ready": "Reorder enabled",
    "obs_showing": "Showing",
    "obs_sort_alpha_asc": "A–Z",
    "obs_sort_alpha_desc": "Z–A",
    "obs_sort_count_asc": "Count ↑",
    "obs_sort_count_desc": "Count ↓",
    "obs_sort_manual": "Manual",
    "obs_sort_prio_asc": "Priority 1–4",
    "preview_finding": "Finding",
    "preview_recommendation": "Recommendation",
    "project_export_no_vba": "Word Export",
    "project_export_running": "Exporting ...",
    "project_logo_import": "Import and resize logo",
    "project_logo_processing": "Processing ...",
    "project_logo_title": "Logo files",
    "project_meta_address": "Address",
    "project_meta_backup": "Auto-backup (minutes)",
    "project_meta_city": "City",
    "project_meta_co_initials": "Co-moderator initials",
    "project_meta_co_moderator": "Co-moderator",
    "project_meta_company": "Company",
    "project_meta_companyid": "Company ID",
    "project_meta_created": "Created",
    "project_meta_locale": "Language",
    "project_meta_moderator": "Moderator",
    "project_meta_moderator_initials": "Moderator initials",
    "project_meta_postal": "Postal code",
    "project_meta_title": "Project details",
    "project_page_nav": "Project",
    "project_spider_apply": "Apply spider adjustments",
    "project_spider_recalc": "Recalculate",
    "project_spider_title": "Spider editor",
    "project_tool_import_self": "Import Self-Assessment",
    "project_tool_library": "Generate or update library",
    "project_tool_log": "Save diagnostic log",
    "project_open_folder_first": "Open project folder first.",
    "project_export_missing": "Word exporter is not available.",
    "project_export_done": "Export complete.",
    "project_logo_missing": "Logo pipeline is not available.",
    "project_logo_done": "Logo assets prepared.",
    "project_spider_saved": "Spider overrides saved.",
    "status_fs_api_unavailable": "File System Access API is not available. Open via http://localhost in Edge or Chrome to enable file access.",
    "status_autosaved": "Autosaved.",
    "status_autosave_failed": "Autosave failed: {error}",
    "status_select_project_folder": "Select a project folder to start.",
    "status_restore_failed": "Saved project folder could not be restored: {error}",
    "status_app_startup_failed": "AutoBericht startup failed: {error}",
    "status_photosorter_startup_failed": "PhotoSorter startup failed: {error}",
    "status_demo_photos_failed": "Demo photos failed: {error}",
    "status_selected_folder": "Selected folder: {name}",
    "status_folder_pick_failed": "Folder selection was canceled or failed: {error}",
    "status_project_load_failed": "Project load failed: {error}",
    "status_sidecar_saved": "Saved project_sidecar.json.",
    "status_save_failed": "Save failed: {error}",
    "status_seed_loaded": "Loaded seed data.",
    "status_seed_load_failed": "Seed load failed: {error}",
    "status_log_saved": "Saved log ({location}): {filename}",
    "status_log_save_failed": "Log save failed: {error}",
    "status_settings_library_saved": "Settings saved. Matching language library loaded and sidecar saved.",
    "status_settings_seed_failed": "Settings saved, but language library setup failed: {error}",
    "status_settings_saved": "Settings saved.",
    "status_settings_sidecar_failed": "Settings saved, but sidecar save failed: {error}",
    "status_settings_unsaved_sidecar": "Settings saved. Remember to save the sidecar.",
    "status_library_update_failed": "Library update failed: {error}",
    "status_spider_recalculating": "Recalculating spider values ...",
    "status_spider_ready": "Spider values ready.",
    "status_spider_calc_failed": "Spider calculation failed: {error}",
    "status_spider_overrides_autosave": "Spider overrides saved; autosave will persist them.",
    "status_spider_save_failed": "Spider save failed: {error}",
    "status_import_capability_missing": "File picker or SheetJS is not available in this browser.",
    "status_validator_missing": "Self-assessment validator is not available.",
    "status_self_assessment_imported": "Imported self-assessment answers ({count}).",
    "status_self_assessment_copied": " Copied {filename} to inputs.",
    "status_self_assessment_copy_failed": " Could not copy {filename} to inputs: {error}",
    "status_import_failed": "Import failed: {error}",
    "status_photo_not_found": "Photo not found: {path}",
    "status_checklist_copied": "Copied: {title}",
    "status_copy_failed": "Copy failed.",
    "status_checklist_load_failed": "Checklist load failed: {error}",
    "status_import_tool_missing": "Import tool is not available.",
    "status_library_tool_missing": "Library tool is not available.",
    "status_log_tool_missing": "Log tool is not available.",
    "status_update_failed": "Update failed: {error}",
    "status_exported_file": "Exported {filename}",
    "status_export_failed": "Export failed: {error}",
    "status_logo_processing_failed": "Logo processing failed: {error}",
    "status_invalid_library": "Invalid library: {error}",
    "status_invalid_seed": "Invalid seed: {error}",
    "status_loading_project_folder": "Loading project folder: {name}",
    "status_loaded_sidecar": "Loaded project_sidecar.json.{warning}",
    "status_new_project_ready": "New project ready. Choose the report language to load its matching library.{warning}",
    "status_library_excel_exported": "Library Excel exported: {filename}",
    "status_knowledge_base_not_found": "Knowledge base seed not found.",
    "status_invalid_knowledge_base": "Invalid knowledge base: {error}",
    "status_library_zero_changes": "Library updated with 0 changes. Set Library to Append or Replace on rows, Chapter 0 context, or chapter positives.",
    "status_library_updated": "Library updated ({count} changes).",
    "status_librarymaker_loaded_my": "Loaded My Library: {filename}",
    "status_librarymaker_load_project_first": "Load your project library first.",
    "status_file_picker_unavailable": "File picker is not available in this browser.",
    "status_librarymaker_loaded_other": "Loaded Other Library: {filename}",
    "status_librarymaker_load_other_failed": "Loading Other Library failed: {error}",
    "status_librarymaker_project_pick_failed": "Project selection failed: {error}",
    "status_librarymaker_load_my_failed": "Loading My Library failed: {error}",
    "status_librarymaker_saved_my": "Saved My Library: {filename}",
    "status_librarymaker_select_project": "Select a project folder to load your library.",
    "status_project_folder": "Project folder: {name}",
    "status_project_pick_canceled": "Project selection canceled.",
    "status_photo_import_module_missing": "Photo import module is not available.",
    "status_imported_photos": "Imported {count} photos",
    "status_skipped_photos": "skipped {count} already present",
    "status_moved_videos": "moved {count} videos to photos/videos",
    "status_photo_export_module_missing": "Photo export module is not available.",
    "status_photo_folder_missing_scan": "No photo folder found. Import or scan photos first.",
    "status_exported_photos": "Exported {label} to {folder}",
    "status_photo_folder_missing_import": "No photo folder found. Import photos first.",
    "status_photo_load_failed": "Photo load failed: {error}",
    "status_photosorter_loaded_sidecar": "Loaded project_sidecar.json.",
    "status_photosorter_saved_sidecar": "Saved photo tags to project_sidecar.json.",
    "status_scanning_photos": "Scanning photos ...",
    "status_loaded_photos": "Loaded {count} photos from {folder}.",
    "status_loading_demo_photos": "Loading demo photos ...",
    "status_loaded_demo_photos": "Loaded {count} demo photos.",
    "status_show_unsorted": "Show Unsorted",
    "status_show_all": "Show All",
    "status_missing_raw_photos": "Missing photos/raw. Expected three-letter subfolders such as pm1, pm2, pm3, or abc.",
    "status_no_raw_images": "No raw images found.",
    "status_importing_photos": "Importing photos {processed}/{total} ({percent}%) • skipped {skipped}",
    "status_preparing_export": "Preparing export ...",
    "status_exporting_photos": "Exporting photos {completed}/{total}",
    "status_seed_loaded_deferred": "Language library loaded. The sidecar will be saved when leaving the Project page.",
    "status_seed_loaded_saved": "Language library loaded and saved to project_sidecar.json.",
    "status_photo_count": "{count} photos",
    "status_photo_copy_count": "{copies} copies ({photos} photos)",
    "status_photosorter_fresh": "Sidecar not found; starting fresh.",
    "status_photosorter_library_tags": "Sidecar not found; loaded tags from the library.",
    "status_photosorter_no_library": "Sidecar and library not found; starting fresh.",
    "photosorter_image_meta": "Image {current} of {filtered} • Total {total} • Unsorted {unsorted}",
    "photosorter_no_photo": "No photo loaded",
    "photosorter_filter_tags": "Filter tags"
};

  // Report locale affects authored report content only. Steering controls stay
  // in simple English so support conversations use one shared vocabulary.
  const reportText = {
    en: { checklist_see_also: "See also" },
    "de-CH": { checklist_see_also: "Siehe auch" },
    "fr-CH": { checklist_see_also: "Voir aussi" },
    "it-CH": { checklist_see_also: "Vedere anche" },
  };

  // Instructional hint paragraphs follow the report language so that a French-
  // or Italian-speaking engineer reads the guidance in their language. Controls,
  // card titles, status messages and errors stay in English (see uiText).
  // Only keys that the UI resolves through tHint() belong here.
  const hintText = {
    "de-CH": {
      project_tool_import_hint: "Datei der Selbstbeurteilung auswählen. Sie wird importiert und in den Projektordner inputs kopiert.",
      project_export_card_hint: "DOCX-Vorlage aus dem Projektordner templates verwenden und mit den aktuellen Sidecar-Daten einen Bericht in outputs erstellen.",
      project_export_ppt_card_hint: "PPTX-Vorlage aus dem Projektordner templates verwenden und mit den aktuellen Sidecar-Daten Folien in outputs erstellen.",
      project_tool_library_hint: "Library aus dem aktuellen Projektinhalt erzeugen oder aktualisieren.",
      project_tool_library_excel_hint: "Aktuelle Library-JSON als Excel-Arbeitsmappe exportieren (lesbar und maschinenverarbeitbar).",
      project_tool_action_plan_hint: "Excel-Vorlage aus dem Projektordner templates verwenden und den Aktionsplan in outputs mit den aktuellen Berichtspunkten erstellen.",
      project_tool_log_hint: "Aktuelles Debug-Log speichern, um Diagnosen bei der Fehlersuche zu teilen.",
      project_logo_hint: "Logo.png/logo.jpg in inputs ablegen und dann die Logo-Datei manuell auswählen. Die App schreibt outputs/logo-large.png und outputs/logo-small.png.",
      project_spider_hint: "Kapitel-Overrides anpassen und die Spider-Grafik live prüfen.",
      chapter0_front_matter_hint: "Dieser Text erscheint vor den Punkten A/B/C. Absätze mit einer Leerzeile trennen.",
      checklist_overlay_hint: "Auf Kopieren klicken, um „Siehe auch: … [Link]“ zu kopieren. Filter anklicken zum Filtern. Quelle: Suva-Checklistenliste 67000.",
      photosorter_report_description: "Kapitel und Unterkapitel (1.x / 1.2 usw.)",
      photosorter_observations_description: "Diese Tags werden als Kapitel 4.8 im Bericht verwendet.",
      photosorter_training_description: "Seminar-/Schulungskategorien.",
    },
    "fr-CH": {
      project_tool_import_hint: "Choisissez le fichier d'autoevaluation. Il sera importe et copie dans le dossier inputs du projet.",
      project_export_card_hint: "Utilisez le modele DOCX du dossier templates du projet et creez un rapport dans outputs avec les donnees sidecar actuelles.",
      project_export_ppt_card_hint: "Utilisez le modele PPTX du dossier templates du projet et creez des diapositives dans outputs avec les donnees sidecar actuelles.",
      project_tool_library_hint: "Generer ou mettre a jour la bibliotheque a partir du contenu actuel du projet.",
      project_tool_library_excel_hint: "Exporter la bibliotheque JSON actuelle vers un classeur Excel (lisible et exploitable).",
      project_tool_action_plan_hint: "Utilisez le modele Excel du dossier templates du projet et creez le plan d'action dans outputs a partir des points actuels du rapport.",
      project_tool_log_hint: "Enregistrez le journal de debug actuel pour partager les diagnostics en cas de probleme.",
      project_logo_hint: "Placez logo.png/logo.jpg dans inputs puis selectionnez manuellement le logo. L'application ecrit outputs/logo-large.png et outputs/logo-small.png.",
      project_spider_hint: "Ajustez les surcharges par chapitre et previsualisez le spider en direct.",
      chapter0_front_matter_hint: "Ce texte apparait avant les points A/B/C. Separez les paragraphes par une ligne vide.",
      checklist_overlay_hint: "Cliquer sur Copier pour copier « Voir aussi : … [lien] ». Utilisez les filtres pour affiner la liste. Source: liste de contrôle Suva 67000.",
      photosorter_report_description: "Chapitres et sous-chapitres (1.x / 1.2 etc.).",
      photosorter_observations_description: "Ces tags sont utilises comme chapitre 4.8 dans le rapport.",
      photosorter_training_description: "Categories de seminaire/formation.",
    },
    "it-CH": {
      project_tool_import_hint: "Scegli il file di autovalutazione. Verra importato e copiato nella cartella inputs del progetto.",
      project_export_card_hint: "Usa il modello DOCX dalla cartella templates del progetto e crea un rapporto in outputs con i dati sidecar correnti.",
      project_export_ppt_card_hint: "Usa il modello PPTX dalla cartella templates del progetto e crea le slide in outputs con i dati sidecar correnti.",
      project_tool_library_hint: "Genera o aggiorna la libreria dal contenuto attuale del progetto.",
      project_tool_library_excel_hint: "Esporta l'attuale libreria JSON in una cartella Excel (leggibile e processabile).",
      project_tool_action_plan_hint: "Usa il modello Excel dalla cartella templates del progetto e crea il piano d'azione in outputs a partire dagli attuali punti del rapporto.",
      project_tool_log_hint: "Salva il log di debug corrente per condividere la diagnostica durante la risoluzione dei problemi.",
      project_logo_hint: "Inserisci logo.png/logo.jpg in inputs e poi seleziona manualmente il logo. L'app scrive outputs/logo-large.png e outputs/logo-small.png.",
      project_spider_hint: "Regola gli override per capitolo e visualizza il grafico spider in tempo reale.",
      chapter0_front_matter_hint: "Questo testo appare prima dei punti A/B/C. Separa i paragrafi con una riga vuota.",
      checklist_overlay_hint: "Fare clic su Copia per copiare « Vedere anche: … [link] ». Usa i filtri per restringere la lista. Fonte: lista di controllo Suva 67000.",
      photosorter_report_description: "Capitoli e sottocapitoli (1.x / 1.2 ecc.).",
      photosorter_observations_description: "Questi tag sono usati come capitolo 4.8 nel rapporto.",
      photosorter_training_description: "Categorie per seminari/formazione.",
    },
  };

  let current = "en";

  const resolveSpellcheckLang = (locale) => {
    const base = String(locale || "").toLowerCase().split("-")[0];
    if (base === "de" || base === "fr" || base === "it" || base === "en") return base;
    return "en";
  };

  const resolveLocale = (locale) => {
    const value = String(locale || "").trim();
    if (!value) return "en";
    if (reportText[value]) return value;
    const base = value.toLowerCase().split("-")[0];
    if (base === "de") return "de-CH";
    if (base === "fr") return "fr-CH";
    if (base === "it") return "it-CH";
    return "en";
  };

  const applyLocaleToDocument = (locale) => {
    if (typeof document === "undefined") return;
    const spellLang = resolveSpellcheckLang(locale);
    document.documentElement?.setAttribute("lang", locale);
    document.querySelectorAll("textarea").forEach((node) => {
      node.setAttribute("lang", spellLang);
      node.setAttribute("spellcheck", "true");
      node.spellcheck = true;
    });
  };

  const setLocale = (locale) => {
    current = resolveLocale(locale);
    applyLocaleToDocument(current);
    apply(typeof document !== "undefined" ? document : null);
  };

  const t = (key, fallback) => uiText[key] ?? fallback ?? key;

  const interpolate = (template, values = {}) => String(template ?? "").replace(/\{([A-Za-z0-9_]+)\}/g, (match, name) => (
    Object.prototype.hasOwnProperty.call(values, name) ? String(values[name] ?? "") : match
  ));

  const tf = (key, fallback, values) => interpolate(t(key, fallback), values);
  const tHint = (key, fallback) => hintText[current]?.[key] ?? t(key, fallback);

  const tReport = (key, fallback) => (
    reportText[current]?.[key] ?? reportText.en?.[key] ?? fallback ?? key
  );

  const apply = (root = document) => {
    if (!root?.querySelectorAll) return;
    root.querySelectorAll("[data-i18n]").forEach((node) => {
      const key = node.getAttribute("data-i18n");
      if (key) node.textContent = t(key, node.textContent);
    });
  };

  window.AutoBerichtI18n = {
    t,
    tf,
    tHint,
    tReport,
    setLocale,
    apply,
    resolveSpellcheckLang,
    resolveLocale,
    locales: { en: uiText },
    reportText,
  };
})();
