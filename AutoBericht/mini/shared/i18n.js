(() => {
  const locales = {
    en: {
      markdown_hint_title: "Markdown:",
      markdown_hint_bold: "**bold**",
      markdown_hint_italic: "*italic*",
      markdown_hint_link: "[text](url)",
      markdown_hint_list: "- List item",
      markdown_hint_paragraph: "Blank line = new paragraph",
      markdown_hint_linebreak: "Line breaks kept",
      checklist_button_title: "Suva checklists",
      checklist_button_label: "Checklists",
      checklist_overlay_title: "Suva checklists",
      checklist_overlay_hint: "Click Copy to copy “See also: … [link]”. Use the filters to narrow the list. Source: Suva checklist list 67000.",
      checklist_overlay_fallback: "No checklist list for this language; showing German list.",
      checklist_copy_title: "Copy checklist link",
      checklist_copy_button: "Copy",
      checklist_filter_all: "All",
      checklist_filter_ekas: "EKAS",
      checklist_filter_vital: "Vital rules",
      checklist_header_title: "Title",
      checklist_header_link: "Link",
      checklist_header_copy: "Copy",
      checklist_see_also: "See also",
      checklist_industry_transport: "Transport & storage",
      checklist_industry_transport_short: "Tr/St",
      checklist_industry_metall: "Metal",
      checklist_industry_metall_short: "Met",
      checklist_industry_buero: "Office",
      checklist_industry_buero_short: "Off",
      checklist_industry_forst: "Forestry",
      checklist_industry_forst_short: "For",
      checklist_industry_holz: "Wood",
      checklist_industry_holz_short: "Wood",
      checklist_industry_bau: "Construction & installations",
      checklist_industry_bau_short: "Constr.",
      checklist_industry_uebrige: "Other",
      checklist_industry_uebrige_short: "Other",
    },
    "de-CH": {
      markdown_hint_title: "Markdown:",
      markdown_hint_bold: "**bold**",
      markdown_hint_italic: "*italic*",
      markdown_hint_link: "[text](url)",
      markdown_hint_list: "- List item",
      markdown_hint_paragraph: "Blank line = new paragraph",
      markdown_hint_linebreak: "Line breaks kept",
      checklist_button_title: "Suva-Checklisten",
      checklist_button_label: "Checklisten",
      checklist_overlay_title: "Suva-Checklisten",
      checklist_overlay_hint: "Auf Kopieren klicken, um „Siehe auch: … [Link]“ zu kopieren. Filter anklicken zum Filtern. Quelle: Suva-Checklistenliste 67000.",
      checklist_overlay_fallback: "Keine Checklistenliste für diese Sprache; deutsche Liste wird angezeigt.",
      checklist_copy_title: "Checkliste kopieren",
      checklist_copy_button: "Kopieren",
      checklist_filter_all: "Alle",
      checklist_filter_ekas: "EKAS",
      checklist_filter_vital: "Lebenswichtige Regeln",
      checklist_header_title: "Titel",
      checklist_header_link: "Link",
      checklist_header_copy: "Kopieren",
      checklist_see_also: "Siehe auch",
      checklist_industry_transport: "Transport und Lagerung",
      checklist_industry_transport_short: "Tr/Lg",
      checklist_industry_metall: "Metall",
      checklist_industry_metall_short: "Met",
      checklist_industry_buero: "Büro",
      checklist_industry_buero_short: "Bü",
      checklist_industry_forst: "Forst",
      checklist_industry_forst_short: "For",
      checklist_industry_holz: "Holz",
      checklist_industry_holz_short: "Holz",
      checklist_industry_bau: "Bau- und Installationsgewerbe",
      checklist_industry_bau_short: "Bau",
      checklist_industry_uebrige: "Übrige",
      checklist_industry_uebrige_short: "Übr",
    },
    "fr-CH": {
      checklist_button_title: "Listes de contrôle Suva",
      checklist_button_label: "Checklists",
      checklist_overlay_title: "Listes de contrôle Suva",
      checklist_overlay_hint: "Cliquer sur Copier pour copier « Voir aussi : … [lien] ». Utilisez les filtres pour affiner la liste. Source: liste de contrôle Suva 67000.",
      checklist_overlay_fallback: "Aucune liste pour cette langue; affichage de la liste allemande.",
      checklist_copy_title: "Copier la check-list",
      checklist_copy_button: "Copier",
      checklist_filter_all: "Tous",
      checklist_filter_ekas: "CFST",
      checklist_filter_vital: "Règles vitales",
      checklist_header_title: "Titre",
      checklist_header_link: "Lien",
      checklist_header_copy: "Copier",
      checklist_see_also: "Voir aussi",
      checklist_industry_transport: "Transport et stockage",
      checklist_industry_transport_short: "Tr/Stock",
      checklist_industry_metall: "Métal",
      checklist_industry_metall_short: "Mét",
      checklist_industry_buero: "Bureau",
      checklist_industry_buero_short: "Bur",
      checklist_industry_forst: "Forêt",
      checklist_industry_forst_short: "For",
      checklist_industry_holz: "Bois",
      checklist_industry_holz_short: "Bois",
      checklist_industry_bau: "Bâtiment et installations",
      checklist_industry_bau_short: "Bât",
      checklist_industry_uebrige: "Autres",
      checklist_industry_uebrige_short: "Aut",
    },
    "it-CH": {
      checklist_button_title: "Liste di controllo Suva",
      checklist_button_label: "Checklist",
      checklist_overlay_title: "Liste di controllo Suva",
      checklist_overlay_hint: "Fare clic su Copia per copiare « Vedere anche: … [link] ». Usa i filtri per restringere la lista. Fonte: lista di controllo Suva 67000.",
      checklist_overlay_fallback: "Nessuna lista per questa lingua; visualizzazione della lista tedesca.",
      checklist_copy_title: "Copia la checklist",
      checklist_copy_button: "Copia",
      checklist_filter_all: "Tutti",
      checklist_filter_ekas: "CFSL",
      checklist_filter_vital: "Regole vitali",
      checklist_header_title: "Titolo",
      checklist_header_link: "Link",
      checklist_header_copy: "Copia",
      checklist_see_also: "Vedere anche",
      checklist_industry_transport: "Trasporto e stoccaggio",
      checklist_industry_transport_short: "Tras",
      checklist_industry_metall: "Metallo",
      checklist_industry_metall_short: "Met",
      checklist_industry_buero: "Ufficio",
      checklist_industry_buero_short: "Uff",
      checklist_industry_forst: "Aziende forestali",
      checklist_industry_forst_short: "For",
      checklist_industry_holz: "Legno",
      checklist_industry_holz_short: "Legno",
      checklist_industry_bau: "Edilizia, installatori",
      checklist_industry_bau_short: "Edil",
      checklist_industry_uebrige: "Altro",
      checklist_industry_uebrige_short: "Alt",
    },
  };

  let current = "en";

  const resolveSpellcheckLang = (locale) => {
    const base = String(locale || "").toLowerCase().split("-")[0];
    if (base === "de" || base === "fr" || base === "it" || base === "en") return base;
    return "en";
  };

  const applyLocaleToDocument = (locale) => {
    if (typeof document === "undefined") return;
    const spellLang = resolveSpellcheckLang(locale);
    if (document.documentElement) {
      document.documentElement.setAttribute("lang", locale);
    }
    document.querySelectorAll("textarea").forEach((node) => {
      node.setAttribute("lang", spellLang);
      node.setAttribute("spellcheck", "true");
      node.spellcheck = true;
    });
  };

  const resolveLocale = (locale) => {
    if (!locale) return "en";
    if (locales[locale]) return locale;
    const base = String(locale).split("-")[0];
    if (base === "de") return "de-CH";
    if (base === "fr") return "fr-CH";
    if (base === "it") return "it-CH";
    if (locales[base]) return base;
    return "en";
  };

  const setLocale = (locale) => {
    current = resolveLocale(locale);
    applyLocaleToDocument(current);
  };

  const t = (key, fallback) => {
    const value = locales[current]?.[key] ?? locales.en?.[key];
    return value ?? fallback ?? key;
  };

  const apply = (root = document) => {
    if (!root?.querySelectorAll) return;
    root.querySelectorAll("[data-i18n]").forEach((node) => {
      const key = node.getAttribute("data-i18n");
      if (!key) return;
      node.textContent = t(key, node.textContent);
    });
  };

  window.AutoBerichtI18n = {
    t,
    setLocale,
    apply,
    resolveSpellcheckLang,
    locales,
  };
})();
