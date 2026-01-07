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
    },
    "de-CH": {
      markdown_hint_title: "Markdown:",
      markdown_hint_bold: "**bold**",
      markdown_hint_italic: "*italic*",
      markdown_hint_link: "[text](url)",
      markdown_hint_list: "- List item",
      markdown_hint_paragraph: "Blank line = new paragraph",
      markdown_hint_linebreak: "Line breaks kept",
    },
    "fr-CH": {},
    "it-CH": {},
  };

  let current = "en";

  const resolveLocale = (locale) => {
    if (!locale) return "en";
    if (locales[locale]) return locale;
    const base = String(locale).split("-")[0];
    if (locales[base]) return base;
    return "en";
  };

  const setLocale = (locale) => {
    current = resolveLocale(locale);
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
    locales,
  };
})();
