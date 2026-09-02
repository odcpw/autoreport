const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const { loadBrowserScripts } = require("./helpers");

test("self-assessment falls back to the first sheet and the Suva column layout", () => {
  const api = loadBrowserScripts(["mini/shared/self-assessment.js"]).AutoBerichtSelfAssessment;
  assert.equal(api.findAssessmentSheetName(["Sheet1"]), "Sheet1");
  assert.equal(api.findAssessmentSheetName(["Anleitung", "Selbstbeurteilung Kunde"]), "Selbstbeurteilung Kunde");
  const parsed = api.parseRows([["Nr", "Question"], ["1.1", "", "Test", "x", ""]]);
  assert.equal(parsed.headerRowIndex, -1);
  assert.deepEqual(Array.from(parsed.entries, (entry) => [entry.id, entry.answer]), [["1.1", 1]]);
});

test("a row marked both Ja and Nein imports as Nein, and a repeated ID keeps the later row", () => {
  const api = loadBrowserScripts(["mini/shared/self-assessment.js"]).AutoBerichtSelfAssessment;
  const both = api.parseRows([["Nr", "Question", "Oui", "Non"], ["1.1", "Q", "x", "x"]]);
  assert.equal(both.entries[0].answer, 0);
  const repeated = api.parseRows([
    ["Nr", "Question", "Ja", "Nein"],
    ["1.1", "Q", "x", ""],
    ["1.1", "Q again", "", "x"],
  ]);
  assert.equal(repeated.entries.length, 1);
  assert.equal(repeated.entries[0].answer, 0);
  assert.deepEqual(Array.from(repeated.structuralIds), ["1.1"]);
});

test("English No. ID header is distinct from the No answer column", () => {
  const api = loadBrowserScripts(["mini/shared/self-assessment.js"]).AutoBerichtSelfAssessment;
  const parsed = api.parseRows([
    ["No.", "Question", "Yes", "No"],
    ["1.1", "Question", "", "x"],
  ]);
  assert.equal(parsed.columns.id, 0);
  assert.equal(parsed.columns.no, 3);
  assert.equal(parsed.entries[0].answer, 0);
});

test("bundled self-assessment workbooks parse sparse Excel header rows", () => {
  const XLSX = require("../libs/sheetjs/xlsx.full.min.js");
  const api = loadBrowserScripts(["mini/shared/self-assessment.js"]).AutoBerichtSelfAssessment;
  const templatesDir = path.resolve(__dirname, "../project-template/templates");
  const workbookNames = [
    "Selbstbeurteilung Integrierte Sicherheit d.V14.xlsx",
    "Selbstbeurteilung Integrierte Sicherheit f.V15.xlsx",
    "Selbstbeurteilung Integrierte Sicherheit i.V14.xlsx",
  ];

  workbookNames.forEach((workbookName) => {
    const workbook = XLSX.read(fs.readFileSync(path.join(templatesDir, workbookName)), { type: "buffer" });
    const sheetName = api.findAssessmentSheetName(workbook.SheetNames);
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, blankrows: false });
    // The bundled templates are blank. Mark the first numbered row so the
    // parser can complete the same path as a filled customer workbook.
    rows[2][3] = "x";
    const parsed = api.parseRows(rows);

    assert.ok(parsed.structuralIds.length > 100, workbookName);
    assert.ok(parsed.entries.length > 0, workbookName);
  });
});

test("self-assessment imports only rows with an answer, comment, or evidence", () => {
  const api = loadBrowserScripts(["mini/shared/self-assessment.js"]).AutoBerichtSelfAssessment;
  const rows = [["Nr", "Question", "Oui", "Non", "Remarque"]];
  for (let index = 1; index <= 10; index += 1) {
    rows.push([`1.${index}`, `Q${index}`, index === 1 ? "x" : "", "", index === 2 ? "note" : ""]);
  }
  const parsed = api.parseRows(rows);
  assert.equal(parsed.structuralIds.length, 10);
  assert.equal(parsed.entries.length, 2);
  assert.deepEqual(Array.from(parsed.entries, (entry) => entry.id), ["1.1", "1.2"]);
});

test("markdown drops unsafe and relative link targets", () => {
  const api = loadBrowserScripts(["mini/shared/markdown.js"]).AutoBerichtMarkdown;
  const unsafe = api.markdownToHtml("[click](javascript:alert(1))");
  assert.equal(unsafe.includes("href="), false);
  assert.equal(unsafe, "<p>click</p>");
  const safe = api.markdownToHtml("[Suva](https://www.suva.ch/)");
  assert.match(safe, /href="https:\/\/www\.suva\.ch\/"/);
  assert.match(safe, /noopener noreferrer/);
});

test("browser Markdown preview uses the same styles and links as Office exports", () => {
  const api = loadBrowserScripts(["mini/shared/markdown.js"]).AutoBerichtMarkdown;
  const html = api.markdownToHtml(
    "***Important*** **bold** *italic* https://example.com/test. <unsafe>",
  );
  assert.match(html, /<strong><em>Important<\/em><\/strong>/);
  assert.match(html, /<strong>bold<\/strong>/);
  assert.match(html, /<em>italic<\/em>/);
  assert.match(html, /href="https:\/\/example\.com\/test"/);
  assert.match(html, /&lt;unsafe&gt;/);
  assert.doesNotMatch(html, /\*\*\*|\*\*bold\*\*|\*italic\*/);
});

test("export markdown parser preserves visible text and inline styles", () => {
  const api = loadBrowserScripts(["mini/shared/markdown.js"]).AutoBerichtMarkdown;
  const segments = api.parseInlineMarkdownSegments(
    "Plain **bold** *italic* ***both*** [Suva](https://www.suva.ch/) https://example.com/test.",
  );
  assert.deepEqual(
    JSON.parse(JSON.stringify(segments.map(({ type, text, url, bold, italic }) => ({ type, text, url, bold, italic })))),
    [
      { type: "text", text: "Plain ", bold: false, italic: false },
      { type: "text", text: "bold", bold: true, italic: false },
      { type: "text", text: " ", bold: false, italic: false },
      { type: "text", text: "italic", bold: false, italic: true },
      { type: "text", text: " ", bold: false, italic: false },
      { type: "text", text: "both", bold: true, italic: true },
      { type: "text", text: " ", bold: false, italic: false },
      { type: "link", text: "Suva", url: "https://www.suva.ch/", bold: false, italic: false },
      { type: "text", text: " ", bold: false, italic: false },
      { type: "link", text: "https://example.com/test", url: "https://example.com/test", bold: false, italic: false },
      { type: "text", text: ".", bold: false, italic: false },
    ],
  );
});

test("PowerPoint text renderer converts Markdown into DrawingML runs", () => {
  const context = loadBrowserScripts([
    "mini/shared/markdown.js",
    "mini/shared/report-rows.js",
    "mini/shared/pptx-export.js",
  ], {
    AutoBerichtWordDocxZip: {},
  });
  const rendered = context.AutoBerichtPptxExport.renderMarkdownTextBodyXml(
    "- **Bold** and *italic* with [Suva](https://www.suva.ch/)",
  );
  assert.match(rendered.xml, /<a:buChar char="•"\/>/);
  assert.match(rendered.xml, /<a:rPr[^>]*b="1"\/><a:t[^>]*>Bold<\/a:t>/);
  assert.match(rendered.xml, /<a:rPr[^>]*i="1"\/><a:t[^>]*>italic<\/a:t>/);
  assert.match(rendered.xml, /<a:hlinkClick r:id="rId2"\/>/);
  assert.doesNotMatch(rendered.xml, /\*\*Bold\*\*|\*italic\*|\[Suva\]\(/);
  assert.deepEqual(Array.from(rendered.hyperlinkTargets), ["https://www.suva.ch/"]);
});

test("report locale changes report wording, hints and spellcheck, while controls stay English", () => {
  let documentLang = "";
  const document = {
    documentElement: { setAttribute(name, value) { if (name === "lang") documentLang = value; } },
    querySelectorAll() { return []; },
  };
  const api = loadBrowserScripts(["mini/shared/i18n.js"], { document }).AutoBerichtI18n;
  const cases = [
    ["de-CH", "Siehe auch", /^Datei der Selbstbeurteilung/],
    ["fr-CH", "Voir aussi", /^Choisissez le fichier/],
    ["it-CH", "Vedere anche", /^Scegli il file/],
  ];
  for (const [locale, seeAlso, importHint] of cases) {
    api.setLocale(locale);
    assert.equal(documentLang, locale);
    assert.equal(api.t("project_meta_locale_select"), "Select language");
    assert.equal(api.t("project_tool_import_title"), "Import Self-Assessment");
    assert.equal(api.t("project_export_card_title"), "Word Export");
    assert.match(api.tHint("project_tool_import_hint"), importHint);
    assert.equal(api.tHint("project_tool_import_title"), "Import Self-Assessment");
    assert.equal(api.tHint("status_autosaved", "Autosaved."), "Autosaved.");
    assert.equal(api.tReport("checklist_see_also"), seeAlso);
    assert.equal(
      api.tf("status_loaded_photos", "Loaded {count} photos from {folder}.", { count: 3, folder: "photos/resized" }),
      "Loaded 3 photos from photos/resized.",
    );
  }
});

test("Chapter 0 front matter is loaded from the user library", () => {
  const context = loadBrowserScripts(["mini/shared/seeds.js"], {
    location: { protocol: "file:", href: "file:///test" },
    AutoBerichtState: {
      compareIdSegments: (a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }),
      toText: (value) => value == null ? "" : String(value),
    },
    AutoBerichtNormalize: {},
  });
  const project = context.AutoBerichtSeeds.buildProjectFromKnowledgeBase({
    schemaVersion: "1.1",
    meta: { locale: "fr-CH" },
    structure: { items: [{ id: "0.1", collapsedId: "0.1", chapterLabel: "Résumé", question: "Point" }] },
    library: {
      entries: [{ id: "0.1", finding: "", recommendation: "Résumé" }],
      chapterFrontMatter: { "0": "Contexte client\n\nDeuxième paragraphe" },
    },
    tags: { report: [], observations: [], training: [] },
  });
  const chapter0 = project.chapters.find((chapter) => chapter.id === "0");
  assert.equal(chapter0.meta.frontMatterText, "Contexte client\n\nDeuxième paragraphe");
});
