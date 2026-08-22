const test = require("node:test");
const assert = require("node:assert/strict");
const { loadBrowserScripts } = require("./helpers");

test("self-assessment requires the named sheet and explicit header columns", () => {
  const api = loadBrowserScripts(["mini/shared/self-assessment.js"]).AutoBerichtSelfAssessment;
  assert.throws(() => api.findAssessmentSheetName(["Sheet1"]), /recognized self-assessment sheet/);
  assert.throws(() => api.parseRows([["Nr", "Question"], ["1.1", "Test"]]), /ID, Yes and No/);
});

test("self-assessment rejects conflicting marks and duplicate IDs", () => {
  const api = loadBrowserScripts(["mini/shared/self-assessment.js"]).AutoBerichtSelfAssessment;
  assert.throws(
    () => api.parseRows([["Nr", "Question", "Oui", "Non"], ["1.1", "Q", "x", "x"]]),
    /Conflicting Yes and No/,
  );
  assert.throws(
    () => api.parseRows([
      ["Nr", "Question", "Ja", "Nein"],
      ["1.1", "Q", "x", ""],
      ["1.1", "Q again", "", "x"],
    ]),
    /Duplicate self-assessment ID/,
  );
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

test("self-assessment counts only meaningful rows and validates project identity", () => {
  const api = loadBrowserScripts(["mini/shared/self-assessment.js"]).AutoBerichtSelfAssessment;
  const rows = [["Nr", "Question", "Oui", "Non", "Remarque"]];
  for (let index = 1; index <= 10; index += 1) {
    rows.push([`1.${index}`, `Q${index}`, index === 1 ? "x" : "", "", index === 2 ? "note" : ""]);
  }
  const parsed = api.parseRows(rows);
  assert.equal(parsed.entries.length, 2);
  assert.deepEqual(Array.from(parsed.entries, (entry) => entry.id), ["1.1", "1.2"]);
  const coverage = api.validateProjectCoverage(parsed, parsed.structuralIds);
  assert.equal(coverage.matched, 10);
  assert.throws(() => api.validateProjectCoverage(parsed, ["9.9"]), /does not match this project/);
});

test("markdown drops unsafe and relative link targets", () => {
  const api = loadBrowserScripts(["mini/shared/markdown.js"]).AutoBerichtMarkdown;
  const unsafe = api.markdownToHtml("[click](javascript:alert(1))");
  assert.equal(unsafe.includes("href="), false);
  const safe = api.markdownToHtml("[Suva](https://www.suva.ch/)");
  assert.match(safe, /href="https:\/\/www\.suva\.ch\/"/);
  assert.match(safe, /noopener noreferrer/);
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

test("locale translation lookup uses the selected locale", () => {
  const document = {
    documentElement: { setAttribute() {} },
    querySelectorAll() { return []; },
  };
  const api = loadBrowserScripts(["mini/shared/i18n.js"], { document }).AutoBerichtI18n;
  api.setLocale("fr-CH");
  assert.equal(api.t("project_meta_locale_select"), "Choisir la langue");
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
