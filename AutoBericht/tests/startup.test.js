const test = require("node:test");
const assert = require("node:assert/strict");
const { loadBrowserScripts, createMemoryDirectory } = require("./helpers");

const functionModule = (...names) => Object.fromEntries(names.map((name) => [name, () => {}]));
const testI18n = {
  setLocale() {},
  t: (key, fallback) => fallback || key,
  tf: (key, fallback, values = {}) => String(fallback || key).replace(/\{(\w+)\}/g, (match, name) => String(values[name] ?? match)),
};

test("dependency guard reports the exact missing module in the page status", () => {
  const status = { textContent: "" };
  const context = loadBrowserScripts(["mini/shared/dependencies.js"], {
    document: { getElementById: () => status },
  });
  assert.throws(
    () => context.AutoBerichtDependencies.requireModules(
      { AutoBerichtMissing: ["init"] },
      { appName: "AutoBericht", statusElementId: "status" },
    ),
    /AutoBericht cannot start: required module AutoBerichtMissing did not load/,
  );
  assert.match(status.textContent, /AutoBerichtMissing/);
});

test("production entry points stop instead of substituting no-op core modules", () => {
  const cases = [
    {
      script: "mini/app.js",
      missing: "AutoBerichtSidecar",
      statusId: "status",
      globals: {
        AutoBerichtElements: functionModule("getElements"),
        AutoReportDebug: functionModule("logLine", "saveLog"),
        AutoBerichtI18n: functionModule("t", "tf", "tHint", "tReport", "setLocale", "resolveSpellcheckLang"),
        AutoBerichtMarkdown: functionModule("escapeHtml", "formatInlineMarkdown", "parseInlineMarkdownSegments"),
        AutoBerichtFsHandle: functionModule("saveHandle", "loadHandle", "requestHandlePermission"),
        AutoBerichtState: functionModule("createState"),
        AutoBerichtNormalize: functionModule("normalizeProject"),
        AutoBerichtSeeds: functionModule("validateKnowledgeBase", "buildProjectFromKnowledgeBase"),
        AutoBerichtImportSelf: functionModule("createHandler"),
        AutoBerichtRender: functionModule("init"),
        AutoBerichtSpider: functionModule("computeSpider"),
        AutoBerichtSpiderUi: functionModule("init"),
        AutoBerichtBindEvents: functionModule("bind"),
      },
    },
    {
      script: "mini/photosorter.js",
      missing: "AutoBerichtPhotoImport",
      statusId: "status-text",
      globals: {
        AutoBerichtPhotoSorterElements: functionModule("getElements"),
        AutoReportDebug: functionModule("logLine", "saveLog"),
        AutoBerichtI18n: functionModule("t", "tf", "tHint", "setLocale", "resolveSpellcheckLang"),
        AutoBerichtFsHandle: functionModule("saveHandle", "loadHandle", "requestHandlePermission"),
        AutoBerichtPhotoSorterState: functionModule("getLayoutConfig", "createState", "createRuntime"),
        AutoBerichtPhotoSorterTags: functionModule("createEmptyTagOptions"),
        AutoBerichtPhotoSorterPhotos: functionModule("init"),
        AutoBerichtPhotoSorterSidecar: functionModule("init"),
        AutoBerichtPhotoSorterRender: functionModule("init"),
        AutoBerichtPhotoSorterBindEvents: functionModule("bind"),
      },
    },
    {
      script: "mini/librarymaker.js",
      missing: "AutoBerichtSeeds",
      statusId: "status",
      globals: {
        AutoReportDebug: functionModule("logLine"),
        AutoBerichtI18n: functionModule("t", "tf", "setLocale", "resolveSpellcheckLang"),
        AutoBerichtFsHandle: functionModule("saveHandle", "loadHandle", "requestHandlePermission"),
        AutoBerichtState: functionModule("compareIdSegments", "formatChapterLabel", "getLibraryFileName", "toText"),
        AutoBerichtNormalize: functionModule("normalizeProject"),
      },
    },
  ];

  cases.forEach(({ script, missing, statusId, globals }) => {
    const status = { textContent: "" };
    assert.throws(
      () => loadBrowserScripts(["mini/shared/dependencies.js", script], {
        ...globals,
        document: { getElementById: (id) => id === statusId ? status : null },
      }),
      new RegExp(`required module ${missing} did not load`),
    );
    assert.match(status.textContent, new RegExp(missing));
  });
});

const createIo = ({ files, project, normalizeProject, seedOverrides }) => {
  const statuses = [];
  let renders = 0;
  const context = loadBrowserScripts([
    "mini/shared/sidecar-storage.js",
    "mini/shared/io-sidecar.js",
  ], {
    location: { href: "http://localhost/mini/index.html", protocol: "http:" },
    fetch: async () => { throw new Error("scaffold unavailable in test"); },
    document: {},
  });
  const state = { project: project || { meta: {}, chapters: [] }, selectedChapterId: "" };
  const runtime = {
    dirHandle: createMemoryDirectory(files),
    sidecarDoc: null,
    saveQueue: Promise.resolve(),
    writerId: "startup-test",
  };
  const io = context.AutoBerichtSidecar.init({
    state,
    runtime,
    elements: {},
    setStatus: (message) => statuses.push(message),
    debug: { logLine() {} },
    i18n: testI18n,
  }, {
    stateHelpers: {
      getLibraryFileName: (meta) => `library_user_${meta.locale}.json`,
      toText: (value) => value == null ? "" : String(value),
    },
    normalizeHelpers: {
      normalizeProject: normalizeProject || ((value) => structuredClone(value)),
      syncObservationChapterRows() {},
    },
    seeds: {
      normalizeTagGroups: (value) => value || { report: [], observations: [], training: [] },
      resolveLocaleKey: (locale) => String(locale).toLowerCase().startsWith("fr") ? "fr" : "de",
      getKnowledgeBaseFilename: (locale) => `knowledge_base_${locale}.json`,
      validateKnowledgeBase: (value) => value,
      readSeedFromProject: async () => null,
      readSeedFromHttp: async () => null,
      buildProjectFromKnowledgeBase: (library) => ({
        meta: { locale: library.meta.locale, libraryMarker: library.marker },
        chapters: [{ id: "0", rows: [] }],
      }),
      ...seedOverrides,
    },
    renderApi: {
      buildPhotoIndex() {},
      render() { renders += 1; },
    },
    spiderModule: {},
  });
  return { io, state, runtime, statuses, get renders() { return renders; } };
};

test("a valid sidecar opens Project even when scaffold setup fails", async () => {
  const sidecar = { report: { project: { meta: { locale: "de-CH" }, chapters: [{ id: "0", rows: [] }] } } };
  const fixture = createIo({ files: { "project_sidecar.json": JSON.stringify(sidecar) } });
  const result = await fixture.io.loadProjectFromFolder();
  assert.equal(result.ok, true);
  assert.equal(result.source, "sidecar");
  assert.equal(fixture.state.selectedChapterId, "__project__");
  assert.equal(fixture.renders > 0, true);
  assert.match(fixture.statuses.at(-1), /Loaded project_sidecar\.json/);
});

test("a wrapped sidecar can load then save even if normalizeProject mutates its input", async () => {
  const sidecar = {
    report: {
      project: {
        meta: { locale: "de-CH" },
        chapters: [{ id: "0", rows: [] }],
      },
    },
  };
  const fixture = createIo({
    files: { "project_sidecar.json": JSON.stringify(sidecar) },
    normalizeProject: (value) => {
      value.meta ||= {};
      value.meta.moderator = "Normalized";
      return value;
    },
  });
  const loaded = await fixture.io.loadProjectFromFolder();
  assert.equal(loaded.ok, true);
  await assert.doesNotReject(() => fixture.io.saveSidecar());
  const saved = JSON.parse(fixture.runtime.dirHandle.read("project_sidecar.json"));
  assert.equal(saved.report.project.meta.moderator, "Normalized");
});

test("language bootstrap selects the matching user library and saves immediately", async () => {
  const de = { meta: { locale: "de-CH" }, marker: "german" };
  const fr = { meta: { locale: "fr-CH" }, marker: "french" };
  const fixture = createIo({
    files: {
      "library_user_de-CH.json": JSON.stringify(de),
      "library_user_fr-CH.json": JSON.stringify(fr),
    },
  });
  await fixture.io.loadProjectFromFolder();
  await fixture.io.bootstrapProjectFromSeed("fr-CH", { deferSave: false });
  assert.equal(fixture.state.project.meta.locale, "fr-CH");
  assert.equal(fixture.state.project.meta.libraryMarker, "french");
  assert.equal(fixture.runtime.pendingBootstrapWrite, false);
  const saved = JSON.parse(fixture.runtime.dirHandle.read("project_sidecar.json"));
  assert.equal(saved.report.project.meta.libraryMarker, "french");
});

test("new-project bootstrap ingests a legacy user library without chapterFrontMatter migration", async () => {
  const fixture = createIo({
    files: {
      "library_user_OLD_de-CH.json": JSON.stringify({
        schemaVersion: "1.1",
        meta: { locale: "de-CH" },
        structure: { items: [{ id: "0.1", collapsedId: "0.1", chapterLabel: "Summary", question: "Lead" }] },
        library: { entries: [{ id: "0.1", finding: "", recommendation: "Base" }] },
        tags: { report: [], observations: [], training: [] },
      }),
    },
    normalizeProject: (value) => {
      value.chapters.forEach((chapter) => {
        chapter.meta ||= {};
        if (chapter.meta.frontMatterText == null) chapter.meta.frontMatterText = "";
        if (chapter.meta.frontMatterLibraryAction == null) chapter.meta.frontMatterLibraryAction = "off";
        if (chapter.meta.frontMatterLibraryHash == null) chapter.meta.frontMatterLibraryHash = "";
      });
      return value;
    },
    seedOverrides: {
      normalizeKnowledgeBaseCompatibility: (library) => {
        const clone = structuredClone(library);
        clone.library = clone.library || { entries: [] };
        clone.library.chapterFrontMatter = {};
        return clone;
      },
      buildProjectFromKnowledgeBase: (library) => ({
        meta: { locale: library.meta.locale },
        chapters: [{
          id: "0",
          rows: [],
          meta: {
            frontMatterText: library.library.chapterFrontMatter?.["0"] || "",
          },
        }],
      }),
    },
  });
  const opened = await fixture.io.loadProjectFromFolder();
  assert.equal(opened.source, "empty");
  await fixture.io.bootstrapProjectFromSeed("de-CH", { deferSave: false });
  const chapter0 = fixture.state.project.chapters.find((chapter) => chapter.id === "0");
  assert.equal(chapter0.meta.frontMatterText, "");
  assert.equal(chapter0.meta.frontMatterLibraryAction, "off");
  assert.equal(chapter0.meta.frontMatterLibraryHash, "");
  const saved = JSON.parse(fixture.runtime.dirHandle.read("project_sidecar.json"));
  assert.equal(saved.report.project.chapters[0].meta.frontMatterText, "");
});

test("new-project bootstrap copies current Chapter 0 customer context from the matching user library", async () => {
  const fixture = createIo({
    files: {
      "library_user_OLD_de-CH.json": JSON.stringify({
        schemaVersion: "1.1",
        meta: { locale: "de-CH" },
        structure: { items: [{ id: "0.1", collapsedId: "0.1", chapterLabel: "Summary", question: "Lead" }] },
        library: {
          entries: [{ id: "0.1", finding: "", recommendation: "Base" }],
          chapterFrontMatter: { "0": "Customer context" },
        },
        tags: { report: [], observations: [], training: [] },
      }),
    },
    normalizeProject: (value) => {
      value.chapters.forEach((chapter) => {
        chapter.meta ||= {};
        if (chapter.meta.frontMatterText == null) chapter.meta.frontMatterText = "";
        if (chapter.meta.frontMatterLibraryAction == null) chapter.meta.frontMatterLibraryAction = "off";
        if (chapter.meta.frontMatterLibraryHash == null) chapter.meta.frontMatterLibraryHash = "";
      });
      return value;
    },
    seedOverrides: {
      normalizeKnowledgeBaseCompatibility: (library) => structuredClone(library),
      buildProjectFromKnowledgeBase: (library) => ({
        meta: { locale: library.meta.locale },
        chapters: [{
          id: "0",
          rows: [],
          meta: {
            frontMatterText: library.library.chapterFrontMatter?.["0"] || "",
          },
        }],
      }),
    },
  });
  await fixture.io.loadProjectFromFolder();
  await fixture.io.bootstrapProjectFromSeed("de-CH", { deferSave: false });
  const chapter0 = fixture.state.project.chapters.find((chapter) => chapter.id === "0");
  assert.equal(chapter0.meta.frontMatterText, "Customer context");
  const saved = JSON.parse(fixture.runtime.dirHandle.read("project_sidecar.json"));
  assert.equal(saved.report.project.chapters[0].meta.frontMatterText, "Customer context");
});

test("Chapter 0 customer context is stored in the user library without duplicate appends", async () => {
  const libraryName = "library_user_fr-CH.json";
  const initialLibrary = {
    meta: { locale: "fr-CH" },
    structure: { items: [] },
    library: {
      entries: [],
      observations: [],
      chapterPositives: {},
      chapterFrontMatter: { 0: "Contexte existant." },
    },
    tags: { report: [], observations: [], training: [] },
  };
  const context = loadBrowserScripts([
    "mini/shared/sidecar-storage.js",
    "mini/shared/io-sidecar.js",
  ], {
    location: { href: "http://localhost/mini/index.html", protocol: "http:" },
    document: {},
  });
  const project = {
    meta: { locale: "fr-CH", moderator: "" },
    chapters: [{
      id: "0",
      rows: [],
      meta: {
        frontMatterText: "Contexte propre à ce client.",
        frontMatterLibraryAction: "append",
        frontMatterLibraryHash: "",
      },
    }],
  };
  const runtime = {
    dirHandle: createMemoryDirectory({ [libraryName]: JSON.stringify(initialLibrary) }),
    sidecarDoc: null,
    saveQueue: Promise.resolve(),
    writerId: "front-matter-test",
    changeVersion: 0,
  };
  const io = context.AutoBerichtSidecar.init({
    state: { project, selectedChapterId: "0" },
    runtime,
    elements: {},
    setStatus() {},
    debug: { logLine() {} },
    i18n: testI18n,
  }, {
    stateHelpers: {
      getLibraryFileName: () => libraryName,
      toText: (value) => value == null ? "" : String(value),
      hashText: (value) => `hash:${String(value)}`,
      getFindingText: () => "",
      getRecommendationText: () => "",
    },
    normalizeHelpers: {
      ensureProjectMeta() {},
      ensureChapterMetaDefaults(chapter) { chapter.meta ||= {}; },
      ensureWorkstateDefaults() {},
    },
    seeds: {
      getKnowledgeBaseFilename: () => "knowledge_base_fr.json",
      validateKnowledgeBase() {},
      readSeedFromProject: async () => null,
      readSeedFromHttp: async () => null,
      normalizeTagGroups: (value) => value || { report: [], observations: [], training: [] },
    },
    renderApi: {},
    spiderModule: {},
  });

  await io.generateLibrary();
  let savedLibrary = JSON.parse(runtime.dirHandle.read(libraryName));
  assert.equal(
    savedLibrary.library.chapterFrontMatter[0],
    "Contexte existant.\n\nContexte propre à ce client.",
  );

  project.chapters[0].meta.frontMatterLibraryHash = "";
  await io.generateLibrary();
  savedLibrary = JSON.parse(runtime.dirHandle.read(libraryName));
  assert.equal(
    savedLibrary.library.chapterFrontMatter[0],
    "Contexte existant.\n\nContexte propre à ce client.",
  );
});
