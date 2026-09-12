const test = require("node:test");
const assert = require("node:assert/strict");
const { loadBrowserScripts, createMemoryDirectory } = require("./helpers");

const compareIdSegments = (left, right) => String(left).localeCompare(String(right), undefined, { numeric: true });
const toText = (value) => {
  if (Array.isArray(value)) return value.join("\n");
  if (value == null) return "";
  return String(value);
};

const createLibraryMakerApi = () => {
  const hooks = {};
  loadBrowserScripts([
    "mini/shared/dependencies.js",
    "mini/librarymaker.js",
  ], {
    __AUTO_BERICHT_TEST__: hooks,
    document: {
      getElementById() { return null; },
      addEventListener() {},
      visibilityState: "visible",
    },
    addEventListener() {},
    showDirectoryPicker: async () => { throw new Error("unused in test"); },
    AutoReportDebug: { logLine() {} },
    AutoBerichtI18n: {
      t: (key, fallback) => fallback || key,
      tf: (key, fallback, values = {}) => String(fallback || key).replace(/\{(\w+)\}/g, (match, name) => String(values[name] ?? match)),
      setLocale() {},
      resolveSpellcheckLang: (locale) => locale || "de-CH",
    },
    AutoBerichtFsHandle: {
      saveHandle: async () => {},
      loadHandle: async () => null,
      requestHandlePermission: async () => false,
    },
    AutoBerichtState: {
      compareIdSegments,
      formatChapterLabel: (chapter) => `Chapter ${chapter.id}`,
      getLibraryFileName: (meta) => `library_user_${meta.locale}.json`,
      toText,
    },
    AutoBerichtNormalize: {
      normalizeProject: (value) => structuredClone(value),
    },
    AutoBerichtSeeds: {
      normalizeTagGroups: (value) => value || { report: [], observations: [], training: [] },
      validateKnowledgeBase(value) {
        if (!value || typeof value !== "object") throw new Error("Knowledge base missing or invalid.");
        if (!value.schemaVersion) throw new Error("Knowledge base missing schemaVersion.");
        if (!value.structure || !Array.isArray(value.structure.items)) throw new Error("Knowledge base structure is missing items.");
        if (!value.library || !Array.isArray(value.library.entries)) throw new Error("Knowledge base library entries missing.");
        if (!value.tags || typeof value !== "object") throw new Error("Knowledge base tags missing.");
      },
    },
  });
  return hooks.librarymaker;
};

const createLibrary = (overrides = {}) => ({
  schemaVersion: "1.1",
  meta: { locale: "de-CH", moderatorInitials: "AB" },
  structure: {
    items: [
      { id: "0.1", collapsedId: "0.1", chapterLabel: "Summary", sectionLabel: "0.1 Summary", question: "Lead" },
      { id: "1.1", collapsedId: "1.1", chapterLabel: "Chapter 1", sectionLabel: "1.1 Topic", question: "Item" },
    ],
  },
  library: {
    entries: [{ id: "0.1", finding: "", recommendation: "Base summary" }],
    observations: [],
    chapterPositives: { "1": "Keep this positive." },
  },
  tags: { report: [], observations: [], training: [] },
  customField: { preserved: true },
  ...overrides,
});

test("legacy library without chapterFrontMatter loads as an empty map", () => {
  const api = createLibraryMakerApi();
  const normalized = api.normalizeKnowledgeBaseForMaker(createLibrary());
  assert.deepEqual(JSON.parse(JSON.stringify(normalized.library.chapterFrontMatter)), {});
});

test("Chapter 0 customer context supports Add and Replace from Other Library", () => {
  const api = createLibraryMakerApi();
  const targetLibrary = api.normalizeKnowledgeBaseForMaker(createLibrary({
    library: {
      entries: [{ id: "0.1", finding: "", recommendation: "Base summary" }],
      observations: [],
      chapterPositives: {},
      chapterFrontMatter: { "0": { text: "My context" } },
    },
  }));
  const sourceLibrary = api.normalizeKnowledgeBaseForMaker(createLibrary({
    library: {
      entries: [{ id: "0.1", finding: "", recommendation: "Base summary" }],
      observations: [],
      chapterPositives: {},
      chapterFrontMatter: { "0": { value: "Other context" } },
    },
  }));

  api.setState({ targetLibrary, sourceLibrary, selectedChapterId: "0" });
  const chapters = Array.from(api.getChapterObjects(), (chapter) => chapter.id);
  assert.deepEqual(chapters, ["0", "1"]);

  const row = api.buildRowsForSelection()[0];
  assert.equal(row.kind, "chapterFrontMatter");

  api.applyButtonAction(row, "text", "append");
  assert.equal(api.getState().targetLibrary.library.chapterFrontMatter["0"], "My context\n\nOther context");

  api.setState({
    targetLibrary: api.normalizeKnowledgeBaseForMaker(createLibrary({
      library: {
        entries: [{ id: "0.1", finding: "", recommendation: "Base summary" }],
        observations: [],
        chapterPositives: {},
        chapterFrontMatter: { "0": { text: "My context" } },
      },
    })),
    sourceLibrary,
    selectedChapterId: "0",
  });
  const replaceRow = api.buildRowsForSelection()[0];
  api.applyButtonAction(replaceRow, "text", "replace");
  assert.equal(api.getState().targetLibrary.library.chapterFrontMatter["0"], "Other context");
});

test("saving My Library upgrades chapterFrontMatter and preserves it on reload", async () => {
  const projectSidecar = {
    report: {
      project: {
        meta: { locale: "de-CH", moderator: "Alice", moderatorInitials: "AL" },
        chapters: [{ id: "0", rows: [] }],
      },
    },
  };
  const libraryFile = createLibrary({
    library: {
      entries: [{ id: "0.1", finding: "", recommendation: "Base summary" }],
      observations: [],
      chapterPositives: { "1": "Keep this positive." },
      chapterFrontMatter: { "0": { text: "Legacy context" } },
    },
  });
  const dir = createMemoryDirectory({
    "project_sidecar.json": JSON.stringify(projectSidecar),
    "library_user_de-CH.json": JSON.stringify(libraryFile),
  });

  const api = createLibraryMakerApi();
  api.setState({ projectHandle: dir });
  await api.loadTargetLibraryFromProject();
  let state = api.getState();
  assert.equal(state.targetLibrary.library.chapterFrontMatter["0"], "Legacy context");

  const row = api.buildRowsForSelection()[0];
  api.setTargetFieldValue(row, "text", "Updated context");
  await api.saveTargetLibrary();

  const saved = JSON.parse(dir.read("library_user_de-CH.json"));
  assert.equal(saved.library.chapterFrontMatter["0"], "Updated context");
  assert.equal(saved.library.chapterPositives["1"], "Keep this positive.");
  assert.equal(saved.customField.preserved, true);

  const reloadApi = createLibraryMakerApi();
  reloadApi.setState({ projectHandle: dir });
  await reloadApi.loadTargetLibraryFromProject();
  state = reloadApi.getState();
  assert.equal(state.targetLibrary.library.chapterFrontMatter["0"], "Updated context");
});
