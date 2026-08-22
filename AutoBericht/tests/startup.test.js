const test = require("node:test");
const assert = require("node:assert/strict");
const { loadBrowserScripts, createMemoryDirectory } = require("./helpers");

const createIo = ({ files, project }) => {
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
    i18n: { setLocale() {} },
  }, {
    stateHelpers: {
      getLibraryFileName: (meta) => `library_user_${meta.locale}.json`,
      toText: (value) => value == null ? "" : String(value),
    },
    normalizeHelpers: {
      normalizeProject: (value) => structuredClone(value),
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
    i18n: { setLocale() {} },
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
