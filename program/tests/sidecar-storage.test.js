const test = require("node:test");
const assert = require("node:assert/strict");
const { loadBrowserScripts, createMemoryDirectory } = require("./helpers");

const loadStorage = () => loadBrowserScripts(["mini/shared/sidecar-storage.js"]).AutoBerichtSidecarStorage;

const testI18n = {
  setLocale() {},
  t: (key, fallback) => fallback || key,
  tf: (key, fallback, values = {}) => String(fallback || key).replace(/\{(\w+)\}/g, (match, name) => String(values[name] ?? match)),
};

test("a failed sidecar write rejects instead of reporting success", async () => {
  const storage = loadStorage();
  const base = { report: { value: "old" }, photos: { value: "old" } };
  const dirHandle = createMemoryDirectory({ "project_sidecar.json": JSON.stringify(base) }, { failWrite: true });
  await assert.rejects(
    storage.saveSidecar({
      dirHandle,
      merge: (latest) => ({ ...latest, report: { value: "new" } }),
    }),
    (error) => error.name === "SidecarWriteError" && /disk full/.test(error.message),
  );
  assert.equal(JSON.parse(dirHandle.read("project_sidecar.json")).report.value, "old");
});

test("a corrupted sidecar is never treated as a missing project", async () => {
  const storage = loadStorage();
  const dirHandle = createMemoryDirectory({ "project_sidecar.json": "{broken" });
  await assert.rejects(
    storage.readSidecar(dirHandle, { allowMissing: true }),
    (error) => error.name === "SidecarParseError" && /corrupted or invalid JSON/.test(error.message),
  );
});

test("the first save creates the sidecar from an empty folder", async () => {
  const storage = loadStorage();
  const dirHandle = createMemoryDirectory({});
  const saved = await storage.saveSidecar({
    dirHandle,
    merge: (latest) => ({ ...(latest || {}), report: { value: "r0" } }),
  });
  assert.equal(saved.report.value, "r0");
  assert.equal(typeof saved.meta.updatedAt, "string");
  assert.equal(JSON.parse(dirHandle.read("project_sidecar.json")).report.value, "r0");
});

test("PhotoSorter's first save retains its library report and includes tagged observations", async () => {
  const browser = loadBrowserScripts([
    "mini/shared/sidecar-storage.js", "mini/shared/normalize.js",
    "mini/photosorter/tags.js", "mini/photosorter/io-sidecar.js",
  ], { document: { documentElement: { getAttribute: () => "fr-CH" } } });
  const dir = createMemoryDirectory({});
  const state = {
    projectHandle: dir,
    sidecarDoc: { report: { project: { meta: { locale: "fr-CH" }, chapters: [{ id: "4.8", rows: [] }] } } },
    projectDoc: {}, photoRootName: "photos/resized",
    tagOptions: { report: [], observations: [{ value: "Escaliers", label: "Escaliers" }], training: [] },
  };
  const io = browser.AutoBerichtPhotoSorterSidecar.init(
    { state, runtime: { saveQueue: Promise.resolve() }, setStatus() {}, debug: { logLine() {} }, elements: {} },
    { tagsApi: browser.AutoBerichtPhotoSorterTags, i18n: testI18n,
      photosApi: { serializePhotos: () => ({ "photos/resized/a.jpg": { tags: { observations: ["Escaliers"] } } }) } },
  );
  await io.saveProjectSidecar();
  const first = JSON.parse(dir.read("project_sidecar.json"));
  assert.equal(first.report.project.meta.locale, "fr-CH");
  assert.equal(first.report.project.chapters[0].rows[0].workstate.includeFinding, true);
  assert.equal(first.report.project.chapters[0].rows[0].workstate.done, false);
  // Another writer's newer report must win over PhotoSorter's cached report.
  first.report.project.meta.company = "Updated in report tab";
  await browser.AutoBerichtSidecarStorage.saveSidecar({ dirHandle: dir, merge: () => first });
  await io.saveProjectSidecar();
  assert.equal(JSON.parse(dir.read("project_sidecar.json")).report.project.meta.company, "Updated in report tab");
});

test("report and photo writers preserve each other's branches", async () => {
  const storage = loadStorage();
  const base = { report: { value: "r0" }, photos: { value: "p0" } };
  const dirHandle = createMemoryDirectory({ "project_sidecar.json": JSON.stringify(base) });
  await storage.saveSidecar({
    dirHandle,
    merge: (latest) => ({ ...latest, report: { value: "r1" } }),
  });
  const result = await storage.saveSidecar({
    dirHandle,
    merge: (latest) => ({ ...latest, photos: { value: "p1" } }),
  });
  assert.equal(result.report.value, "r1");
  assert.equal(result.photos.value, "p1");
});

// Regression for the September 2026 field failure: PhotoSorter and AutoBericht
// open in two tabs. AutoBericht saves (an edit, or the 30-minute auto-backup),
// then PhotoSorter must still be able to save its tags without a reload.
test("PhotoSorter keeps saving after AutoBericht saved in another tab", async () => {
  const initial = {
    report: {
      project: {
        meta: { locale: "de-CH", company: "ACME" },
        chapters: [{ id: "0", rows: [] }, { id: "4.8", rows: [] }],
      },
    },
    photos: {
      meta: { projectId: "", createdAt: "2026-08-01T00:00:00.000Z" },
      photos: { "photos/resized/pm1_0001.jpg": { notes: "", tags: { report: ["1.1"], observations: [], training: [] } } },
      photoTagOptions: {
        report: [{ value: "1.1", label: "1.1 Leitbild" }],
        observations: [{ value: "Ordnung", label: "Ordnung" }],
        training: [{ value: "Basics", label: "Basics" }],
      },
      photoRoot: "photos/resized",
    },
  };
  const dir = createMemoryDirectory({ "project_sidecar.json": JSON.stringify(initial, null, 2) });

  // PhotoSorter tab.
  const psContext = loadBrowserScripts([
    "mini/shared/sidecar-storage.js",
    "mini/shared/normalize.js",
    "mini/photosorter/tags.js",
    "mini/photosorter/io-sidecar.js",
  ], {
    document: { documentElement: { getAttribute: () => "de-CH" } },
    location: { href: "http://localhost/mini/photosorter.html" },
    fetch: async () => { throw new Error("no http in test"); },
  });
  const psState = { projectHandle: dir, photos: [], tagOptions: null, projectDoc: null, sidecarDoc: null, photoRootName: "" };
  const psRuntime = { saveQueue: Promise.resolve(), hasUnsavedChanges: false, autosaveTimer: null };
  const tagsApi = psContext.AutoBerichtPhotoSorterTags;
  const psIo = psContext.AutoBerichtPhotoSorterSidecar.init(
    { state: psState, runtime: psRuntime, setStatus() {}, debug: { logLine() {} }, elements: {} },
    {
      tagsApi,
      photosApi: {
        serializePhotos: () => Object.fromEntries(psState.photos.map((photo) => [photo.path, { notes: photo.notes || "", tags: photo.tags }])),
        maybeAutoScan: async () => false,
      },
      i18n: testI18n,
    },
  );
  const parsed = await psContext.AutoBerichtSidecarStorage.readSidecar(dir, { allowMissing: true });
  psState.sidecarDoc = parsed;
  psState.projectDoc = psIo.normalizePhotoDoc(parsed.photos);
  psState.photoRootName = psState.projectDoc.photoRoot;
  psState.tagOptions = tagsApi.ensureTagOptions(psState.projectDoc.photoTagOptions);
  psState.photos = Object.entries(psState.projectDoc.photos).map(([path, value]) => ({ path, notes: value.notes, tags: structuredClone(value.tags) }));

  // AutoBericht tab.
  const abContext = loadBrowserScripts([
    "mini/shared/sidecar-storage.js",
    "mini/shared/state.js",
    "mini/shared/normalize.js",
    "mini/shared/seeds.js",
    "mini/shared/io-sidecar.js",
  ], {
    location: { href: "http://localhost/mini/index.html", protocol: "http:" },
    fetch: async () => { throw new Error("no http in test"); },
    document: {},
  });
  const abState = { project: { meta: {}, chapters: [] }, selectedChapterId: "", spiderOverrides: {} };
  const abRuntime = { dirHandle: dir, sidecarDoc: null, saveQueue: Promise.resolve() };
  const abIo = abContext.AutoBerichtSidecar.init(
    { state: abState, runtime: abRuntime, elements: {}, setStatus() {}, debug: { logLine() {} }, i18n: testI18n },
    {
      stateHelpers: abContext.AutoBerichtState,
      normalizeHelpers: abContext.AutoBerichtNormalize,
      seeds: abContext.AutoBerichtSeeds,
      renderApi: { buildPhotoIndex() {}, render() {} },
      spiderModule: {},
    },
  );

  psState.photos[0].tags.observations.push("Ordnung");
  await psIo.saveProjectSidecar();
  const taggedRow = JSON.parse(dir.read("project_sidecar.json")).report.project.chapters.find((c) => c.id === "4.8").rows[0];
  assert.equal(taggedRow.workstate.includeFinding, true);
  assert.equal(taggedRow.workstate.includeRecommendation, true);
  assert.equal(taggedRow.workstate.done, false);

  const loaded = await abIo.loadProjectFromFolder();
  assert.equal(loaded.ok, true);
  abState.project.meta.company = "ACME edited";
  await abIo.saveSidecar();

  psState.photos[0].tags.training.push("Basics");
  await assert.doesNotReject(() => psIo.saveProjectSidecar());

  abState.project.meta.company = "ACME edited twice";
  await assert.doesNotReject(() => abIo.saveSidecar());

  const onDisk = JSON.parse(dir.read("project_sidecar.json"));
  const photo = onDisk.photos.photos["photos/resized/pm1_0001.jpg"];
  assert.deepEqual(photo.tags.observations, ["Ordnung"]);
  assert.deepEqual(photo.tags.training, ["Basics"]);
  assert.equal(onDisk.report.project.meta.company, "ACME edited twice");
  assert.equal(onDisk.photos.photoRoot, "photos/resized");
  assert.equal(onDisk.report.project.chapters.find((c) => c.id === "4.8").rows[0].workstate.includeFinding, true);
  const localRow = abState.project.chapters.find((c) => c.id === "4.8").rows[0];
  localRow.workstate.includeFinding = false;
  localRow.workstate.done = true;
  await abIo.saveSidecar();
  psState.photos[0].tags.observations = [];
  await psIo.saveProjectSidecar();
  psState.photos.push({ path: "photos/resized/new.jpg", tags: { observations: ["Ordnung"], report: [], training: [] } });
  await psIo.saveProjectSidecar();
  // The report tab still has its old unchecked row and must reconcile the
  // newly tagged photo. A second save must retain that activated draft.
  await abIo.saveSidecar();
  await abIo.saveSidecar();
  const updatedRow = JSON.parse(dir.read("project_sidecar.json")).report.project.chapters.find((c) => c.id === "4.8").rows[0];
  assert.equal(updatedRow.workstate.includeFinding, true);
  assert.equal(updatedRow.workstate.done, false);
});

test("new observation assignments include drafts, but unchanged tags respect manual exclusion", () => {
  const ctx = loadBrowserScripts(["mini/shared/normalize.js"]);
  const row = { type: "field_observation", tag: "Regale", titleOverride: "Rayonnages", workstate: { includeFinding: false, includeRecommendation: false, done: true, findingText: "My draft" } };
  const project = { chapters: [{ id: "4.8", rows: [row] }] };
  const before = { photos: { photos: { "a.jpg": { tags: { observations: [] } } } } };
  const after = { photos: { photoTagOptions: { observations: [{ value: "Regale", label: "Rayonnages" }] }, photos: { "a.jpg": { tags: { observations: ["Regale"] } } } } };
  ctx.AutoBerichtNormalize.includeNewObservationAssignments(project, after, before);
  assert.equal(row.workstate.includeFinding, true);
  assert.equal(row.workstate.includeRecommendation, true);
  assert.equal(row.workstate.done, false);
  assert.equal(row.workstate.findingText, "My draft");
  row.workstate.includeFinding = false;
  row.workstate.done = true;
  ctx.AutoBerichtNormalize.includeNewObservationAssignments(project, after, after);
  assert.equal(row.workstate.includeFinding, false);
  assert.equal(row.workstate.done, true);
  ctx.AutoBerichtNormalize.includeNewObservationAssignments(project, before, after);
  assert.equal(row.workstate.includeFinding, false);
});

// Sidecars written before 24 February 2026 kept the project and the photo map
// at the root. Both apps must read them and lift them into the wrapped layout.
test("a flat pre-February sidecar keeps its photo tags in both apps", async () => {
  const tags = {
    report: [{ value: "1.1", label: "1.1 Leitbild" }],
    observations: [{ value: "Ordnung", label: "Ordnung" }],
    training: [],
  };
  const flat = {
    meta: { locale: "de-CH", company: "Legacy AG" },
    chapters: [{ id: "0", rows: [] }, { id: "1", rows: [] }],
    photoRoot: "photos/resized",
    photoTagOptions: tags,
    photos: { "photos/resized/a.jpg": { notes: "", tags: { report: ["1.1"], observations: ["Ordnung"], training: [] } } },
  };
  const notFound = (name) => { const error = new Error(`${name} not found`); error.name = "NotFoundError"; return error; };
  const dir = createMemoryDirectory({ "project_sidecar.json": JSON.stringify(flat) });
  dir.getDirectoryHandle = async (name) => { throw notFound(name); };

  const bootPhotoSorter = () => {
    const context = loadBrowserScripts([
      "mini/shared/sidecar-storage.js",
      "mini/shared/state.js",
      "mini/shared/normalize.js",
      "mini/shared/seeds.js",
      "mini/photosorter/tags.js",
      "mini/photosorter/io-sidecar.js",
    ], {
      document: { documentElement: { getAttribute: () => "de-CH" } },
      location: { href: "http://localhost/mini/photosorter.html" },
      fetch: async () => { throw new Error("no http in test"); },
    });
    const state = { projectHandle: dir, photos: [], tagOptions: null, projectDoc: null, sidecarDoc: null, photoRootName: "", activeTagFilters: {}, filterMode: "all", photoHandle: null };
    const io = context.AutoBerichtPhotoSorterSidecar.init(
      { state, runtime: { saveQueue: Promise.resolve() }, setStatus() {}, debug: { logLine() {} }, elements: {} },
      {
        renderApi: { renderAll() {}, renderPanels() {} },
        tagsApi: context.AutoBerichtPhotoSorterTags,
        photosApi: { serializePhotos: () => Object.fromEntries(state.photos.map((photo) => [photo.path, { notes: "", tags: photo.tags }])), maybeAutoScan: async () => false },
        i18n: testI18n,
      },
    );
    return { io, state };
  };

  const first = bootPhotoSorter();
  await first.io.loadProjectSidecar();
  assert.deepEqual(Object.keys(first.state.projectDoc.photos), ["photos/resized/a.jpg"]);
  assert.equal(first.state.tagOptions.observations.some((option) => option.value === "Ordnung"), true);
  assert.equal(first.state.photoRootName, "photos/resized");

  const abContext = loadBrowserScripts([
    "mini/shared/sidecar-storage.js",
    "mini/shared/state.js",
    "mini/shared/normalize.js",
    "mini/shared/seeds.js",
    "mini/shared/io-sidecar.js",
  ], {
    location: { href: "http://localhost/mini/index.html", protocol: "http:" },
    fetch: async () => { throw new Error("no http in test"); },
    document: {},
  });
  const abState = { project: { meta: {}, chapters: [] }, selectedChapterId: "", spiderOverrides: {} };
  const abIo = abContext.AutoBerichtSidecar.init(
    { state: abState, runtime: { dirHandle: dir, sidecarDoc: null, saveQueue: Promise.resolve() }, elements: {}, setStatus() {}, debug: { logLine() {} }, i18n: testI18n },
    {
      stateHelpers: abContext.AutoBerichtState,
      normalizeHelpers: abContext.AutoBerichtNormalize,
      seeds: abContext.AutoBerichtSeeds,
      renderApi: { buildPhotoIndex() {}, render() {} },
      spiderModule: {},
    },
  );
  const loaded = await abIo.loadProjectFromFolder();
  assert.equal(loaded.source, "sidecar");
  assert.equal(abState.project.meta.company, "Legacy AG");
  await abIo.saveSidecar();

  const saved = JSON.parse(dir.read("project_sidecar.json"));
  assert.equal(saved.report.project.meta.company, "Legacy AG");
  assert.equal("chapters" in saved, false);
  assert.equal("photoTagOptions" in saved, false);
  assert.deepEqual(Object.keys(saved.photos.photos), ["photos/resized/a.jpg"]);
  assert.equal(saved.photos.photoRoot, "photos/resized");

  const second = bootPhotoSorter();
  await second.io.loadProjectSidecar();
  assert.deepEqual(Object.keys(second.state.projectDoc.photos), ["photos/resized/a.jpg"]);
  assert.equal(second.state.tagOptions.observations.some((option) => option.value === "Ordnung"), true);
});
