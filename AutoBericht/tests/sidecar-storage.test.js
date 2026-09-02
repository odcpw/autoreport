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
});
