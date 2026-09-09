const test = require("node:test");
const assert = require("node:assert/strict");
const { loadBrowserScripts, createMemoryDirectory } = require("./helpers");

const tags = () => ({ report: [], observations: [], training: [] });
const photo = (extra = {}) => ({ notes: "", tags: tags(), ...extra });
const i18n = {
  setLocale() {},
  t: (key, fallback) => fallback || key,
  tf: (key, fallback, values = {}) => String(fallback || key).replace(/\{(\w+)\}/g, (match, name) => String(values[name] ?? match)),
};
const project = (records, lastPhotoNumber) => {
  const original = {
    report: { project: { meta: { locale: "fr-CH" }, chapters: [] } },
    photos: {
      meta: { ...(lastPhotoNumber === undefined ? {} : { lastPhotoNumber }) },
      photos: Object.fromEntries(Object.entries(records).map(([name, record]) => [`photos/resized/${name}`, record])),
      photoRoot: "photos/resized",
      photoTagOptions: {
        report: [{ value: "1.3", label: "Leadership" }],
        observations: [{ value: "Escaliers", label: "Escaliers" }],
        training: [{ value: "Example", label: "Example" }],
      },
    },
  };
  const dir = createMemoryDirectory({ "project_sidecar.json": JSON.stringify(original) });
  const files = new Set(Object.keys(records));
  const imageDir = {
    async *values() {
      // Deliberately unsorted: migration must follow the UI sort, not disk order.
      for (const name of [...files].reverse()) yield { kind: "file", name };
    },
  };
  dir.getDirectoryHandle = async (name) => {
    assert.equal(name, "photos");
    return { getDirectoryHandle: async (child) => { assert.equal(child, "resized"); return imageDir; } };
  };
  return { dir, files, original, read: () => JSON.parse(dir.read("project_sidecar.json")) };
};
const open = async (fixture) => {
  const browser = loadBrowserScripts([
    "mini/shared/sidecar-storage.js", "mini/photosorter/state.js",
    "mini/photosorter/tags.js", "mini/photosorter/photos.js",
    "mini/photosorter/io-sidecar.js", "mini/photosorter/render.js",
  ], { document: { documentElement: { getAttribute: () => "fr" } } });
  const state = browser.AutoBerichtPhotoSorterState.createState();
  const runtime = browser.AutoBerichtPhotoSorterState.createRuntime();
  state.projectHandle = fixture.dir;
  const ctx = { state, runtime, elements: {}, i18n, setStatus() {}, debug: { logLine() {} } };
  const tagsApi = browser.AutoBerichtPhotoSorterTags;
  let io;
  const photos = browser.AutoBerichtPhotoSorterPhotos.init(ctx, { tagsApi, notifyChange: () => io.scheduleAutosave() });
  io = browser.AutoBerichtPhotoSorterSidecar.init(ctx, { tagsApi, photosApi: photos, i18n, renderApi: { renderAll() {} } });
  await io.loadProjectSidecar();
  return { browser, ctx, state, runtime, photos, io };
};
const numbers = (session) => Object.fromEntries(session.state.photos.map(p => [p.path.split("/").pop(), p.photoNumber]));

test("photo numbers survive migration, save/reopen, filters, additions and removal of the highest number", async () => {
  const fixture = project({ "b.jpg": photo({ notes: "Keep this note", custom: { keep: true } }), "c.jpg": photo() });
  const first = await open(fixture);
  assert.deepEqual(numbers(first), { "b.jpg": 1, "c.jpg": 2 });
  assert.equal(first.runtime.hasUnsavedChanges, true, "number assignment schedules a save without tagging");
  await first.io.flushAutosave();
  assert.equal(fixture.read().photos.photos["photos/resized/b.jpg"].photoNumber, 1);
  assert.deepEqual(fixture.read().photos.photos["photos/resized/b.jpg"].custom, { keep: true });
  assert.deepEqual(fixture.read().report, fixture.original.report);

  const second = await open(fixture);
  assert.deepEqual(numbers(second), { "b.jpg": 1, "c.jpg": 2 });
  second.state.photos[0].tags.observations.push("Escaliers");
  second.state.photos[0].notes = "Edited before rescan";
  second.state.filterMode = "unsorted";
  assert.equal(second.photos.getCurrentPhoto().photoNumber, 2);
  second.state.filterMode = "all";
  second.state.activeTagFilters.observations = ["Escaliers"];
  assert.equal(second.photos.getCurrentPhoto().photoNumber, 1);
  await second.photos.scanPhotos();
  assert.equal(second.state.photos[0].notes, "Edited before rescan");
  assert.equal(second.state.photos[0].tags.observations[0], "Escaliers");

  fixture.files.delete("c.jpg");
  await second.photos.scanPhotos();
  await second.io.flushAutosave();
  assert.equal(fixture.read().photos.meta.lastPhotoNumber, 2);
  fixture.files.add("a.jpg"); // sorts before b; must NOT take number 1 or retired 2
  const third = await open(fixture);
  assert.deepEqual(numbers(third), { "a.jpg": 3, "b.jpg": 1 });
  await third.io.flushAutosave();

  fixture.files.clear();
  await third.photos.scanPhotos();
  await third.io.flushAutosave();
  assert.equal(fixture.read().photos.meta.lastPhotoNumber, 3);
  fixture.files.add("d.jpg");
  const fourth = await open(fixture);
  assert.deepEqual(numbers(fourth), { "d.jpg": 4 });
  await fourth.io.flushAutosave();

  // The same filename in another project has that project's number and content.
  const other = project({ "d.jpg": photo({ photoNumber: 38, notes: "Another project" }) }, 38);
  fourth.state.projectHandle = other.dir;
  await fourth.io.loadProjectSidecar();
  assert.deepEqual(numbers(fourth), { "d.jpg": 38 });
  assert.equal(fourth.state.photos[0].notes, "Another project");
  await fourth.io.flushAutosave();
});

test("duplicate saved photo numbers reject the scan without silently changing identities", async () => {
  const fixture = project({ "a.jpg": photo({ photoNumber: 7 }), "b.jpg": photo({ photoNumber: 7 }) }, 7);
  await assert.rejects(open(fixture), /Photo 7 is assigned to both/);
  assert.deepEqual(fixture.read(), fixture.original);
});

test("a queued save remains bound to the project whose photos it captured", async () => {
  const first = project({ "a.jpg": photo() });
  const second = project({ "b.jpg": photo({ photoNumber: 38 }) }, 38);
  const session = await open(first);
  const pending = session.io.flushAutosave();
  session.state.projectHandle = second.dir;
  const secondDoc = session.io.normalizePhotoDoc(second.original.photos);
  session.state.projectDoc = secondDoc;
  await pending;
  assert.equal(first.read().photos.photos["photos/resized/a.jpg"].photoNumber, 1);
  assert.deepEqual(second.read(), second.original);
  assert.equal(session.state.projectDoc, secondDoc);
});

test("an older queued save cannot release numbers allocated during a rescan", async () => {
  const fixture = project({ "a.jpg": photo() });
  const session = await open(fixture);
  await session.io.flushAutosave();
  let release;
  session.runtime.saveQueue = new Promise((resolve) => { release = resolve; });
  const pending = session.io.saveProjectSidecar(); // captures lastPhotoNumber = 1
  fixture.files.add("b.jpg");
  await session.photos.scanPhotos(); // allocates 2
  fixture.files.delete("b.jpg");
  await session.photos.scanPhotos(); // 2 is retired while the old write is queued
  release();
  await pending;
  await session.io.flushAutosave();
  assert.equal(fixture.read().photos.meta.lastPhotoNumber, 2);
  fixture.files.add("c.jpg");
  await session.photos.scanPhotos();
  assert.deepEqual(numbers(session), { "a.jpg": 1, "c.jpg": 3 });
  await session.io.flushAutosave();
});

test("viewer separates persistent identity from filtered position and retains the filename", async () => {
  const fixture = project({ "long_original_gf1_0034.jpg": photo({ photoNumber: 38 }), "z.jpg": photo({ photoNumber: 1123 }) }, 1123);
  const session = await open(fixture);
  await session.io.flushAutosave();
  const { browser, ctx, state, photos } = session;
  const elements = { photoNumberEl: {}, photoMetaEl: {}, photoFilenameEl: {} };
  const render = browser.AutoBerichtPhotoSorterRender.init(ctx, {
    elements, photosApi: { ...photos, loadPhotoUrl: async () => null }, actions: {}, tagsApi: {},
  });
  render.renderViewer();
  assert.equal(elements.photoNumberEl.textContent, "Photo 038");
  assert.equal(elements.photoMetaEl.textContent, "1 of 2 • Unsorted 2");
  assert.equal(elements.photoFilenameEl.textContent, "long_original_gf1_0034.jpg");
  state.photos[0].tags.observations = ["Escaliers"];
  state.activeTagFilters.observations = ["Escaliers"];
  render.renderViewer();
  assert.equal(elements.photoNumberEl.textContent, "Photo 038");
  assert.equal(elements.photoMetaEl.textContent, "1 of 1 filtered • Total 2 • Unsorted 1");
  state.activeTagFilters.observations = ["No match"];
  render.renderViewer();
  assert.equal(elements.photoNumberEl.hidden, true);
  assert.equal(elements.photoFilenameEl.textContent, "");
  state.activeTagFilters.observations = [];
  state.currentIndex = 1;
  render.renderViewer();
  assert.equal(elements.photoNumberEl.textContent, "Photo 1123");
});
