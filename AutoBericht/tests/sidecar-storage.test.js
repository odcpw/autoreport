const test = require("node:test");
const assert = require("node:assert/strict");
const { loadBrowserScripts, createMemoryDirectory } = require("./helpers");

const loadStorage = () => loadBrowserScripts(["mini/shared/sidecar-storage.js"]).AutoBerichtSidecarStorage;

test("a failed sidecar write rejects instead of reporting success", async () => {
  const storage = loadStorage();
  const base = { report: { value: "old" }, photos: { value: "old" } };
  const dirHandle = createMemoryDirectory({ "project_sidecar.json": JSON.stringify(base) }, { failWrite: true });
  await assert.rejects(
    storage.saveBranch({
      dirHandle,
      baseDoc: base,
      branch: "report",
      merge: (latest) => ({ ...latest, report: { value: "new" } }),
    }),
    (error) => error.name === "SidecarWriteError" && /disk full/.test(error.message),
  );
});

test("a corrupted sidecar is never treated as a missing project", async () => {
  const storage = loadStorage();
  const dirHandle = createMemoryDirectory({ "project_sidecar.json": "{broken" });
  await assert.rejects(
    storage.readSidecar(dirHandle, { allowMissing: true }),
    (error) => error.name === "SidecarParseError" && /corrupted or invalid JSON/.test(error.message),
  );
});

test("report and photo writers preserve each other's branches", async () => {
  const storage = loadStorage();
  const base = { report: { value: "r0" }, photos: { value: "p0" } };
  const dirHandle = createMemoryDirectory({ "project_sidecar.json": JSON.stringify(base) });
  await storage.saveBranch({
    dirHandle,
    baseDoc: base,
    branch: "report",
    writerId: "report-test",
    merge: (latest) => ({ ...latest, report: { value: "r1" } }),
  });
  const result = await storage.saveBranch({
    dirHandle,
    baseDoc: base,
    branch: "photos",
    writerId: "photos-test",
    merge: (latest) => ({ ...latest, photos: { value: "p1" } }),
  });
  assert.equal(result.report.value, "r1");
  assert.equal(result.photos.value, "p1");
  assert.equal(result.meta.revision, 2);
});

test("a stale writer on the same branch gets a conflict", async () => {
  const storage = loadStorage();
  const base = { report: { value: "r0" } };
  const dirHandle = createMemoryDirectory({ "project_sidecar.json": JSON.stringify(base) });
  await storage.saveBranch({
    dirHandle,
    baseDoc: base,
    branch: "report",
    merge: (latest) => ({ ...latest, report: { value: "r1" } }),
  });
  await assert.rejects(
    storage.saveBranch({
      dirHandle,
      baseDoc: base,
      branch: "report",
      merge: (latest) => ({ ...latest, report: { value: "r2" } }),
    }),
    (error) => error.name === "SidecarConflictError",
  );
});

test("readback verification detects a changed saved branch", async () => {
  const storage = loadStorage();
  const base = { report: { value: "r0" } };
  const dirHandle = createMemoryDirectory(
    { "project_sidecar.json": JSON.stringify(base) },
    { mutateAfterWrite: () => JSON.stringify({ report: { value: "tampered" } }) },
  );
  await assert.rejects(
    storage.saveBranch({
      dirHandle,
      baseDoc: base,
      branch: "report",
      merge: (latest) => ({ ...latest, report: { value: "r1" } }),
    }),
    (error) => error.name === "SidecarVerificationError",
  );
});

test("readback verification also detects loss from the other editor branch", async () => {
  const storage = loadStorage();
  const base = { report: { value: "r0" }, photos: { value: "p0" } };
  const dirHandle = createMemoryDirectory(
    { "project_sidecar.json": JSON.stringify(base) },
    { mutateAfterWrite: (saved) => {
      const value = JSON.parse(saved);
      delete value.photos;
      return JSON.stringify(value);
    } },
  );
  await assert.rejects(
    storage.saveBranch({
      dirHandle,
      baseDoc: base,
      branch: "report",
      merge: (latest) => ({ ...latest, report: { value: "r1" } }),
    }),
    (error) => error.name === "SidecarVerificationError",
  );
});
