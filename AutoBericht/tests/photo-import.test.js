const test = require("node:test");
const assert = require("node:assert/strict");

const { loadBrowserScripts } = require("./helpers");

const createDirectory = (name, options = {}) => {
  const entries = new Map();
  const directory = {
    kind: "directory",
    name,
    entries,
    async getDirectoryHandle(childName, handleOptions = {}) {
      const existing = entries.get(childName);
      if (existing?.kind === "directory") return existing;
      if (!handleOptions.create) {
        const error = new Error(`${childName} not found`);
        error.name = "NotFoundError";
        throw error;
      }
      const child = createDirectory(childName, options);
      entries.set(childName, child);
      return child;
    },
    async getFileHandle(fileName, handleOptions = {}) {
      const existing = entries.get(fileName);
      if (existing?.kind === "file") return existing;
      if (!handleOptions.create) {
        const error = new Error(`${fileName} not found`);
        error.name = "NotFoundError";
        throw error;
      }
      const handle = createFileHandle(fileName, Buffer.alloc(0), options);
      entries.set(fileName, handle);
      return handle;
    },
    async removeEntry(childName) {
      if (!entries.delete(childName)) {
        const error = new Error(`${childName} not found`);
        error.name = "NotFoundError";
        throw error;
      }
    },
    async *values() {
      yield* entries.values();
    },
  };
  return directory;
};

const toBuffer = async (value) => {
  if (Buffer.isBuffer(value)) return value;
  if (value?.arrayBuffer) return Buffer.from(await value.arrayBuffer());
  if (value?.bytes) return Buffer.from(value.bytes);
  return Buffer.from(value || "");
};

const createFileHandle = (name, initialBytes, options = {}) => {
  let bytes = Buffer.from(initialBytes);
  return {
    kind: "file",
    name,
    async getFile() {
      const snapshot = Buffer.from(bytes);
      return {
        name,
        size: snapshot.length,
        bytes: snapshot,
        async arrayBuffer() {
          return snapshot.buffer.slice(snapshot.byteOffset, snapshot.byteOffset + snapshot.byteLength);
        },
      };
    },
    async createWritable() {
      let pending = Buffer.alloc(0);
      return {
        async write(value) {
          pending = await toBuffer(value);
        },
        async close() {
          bytes = options.truncateWrites && pending.length ? pending.subarray(0, pending.length - 1) : pending;
        },
      };
    },
  };
};

const addFile = (directory, name, bytes, options = {}) => {
  const handle = createFileHandle(name, bytes, options);
  directory.entries.set(name, handle);
  return handle;
};

const getNestedDirectory = async (root, parts, options = {}) => {
  let current = root;
  for (const part of parts) {
    current = await current.getDirectoryHandle(part, options);
  }
  return current;
};

const createVideoProject = (options = {}) => {
  const root = createDirectory("project", options);
  const photos = createDirectory("photos", options);
  const raw = createDirectory("raw", options);
  const owner = createDirectory("abc", options);
  root.entries.set("photos", photos);
  photos.entries.set("raw", raw);
  raw.entries.set("abc", owner);
  addFile(owner, "walkthrough.MOV", Buffer.from("video-payload"), options);
  return { root, photos, owner };
};

test("raw videos are copied into photos/videos and the intake originals remain untouched", async () => {
  const browser = loadBrowserScripts(["mini/shared/photo-import.js"]);
  const { root, photos, owner } = createVideoProject();

  const result = await browser.AutoBerichtPhotoImport.importRawPhotos({
    projectHandle: root,
    getNestedDirectory,
    isImageFile: () => false,
  });

  assert.equal(result.copiedVideoCount, 1);
  assert.equal(owner.entries.has("walkthrough.MOV"), true);
  const videos = photos.entries.get("videos");
  const copied = await videos.entries.get("abc_walkthrough.MOV").getFile();
  assert.equal(Buffer.from(copied.bytes).toString(), "video-payload");
});

test("an incomplete video copy leaves the raw source untouched", async () => {
  const browser = loadBrowserScripts(["mini/shared/photo-import.js"]);
  const { root, photos, owner } = createVideoProject();
  const videos = createDirectory("videos", { truncateWrites: true });
  photos.entries.set("videos", videos);

  await assert.rejects(
    browser.AutoBerichtPhotoImport.importRawPhotos({
      projectHandle: root,
      getNestedDirectory,
      isImageFile: () => false,
    }),
    /Video copy verification failed/,
  );

  assert.equal(owner.entries.has("walkthrough.MOV"), true);
});
