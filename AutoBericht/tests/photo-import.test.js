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
        lastModified: options.lastModified ?? new Date(2026, 8, 10, 9, 0).getTime(),
        slice(start, end) {
          const part = snapshot.subarray(start, end);
          return { async arrayBuffer() { return part.buffer.slice(part.byteOffset, part.byteOffset + part.byteLength); } };
        },
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

test("raw videos are moved into photos/videos once the copy is verified", async () => {
  const browser = loadBrowserScripts(["mini/shared/photo-import.js"]);
  const { root, photos, owner } = createVideoProject();

  const result = await browser.AutoBerichtPhotoImport.importRawPhotos({
    projectHandle: root,
    getNestedDirectory,
    isImageFile: () => false,
  });

  assert.equal(result.movedVideoCount, 1);
  assert.equal(owner.entries.has("walkthrough.MOV"), false);
  const videos = photos.entries.get("videos");
  const moved = await videos.entries.get("abc_walkthrough.MOV").getFile();
  assert.equal(Buffer.from(moved.bytes).toString(), "video-payload");
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

// Minimal JPEG APP1 with a TIFF IFD0 (ModifyDate + ExifIFD pointer) and an
// ExifIFD DateTimeOriginal. Pixel decoding is stubbed; the real importer reads
// these bytes and writes the output name through its normal directory API.
const datedJpeg = ({ original = "2026:09:08 14:32:17", modified = "2026:09:09 18:45:00", littleEndian = true, xmpFirst = false } = {}) => {
  const tiff = Buffer.alloc(180);
  const u16 = (value, offset) => littleEndian ? tiff.writeUInt16LE(value, offset) : tiff.writeUInt16BE(value, offset);
  const u32 = (value, offset) => littleEndian ? tiff.writeUInt32LE(value, offset) : tiff.writeUInt32BE(value, offset);
  tiff.write(littleEndian ? "II" : "MM", 0, "ascii");
  u16(42, 2); u32(8, 4);
  u16(2, 8);
  u16(0x0132, 10); u16(2, 12); u32(20, 14); u32(140, 18);
  u16(0x8769, 22); u16(4, 24); u32(1, 26); u32(38, 30);
  u16(original === null ? 0 : 1, 38);
  if (original !== null) {
    u16(0x9003, 40); u16(2, 42); u32(20, 44); u32(100, 48);
    tiff.write(original, 100, "ascii");
  }
  tiff.write(modified, 140, "ascii");
  const app1 = (payload) => {
    const header = Buffer.from([0xff, 0xe1, 0, 0]);
    header.writeUInt16BE(payload.length + 2, 2);
    return Buffer.concat([header, payload]);
  };
  return Buffer.concat([
    Buffer.from([0xff, 0xd8]),
    ...(xmpFirst ? [app1(Buffer.from("http://ns.adobe.com/xap/1.0/\0<xmp/>"))] : []),
    app1(Buffer.concat([Buffer.from("Exif\0\0", "ascii"), tiff])),
    Buffer.from([0xff, 0xd9]),
  ]);
};

const imageBrowser = () => loadBrowserScripts(["mini/shared/photo-import.js"], {
  createImageBitmap: async () => ({ width: 4, height: 3, close() {} }),
  document: { createElement: () => ({
    getContext: () => ({ drawImage() {} }),
    toBlob(callback) { callback(Buffer.from("resized-jpeg")); },
  }) },
});

const importImages = async (files, options = {}) => {
  const root = createDirectory("project");
  const raw = await getNestedDirectory(root, ["photos", "raw", "abc"], { create: true });
  files.forEach(([name, bytes]) => addFile(raw, name, bytes, options));
  const context = { projectHandle: root, getNestedDirectory, isImageFile: (name) => /\.jpg$/i.test(name) };
  const api = imageBrowser().AutoBerichtPhotoImport;
  const result = await api.importRawPhotos(context);
  return { result, raw, api, context, names: Array.from(result.resizedHandle.entries.keys()) };
};

test("JPEG filenames use DateTimeOriginal, not EXIF modification or download date", async () => {
  const jpeg = datedJpeg();
  const { result, names, raw, api, context } = await importImages([["phone.jpg", jpeg]]);
  assert.deepEqual(names, ["2026-09-08-14-32_abc_0001.jpg"]);
  assert.equal(result.fileDateCount, 0);
  assert.deepEqual((await raw.entries.get("phone.jpg").getFile()).bytes, jpeg);
  const resumed = await api.importRawPhotos(context);
  assert.equal(resumed.importedCount, 0);
  assert.equal(resumed.skippedCount, 1);
});

test("JPEG EXIF can follow another APP1 segment and use either TIFF byte order", async () => {
  for (const littleEndian of [true, false]) {
    const { names, result } = await importImages([["phone.jpg", datedJpeg({ littleEndian, xmpFirst: true })]]);
    assert.deepEqual(names, ["2026-09-08-14-32_abc_0001.jpg"]);
    assert.equal(result.fileDateCount, 0);
  }
});

test("absent, invalid or truncated capture metadata uses the file date and reports it", async () => {
  const malformed = datedJpeg();
  malformed.writeUInt32LE(0xfffffff0, 16); // Invalid TIFF IFD0 offset.
  for (const bytes of [datedJpeg({ original: null }), datedJpeg({ original: "2026:02:30 14:32:00" }), malformed, datedJpeg().subarray(0, 55)]) {
    const { names, result } = await importImages([["phone.jpg", bytes]]);
    assert.deepEqual(names, ["2026-09-10-09-00_abc_0001.jpg"]);
    assert.equal(result.fileDateCount, 1);
  }
});

test("sequences follow source filenames while capture dates can span several days", async () => {
  const { names, result } = await importImages([
    ["IMG_0002.jpg", datedJpeg({ original: "2026:09:08 23:55:00" })],
    ["IMG_0001.jpg", datedJpeg({ original: "2026:09:09 00:05:00" })],
  ]);
  assert.deepEqual(names, ["2026-09-09-00-05_abc_0001.jpg", "2026-09-08-23-55_abc_0002.jpg"]);
  assert.equal(result.fileDateCount, 0);
});

test("capture wall-clock time does not depend on the importing computer's daylight-saving rules", async () => {
  const previous = process.env.TZ;
  process.env.TZ = "Europe/Zurich";
  try {
    // This time is possible in another camera timezone, but falls in the
    // Swiss DST gap. It must not become 03:30 or a fallback download date.
    const { names, result } = await importImages([["phone.jpg", datedJpeg({ original: "2026:03:29 02:30:00" })]]);
    assert.deepEqual(names, ["2026-03-29-02-30_abc_0001.jpg"]);
    assert.equal(result.fileDateCount, 0);
  } finally {
    if (previous === undefined) delete process.env.TZ;
    else process.env.TZ = previous;
  }
});

test("corrected date naming does not rename or overwrite an already imported project", async () => {
  const jpeg = datedJpeg();
  const { api, context, result, raw, names } = await importImages([["phone.jpg", jpeg]]);
  const existing = result.resizedHandle;
  existing.entries.delete(names[0]);
  const oldName = "2026-09-10-09-00_abc_0001.jpg";
  addFile(existing, oldName, Buffer.from("existing-photo"));
  await assert.rejects(api.importRawPhotos(context), /Unexpected files/);
  assert.deepEqual(Array.from(existing.entries.keys()), [oldName]);
  assert.equal((await existing.entries.get(oldName).getFile()).bytes.toString(), "existing-photo");
  assert.deepEqual((await raw.entries.get("phone.jpg").getFile()).bytes, jpeg);
});

test("the PhotoSorter import action reports when file dates were used", async () => {
  const root = createDirectory("project");
  const raw = await getNestedDirectory(root, ["photos", "raw", "abc"], { create: true });
  addFile(raw, "with-capture.jpg", datedJpeg());
  addFile(raw, "without-capture.jpg", datedJpeg({ original: null }));
  const bindings = loadBrowserScripts(["mini/photosorter/bind-events.js"], { addEventListener() {} });
  let click;
  const messages = [];
  let saved = 0;
  let scanned = 0;
  const state = { projectHandle: root, projectDoc: {} };
  const tf = (key, fallback, values = {}) => fallback.replace(/\{(\w+)\}/g, (match, name) => String(values[name] ?? match));
  bindings.AutoBerichtPhotoSorterBindEvents.bind({
    state, runtime: {}, debug: {}, setStatus: (message) => messages.push(message),
    i18n: { t: (key) => key, tf }, constants: { RESIZE_MAX: 1920, RESIZE_QUALITY: 0.85 },
    photoImport: imageBrowser().AutoBerichtPhotoImport,
  }, {
    elements: { importActionBtn: { addEventListener(event, callback) { if (event === "click") click = callback; } } },
    tagsApi: {}, renderApi: {}, actions: {},
    ioApi: { getNestedDirectory, async saveProjectSidecar() { saved += 1; } },
    photosApi: { isImageFile: (name) => name.endsWith(".jpg"), async scanPhotos() { scanned += 1; } },
  });
  await click();
  assert.equal(messages.at(-1), "Imported 2 photos, 1 with file date (capture time unavailable).");
  assert.equal(saved, 1);
  assert.equal(scanned, 1);
  assert.equal(state.projectDoc.photoRoot, "photos/resized");
});
