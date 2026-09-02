const test = require("node:test");
const assert = require("node:assert/strict");
const XLSX = require("../libs/sheetjs/xlsx.full.min.js");
const { loadBrowserScripts } = require("./helpers");

const testI18n = {
  setLocale() {},
  t: (key, fallback) => fallback || key,
  tf: (key, fallback, values = {}) => String(fallback || key).replace(/\{(\w+)\}/g, (match, name) => String(values[name] ?? match)),
};

const notFound = (name) => {
  const error = new Error(`${name} not found`);
  error.name = "NotFoundError";
  return error;
};

const toBytes = (value) => {
  if (typeof value === "string") return new TextEncoder().encode(value);
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  return new Uint8Array(value.buffer ? value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength) : value);
};

class MemoryFile {
  constructor(name, bytes = new Uint8Array(), options = {}) {
    this.kind = "file";
    this.name = name;
    this.bytes = bytes;
    this.options = options;
  }

  async getFile() {
    const bytes = this.bytes;
    return {
      name: this.name,
      size: bytes.byteLength,
      async arrayBuffer() { return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength); },
    };
  }

  async createWritable() {
    if (this.options.failWrites) throw new Error("disk full");
    let pending = new Uint8Array();
    return {
      write: async (value) => { pending = toBytes(value); },
      close: async () => { this.bytes = pending; },
    };
  }
}

class MemoryDir {
  constructor(name, options = {}) {
    this.kind = "directory";
    this.name = name;
    this.options = options;
    this.children = new Map();
  }

  async getDirectoryHandle(name, handleOptions = {}) {
    const existing = this.children.get(name);
    if (existing?.kind === "directory") return existing;
    if (!handleOptions.create) throw notFound(name);
    const created = new MemoryDir(name, this.options);
    this.children.set(name, created);
    return created;
  }

  async getFileHandle(name, handleOptions = {}) {
    const existing = this.children.get(name);
    if (existing?.kind === "file") return existing;
    if (!handleOptions.create) throw notFound(name);
    const created = new MemoryFile(name, new Uint8Array(), this.options);
    this.children.set(name, created);
    return created;
  }

  async *entries() { yield* this.children.entries(); }
}

const buildWorkbook = () => {
  const sheet = XLSX.utils.aoa_to_sheet([
    ["Nr", "", "Frage", "Ja", "Nein", "Bemerkung"],
    ["1.1.1", "", "Frage eins", "x", "", ""],
    ["1.1.2", "", "Frage zwei", "", "x", "Kommentar"],
  ]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Selbstbeurteilung Kunde");
  return new Uint8Array(XLSX.write(workbook, { type: "array", bookType: "xlsx" }));
};

const runImport = async ({ root }) => {
  const workbookBytes = buildWorkbook();
  const picked = new MemoryFile("Selbstbeurteilung.xlsx", workbookBytes);
  const context = loadBrowserScripts([
    "mini/shared/self-assessment.js",
    "mini/shared/import-self.js",
  ], {
    XLSX,
    showOpenFilePicker: async () => [picked],
    document: {},
  });
  const project = {
    meta: { locale: "de-CH" },
    chapters: [{
      id: "1",
      rows: [
        { id: "1.1.1", kind: "item", master: {}, workstate: {} },
        { id: "1.1.2", kind: "item", master: {}, workstate: {} },
      ],
    }],
  };
  const statuses = [];
  let saves = 0;
  let renders = 0;
  const handler = context.AutoBerichtImportSelf.createHandler(
    { runtime: { dirHandle: root }, debug: { logLine() {} }, setStatus: (message) => statuses.push(message), state: { project }, i18n: testI18n },
    { renderRows() { renders += 1; }, saveSidecar: async () => { saves += 1; } },
  );
  await handler();
  return { project, statuses, saves, renders, workbookBytes };
};

test("self-assessment import applies answers and copies the workbook into inputs", async () => {
  const root = new MemoryDir("project");
  const result = await runImport({ root });

  const rows = result.project.chapters[0].rows;
  assert.equal(rows[0].customer.items[0].answer, 1);
  assert.equal(rows[1].customer.items[0].answer, 0);
  assert.equal(rows[1].customer.items[0].comment, "Kommentar");
  assert.equal(result.saves, 1);
  assert.equal(result.renders, 1);
  assert.match(result.statuses.at(-1), /^Imported self-assessment answers \(2\)\. Copied Selbstbeurteilung\.xlsx to inputs\.$/);

  const inputs = await root.getDirectoryHandle("inputs");
  const copy = await inputs.getFileHandle("Selbstbeurteilung.xlsx");
  assert.equal(Buffer.compare(Buffer.from(copy.bytes), Buffer.from(result.workbookBytes)), 0);
});

test("a failed copy into inputs keeps the imported answers and says so", async () => {
  const root = new MemoryDir("project");
  root.children.set("inputs", new MemoryDir("inputs", { failWrites: true }));
  const result = await runImport({ root });

  assert.equal(result.project.chapters[0].rows[0].customer.items[0].answer, 1);
  assert.equal(result.saves, 1);
  assert.match(result.statuses.at(-1), /^Imported self-assessment answers \(2\)\. Could not copy Selbstbeurteilung\.xlsx to inputs: disk full$/);
  assert.equal((await root.getDirectoryHandle("inputs")).children.has("Selbstbeurteilung.xlsx"), true);
});
