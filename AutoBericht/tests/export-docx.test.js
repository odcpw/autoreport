const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const { ROOT, loadBrowserScripts } = require("./helpers");

const asBytes = async (value) => {
  if (typeof value === "string") return new TextEncoder().encode(value);
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (value?.arrayBuffer) return new Uint8Array(await value.arrayBuffer());
  return new Uint8Array(value);
};

class MemoryFileHandle {
  constructor(name, bytes = new Uint8Array()) {
    this.kind = "file";
    this.name = name;
    this.bytes = bytes;
  }

  async getFile() {
    const blob = new Blob([this.bytes]);
    Object.defineProperty(blob, "name", { value: this.name });
    return blob;
  }

  async createWritable() {
    let pending = new Uint8Array();
    return {
      write: async (value) => { pending = await asBytes(value); },
      close: async () => { this.bytes = pending; },
    };
  }
}

class MemoryDirectoryHandle {
  constructor(name) {
    this.kind = "directory";
    this.name = name;
    this.children = new Map();
  }

  async getDirectoryHandle(name, options = {}) {
    const child = this.children.get(name);
    if (child?.kind === "directory") return child;
    if (!options.create) {
      const error = new Error(`${name} not found`);
      error.name = "NotFoundError";
      throw error;
    }
    const created = new MemoryDirectoryHandle(name);
    this.children.set(name, created);
    return created;
  }

  async getFileHandle(name, options = {}) {
    const child = this.children.get(name);
    if (child?.kind === "file") return child;
    if (!options.create) {
      const error = new Error(`${name} not found`);
      error.name = "NotFoundError";
      throw error;
    }
    const created = new MemoryFileHandle(name);
    this.children.set(name, created);
    return created;
  }

  async *entries() {
    yield* this.children.entries();
  }
}

test("French Word export inserts Chapter 0 customer context at the real DOCX boundary", async () => {
  const root = new MemoryDirectoryHandle("project");
  const templates = await root.getDirectoryHandle("templates", { create: true });
  const templateName = "Vorlage IST-Aufnahme-Bericht f.V01.docx";
  templates.children.set(
    templateName,
    new MemoryFileHandle(
      templateName,
      new Uint8Array(fs.readFileSync(path.join(ROOT, "project-template", "templates", templateName))),
    ),
  );

  const context = loadBrowserScripts([
    "mini/shared/word-docx-zip.js",
    "mini/shared/word-docx-xml.js",
    "mini/shared/word-export.js",
  ], {
    Blob,
    Response,
    DecompressionStream,
    location: { href: "http://localhost/mini/index.html" },
    fetch: async () => new Response(
      fs.readFileSync(path.join(ROOT, "project-template", "templates", templateName)),
      { status: 200 },
    ),
    createImageBitmap: async () => ({ width: 760, height: 500, close() {} }),
    document: { createElement() { throw new Error("Canvas fallback was not expected."); } },
    AutoBerichtState: {
      formatChapterLabel: (chapter) => String(chapter?.id || ""),
    },
    AutoBerichtSpiderChart: {
      drawToBlob: async () => new Blob([new Uint8Array([137, 80, 78, 71])], { type: "image/png" }),
    },
  });

  const chapterIds = ["0", "1", "2", "3", "4", "4.8", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14"];
  const project = {
    meta: {
      locale: "fr-CH",
      company: "Client test",
      moderator: "Consultant",
      createdAt: "2026-08-22T00:00:00.000Z",
    },
    chapters: chapterIds.map((id) => ({
      id,
      title: { fr: `Chapitre ${id}` },
      meta: id === "0" ? { frontMatterText: "Premier paragraphe client.\n\nDeuxième paragraphe client." } : {},
      rows: [],
    })),
  };
  const result = await context.AutoBerichtWordExport.exportReportDocx({
    project,
    projectHandle: root,
    computeSpider: async () => ({ effective: { chapters_1_11: [] } }),
    compareIdSegments: (a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }),
    toText: (value) => value == null ? "" : String(value),
  });
  const outputName = result.savedAs.split("/").pop();
  const outputHandle = await (await root.getDirectoryHandle("outputs")).getFileHandle(outputName);
  const outputBuffer = outputHandle.bytes.buffer.slice(
    outputHandle.bytes.byteOffset,
    outputHandle.bytes.byteOffset + outputHandle.bytes.byteLength,
  );
  const entries = await context.AutoBerichtWordDocxZip.unzipAllEntries(outputBuffer);
  const documentEntry = entries.find((entry) => entry.name === "word/document.xml");
  const documentXml = new TextDecoder().decode(documentEntry.data);
  assert.match(documentXml, /Premier paragraphe client\./);
  assert.match(documentXml, /Deuxième paragraphe client\./);
  assert.doesNotMatch(documentXml, /CHAPTER(?:0_FRONT_MATTER|[0-9.]+)\$\$/);
  assert.doesNotMatch(documentXml, /SPIDER\$\$/);

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "autobericht-docx-test-"));
  const outputPath = path.join(tempDir, outputName);
  fs.writeFileSync(outputPath, outputHandle.bytes);
  const validation = spawnSync("ooxml", ["validate", "--strict", outputPath], { encoding: "utf8" });
  assert.equal(validation.status, 0, validation.stderr || validation.stdout);
  fs.rmSync(tempDir, { recursive: true, force: true });

  await assert.rejects(
    context.AutoBerichtWordExport.exportReportDocx({
      project: { ...project, chapters: project.chapters.filter((chapter) => chapter.id !== "14") },
      projectHandle: root,
      computeSpider: async () => ({ effective: { chapters_1_11: [] } }),
      compareIdSegments: (a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }),
      toText: (value) => value == null ? "" : String(value),
    }),
    /Project data is missing a chapter required by the Word template \(CHAPTER14\$\$\)/,
  );

  const templateBytes = new Uint8Array(fs.readFileSync(path.join(ROOT, "project-template", "templates", templateName)));
  const templateBuffer = templateBytes.buffer.slice(templateBytes.byteOffset, templateBytes.byteOffset + templateBytes.byteLength);
  const legacyEntries = await context.AutoBerichtWordDocxZip.unzipAllEntries(templateBuffer);
  const legacyDocument = legacyEntries.find((entry) => entry.name === "word/document.xml");
  legacyDocument.data = new TextEncoder().encode(
    new TextDecoder().decode(legacyDocument.data).replace("CHAPTER0_FRONT_MATTER$$", "Legacy hard-coded customer text"),
  );
  const legacyRoot = new MemoryDirectoryHandle("legacy-project");
  const legacyTemplates = await legacyRoot.getDirectoryHandle("templates", { create: true });
  legacyTemplates.children.set(templateName, new MemoryFileHandle(templateName, context.AutoBerichtWordDocxZip.buildZipStore(legacyEntries)));
  const notices = [];
  const fallbackResult = await context.AutoBerichtWordExport.exportReportDocx({
    project,
    projectHandle: legacyRoot,
    computeSpider: async () => ({ effective: { chapters_1_11: [] } }),
    compareIdSegments: (a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }),
    toText: (value) => value == null ? "" : String(value),
    notify: (message) => notices.push(message),
  });
  assert.match(fallbackResult.savedAs, /^outputs\//);
  assert.equal(notices.some((message) => /template is outdated; using the bundled/.test(message)), true);
});
