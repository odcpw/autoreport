const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const { ROOT, loadBrowserScripts } = require("./helpers");

// Strict OOXML validation needs the optional `ooxml` CLI; the structural
// assertions below run everywhere.
const ooxmlAvailable = spawnSync("ooxml", ["--json", "capabilities"], { encoding: "utf8" }).status === 0;

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

  const context = loadBrowserScripts([
    "mini/shared/markdown.js",
    "mini/shared/report-rows.js",
    "mini/shared/word-docx-zip.js",
    "mini/shared/word-docx-xml.js",
    "mini/shared/word-export.js",
  ], {
    Blob,
    Response,
    DecompressionStream,
    location: { href: "http://localhost/mini/index.html" },
    fetch: async (url) => {
      const requestPath = decodeURIComponent(new URL(url).pathname);
      assert.equal(requestPath, `/templates/${templateName}`);
      return new Response(fs.readFileSync(path.join(ROOT, "..", requestPath)), { status: 200 });
    },
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
      meta: id === "0" ? {
        frontMatterText: "Premier **paragraphe** client.\n\nDeuxième *paragraphe* client avec [Suva](https://www.suva.ch/).",
      } : {},
      rows: id === "0" ? [{
        id: "0.1",
        master: { finding: "", recommendation: "" },
        workstate: {
          includeFinding: true,
          done: true,
          findingText: "",
          recommendationText: "Résumé avec **priorité** et *responsabilité*.",
        },
      }] : [],
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
  assert.match(documentXml, /Premier /);
  assert.match(documentXml, /<w:rPr><w:b\/><\/w:rPr><w:t[^>]*>paragraphe<\/w:t>/);
  assert.match(documentXml, /<w:rPr><w:i\/><\/w:rPr><w:t[^>]*>paragraphe<\/w:t>/);
  assert.match(documentXml, /w:instr="HYPERLINK &quot;https:\/\/www\.suva\.ch\/&quot;"/);
  assert.match(documentXml, /<w:rPr><w:b\/><\/w:rPr><w:t[^>]*>priorité<\/w:t>/);
  assert.match(documentXml, /<w:rPr><w:i\/><\/w:rPr><w:t[^>]*>responsabilité<\/w:t>/);
  assert.doesNotMatch(documentXml, /\*\*(?:paragraphe|priorité)\*\*/);
  assert.doesNotMatch(documentXml, /\*(?:paragraphe|responsabilité)\*/);
  assert.doesNotMatch(documentXml, /CHAPTER(?:0_FRONT_MATTER|[0-9.]+)\$\$/);
  assert.doesNotMatch(documentXml, /SPIDER\$\$/);

  if (ooxmlAvailable) {
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "autobericht-docx-test-"));
    const outputPath = path.join(tempDir, outputName);
    fs.writeFileSync(outputPath, outputHandle.bytes);
    const validation = spawnSync("ooxml", ["validate", "--strict", outputPath], { encoding: "utf8" });
    assert.equal(validation.status, 0, validation.stderr || validation.stdout);
    fs.rmSync(tempDir, { recursive: true, force: true });
  }

  // A project without chapter 14 still exports; the unused marker is removed
  // and reported instead of aborting the export.
  const partialNotices = [];
  const partialResult = await context.AutoBerichtWordExport.exportReportDocx({
    project: { ...project, chapters: project.chapters.filter((chapter) => chapter.id !== "14") },
    projectHandle: root,
    computeSpider: async () => { throw new Error("weights unavailable"); },
    compareIdSegments: (a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }),
    toText: (value) => value == null ? "" : String(value),
    notify: (message) => partialNotices.push(message),
  });
  assert.match(partialResult.savedAs, /^outputs\//);
  assert.equal(partialNotices.some((message) => /CHAPTER14\$\$ has no matching chapter/.test(message)), true);
  assert.equal(partialNotices.some((message) => /spider picture skipped \(weights unavailable\)/.test(message)), true);
  const partialHandle = await (await root.getDirectoryHandle("outputs")).getFileHandle(partialResult.savedAs.split("/").pop());
  const partialEntries = await context.AutoBerichtWordDocxZip.unzipAllEntries(partialHandle.bytes.buffer.slice(
    partialHandle.bytes.byteOffset,
    partialHandle.bytes.byteOffset + partialHandle.bytes.byteLength,
  ));
  const partialXml = new TextDecoder().decode(partialEntries.find((entry) => entry.name === "word/document.xml").data);
  assert.doesNotMatch(partialXml, /SPIDER\$\$|THERMO[0-9A-Za-z_.]*\$\$|CHAPTER(?:0_FRONT_MATTER|[0-9.]+)\$\$/);

  const templateBytes = new Uint8Array(fs.readFileSync(path.join(ROOT, "..", "templates", templateName)));
  const templateBuffer = templateBytes.buffer.slice(templateBytes.byteOffset, templateBytes.byteOffset + templateBytes.byteLength);
  const legacyEntries = await context.AutoBerichtWordDocxZip.unzipAllEntries(templateBuffer);
  const legacyDocument = legacyEntries.find((entry) => entry.name === "word/document.xml");
  legacyDocument.data = new TextEncoder().encode(
    new TextDecoder().decode(legacyDocument.data).replace("CHAPTER0_FRONT_MATTER$$", "Legacy hard-coded customer text"),
  );
  const legacyRoot = new MemoryDirectoryHandle("legacy-project");
  const legacyTemplates = await legacyRoot.getDirectoryHandle("templates", { create: true });
  legacyTemplates.children.set(templateName, new MemoryFileHandle(templateName, context.AutoBerichtWordDocxZip.buildZipStore(legacyEntries)));
  // A project template from before the front-matter marker existed is still
  // used as-is; only the customer context is skipped.
  const notices = [];
  const legacyResult = await context.AutoBerichtWordExport.exportReportDocx({
    project,
    projectHandle: legacyRoot,
    computeSpider: async () => ({ effective: { chapters_1_11: [] } }),
    compareIdSegments: (a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }),
    toText: (value) => value == null ? "" : String(value),
    notify: (message) => notices.push(message),
  });
  assert.match(legacyResult.savedAs, /^outputs\//);
  assert.equal(notices.some((message) => /using project template/.test(message)), true);
  assert.equal(notices.some((message) => /no CHAPTER0_FRONT_MATTER\$\$ marker; customer context was skipped/.test(message)), true);
  assert.equal(notices.some((message) => /using the bundled/.test(message)), false);
});
