const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const http = require("node:http");
const os = require("node:os");
const path = require("node:path");
const { spawn, spawnSync } = require("node:child_process");
const { ROOT, loadBrowserScripts } = require("./helpers");

const ooxmlAvailable = spawnSync("ooxml", ["--json", "capabilities"], { encoding: "utf8" }).status === 0;
const firefoxAvailable = spawnSync("firefox", ["--version"], { encoding: "utf8" }).status === 0;

const contentType = (filename) => {
  const extension = path.extname(filename).toLowerCase();
  if (extension === ".html") return "text/html; charset=utf-8";
  if (extension === ".js") return "text/javascript; charset=utf-8";
  if (extension === ".pptx") return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  return "application/octet-stream";
};

const harnessHtml = String.raw`<!doctype html>
<html><body>
<script src="/mini/shared/markdown.js"></script>
<script src="/mini/shared/report-rows.js"></script>
<script src="/mini/shared/word-docx-zip.js"></script>
<script src="/mini/shared/pptx-export.js"></script>
<script>
(() => {
  const asBytes = async (value) => {
    if (typeof value === "string") return new TextEncoder().encode(value);
    if (value instanceof Uint8Array) return value;
    if (value instanceof ArrayBuffer) return new Uint8Array(value);
    if (value && typeof value.arrayBuffer === "function") return new Uint8Array(await value.arrayBuffer());
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
      if (child && child.kind === "directory") return child;
      if (!options.create) {
        const error = new Error(name + " not found");
        error.name = "NotFoundError";
        throw error;
      }
      const created = new MemoryDirectoryHandle(name);
      this.children.set(name, created);
      return created;
    }
    async getFileHandle(name, options = {}) {
      const child = this.children.get(name);
      if (child && child.kind === "file") return child;
      if (!options.create) {
        const error = new Error(name + " not found");
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

  const run = async () => {
    const root = new MemoryDirectoryHandle("project");
    window.AutoBerichtState = {
      formatChapterLabel: (chapter) => String(chapter && chapter.id || ""),
    };
    window.AutoBerichtSpiderChart = {
      drawToBlob: async () => {
        const source = await (await fetch("/mini/demo-photos/1.jpg")).blob();
        const bitmap = await createImageBitmap(source);
        const canvas = document.createElement("canvas");
        canvas.width = bitmap.width;
        canvas.height = bitmap.height;
        canvas.getContext("2d").drawImage(bitmap, 0, 0);
        bitmap.close();
        return new Promise((resolve, reject) => canvas.toBlob(
          (blob) => blob ? resolve(blob) : reject(new Error("PNG conversion failed")),
          "image/png",
        ));
      },
    };
    const project = {
      meta: { locale: "fr-CH", company: "Client test" },
      chapters: [{
        id: "0",
        title: { fr: "0 Synthèse" },
        rows: [{
          id: "0.1",
          master: { finding: "", recommendation: "" },
          workstate: {
            includeFinding: true,
            done: true,
            recommendationText: "Mesure **prioritaire**, *responsable* et [Suva](https://www.suva.ch/).",
          },
        }],
      }],
    };
    const result = await window.AutoBerichtPptxExport.exportReportPptx({
      project,
      sidecarDoc: { photos: { photos: {} } },
      projectHandle: root,
      compareIdSegments: (left, right) => String(left).localeCompare(String(right), undefined, { numeric: true }),
      toText: (value) => value == null ? "" : String(value),
      computeSpider: async () => ({ effective: { chapters_1_11: [] } }),
    });
    const outputs = await root.getDirectoryHandle("outputs");
    const outputName = result.savedAs.split("/").pop();
    const output = await outputs.getFileHandle(outputName);
    const response = await fetch("/__pptx_result", {
      method: "POST",
      headers: {
        "Content-Type": "application/octet-stream",
        "X-Slide-Count": String(result.slideCount),
        "X-Output-Name": outputName,
      },
      body: output.bytes,
    });
    if (!response.ok) throw new Error("Result upload failed: " + response.status);
  };

  run().catch(async (error) => {
    await fetch("/__pptx_error", { method: "POST", body: String(error && error.stack || error) });
  });
})();
</script>
</body></html>`;

test("French PowerPoint export produces a valid, rendered PPTX with Markdown runs", {
  skip: !firefoxAvailable || !ooxmlAvailable,
  timeout: 45000,
}, async (t) => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "autobericht-pptx-test-"));
  const profileDir = path.join(tempDir, "firefox-profile");
  const outputPath = path.join(tempDir, "generated.pptx");
  const renderDir = path.join(tempDir, "rendered");
  fs.mkdirSync(profileDir);
  t.after(() => fs.rmSync(tempDir, { recursive: true, force: true }));

  let settleResult;
  let settleError;
  const browserResult = new Promise((resolve, reject) => {
    settleResult = resolve;
    settleError = reject;
  });

  const server = http.createServer((request, response) => {
    const requestUrl = new URL(request.url, "http://127.0.0.1");
    if (request.method === "POST" && requestUrl.pathname === "/__pptx_result") {
      const chunks = [];
      request.on("data", (chunk) => chunks.push(chunk));
      request.on("end", () => {
        const bytes = Buffer.concat(chunks);
        fs.writeFileSync(outputPath, bytes);
        response.writeHead(204);
        response.end();
        settleResult({
          bytes: bytes.length,
          slideCount: Number(request.headers["x-slide-count"] || 0),
          outputName: String(request.headers["x-output-name"] || ""),
        });
      });
      return;
    }
    if (request.method === "POST" && requestUrl.pathname === "/__pptx_error") {
      const chunks = [];
      request.on("data", (chunk) => chunks.push(chunk));
      request.on("end", () => {
        const message = Buffer.concat(chunks).toString("utf8");
        response.writeHead(204);
        response.end();
        settleError(new Error(message));
      });
      return;
    }
    if (requestUrl.pathname === "/__pptx_harness.html") {
      response.writeHead(200, { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" });
      response.end(harnessHtml);
      return;
    }

    const relative = decodeURIComponent(requestUrl.pathname).replace(/^\/+/, "");
    const servingRoot = relative.startsWith("templates/") ? path.dirname(ROOT) : ROOT;
    const filename = path.resolve(servingRoot, relative);
    const rootBoundary = `${servingRoot}${path.sep}`;
    if (filename !== servingRoot && !filename.startsWith(rootBoundary)) {
      response.writeHead(400);
      response.end("Bad request");
      return;
    }
    if (!fs.existsSync(filename) || !fs.statSync(filename).isFile()) {
      response.writeHead(404);
      response.end("Not found");
      return;
    }
    response.writeHead(200, { "Content-Type": contentType(filename), "Cache-Control": "no-store" });
    fs.createReadStream(filename).pipe(response);
  });

  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  t.after(() => {
    server.closeAllConnections?.();
    server.close();
  });
  const address = server.address();
  const harnessUrl = `http://127.0.0.1:${address.port}/__pptx_harness.html`;
  const firefox = spawn("firefox", ["--headless", "--no-remote", "--profile", profileDir, harnessUrl], {
    stdio: ["ignore", "ignore", "pipe"],
  });
  let firefoxError = "";
  firefox.stderr.on("data", (chunk) => { firefoxError += chunk.toString(); });
  t.after(() => {
    if (firefox.exitCode == null) firefox.kill("SIGTERM");
  });

  let timeoutId;
  const timeout = new Promise((_, reject) => {
    timeoutId = setTimeout(
      () => reject(new Error(`Timed out waiting for Firefox export. ${firefoxError}`)),
      30000,
    );
  });
  const result = await Promise.race([browserResult, timeout]);
  clearTimeout(timeoutId);
  if (firefox.exitCode == null) firefox.kill("SIGTERM");
  server.closeAllConnections?.();
  assert.ok(result.bytes > 1000, `Unexpected PPTX size: ${result.bytes}`);
  assert.ok(result.slideCount >= 3, `Unexpected generated slide count: ${result.slideCount}`);
  assert.match(result.outputName, /Client-test-Bericht-Besprechung\.pptx$/);

  const zipContext = loadBrowserScripts(["mini/shared/word-docx-zip.js"], {
    Blob,
    Response,
    DecompressionStream,
  });
  const outputBytes = fs.readFileSync(outputPath);
  const outputBuffer = outputBytes.buffer.slice(outputBytes.byteOffset, outputBytes.byteOffset + outputBytes.byteLength);
  const entries = await zipContext.AutoBerichtWordDocxZip.unzipAllEntries(outputBuffer);
  const decoder = new TextDecoder();
  const markdownSlide = entries.find((entry) => (
    /^ppt\/slides\/slide\d+\.xml$/.test(entry.name)
    && decoder.decode(entry.data).includes("prioritaire")
  ));
  assert.ok(markdownSlide, "Generated Markdown slide was not found in the PPTX package.");
  const slideXml = decoder.decode(markdownSlide.data);
  assert.match(slideXml, /<a:rPr[^>]*b="1"\/><a:t[^>]*>prioritaire<\/a:t>/);
  assert.match(slideXml, /<a:rPr[^>]*i="1"\/><a:t[^>]*>responsable<\/a:t>/);
  assert.doesNotMatch(slideXml, /\*\*prioritaire\*\*|\*responsable\*|\[Suva\]\(/);
  const partNumber = markdownSlide.name.match(/slide(\d+)\.xml$/)[1];
  const relationship = entries.find((entry) => entry.name === `ppt/slides/_rels/slide${partNumber}.xml.rels`);
  assert.ok(relationship, "Generated Markdown slide relationships were not found.");
  assert.match(decoder.decode(relationship.data), /Target="https:\/\/www\.suva\.ch\/"[^>]*TargetMode="External"/);

  const validation = spawnSync("ooxml", ["validate", "--strict", outputPath], { encoding: "utf8" });
  assert.equal(validation.status, 0, validation.stderr || validation.stdout);

  const slideList = spawnSync("ooxml", ["--json", "pptx", "slides", "list", outputPath], { encoding: "utf8" });
  assert.equal(slideList.status, 0, slideList.stderr || slideList.stdout);
  const slides = JSON.parse(slideList.stdout).slides;
  const slideInfo = slides.find((slide) => slide.partUri === `/${markdownSlide.name}`);
  assert.ok(slideInfo, `Generated slide part ${markdownSlide.name} is not in the presentation order.`);
  const render = spawnSync("ooxml", [
    "--json", "pptx", "render", outputPath,
    "--out", renderDir,
    "--slides", String(slideInfo.number),
  ], { encoding: "utf8" });
  assert.equal(render.status, 0, render.stderr || render.stdout);
  const rendered = fs.readdirSync(renderDir).filter((name) => name.toLowerCase().endsWith(".png"));
  assert.ok(rendered.length >= 1, "Generated Markdown slide did not render to PNG.");
});
