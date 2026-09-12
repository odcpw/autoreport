const test = require("node:test");
const assert = require("node:assert/strict");
const { loadBrowserScripts } = require("./helpers");

test("DOCX/PPTX packager refuses ZIP64-sized output instead of emitting a corrupt archive", () => {
  const api = loadBrowserScripts(["mini/shared/word-docx-zip.js"], { Blob, Response, DecompressionStream })
    .AutoBerichtWordDocxZip;
  const entries = Array.from({ length: 0x10000 }, (_, index) => ({
    name: `entry-${index}`,
    data: new Uint8Array(),
  }));
  assert.throws(() => api.buildZipStore(entries), /ZIP64 output is not supported/);
});
