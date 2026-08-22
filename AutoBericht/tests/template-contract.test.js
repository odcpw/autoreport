const test = require("node:test");
const assert = require("node:assert/strict");
const path = require("node:path");
const { spawnSync } = require("node:child_process");
const { ROOT } = require("./helpers");

const templates = [
  "Vorlage IST-Aufnahme-Bericht d.V01.docx",
  "Vorlage IST-Aufnahme-Bericht f.V01.docx",
  "Vorlage IST-Aufnahme-Bericht i.V01.docx",
];

test("Word templates are valid and expose the required Chapter 0 markers", () => {
  templates.forEach((name) => {
    const file = path.join(ROOT, "project-template", "templates", name);
    const validation = spawnSync("ooxml", ["validate", "--strict", file], { encoding: "utf8" });
    assert.equal(validation.status, 0, validation.stderr || validation.stdout);
    const extraction = spawnSync("ooxml", ["docx", "text", file, "--json"], { encoding: "utf8" });
    assert.equal(extraction.status, 0, extraction.stderr || extraction.stdout);
    assert.match(extraction.stdout, /CHAPTER0_FRONT_MATTER\$\$/);
    assert.match(extraction.stdout, /CHAPTER0\$\$/);
    assert.match(extraction.stdout, /SPIDER\$\$/);
  });
});
