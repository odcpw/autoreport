#!/usr/bin/env node

const fs = require("fs");
const path = require("path");

const usage = () => {
  console.error("Usage: node AutoBericht/tools/migrate-sidecar-schema.js <project_sidecar.json> [--write]");
};

const args = process.argv.slice(2);
if (!args.length || args.includes("-h") || args.includes("--help")) {
  usage();
  process.exit(args.length ? 0 : 1);
}

const inputPath = path.resolve(process.cwd(), args[0]);
const writeChanges = args.includes("--write");

const isPlainObject = (value) => (
  !!value
  && typeof value === "object"
  && !Array.isArray(value)
);

const clone = (value) => JSON.parse(JSON.stringify(value));

const normalizeTagGroup = (value) => {
  if (!Array.isArray(value)) return [];
  return clone(value);
};

const sanitizePhotoDoc = (doc) => {
  const source = isPlainObject(doc) ? doc : {};
  const meta = isPlainObject(source.meta) ? clone(source.meta) : {};
  if (!Object.prototype.hasOwnProperty.call(meta, "projectId")) meta.projectId = "";
  if (!meta.createdAt) meta.createdAt = new Date().toISOString();
  if (!Object.prototype.hasOwnProperty.call(meta, "updatedAt")) meta.updatedAt = "";
  const photos = isPlainObject(source.photos) ? clone(source.photos) : {};
  return {
    meta,
    photoRoot: typeof source.photoRoot === "string" ? source.photoRoot : "",
    photoTagOptions: {
      report: normalizeTagGroup(source.photoTagOptions?.report),
      observations: normalizeTagGroup(source.photoTagOptions?.observations),
      training: normalizeTagGroup(source.photoTagOptions?.training),
    },
    photos,
  };
};

const looksLikeLegacyPhotoDoc = (doc) => (
  isPlainObject(doc)
  && !isPlainObject(doc.report)
  && (
    Object.prototype.hasOwnProperty.call(doc, "photoRoot")
    || Object.prototype.hasOwnProperty.call(doc, "photoTagOptions")
    || isPlainObject(doc.photos)
  )
);

const migrate = (rawDoc) => {
  if (!isPlainObject(rawDoc)) {
    throw new Error("Sidecar root must be a JSON object.");
  }
  const doc = clone(rawDoc);
  const notes = [];

  if (isPlainObject(doc.photos)) {
    const before = JSON.stringify(doc.photos);
    doc.photos = sanitizePhotoDoc(doc.photos);
    if (JSON.stringify(doc.photos) !== before) {
      notes.push("Sanitized photos branch to canonical keys (meta/photoRoot/photoTagOptions/photos).");
    }
  } else if (looksLikeLegacyPhotoDoc(doc)) {
    const converted = sanitizePhotoDoc(doc);
    doc.photos = converted;
    notes.push("Converted legacy top-level photo fields into sidecar.photos.");
  }

  if (Object.prototype.hasOwnProperty.call(doc, "photoRoot")) {
    delete doc.photoRoot;
    notes.push("Removed legacy top-level photoRoot.");
  }
  if (Object.prototype.hasOwnProperty.call(doc, "photoTagOptions")) {
    delete doc.photoTagOptions;
    notes.push("Removed legacy top-level photoTagOptions.");
  }

  const normalized = JSON.stringify(doc, null, 2);
  const changed = normalized !== JSON.stringify(rawDoc, null, 2);
  return { doc, changed, notes };
};

let originalText = "";
let parsed = null;
try {
  originalText = fs.readFileSync(inputPath, "utf8");
} catch (err) {
  console.error(`Failed to read ${inputPath}: ${err.message || err}`);
  process.exit(1);
}

try {
  parsed = JSON.parse(originalText);
} catch (err) {
  console.error(`Failed to parse JSON: ${err.message || err}`);
  process.exit(1);
}

let result = null;
try {
  result = migrate(parsed);
} catch (err) {
  console.error(`Migration failed: ${err.message || err}`);
  process.exit(1);
}

if (!result.changed) {
  console.log("No schema changes required.");
  process.exit(0);
}

console.log("Schema updates detected:");
result.notes.forEach((note) => console.log(`- ${note}`));

if (!writeChanges) {
  console.log("Dry run complete. Re-run with --write to apply changes.");
  process.exit(0);
}

const stamp = new Date().toISOString().replace(/[:.]/g, "-");
const backupPath = `${inputPath}.bak-${stamp}.json`;

try {
  fs.writeFileSync(backupPath, originalText, "utf8");
  fs.writeFileSync(inputPath, `${JSON.stringify(result.doc, null, 2)}\n`, "utf8");
} catch (err) {
  console.error(`Failed to write files: ${err.message || err}`);
  process.exit(1);
}

console.log(`Backup written: ${backupPath}`);
console.log(`Updated file: ${inputPath}`);
