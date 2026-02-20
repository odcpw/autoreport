#!/usr/bin/env node
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const EXP_DIR = path.resolve(__dirname, "..");
const RESULTS_DIR = path.join(EXP_DIR, "qa-corpus", "results");

function stampIso() {
  return new Date().toISOString();
}

function findLiteralFetchUrls(source) {
  const urls = [];
  const re = /fetch\((['"`])([^'"`]+)\1/g;
  let match = null;
  while ((match = re.exec(source)) !== null) {
    urls.push(match[2]);
  }
  return urls;
}

function isAbsoluteRemote(url) {
  const value = String(url || "").trim().toLowerCase();
  return value.startsWith("http://") || value.startsWith("https://") || value.startsWith("//");
}

async function main() {
  await fs.mkdir(RESULTS_DIR, { recursive: true });

  const filesToScan = [
    path.join(EXP_DIR, "app.js"),
    path.join(EXP_DIR, "index.html"),
    path.join(EXP_DIR, "workers", "ingest.worker.js"),
    path.join(EXP_DIR, "workers", "embed.worker.js"),
    path.join(EXP_DIR, "workers", "match.worker.js"),
  ];

  const content = new Map();
  for (const file of filesToScan) {
    content.set(file, await fs.readFile(file, "utf8"));
  }

  const findings = [];

  for (const [file, text] of content.entries()) {
    const hasAbsoluteUrl = /https?:\/\//i.test(text);
    if (hasAbsoluteUrl) {
      findings.push({
        level: "error",
        file: path.relative(EXP_DIR, file),
        message: "Contains absolute http/https URL literal.",
      });
    }

    const fetchUrls = findLiteralFetchUrls(text);
    fetchUrls.forEach((url) => {
      if (isAbsoluteRemote(url)) {
        findings.push({
          level: "error",
          file: path.relative(EXP_DIR, file),
          message: `Fetch uses remote URL: ${url}`,
        });
      }
    });
  }

  const appSource = content.get(path.join(EXP_DIR, "app.js")) || "";

  if (!/allowRemoteModels\s*=\s*false/.test(appSource)) {
    findings.push({
      level: "error",
      file: "app.js",
      message: "Missing allowRemoteModels=false guard.",
    });
  }

  if (!/allowLocalModels\s*=\s*true/.test(appSource)) {
    findings.push({
      level: "error",
      file: "app.js",
      message: "Missing allowLocalModels=true guard.",
    });
  }

  if (!/local_files_only\s*:\s*true/.test(appSource)) {
    findings.push({
      level: "error",
      file: "app.js",
      message: "Missing local_files_only:true in model pipeline options.",
    });
  }

  const ok = findings.filter((f) => f.level === "error").length === 0;
  const report = {
    created_at: stampIso(),
    generator: "ai-evidence-lab-offline-check",
    ok,
    files_scanned: filesToScan.map((f) => path.relative(EXP_DIR, f)),
    findings,
  };

  const reportPath = path.join(RESULTS_DIR, "offline_mode_report.json");
  await fs.writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`, "utf8");

  if (!ok) {
    console.error("Offline verification failed. See qa-corpus/results/offline_mode_report.json");
    process.exitCode = 2;
    return;
  }

  console.log("Offline verification passed.");
}

main().catch((err) => {
  console.error("Offline verification crashed:", err);
  process.exitCode = 1;
});
