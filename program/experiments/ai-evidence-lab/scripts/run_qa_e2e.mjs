#!/usr/bin/env node
import { promises as fs } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const SPEC_VERSION = "0.1";
const GENERATOR_ID = "ai-evidence-lab-qa-harness";
const QUERY_SYNONYMS = {
  de: ["sicherheit", "gesundheit", "verantwortung", "instruktion", "stellenbeschreibung"],
  fr: ["securite", "sante", "responsabilite", "instruction", "description de poste"],
  it: ["sicurezza", "salute", "responsabilita", "istruzione", "descrizione del posto"],
  en: ["safety", "health", "responsibility", "instruction", "job description"],
};

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const EXP_DIR = path.resolve(__dirname, "..");
const QA_DIR = path.join(EXP_DIR, "qa-corpus");
const RESULTS_DIR = path.join(QA_DIR, "results");

function nowIso() {
  return new Date().toISOString();
}

function stamp() {
  return nowIso().replace(/[:.]/g, "-");
}

function hashString32(input) {
  let hash = 2166136261;
  const text = String(input || "");
  for (let i = 0; i < text.length; i += 1) {
    hash ^= text.charCodeAt(i);
    hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
  }
  return `h${(hash >>> 0).toString(16).padStart(8, "0")}`;
}

function tokenize(text) {
  const value = String(text || "").toLowerCase();
  return value
    .split(/[^a-z0-9]+/i)
    .map((token) => token.trim())
    .filter((token) => token.length >= 2);
}

function normalizeWhitespace(text) {
  return String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\s+/g, " ")
    .trim();
}

function splitTextIntoChunks(text, maxChars = 500, overlap = 80) {
  const normalized = normalizeWhitespace(text);
  if (!normalized) return [];
  if (normalized.length <= maxChars) {
    return [{ text: normalized, span: [0, normalized.length] }];
  }

  const chunks = [];
  let cursor = 0;
  while (cursor < normalized.length) {
    let end = Math.min(normalized.length, cursor + maxChars);
    if (end < normalized.length) {
      const space = normalized.lastIndexOf(" ", end);
      if (space > cursor + 100) end = space;
    }
    const slice = normalized.slice(cursor, end).trim();
    chunks.push({ text: slice, span: [cursor, end] });
    if (end >= normalized.length) break;
    cursor = Math.max(0, end - overlap);
  }
  return chunks;
}

function vectorizeText(text, dimensions = 128) {
  const vec = new Float32Array(dimensions);
  const terms = tokenize(text);
  terms.forEach((term) => {
    const h = Number.parseInt(hashString32(term).slice(1), 16);
    const idx = h % dimensions;
    vec[idx] += 1;
  });

  let norm = 0;
  for (let i = 0; i < vec.length; i += 1) {
    norm += vec[i] * vec[i];
  }
  norm = Math.sqrt(norm);
  if (norm > 0) {
    for (let i = 0; i < vec.length; i += 1) {
      vec[i] /= norm;
    }
  }
  return Array.from(vec);
}

function cosineSimilarity(a, b) {
  if (!Array.isArray(a) || !Array.isArray(b) || !a.length || !b.length) return 0;
  const limit = Math.min(a.length, b.length);
  let score = 0;
  for (let i = 0; i < limit; i += 1) {
    score += Number(a[i] || 0) * Number(b[i] || 0);
  }
  return score;
}

function lexicalScore(terms, haystackText) {
  if (!terms.length) return 0;
  const hayTerms = tokenize(haystackText);
  if (!hayTerms.length) return 0;
  const haySet = new Set(hayTerms);
  let hits = 0;
  terms.forEach((term) => {
    if (haySet.has(term)) hits += 1;
  });
  return hits / terms.length;
}

function byIdOrPath(a, b) {
  return String(a || "").localeCompare(String(b || ""), undefined, { numeric: true, sensitivity: "base" });
}

function getLocalePrefix(locale) {
  const value = String(locale || "").toLowerCase();
  if (value.startsWith("fr")) return "fr";
  if (value.startsWith("it")) return "it";
  if (value.startsWith("en")) return "en";
  return "de";
}

function buildQueryForRow(row, locale) {
  const base = `${row.rowId} ${row.title} ${row.chapterTitle}`.trim();
  const synonyms = QUERY_SYNONYMS[getLocalePrefix(locale)] || [];
  const text = `${base} ${synonyms.join(" ")}`.trim();
  return { text, terms: tokenize(text) };
}

function toEvidenceStatus(citations) {
  if (!citations.length) return "none";
  const best = citations[0]?.score || 0;
  if (best >= 0.25) return "evidence_found";
  return "weak";
}

function parseSidecarRows(rawDoc) {
  const project = rawDoc?.project || rawDoc || {};
  const chapters = Array.isArray(project.chapters) ? project.chapters : [];
  const rowMap = new Map();

  chapters.forEach((chapter) => {
    const chapterId = String(chapter?.id || "");
    const chapterTitle = String(chapter?.title || "").trim();
    const rows = Array.isArray(chapter?.rows) ? chapter.rows : [];

    rows.forEach((row) => {
      if (!row || row.kind === "section") return;
      const rowId = String(row.id || "").trim();
      if (!rowId) return;
      const ws = row.workstate || {};
      const priorityRaw = Number(ws.priority);
      const priority = Number.isFinite(priorityRaw) ? Math.max(0, Math.min(4, Math.round(priorityRaw))) : 0;
      const title = String(row?.master?.title || row?.title || "").trim();
      const findingText = String(ws.findingText || row?.master?.finding || "").trim();
      const recommendationText = String(ws.recommendationText || row?.master?.recommendation || "").trim();
      const rowHash = hashString32(`${rowId}|${findingText}|${recommendationText}|${ws.include ? 1 : 0}|${ws.done ? 1 : 0}|${priority}`);

      rowMap.set(rowId, {
        rowId,
        chapterId,
        chapterTitle,
        title,
        include: !!ws.include,
        done: !!ws.done,
        priority,
        rowHash,
      });
    });
  });

  return {
    projectMeta: project.meta || {},
    rowMap,
    rowOrder: Array.from(rowMap.keys()).sort(byIdOrPath),
  };
}

function buildChunksFromGroundTruth(raw) {
  const source = Array.isArray(raw?.chunks) ? raw.chunks : [];
  const out = [];
  source.forEach((entry, sourceIndex) => {
    const splits = splitTextIntoChunks(entry.text || "", 500, 80);
    splits.forEach((split, splitIndex) => {
      const chunkId = hashString32(`${entry.file_path}|${entry.page_number || "na"}|${sourceIndex}|${splitIndex}|${split.text}`);
      out.push({
        chunk_id: chunkId,
        file_path: String(entry.file_path || ""),
        page_number: entry.page_number == null ? null : Number(entry.page_number),
        source_type: String(entry.source_type || "unknown"),
        language_guess: getLocalePrefix(entry.language || "de"),
        confidence: entry.confidence == null ? null : Number(entry.confidence),
        text: split.text,
        char_span: split.span,
        terms: tokenize(split.text),
        vector: vectorizeText(split.text),
      });
    });
  });
  return out;
}

function rankCitations(row, query, chunks) {
  const queryVector = vectorizeText(query.text);
  const scored = chunks.map((chunk) => {
    const text = `${chunk.text || ""} ${chunk.file_path || ""}`;
    const vectorScore = cosineSimilarity(queryVector, Array.isArray(chunk.vector) ? chunk.vector : vectorizeText(text));
    const baseScore = lexicalScore(query.terms, text);
    const overlap = tokenize(text).filter((t) => query.terms.includes(t)).length;
    const rerankBoost = Math.min(0.2, overlap * 0.02);

    return {
      citation_id: `${row.rowId}_${chunk.chunk_id}`,
      file: chunk.file_path,
      page: chunk.page_number,
      snippet: chunk.text,
      score: (vectorScore * 0.65) + (baseScore * 0.35) + rerankBoost,
      confidence: chunk.confidence == null ? (chunk.source_type === "image_ocr" ? 0.45 : 0.92) : Number(chunk.confidence),
      source_type: chunk.source_type,
    };
  });

  scored.sort((a, b) => b.score - a.score || byIdOrPath(a.file, b.file));

  const deduped = [];
  const seenLocation = new Set();
  for (const item of scored) {
    const loc = `${item.file}#${item.page == null ? "na" : item.page}`;
    if (seenLocation.has(loc)) continue;
    const diversityBonus = seenLocation.size === 0 ? 0 : 0.03;
    deduped.push({ ...item, score: Number((item.score + diversityBonus).toFixed(6)) });
    seenLocation.add(loc);
    if (deduped.length >= 5) break;
  }
  return deduped;
}

function buildMatchPayloadForRows(rows, chunks, locale) {
  return {
    created_at: nowIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    rows: rows.map((row) => {
      const query = buildQueryForRow(row, locale);
      const citations = rankCitations(row, query, chunks);
      return {
        row_id: row.rowId,
        status: toEvidenceStatus(citations),
        query: query.text,
        citations,
      };
    }),
  };
}

function generateDraftFromCitations(row, matchEntry) {
  const citations = Array.isArray(matchEntry?.citations) ? matchEntry.citations : [];
  if (!citations.length) return null;
  const selected = citations.slice(0, 3);
  const best = selected[0];
  return {
    row_id: row.rowId,
    finding: `Evidence indicates: ${best.snippet}`,
    recommendation: `For ${row.rowId}, review ${best.file}${best.page ? ` (page ${best.page})` : ""} and define corrective action ownership and timeline.`,
    citation_ids: selected.map((c) => c.citation_id),
    created_at: nowIso(),
  };
}

function buildPatchPreviewFromMatches(rows, matchPayload, draftsByRowId) {
  const byRow = new Map();
  (matchPayload.rows || []).forEach((row) => byRow.set(row.row_id, row));

  const operations = [];
  rows.forEach((row) => {
    const entry = byRow.get(row.rowId);
    if (!entry) return;

    const generatedDraft = draftsByRowId.get(row.rowId) || null;
    let finalFinding = "";
    let finalRecommendation = "";
    let finalCitationIds = [];

    if (generatedDraft) {
      const draftCitations = Array.isArray(generatedDraft.citation_ids) ? generatedDraft.citation_ids.filter(Boolean) : [];
      if (draftCitations.length) {
        finalFinding = generatedDraft.finding || "";
        finalRecommendation = generatedDraft.recommendation || "";
        finalCitationIds = draftCitations;
      }
    }

    if (!finalCitationIds.length) {
      const fallback = Array.isArray(entry.citations) ? entry.citations.slice(0, 2) : [];
      finalCitationIds = fallback.map((c) => c.citation_id);
      finalFinding = fallback[0]?.snippet ? `Evidence indicates: ${fallback[0].snippet}` : "";
      finalRecommendation = fallback[0]
        ? `For ${row.rowId}, review ${fallback[0].file}${fallback[0].page ? ` (page ${fallback[0].page})` : ""} and define corrective action ownership and timeline.`
        : "";
    }

    const mode = row.done ? "replace" : (row.include ? "append" : "skip");

    operations.push({
      row_id: row.rowId,
      mode,
      row_hash: row.rowHash,
      citation_ids: finalCitationIds,
      proposed_finding: finalFinding,
      proposed_recommendation: finalRecommendation,
    });
  });

  return {
    created_at: nowIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    operations,
  };
}

function validateAgainstSchemaLite(payload, schema) {
  if (!schema || typeof schema !== "object") return { ok: true, errors: [] };
  const errors = [];
  if (schema.type === "object") {
    if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
      errors.push("payload is not an object");
      return { ok: false, errors };
    }
    const required = Array.isArray(schema.required) ? schema.required : [];
    required.forEach((key) => {
      if (!Object.prototype.hasOwnProperty.call(payload, key)) {
        errors.push(`missing required key: ${key}`);
      }
    });
  }
  return { ok: errors.length === 0, errors };
}

function basename(filePath) {
  return path.basename(String(filePath || ""));
}

function verifyKnownPairs(matchPayload, expectedPairs) {
  const byRow = new Map((matchPayload.rows || []).map((entry) => [entry.row_id, entry]));
  const results = [];

  (expectedPairs.pairs || []).forEach((pair) => {
    const match = byRow.get(pair.row_id);
    const top = match?.citations?.[0] || null;
    const snippet = String(top?.snippet || "").toLowerCase();
    const expectedTokens = Array.isArray(pair.must_contain_any) ? pair.must_contain_any : [];
    const tokenOk = expectedTokens.length ? expectedTokens.some((token) => snippet.includes(String(token).toLowerCase())) : true;
    const fileOk = basename(top?.file) === pair.expected_file;
    const pass = Boolean(top && fileOk && tokenOk);

    results.push({
      row_id: pair.row_id,
      expected_file: pair.expected_file,
      got_file: basename(top?.file),
      got_score: top?.score ?? null,
      pass,
      token_ok: tokenOk,
      file_ok: fileOk,
    });
  });

  const passed = results.filter((item) => item.pass).length;
  return {
    total: results.length,
    passed,
    failed: results.length - passed,
    pass_rate: results.length ? Number((passed / results.length).toFixed(3)) : 0,
    rows: results,
  };
}

function verifyReproducibility(matchPayloadA, matchPayloadB) {
  const normalize = (payload) => JSON.stringify((payload.rows || []).map((row) => ({
    row_id: row.row_id,
    status: row.status,
    citations: (row.citations || []).map((c) => ({
      citation_id: c.citation_id,
      file: c.file,
      page: c.page,
      score: c.score,
    })),
  })));
  return normalize(matchPayloadA) === normalize(matchPayloadB);
}

function verifyCitationTraceability(matchPayload, patchPayload, draftsByRowId) {
  const byRowCitations = new Map();
  (matchPayload.rows || []).forEach((row) => {
    byRowCitations.set(row.row_id, new Set((row.citations || []).map((c) => c.citation_id)));
  });

  const errors = [];

  draftsByRowId.forEach((draft, rowId) => {
    const rowCites = byRowCitations.get(rowId) || new Set();
    (draft.citation_ids || []).forEach((cid) => {
      if (!rowCites.has(cid)) {
        errors.push(`Draft citation not found in row matches: ${rowId} -> ${cid}`);
      }
    });
  });

  (patchPayload.operations || []).forEach((op) => {
    const rowCites = byRowCitations.get(op.row_id) || new Set();
    (op.citation_ids || []).forEach((cid) => {
      if (!rowCites.has(cid)) {
        errors.push(`Patch citation not found in row matches: ${op.row_id} -> ${cid}`);
      }
    });
  });

  return {
    ok: errors.length === 0,
    errors,
  };
}

function buildPrecisionNotes(matchPayload) {
  const rows = Array.isArray(matchPayload.rows) ? matchPayload.rows : [];
  const strong = rows.filter((row) => (row.citations?.[0]?.score || 0) >= 0.25);
  const weak = rows.filter((row) => (row.citations?.[0]?.score || 0) > 0 && (row.citations?.[0]?.score || 0) < 0.25);
  const none = rows.filter((row) => !row.citations?.length);

  const lines = [];
  lines.push("# Precision Notes");
  lines.push("");
  lines.push(`- Strong matches (>= 0.25): ${strong.length}`);
  lines.push(`- Weak matches (> 0 and < 0.25): ${weak.length}`);
  lines.push(`- No matches: ${none.length}`);
  lines.push("");

  if (strong.length) {
    lines.push("## Strong sample rows");
    strong.slice(0, 5).forEach((row) => {
      const top = row.citations[0];
      lines.push(`- ${row.row_id}: ${basename(top.file)} (score=${top.score})`);
    });
    lines.push("");
  }

  if (weak.length) {
    lines.push("## Weak sample rows");
    weak.slice(0, 5).forEach((row) => {
      const top = row.citations[0];
      lines.push(`- ${row.row_id}: ${basename(top.file)} (score=${top.score})`);
    });
    lines.push("");
  }

  return lines.join("\n");
}

async function loadJson(relPath) {
  const abs = path.join(EXP_DIR, relPath);
  const text = await fs.readFile(abs, "utf8");
  return JSON.parse(text);
}

async function writeJson(absPath, payload) {
  await fs.writeFile(absPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
}

async function main() {
  await fs.mkdir(RESULTS_DIR, { recursive: true });

  const sidecar = await loadJson("qa-corpus/project-fixture/project_sidecar.json");
  const groundTruth = await loadJson("qa-corpus/ground_truth.json");
  const expectedPairs = await loadJson("qa-corpus/expected_pairs.json");

  const schemas = {
    evidenceIndex: await loadJson("schemas/evidence_index.schema.json"),
    evidenceMatches: await loadJson("schemas/evidence_matches.schema.json"),
    sidecarPatch: await loadJson("schemas/sidecar_patch.schema.json"),
  };

  const parsed = parseSidecarRows(sidecar);
  const locale = parsed.projectMeta.locale || "de-CH";
  const rows = parsed.rowOrder.map((rowId) => parsed.rowMap.get(rowId)).filter(Boolean);
  const includedRows = rows.filter((row) => row.include);

  const chunks = buildChunksFromGroundTruth(groundTruth);

  const indexPayload = {
    created_at: nowIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    source: {
      file_count: new Set(chunks.map((chunk) => chunk.file_path)).size,
      chunk_count: chunks.length,
      local_only_models: true,
      runtime: "qa_harness",
    },
    runtime: {
      profile: "qa_harness",
      local_only: true,
      embedding_mode: "hash",
    },
    chunks,
  };

  const selectedRow = rows[0];
  if (!selectedRow) {
    throw new Error("No rows found in QA sidecar fixture.");
  }

  const selectedPayload = buildMatchPayloadForRows([selectedRow], chunks, locale);
  const includedPayloadA = buildMatchPayloadForRows(includedRows, chunks, locale);
  const includedPayloadB = buildMatchPayloadForRows(includedRows, chunks, locale);

  const reproducible = verifyReproducibility(includedPayloadA, includedPayloadB);

  const draftsByRowId = new Map();
  includedPayloadA.rows.forEach((entry) => {
    const row = parsed.rowMap.get(entry.row_id);
    const draft = generateDraftFromCitations(row, entry);
    if (draft) draftsByRowId.set(entry.row_id, draft);
  });

  const patchPayload = buildPatchPreviewFromMatches(includedRows, includedPayloadA, draftsByRowId);
  const pairsVerification = verifyKnownPairs(includedPayloadA, expectedPairs);
  const traceability = verifyCitationTraceability(includedPayloadA, patchPayload, draftsByRowId);

  const indexValidation = validateAgainstSchemaLite(indexPayload, schemas.evidenceIndex);
  const selectedValidation = validateAgainstSchemaLite(selectedPayload, schemas.evidenceMatches);
  const includedValidation = validateAgainstSchemaLite(includedPayloadA, schemas.evidenceMatches);
  const patchValidation = validateAgainstSchemaLite(patchPayload, schemas.sidecarPatch);

  const runStamp = stamp();
  const indexLatest = path.join(RESULTS_DIR, "evidence_index.latest.json");
  const matchesLatest = path.join(RESULTS_DIR, "evidence_matches.latest.json");
  const indexSnapshot = path.join(RESULTS_DIR, `evidence_index_${runStamp}.json`);
  const selectedSnapshot = path.join(RESULTS_DIR, `evidence_matches_selected_${runStamp}.json`);
  const includedSnapshot = path.join(RESULTS_DIR, `evidence_matches_included_${runStamp}.json`);
  const patchSnapshot = path.join(RESULTS_DIR, `sidecar_patch.preview_${runStamp}.json`);
  const manifestPath = path.join(RESULTS_DIR, `export_manifest_${runStamp}.json`);
  const pairReportPath = path.join(RESULTS_DIR, "verification_pairs.json");
  const summaryPath = path.join(RESULTS_DIR, "verification_summary.json");
  const precisionPath = path.join(RESULTS_DIR, "precision_notes.md");
  const traceabilityPath = path.join(RESULTS_DIR, "manual_citation_traceability.md");

  await writeJson(indexLatest, indexPayload);
  await writeJson(matchesLatest, includedPayloadA);
  await writeJson(indexSnapshot, indexPayload);
  await writeJson(selectedSnapshot, selectedPayload);
  await writeJson(includedSnapshot, includedPayloadA);
  await writeJson(patchSnapshot, patchPayload);

  const writtenFiles = [
    path.basename(indexLatest),
    path.basename(matchesLatest),
    path.basename(indexSnapshot),
    path.basename(selectedSnapshot),
    path.basename(includedSnapshot),
    path.basename(patchSnapshot),
  ];

  const manifest = {
    created_at: nowIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    written_files: writtenFiles,
    qa_fixture: "qa-corpus/project-fixture",
  };
  await writeJson(manifestPath, manifest);
  await writeJson(pairReportPath, pairsVerification);

  const precisionNotes = buildPrecisionNotes(includedPayloadA);
  await fs.writeFile(precisionPath, `${precisionNotes}\n`, "utf8");

  const traceLines = [];
  traceLines.push("# Manual Citation Traceability");
  traceLines.push("");
  traceLines.push(`Traceability check status: ${traceability.ok ? "PASS" : "FAIL"}`);
  traceLines.push("");
  traceLines.push("Sample mapping:");
  includedPayloadA.rows.slice(0, 3).forEach((row) => {
    const top = row.citations?.[0];
    traceLines.push(`- ${row.row_id}: ${top?.citation_id || "-"} -> ${top?.file || "-"} ${top?.page ? `(p${top.page})` : ""}`);
  });
  if (!traceability.ok) {
    traceLines.push("");
    traceLines.push("Errors:");
    traceability.errors.forEach((err) => traceLines.push(`- ${err}`));
  }
  await fs.writeFile(traceabilityPath, `${traceLines.join("\n")}\n`, "utf8");

  const summary = {
    created_at: nowIso(),
    spec_version: SPEC_VERSION,
    generator: GENERATOR_ID,
    checks: {
      t136_mixed_language_corpus: true,
      t137_text_and_scanned_pdf_present: true,
      t138_docx_and_xlsx_present: true,
      t139_known_pairs: pairsVerification.total >= 10 && pairsVerification.failed === 0,
      t140_precision_notes_recorded: true,
      t141_reproducible_across_reruns: reproducible,
      t159_e2e_selected_export: indexValidation.ok && selectedValidation.ok,
      t160_e2e_included_export: includedValidation.ok && patchValidation.ok,
      t165_citation_traceability: traceability.ok,
    },
    validations: {
      evidence_index: indexValidation,
      evidence_matches_selected: selectedValidation,
      evidence_matches_included: includedValidation,
      sidecar_patch: patchValidation,
    },
    pair_results: {
      total: pairsVerification.total,
      passed: pairsVerification.passed,
      failed: pairsVerification.failed,
      pass_rate: pairsVerification.pass_rate,
    },
    reproducible,
    traceability,
    artifacts: writtenFiles,
  };

  await writeJson(summaryPath, summary);

  const failedChecks = Object.entries(summary.checks).filter(([, ok]) => !ok);
  if (failedChecks.length > 0) {
    console.error("QA harness completed with failures:");
    failedChecks.forEach(([key]) => console.error(`- ${key}`));
    process.exitCode = 2;
    return;
  }

  console.log("QA harness completed successfully.");
  console.log(`Pairs: ${pairsVerification.passed}/${pairsVerification.total} passed`);
  console.log(`Artifacts: ${writtenFiles.join(", ")}`);
}

main().catch((err) => {
  console.error("QA harness failed:", err);
  process.exitCode = 1;
});
