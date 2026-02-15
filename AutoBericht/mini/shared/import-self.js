(() => {
  const createHandler = (ctx, deps) => async () => {
    const { runtime, debug, setStatus, state } = ctx;
    const { renderRows, saveSidecar } = deps;
    const stateHelpers = window.AutoBerichtState || {};
    const compareIdSegments = stateHelpers.compareIdSegments || ((a, b) => String(a || "").localeCompare(String(b || ""), "de", { numeric: true }));
    const normalizeId = (raw) => String(raw || "")
      .trim()
      .toLowerCase()
      // Some Selbstbeurteilung sources contain IDs like `3.2.4. a`.
      // Collapse whitespace so the import can still map answers correctly.
      .replace(/\s+/g, "")
      .replace(/\.+$/g, "");

    const toGroupId = (normalizedId) => {
      const match = String(normalizedId || "").match(/^(\d+(?:\.\d+)*)(?:\.[a-z])$/i);
      return match ? match[1] : normalizedId;
    };
    if (!runtime.dirHandle) return;
    if (!window.showOpenFilePicker || !window.XLSX) {
      setStatus("File picker or SheetJS not available in this browser.");
      return;
    }
    try {
      const [fileHandle] = await window.showOpenFilePicker({
        types: [
          {
            description: "Excel workbook",
            accept: {
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [
                ".xlsx",
                ".xlsm",
              ],
            },
          },
        ],
        multiple: false,
      });
      const file = await fileHandle.getFile();
      const buffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(buffer, { type: "array" });
      const sheetName = workbook.SheetNames.find((name) => {
        const lower = name.toLowerCase();
        return lower.includes("selbstbeurteilung")
          || lower.includes("autoévaluation")
          || lower.includes("autoevaluation")
          || lower.includes("autovalutazione");
      }) || workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
      const headerRowIndex = rows.findIndex((row) => String(row[0] || "").toLowerCase().includes("nr"));
      const headerRow = headerRowIndex >= 0 ? rows[headerRowIndex] : [];
      const findCol = (label) => headerRow.findIndex((cell) => String(cell || "").toLowerCase().includes(label));
      const pickCol = (...labels) => {
        for (const label of labels) {
          const idx = findCol(label);
          if (idx >= 0) return idx;
        }
        return -1;
      };
      const colId = headerRowIndex >= 0 ? findCol("nr") : 0;
      const colQuestion = headerRowIndex >= 0 ? (() => {
        const idx = pickCol("frage", "question", "domanda");
        return idx >= 0 ? idx : -1;
      })() : -1;
      const colYes = headerRowIndex >= 0 ? findCol("ja") : 3;
      const colNo = headerRowIndex >= 0 ? findCol("nein") : 4;
      const colComment = headerRowIndex >= 0 ? findCol("bemerk") : 5;
      const colEvidence = headerRowIndex >= 0
        ? pickCol("nachweis", "beleg", "evidence", "document")
        : -1;
      const dataRows = headerRowIndex >= 0 ? rows.slice(headerRowIndex + 1) : rows;
      const answerMap = new Map();

      dataRows.forEach((row) => {
        const rawId = String(row[colId] || "").trim();
        if (!rawId) return;
        const normalized = normalizeId(rawId);
        if (!/^[0-9]/.test(normalized)) return;
        const question = colQuestion >= 0 ? String(row[colQuestion] || "").trim() : "";
        const yesVal = String(row[colYes] || "").trim().toLowerCase();
        const noVal = String(row[colNo] || "").trim().toLowerCase();
        const comment = String(row[colComment] || "").trim();
        const evidence = colEvidence >= 0 ? String(row[colEvidence] || "").trim() : "";
        let answer = null;
        if (yesVal === "x" || yesVal === "ja") answer = 1;
        if (noVal === "x" || noVal === "nein") answer = 0;
        answerMap.set(normalized, {
          originalId: rawId,
          question,
          answer,
          comment,
          evidence,
        });
      });

      const idMap = new Map();
      const rowMap = new Map();
      state.project.chapters.forEach((chapter) => {
        chapter.rows.forEach((row) => {
          if (row.kind === "section") return;
          rowMap.set(normalizeId(row.id), row);
          row.customer = row.customer || { items: [] };
          row.customer.items = row.customer.items || [];
          row.customer.items.forEach((item) => {
            const key = normalizeId(item.id);
            if (key) idMap.set(key, item);
            const original = normalizeId(item.originalId);
            if (original) idMap.set(original, item);
          });
        });
      });

      const ensureSelfItem = (id, payload) => {
        const groupId = toGroupId(id);
        const row = rowMap.get(normalizeId(groupId));
        if (!row) return null;
        row.customer = row.customer || { items: [] };
        row.customer.items = row.customer.items || [];
        const existing = row.customer.items.find((it) => normalizeId(it.id) === id);
        if (existing) return existing;

        const base = row.customer.items[0] || {};
        const seed = {
          // Keep structural fields from siblings if present (chapter label, section label, etc.)
          chapter: base.chapter,
          chapterLabel: base.chapterLabel,
          sectionLabel: base.sectionLabel,
          id,
          groupId,
          collapsedId: groupId,
          question: "",
          answer: null,
          comment: "",
          evidence: "",
        };
        if (payload?.originalId) seed.originalId = payload.originalId;
        if (payload?.question) seed.question = payload.question;
        if (payload?.answer === 0 || payload?.answer === 1) seed.answer = payload.answer;
        if (payload?.comment) seed.comment = payload.comment;
        if (payload?.evidence) seed.evidence = payload.evidence;

        row.customer.items.push(seed);
        row.customer.items.sort((a, b) => compareIdSegments(a.id, b.id));
        if (row.customer.items.length > 1 && row.customer.items[0]?.question) {
          // For tiroir questions, the first sub-question carries the full stem.
          row.titleOverride = row.customer.items[0].question;
        }

        // Update lookup tables so later cells can map to the newly inserted item.
        idMap.set(id, seed);
        if (seed.originalId) idMap.set(normalizeId(seed.originalId), seed);
        return seed;
      };

      let applied = 0;
      answerMap.forEach((payload, key) => {
        let item = idMap.get(key);
        if (!item) {
          item = ensureSelfItem(key, payload);
        }
        if (!item) return;
        if (payload.answer === 0 || payload.answer === 1) {
          item.answer = payload.answer;
        }
        if (payload.question && !item.question) {
          item.question = payload.question;
        }
        if (payload.comment) {
          item.comment = payload.comment;
        }
        if (payload.evidence) {
          item.evidence = payload.evidence;
        }
        applied += 1;
      });

      setStatus(`Imported self-assessment answers (${applied}).`);
      debug.logLine("info", `Imported self-assessment answers (${applied}).`);
      renderRows();
      await saveSidecar();
    } catch (err) {
      setStatus(`Import failed: ${err.message}`);
      debug.logLine("error", `Import failed: ${err.message || err}`);
    }
  };

  window.AutoBerichtImportSelf = { createHandler };
})();
