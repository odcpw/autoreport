(() => {
  const createHandler = (ctx, deps) => async () => {
    const { runtime, debug, setStatus, state } = ctx;
    const { t, tf } = ctx.i18n;
    const { renderRows, saveSidecar } = deps;
    const stateHelpers = window.AutoBerichtState || {};
    const assessment = window.AutoBerichtSelfAssessment || {};
    const compareIdSegments = stateHelpers.compareIdSegments || ((a, b) => String(a || "").localeCompare(String(b || ""), "de", { numeric: true }));
    const normalizeId = assessment.normalizeId || ((raw) => String(raw || "").trim().toLowerCase());

    const toGroupId = (normalizedId) => {
      const match = String(normalizedId || "").match(/^(\d+(?:\.\d+)*)(?:\.[a-z])$/i);
      return match ? match[1] : normalizedId;
    };
    if (!runtime.dirHandle) return;
    if (!window.showOpenFilePicker || !window.XLSX) {
      setStatus(t("status_import_capability_missing"));
      return;
    }
    if (!assessment.findAssessmentSheetName || !assessment.parseRows) {
      setStatus(t("status_validator_missing"));
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
      const sheetName = assessment.findAssessmentSheetName(workbook.SheetNames);
      const sheet = workbook.Sheets[sheetName];
      const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
      const parsed = assessment.parseRows(rows);

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
      const answerMap = new Map(parsed.entries.map((entry) => [entry.id, entry]));

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
        if (payload.answer === 0 || payload.answer === 1 || payload.comment || payload.evidence) {
          applied += 1;
        }
      });

      if (applied === 0) {
        throw new Error("No meaningful self-assessment rows matched the current project.");
      }

      setStatus(tf("status_self_assessment_imported", "Imported self-assessment answers ({count}).", { count: applied }));
      debug.logLine("info", `Imported self-assessment answers (${applied}).`);
      renderRows();
      await saveSidecar();
    } catch (err) {
      setStatus(tf("status_import_failed", "Import failed: {error}", { error: err.message || err }));
      debug.logLine("error", `Import failed: ${err.message || err}`);
    }
  };

  window.AutoBerichtImportSelf = { createHandler };
})();
