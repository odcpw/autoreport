(() => {
  const createHandler = (ctx, deps) => async () => {
    const { runtime, debug, setStatus, state } = ctx;
    const { renderRows, saveSidecar } = deps;
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
      const colId = headerRowIndex >= 0 ? findCol("nr") : 0;
      const colYes = headerRowIndex >= 0 ? findCol("ja") : 3;
      const colNo = headerRowIndex >= 0 ? findCol("nein") : 4;
      const colComment = headerRowIndex >= 0 ? findCol("bemerk") : 5;
      const pickCol = (...labels) => {
        for (const label of labels) {
          const idx = findCol(label);
          if (idx >= 0) return idx;
        }
        return -1;
      };
      const colEvidence = headerRowIndex >= 0
        ? pickCol("nachweis", "beleg", "evidence", "document")
        : -1;
      const dataRows = headerRowIndex >= 0 ? rows.slice(headerRowIndex + 1) : rows;
      const answerMap = new Map();

      dataRows.forEach((row) => {
        const rawId = String(row[colId] || "").trim();
        if (!rawId) return;
        const normalized = rawId.replace(/[.\s]+$/g, "").toLowerCase();
        if (!/^[0-9]/.test(normalized)) return;
        const yesVal = String(row[colYes] || "").trim().toLowerCase();
        const noVal = String(row[colNo] || "").trim().toLowerCase();
        const comment = String(row[colComment] || "").trim();
        const evidence = colEvidence >= 0 ? String(row[colEvidence] || "").trim() : "";
        let answer = null;
        if (yesVal === "x" || yesVal === "ja") answer = 1;
        if (noVal === "x" || noVal === "nein") answer = 0;
        answerMap.set(normalized, { answer, comment, evidence });
      });

      const idMap = new Map();
      state.project.chapters.forEach((chapter) => {
        chapter.rows.forEach((row) => {
          if (row.kind === "section") return;
          row.customer = row.customer || { items: [] };
          row.customer.items = row.customer.items || [];
          row.customer.items.forEach((item) => {
            const key = String(item.id || "").trim().toLowerCase();
            if (key) idMap.set(key, item);
            const original = String(item.originalId || "").trim().toLowerCase();
            if (original) idMap.set(original, item);
          });
        });
      });

      let applied = 0;
      answerMap.forEach((payload, key) => {
        const item = idMap.get(key);
        if (!item) return;
        if (payload.answer === 0 || payload.answer === 1) {
          item.answer = payload.answer;
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
