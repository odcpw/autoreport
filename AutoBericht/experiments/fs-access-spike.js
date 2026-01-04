(() => {
  const pickFolderBtn = document.getElementById("pick-folder");
  const loadSidecarBtn = document.getElementById("load-sidecar");
  const saveSidecarBtn = document.getElementById("save-sidecar");
  const writeXlsxBtn = document.getElementById("write-xlsx");
  const statusEl = document.getElementById("status");
  const sidecarEl = document.getElementById("sidecar");

  let dirHandle = null;

  const defaultSidecar = {
    meta: {
      projectId: "2026-TEST-001",
      company: "ACME AG",
      locale: "de-CH",
      author: "consultant@example.com",
      createdAt: new Date().toISOString(),
    },
    chapters: [
      {
        id: "1",
        title: "Leitbild",
        rows: [
          {
            id: "1.1.1",
            master: {
              finding: "Das Unternehmen verfuegt nicht ueber ein Leitbild.",
              levels: {
                "1": "Empfehlung Stufe 1",
                "2": "Empfehlung Stufe 2",
                "3": "Empfehlung Stufe 3",
                "4": "Empfehlung Stufe 4",
              },
            },
            customer: { answer: 0, remark: "" },
            workstate: {
              selectedLevel: 2,
              includeFinding: true,
              includeRecommendation: true,
              useFindingOverride: false,
              findingOverride: "",
              useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
              levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
            },
          },
        ],
      },
    ],
  };

  const setStatus = (message) => {
    statusEl.textContent = message;
  };

  const enableActions = () => {
    const enabled = !!dirHandle;
    loadSidecarBtn.disabled = !enabled;
    saveSidecarBtn.disabled = !enabled;
    writeXlsxBtn.disabled = !enabled;
  };

  const ensureFsAccess = () => {
    if (!window.showDirectoryPicker) {
      setStatus("File System Access API is not available in this browser.");
      pickFolderBtn.disabled = true;
      return false;
    }
    return true;
  };

  const readFileText = async (fileHandle) => {
    const file = await fileHandle.getFile();
    return await file.text();
  };

  const getSidecarHandle = async (options = {}) => {
    if (!dirHandle) return null;
    return await dirHandle.getFileHandle("project_sidecar.json", options);
  };

  const parseSidecar = () => {
    try {
      return JSON.parse(sidecarEl.value);
    } catch (err) {
      setStatus(`Invalid JSON: ${err.message}`);
      return null;
    }
  };

  const buildWorkbook = (data) => {
    const wb = XLSX.utils.book_new();
    const metaRows = Object.entries(data.meta || {}).map(([key, value]) => ({ key, value }));
    const metaSheet = XLSX.utils.json_to_sheet(metaRows);
    XLSX.utils.book_append_sheet(wb, metaSheet, "Meta");

    const rowRows = [];
    (data.chapters || []).forEach((chapter) => {
      (chapter.rows || []).forEach((row) => {
        const finding = row.workstate?.useFindingOverride
          ? row.workstate.findingOverride
          : row.master?.finding || row.finding || "";
        rowRows.push({
          chapterId: chapter.id,
          rowId: row.id,
          finding,
          selectedLevel: row.workstate?.selectedLevel ?? "",
          includeFinding: row.workstate?.includeFinding ?? "",
          includeRecommendation: row.workstate?.includeRecommendation ?? "",
        });
      });
    });

    const rowsSheet = XLSX.utils.json_to_sheet(rowRows);
    XLSX.utils.book_append_sheet(wb, rowsSheet, "Rows");
    return wb;
  };

  pickFolderBtn.addEventListener("click", async () => {
    if (!ensureFsAccess()) return;
    try {
      dirHandle = await window.showDirectoryPicker();
      enableActions();
      setStatus(`Selected folder: ${dirHandle.name}`);
    } catch (err) {
      setStatus(`Folder pick canceled or failed: ${err.message}`);
    }
  });

  loadSidecarBtn.addEventListener("click", async () => {
    try {
      const handle = await getSidecarHandle();
      const text = await readFileText(handle);
      sidecarEl.value = text;
      setStatus("Loaded project_sidecar.json");
    } catch (err) {
      sidecarEl.value = JSON.stringify(defaultSidecar, null, 2);
      setStatus("Sidecar not found; loaded default template.");
    }
  });

  saveSidecarBtn.addEventListener("click", async () => {
    const data = parseSidecar();
    if (!data) return;
    try {
      const handle = await getSidecarHandle({ create: true });
      const writable = await handle.createWritable();
      await writable.write(JSON.stringify(data, null, 2));
      await writable.close();
      setStatus("Saved project_sidecar.json");
    } catch (err) {
      setStatus(`Save failed: ${err.message}`);
    }
  });

  writeXlsxBtn.addEventListener("click", async () => {
    const data = parseSidecar();
    if (!data) return;
    if (!window.XLSX) {
      setStatus("SheetJS not loaded.");
      return;
    }
    try {
      const wb = buildWorkbook(data);
      const buffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const handle = await dirHandle.getFileHandle("project_db.xlsx", { create: true });
      const writable = await handle.createWritable();
      await writable.write(buffer);
      await writable.close();
      setStatus("Wrote project_db.xlsx");
    } catch (err) {
      setStatus(`Excel write failed: ${err.message}`);
    }
  });

  ensureFsAccess();
})();
