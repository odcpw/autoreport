(() => {
  const pickFolderBtn = document.getElementById("pick-folder");
  const loadSidecarBtn = document.getElementById("load-sidecar");
  const saveSidecarBtn = document.getElementById("save-sidecar");
  const loadSeedsBtn = document.getElementById("load-seeds");
  const saveLogBtn = document.getElementById("save-log");
  const statusEl = document.getElementById("status");
  const chapterListEl = document.getElementById("chapter-list");
  const chapterTitleEl = document.getElementById("chapter-title");
  const rowsEl = document.getElementById("rows");

  const filterHideYesEl = document.getElementById("filter-hide-yes");
  const filterIncludeOnlyEl = document.getElementById("filter-include-only");
  const filterDoneOnlyEl = document.getElementById("filter-done-only");

  let dirHandle = null;
  const debug = window.AutoReportDebug || {
    logLine: () => {},
    saveLog: async () => ({ location: "none", filename: "" }),
  };

  const defaultProject = {
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
        title: { de: "Leitbild" },
        rows: [
          {
            id: "1.1.1",
            titleOverride: "Unternehmensleitbild",
            master: {
              finding: "Das Unternehmen verfuegt nicht ueber ein Leitbild.",
              levels: {
                "1": "Eine Sicherheitscharta als ersten Schritt etablieren.",
                "2": "Die vorhandenen Werte in die Fuhrung integrieren.",
                "3": "Die dokumentierte Charta verbreiten und leben.",
                "4": "Die Charta als aktives Fuehrungsinstrument nutzen.",
              },
            },
            customer: {
              answer: 1,
              remark: "Leitbild vorhanden",
              items: [
                {
                  id: "1.1.1",
                  question: "Gibt es ein Sicherheitsleitbild?",
                  collapsedId: "1.1.1",
                },
              ],
            },
            workstate: {
              selectedLevel: 2,
              includeFinding: true,
              includeRecommendation: true,
              done: false,
              useFindingOverride: false,
              findingOverride: "",
              useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
              levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
            },
          },
          {
            id: "1.1.2",
            titleOverride: "Strategie",
            master: {
              finding: "Es fehlt eine dokumentierte Sicherheitsstrategie.",
              levels: {
                "1": "Strategie-Grundsaetze definieren.",
                "2": "Strategie in Ziele uebersetzen.",
                "3": "Strategie regelmaessig pruefen.",
                "4": "Strategie in allen Bereichen verankern.",
              },
            },
            customer: {
              answer: 0,
              remark: "",
              items: [
                {
                  id: "1.1.2",
                  question: "Gibt es eine dokumentierte Sicherheitsstrategie?",
                  collapsedId: "1.1.2",
                },
              ],
            },
            workstate: {
              selectedLevel: 3,
              includeFinding: true,
              includeRecommendation: true,
              done: true,
              useFindingOverride: true,
              findingOverride: "Eine Strategie besteht, wird aber nicht aktiv kommuniziert.",
              useLevelOverride: { "1": false, "2": true, "3": false, "4": false },
              levelOverrides: { "1": "", "2": "Strategie sichtbar machen.", "3": "", "4": "" },
            },
          },
        ],
      },
      {
        id: "4.6",
        title: { de: "Feldbeobachtungen" },
        rows: [
          {
            id: "4.6.1",
            titleOverride: "Regale",
            master: {
              finding: "Regale sind nicht gegen Kippen gesichert.",
              levels: {
                "1": "Regale sichern und Sichtkontrolle definieren.",
                "2": "Regalinspektionen regelmaessig durchfuehren.",
                "3": "Sicherheitschecks dokumentieren.",
                "4": "Regalmanagement im Sicherheitsprogramm verankern.",
              },
            },
            customer: {
              answer: 0,
              remark: "",
              items: [
                {
                  id: "4.6.1",
                  question: "Sind Regale gegen Kippen gesichert?",
                  collapsedId: "4.6.1",
                },
              ],
            },
            workstate: {
              selectedLevel: 1,
              includeFinding: true,
              includeRecommendation: true,
              done: false,
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

  const state = {
    project: structuredClone(defaultProject),
    selectedChapterId: defaultProject.chapters[0].id,
    filters: {
      hideYes: false,
      includeOnly: false,
      doneOnly: false,
    },
  };

  const setStatus = (message) => {
    statusEl.textContent = message;
    debug.logLine("info", message);
  };

  const readJsonFromHttp = async (path) => {
    const response = await fetch(path);
    if (!response.ok) {
      throw new Error(`Failed to fetch ${path}`);
    }
    return response.json();
  };

  const autoLoadSeeds = async () => {
    if (window.location.protocol !== "http:" && window.location.protocol !== "https:") {
      debug.logLine("info", "Seed auto-load skipped (not on http).");
      return;
    }
    try {
      const selbst = await readJsonFromHttp("/data/seed/selbstbeurteilung_ids.json");
      const library = await readJsonFromHttp("/data/seed/library_master.json");
      state.project = buildProjectFromSeeds(selbst, library);
      state.selectedChapterId = state.project.chapters[0]?.id || "";
      render();
      setStatus("Loaded seed data from /data/seed.");
      debug.logLine("info", "Auto-loaded seed data from /data/seed.");
    } catch (err) {
      debug.logLine("warn", `Seed auto-load skipped: ${err.message}`);
    }
  };

  const ensureFsAccess = () => {
    if (!window.showDirectoryPicker) {
      setStatus("File System Access API is not available in this browser.");
      pickFolderBtn.disabled = true;
      return false;
    }
    return true;
  };

  const enableActions = () => {
    const enabled = !!dirHandle;
    loadSidecarBtn.disabled = !enabled;
    saveSidecarBtn.disabled = !enabled;
    loadSeedsBtn.disabled = !enabled;
  };

  const getChapterTitle = (chapter) => {
    if (!chapter) return "";
    if (typeof chapter.title === "string") return chapter.title;
    if (chapter.title && chapter.title.de) return chapter.title.de;
    return chapter.id || "";
  };

  const toText = (value) => {
    if (Array.isArray(value)) return value.join("\n");
    if (value == null) return "";
    return String(value);
  };

  const ensureWorkstateDefaults = (row) => {
    if (!row.workstate) row.workstate = {};
    const ws = row.workstate;
    if (ws.selectedLevel == null) ws.selectedLevel = 1;
    if (ws.includeFinding == null) ws.includeFinding = true;
    if (ws.includeRecommendation == null) ws.includeRecommendation = true;
    if (ws.done == null) ws.done = false;
    if (!ws.useFindingOverride) ws.useFindingOverride = false;
    if (!ws.findingOverride) ws.findingOverride = "";
    ws.useLevelOverride = ws.useLevelOverride || { "1": false, "2": false, "3": false, "4": false };
    ws.levelOverrides = ws.levelOverrides || { "1": "", "2": "", "3": "", "4": "" };
  };

  const buildProjectFromSeeds = (selbstData, libraryData) => {
    const libraryMap = new Map(
      (libraryData.entries || []).map((entry) => [entry.id, entry])
    );
    const groups = new Map();

    (selbstData.items || []).forEach((item) => {
      const collapsedId = item.collapsedId || item.groupId || item.id;
      if (!collapsedId) return;
      const chapterId = String(collapsedId).split(".")[0];
      const group = groups.get(collapsedId) || {
        id: collapsedId,
        chapterId,
        items: [],
        chapterLabel: item.chapterLabel || "",
      };
      group.items.push(item);
      if (!group.chapterLabel && item.chapterLabel) {
        group.chapterLabel = item.chapterLabel;
      }
      groups.set(collapsedId, group);
    });

    const chaptersById = new Map();
    groups.forEach((group) => {
      const chapterId = group.chapterId || "0";
      const chapter = chaptersById.get(chapterId) || {
        id: chapterId,
        title: { de: group.chapterLabel || `Kapitel ${chapterId}` },
        rows: [],
      };
      const master = libraryMap.get(group.id) || null;
      chapter.rows.push({
        id: group.id,
        titleOverride: group.items[0]?.question || "",
        master,
        customer: {
          answer: null,
          remark: "",
          items: group.items,
        },
        workstate: {
          selectedLevel: 1,
          includeFinding: true,
          includeRecommendation: true,
          done: false,
          useFindingOverride: false,
          findingOverride: "",
          useLevelOverride: { "1": false, "2": false, "3": false, "4": false },
          levelOverrides: { "1": "", "2": "", "3": "", "4": "" },
        },
      });
      chaptersById.set(chapterId, chapter);
    });

    const chapters = Array.from(chaptersById.values()).sort((a, b) => {
      const aNum = Number(a.id);
      const bNum = Number(b.id);
      if (Number.isNaN(aNum) || Number.isNaN(bNum)) return String(a.id).localeCompare(String(b.id));
      return aNum - bNum;
    });
    chapters.forEach((chapter) => {
      chapter.rows.sort((a, b) => String(a.id).localeCompare(String(b.id)));
    });

    return {
      meta: {
        projectId: "seed-import",
        company: "",
        locale: "de-CH",
        author: "",
        createdAt: new Date().toISOString(),
      },
      chapters,
    };
  };

  const ensureWorkstate = (row) => {
    ensureWorkstateDefaults(row);
  };

  const getFindingText = (row) => {
    const ws = row.workstate;
    if (ws.useFindingOverride && ws.findingOverride) return ws.findingOverride;
    return toText(row.master?.finding);
  };

  const getRecommendationText = (row, level) => {
    const ws = row.workstate;
    const levelKey = String(level);
    if (ws.useLevelOverride?.[levelKey] && ws.levelOverrides?.[levelKey]) {
      return ws.levelOverrides[levelKey];
    }
    return toText(row.master?.levels?.[levelKey]);
  };

  const renderChapterList = () => {
    chapterListEl.innerHTML = "";
    state.project.chapters.forEach((chapter) => {
      const button = document.createElement("button");
      button.textContent = `${chapter.id} ${getChapterTitle(chapter)}`;
      button.className = chapter.id === state.selectedChapterId ? "active" : "";
      button.addEventListener("click", () => {
        state.selectedChapterId = chapter.id;
        render();
      });
      chapterListEl.appendChild(button);
    });
  };

  const renderRows = () => {
    rowsEl.innerHTML = "";
    const chapter = state.project.chapters.find((c) => c.id === state.selectedChapterId);
    if (!chapter) return;
    chapterTitleEl.textContent = `${chapter.id} ${getChapterTitle(chapter)}`;

    chapter.rows.forEach((row) => {
      ensureWorkstate(row);
      const ws = row.workstate;

      if (state.filters.hideYes && row.customer?.answer === 1) return;
      if (state.filters.includeOnly && !ws.includeFinding) return;
      if (state.filters.doneOnly && !ws.done) return;

      const card = document.createElement("div");
      card.className = "row-card";

      const header = document.createElement("div");
      header.className = "row-header";
      const meta = document.createElement("div");
      meta.className = "row-meta";
      meta.textContent = `${row.id} ${row.titleOverride || ""}`.trim();
      const badge = document.createElement("span");
      badge.className = "badge";
      if (row.customer?.answer === 1) {
        badge.textContent = "Answer: Yes";
      } else if (row.customer?.answer === 0) {
        badge.textContent = "Answer: No";
      } else {
        badge.textContent = "Answer: —";
      }
      header.appendChild(meta);
      header.appendChild(badge);
      card.appendChild(header);

      const selfItems = row.customer?.items || [];
      if (selfItems.length) {
        const details = document.createElement("details");
        details.className = "self-details";
        const summary = document.createElement("summary");
        summary.textContent = `Selbstbeurteilung (${selfItems.length})`;
        details.appendChild(summary);
        const list = document.createElement("ul");
        list.className = "self-list";
        selfItems.forEach((item) => {
          const li = document.createElement("li");
          li.textContent = `${item.id} — ${item.question || ""}`.trim();
          list.appendChild(li);
        });
        details.appendChild(list);
        card.appendChild(details);
      }

      const checkboxRow = document.createElement("div");
      checkboxRow.className = "checkbox-row";
      const includeLabel = document.createElement("label");
      const includeCheckbox = document.createElement("input");
      includeCheckbox.type = "checkbox";
      includeCheckbox.checked = ws.includeFinding;
      includeCheckbox.addEventListener("change", () => {
        ws.includeFinding = includeCheckbox.checked;
        renderRows();
      });
      includeLabel.appendChild(includeCheckbox);
      includeLabel.appendChild(document.createTextNode("Include"));

      const doneLabel = document.createElement("label");
      const doneCheckbox = document.createElement("input");
      doneCheckbox.type = "checkbox";
      doneCheckbox.checked = ws.done;
      doneCheckbox.addEventListener("change", () => {
        ws.done = doneCheckbox.checked;
        renderRows();
      });
      doneLabel.appendChild(doneCheckbox);
      doneLabel.appendChild(document.createTextNode("Done"));

      checkboxRow.appendChild(includeLabel);
      checkboxRow.appendChild(doneLabel);
      card.appendChild(checkboxRow);

      const findingField = document.createElement("div");
      findingField.className = "field";
      const findingLabel = document.createElement("label");
      findingLabel.textContent = "Finding";
      const findingToggle = document.createElement("label");
      findingToggle.className = "checkbox-row";
      const findingCheckbox = document.createElement("input");
      findingCheckbox.type = "checkbox";
      findingCheckbox.checked = ws.useFindingOverride;
      findingCheckbox.addEventListener("change", () => {
        ws.useFindingOverride = findingCheckbox.checked;
        if (ws.useFindingOverride && !ws.findingOverride) {
          ws.findingOverride = toText(row.master?.finding);
        }
        renderRows();
      });
      findingToggle.appendChild(findingCheckbox);
      findingToggle.appendChild(document.createTextNode("Override"));

      const findingArea = document.createElement("textarea");
      findingArea.value = getFindingText(row);
      findingArea.disabled = !ws.useFindingOverride;
      findingArea.addEventListener("input", () => {
        ws.findingOverride = findingArea.value;
      });

      findingField.appendChild(findingLabel);
      findingField.appendChild(findingToggle);
      findingField.appendChild(findingArea);
      card.appendChild(findingField);

      const recField = document.createElement("div");
      recField.className = "field";
      const recLabel = document.createElement("label");
      recLabel.textContent = "Recommendation";
      const controls = document.createElement("div");
      controls.className = "controls";

      const levelSelect = document.createElement("select");
      levelSelect.className = "level-select";
      [1, 2, 3, 4].forEach((level) => {
        const option = document.createElement("option");
        option.value = String(level);
        option.textContent = `Level ${level}`;
        if (ws.selectedLevel === level) option.selected = true;
        levelSelect.appendChild(option);
      });
      levelSelect.addEventListener("change", () => {
        ws.selectedLevel = Number(levelSelect.value);
        renderRows();
      });

      const overrideLabel = document.createElement("label");
      const overrideCheckbox = document.createElement("input");
      overrideCheckbox.type = "checkbox";
      const levelKey = String(ws.selectedLevel);
      overrideCheckbox.checked = !!ws.useLevelOverride?.[levelKey];
      overrideCheckbox.addEventListener("change", () => {
        ws.useLevelOverride[levelKey] = overrideCheckbox.checked;
        if (ws.useLevelOverride[levelKey] && !ws.levelOverrides[levelKey]) {
          ws.levelOverrides[levelKey] = toText(row.master?.levels?.[levelKey]);
        }
        renderRows();
      });
      overrideLabel.appendChild(overrideCheckbox);
      overrideLabel.appendChild(document.createTextNode("Override"));

      controls.appendChild(levelSelect);
      controls.appendChild(overrideLabel);

      const recArea = document.createElement("textarea");
      recArea.value = getRecommendationText(row, ws.selectedLevel);
      recArea.disabled = !ws.useLevelOverride?.[levelKey];
      recArea.addEventListener("input", () => {
        ws.levelOverrides[levelKey] = recArea.value;
      });

      recField.appendChild(recLabel);
      recField.appendChild(controls);
      recField.appendChild(recArea);
      card.appendChild(recField);

      const children = row.master?.children || [];
      if (children.length) {
        const childDetails = document.createElement("details");
        childDetails.className = "child-details";
        const childSummary = document.createElement("summary");
        childSummary.textContent = `Combined sub-blocks (${children.length})`;
        childDetails.appendChild(childSummary);

        children.forEach((child) => {
          const childCard = document.createElement("div");
          childCard.className = "child-card";
          const childHeader = document.createElement("div");
          childHeader.className = "child-header";
          const childTitle = document.createElement("div");
          childTitle.textContent = child.idRange || "Sub-block";
          const childIds = document.createElement("div");
          childIds.className = "child-ids";
          childIds.textContent = child.ids?.length ? child.ids.join(", ") : "";
          childHeader.appendChild(childTitle);
          childHeader.appendChild(childIds);
          childCard.appendChild(childHeader);

          const childFinding = document.createElement("div");
          childFinding.className = "child-finding";
          childFinding.textContent = child.finding || "";
          childCard.appendChild(childFinding);

          const childRec = document.createElement("div");
          childRec.className = "child-rec";
          const childLevelText = toText(child.levels?.[levelKey]);
          childRec.textContent = childLevelText || "";
          childCard.appendChild(childRec);
          childDetails.appendChild(childCard);
        });

        card.appendChild(childDetails);
      }

      rowsEl.appendChild(card);
    });
  };

  const render = () => {
    renderChapterList();
    renderRows();
  };

  pickFolderBtn.addEventListener("click", async () => {
    if (!ensureFsAccess()) return;
    try {
      dirHandle = await window.showDirectoryPicker();
      enableActions();
      setStatus(`Selected folder: ${dirHandle.name}`);
      debug.logLine("info", `Selected folder: ${dirHandle.name}`);
    } catch (err) {
      setStatus(`Folder pick canceled or failed: ${err.message}`);
      debug.logLine("warn", `Folder pick canceled or failed: ${err.message}`);
    }
  });

  loadSidecarBtn.addEventListener("click", async () => {
    if (!dirHandle) return;
    try {
      const handle = await dirHandle.getFileHandle("project_sidecar.json");
      const file = await handle.getFile();
      const text = await file.text();
      state.project = JSON.parse(text);
      state.selectedChapterId = state.project.chapters[0]?.id || "";
      render();
      setStatus("Loaded project_sidecar.json");
      debug.logLine("info", "Loaded project_sidecar.json");
    } catch (err) {
      state.project = structuredClone(defaultProject);
      state.selectedChapterId = state.project.chapters[0].id;
      render();
      setStatus("Sidecar not found; loaded default template.");
      debug.logLine("warn", `Sidecar not found: ${err.message || err}`);
    }
  });

  saveSidecarBtn.addEventListener("click", async () => {
    if (!dirHandle) return;
    try {
      const handle = await dirHandle.getFileHandle("project_sidecar.json", { create: true });
      const writable = await handle.createWritable();
      await writable.write(JSON.stringify(state.project, null, 2));
      await writable.close();
      setStatus("Saved project_sidecar.json");
      debug.logLine("info", "Saved project_sidecar.json");
    } catch (err) {
      setStatus(`Save failed: ${err.message}`);
      debug.logLine("error", `Save failed: ${err.message}`);
    }
  });

  const readJsonFromPath = async (rootHandle, pathParts) => {
    let currentHandle = rootHandle;
    for (let i = 0; i < pathParts.length - 1; i += 1) {
      currentHandle = await currentHandle.getDirectoryHandle(pathParts[i]);
    }
    const fileHandle = await currentHandle.getFileHandle(pathParts[pathParts.length - 1]);
    const file = await fileHandle.getFile();
    const text = await file.text();
    return JSON.parse(text);
  };

  loadSeedsBtn.addEventListener("click", async () => {
    if (!dirHandle) return;
    try {
      const selbst = await readJsonFromPath(dirHandle, ["data", "seed", "selbstbeurteilung_ids.json"]);
      const library = await readJsonFromPath(dirHandle, ["data", "seed", "library_master.json"]);
      state.project = buildProjectFromSeeds(selbst, library);
      state.selectedChapterId = state.project.chapters[0]?.id || "";
      render();
      setStatus("Loaded seed data from data/seed.");
      debug.logLine("info", "Loaded seed data from data/seed.");
    } catch (err) {
      setStatus(`Seed load failed: ${err.message}`);
      debug.logLine("error", `Seed load failed: ${err.message}`);
    }
  });

  saveLogBtn.addEventListener("click", async () => {
    try {
      const result = await debug.saveLog({
        suggestedName: `mini-editor-log-${new Date().toISOString().replace(/[:.]/g, "-")}.txt`,
        dirHandle,
      });
      setStatus(`Saved log (${result.location}): ${result.filename}`);
    } catch (err) {
      setStatus(`Log save failed: ${err.message}`);
    }
  });

  filterHideYesEl.addEventListener("change", () => {
    state.filters.hideYes = filterHideYesEl.checked;
    renderRows();
  });

  filterIncludeOnlyEl.addEventListener("change", () => {
    state.filters.includeOnly = filterIncludeOnlyEl.checked;
    renderRows();
  });

  filterDoneOnlyEl.addEventListener("change", () => {
    state.filters.doneOnly = filterDoneOnlyEl.checked;
    renderRows();
  });

  ensureFsAccess();
  render();
  autoLoadSeeds();
})();
