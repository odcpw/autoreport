(() => {
  const init = (ctx, deps) => {
    const { stateHelpers, normalizeHelpers } = deps;
    const { elements, state, runtime, debug, setStatus } = ctx;
    const { t } = ctx.i18n;
    const { escapeHtml = (value) => value, formatInlineMarkdown = (value) => value, markdownToHtml = (value) => value } = ctx.markdown || {};

    const {
      chapterListEl,
      chapterTitleEl,
      rowsEl,
      photoOverlayEl,
      photoOverlayClose,
      photoOverlayCloseBtn,
      photoOverlayPrevBtn,
      photoOverlayNextBtn,
      photoOverlayTitle,
      photoOverlayImage,
      checklistOverlayEl,
      checklistOverlayClose,
      checklistOverlayCloseBtn,
      checklistOverlayTitle,
      checklistOverlayHint,
      checklistFiltersEl,
      checklistListEl,
      checklistSearchEl,
      chapterPreviewModal,
      chapterPreviewTitle,
      chapterPreviewBody,
      observationsOrganizeModal,
      observationsOrganizeSearchEl,
      observationsOrganizeFiltersEl,
      observationsOrganizeSortsEl,
      observationsOrganizeStatusEl,
      observationsOrganizeListEl,
      observationsOrganizeApplyBtn,
    } = elements;

    let scheduleAutosave = () => {};

    const setScheduleAutosave = (fn) => {
      scheduleAutosave = typeof fn === "function" ? fn : () => {};
    };

    const autosizeTextarea = (textarea) => {
      if (!textarea) return;
      const minHeight = 120;
      const maxHeight = 360;
      textarea.style.height = "auto";
      const scroll = textarea.scrollHeight || minHeight;
      const nextHeight = Math.max(minHeight, Math.min(maxHeight, scroll));
      textarea.style.height = `${nextHeight}px`;
      textarea.style.overflowY = scroll > maxHeight ? "auto" : "hidden";
    };

    const buildPhotoIndex = () => {
      const reportIndex = new Map();
      const observationIndex = new Map();
      const photoDoc = runtime.sidecarDoc?.photos;
      const photos = photoDoc?.photos || {};
      state.photoRoot = photoDoc?.photoRoot || "";
      Object.entries(photos).forEach(([path, data]) => {
        const reportTags = data?.tags?.report || [];
        reportTags.forEach((tag) => {
          const key = String(tag || "").trim();
          if (!key) return;
          if (!reportIndex.has(key)) reportIndex.set(key, []);
          reportIndex.get(key).push({
            path,
            notes: data?.notes || "",
          });
        });
        const observationTags = data?.tags?.observations || [];
        observationTags.forEach((tag) => {
          const key = String(tag || "").trim();
          if (!key) return;
          if (!observationIndex.has(key)) observationIndex.set(key, []);
          observationIndex.get(key).push({
            path,
            notes: data?.notes || "",
          });
        });
      });
      reportIndex.forEach((items) => {
        items.sort((a, b) => a.path.localeCompare(b.path, "de", { numeric: true }));
      });
      observationIndex.forEach((items) => {
        items.sort((a, b) => a.path.localeCompare(b.path, "de", { numeric: true }));
      });
      state.photoIndex = {
        report: reportIndex,
        observations: observationIndex,
      };
    };

    const getPhotosForTag = (tag, kind = "report") => {
      if (!tag) return [];
      const index = state.photoIndex?.[kind];
      if (!index) return [];
      return index.get(tag) || [];
    };

    const resolvePhotoFile = async (path) => {
      if (!runtime.dirHandle) throw new Error("Project folder not selected.");
      const parts = String(path || "").split("/").filter(Boolean);
      let current = runtime.dirHandle;
      for (let i = 0; i < parts.length - 1; i += 1) {
        current = await current.getDirectoryHandle(parts[i]);
      }
      const fileHandle = await current.getFileHandle(parts[parts.length - 1]);
      return fileHandle.getFile();
    };

    const loadOverlayImage = async (path) => {
      if (!photoOverlayImage) return;
      if (state.photoOverlay.url) {
        URL.revokeObjectURL(state.photoOverlay.url);
        state.photoOverlay.url = "";
      }
      try {
        const file = await resolvePhotoFile(path);
        const url = URL.createObjectURL(file);
        state.photoOverlay.url = url;
        photoOverlayImage.src = url;
        photoOverlayImage.alt = path;
      } catch (err) {
        photoOverlayImage.removeAttribute("src");
        photoOverlayImage.alt = "Photo not available";
        setStatus(`Photo not found: ${path}`);
      }
    };

    const renderPhotoOverlay = () => {
      if (!photoOverlayEl) return;
      const { items, index, tag } = state.photoOverlay;
      if (!items.length) return;
      const current = items[index];
      const filename = current.path.split("/").pop();
      if (photoOverlayTitle) {
        photoOverlayTitle.textContent = `${tag} • Image ${index + 1} of ${items.length} • ${filename}`;
      }
      if (photoOverlayPrevBtn) {
        photoOverlayPrevBtn.disabled = items.length <= 1;
      }
      if (photoOverlayNextBtn) {
        photoOverlayNextBtn.disabled = items.length <= 1;
      }
      loadOverlayImage(current.path);
    };

    const openPhotoOverlay = (tag, kind = "report") => {
      const items = getPhotosForTag(tag, kind);
      if (!items.length) return;
      state.photoOverlay = {
        tag,
        items,
        index: 0,
        url: "",
        kind,
      };
      if (photoOverlayEl) {
        photoOverlayEl.classList.add("is-open");
        photoOverlayEl.setAttribute("aria-hidden", "false");
      }
      renderPhotoOverlay();
    };

    const closePhotoOverlay = () => {
      if (!photoOverlayEl) return;
      photoOverlayEl.classList.remove("is-open");
      photoOverlayEl.setAttribute("aria-hidden", "true");
      if (state.photoOverlay.url) {
        URL.revokeObjectURL(state.photoOverlay.url);
        state.photoOverlay.url = "";
      }
    };

    const stepPhotoOverlay = (delta) => {
      const total = state.photoOverlay.items.length;
      if (!total) return;
      state.photoOverlay.index = (state.photoOverlay.index + delta + total) % total;
      renderPhotoOverlay();
    };

    const checklistCache = new Map();

    const getChecklistLocale = (locale) => {
      const requested = String(locale || "").toLowerCase();
      if (!requested) return null;
      const base = requested.split("-")[0];
      if (["de", "fr", "it"].includes(requested)) return requested;
      if (["de", "fr", "it"].includes(base)) return base;
      return null;
    };

    const loadChecklistData = async () => {
      const locale = state.project?.meta?.locale || "de-CH";
      const resolved = getChecklistLocale(locale);
      if (!resolved) {
        throw new Error(`Checklist locale not supported: ${locale}`);
      }
      if (checklistCache.has(resolved)) return {
        items: checklistCache.get(resolved),
        locale: resolved,
        requested: locale,
        fallback: false,
      };
      if (window.location.protocol !== "http:" && window.location.protocol !== "https:") {
        throw new Error("Checklists require http(s) context.");
      }
      const url = new URL(`../data/checklists/checklists_${resolved}.json`, window.location.href).toString();
      const response = await fetch(url);
      if (!response.ok) throw new Error(`Failed to load ${url}`);
      const items = await response.json();
      checklistCache.set(resolved, items);
      return {
        items,
        locale: resolved,
        requested: locale,
        fallback: false,
      };
    };

    const formatChecklistLocaleLabel = (source) => {
      const resolved = String(source.locale || "de").toUpperCase();
      return resolved;
    };

    const normalizeChecklistUrl = (url) => {
      const value = String(url || "").trim();
      if (!value) return "";
      if (/^https?:\/\//i.test(value)) return value;
      return `https://${value}`;
    };

    const normalizeChecklistText = (value) => (
      String(value || "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
    );

    const isEkasSection = (value) => {
      const normalized = normalizeChecklistText(value);
      return (
        normalized === "publikationen der ekas"
        || normalized === "publications cfst"
        || normalized === "pubblicazioni cfsl"
      );
    };

    const isVitalRulesTitle = (value) => {
      const normalized = normalizeChecklistText(value);
      return (
        normalized.includes("lebenswichtige regeln")
        || normalized.includes("regles vitales")
        || normalized.includes("regole vitali")
      );
    };

    const getChecklistCategory = (item) => {
      if (!item) return "";
      const section = String(item.section || "").trim();
      if (section && isEkasSection(section)) return "ekas";
      const title = String(item.title || "");
      if (title && isVitalRulesTitle(title)) return "vital";
      return "";
    };

    const getChecklistIndustries = () => ([
      {
        key: "transport",
        label: t("checklist_industry_transport", "Transport und Lagerung"),
        short: t("checklist_industry_transport_short", "Tr/Lg"),
      },
      {
        key: "metall",
        label: t("checklist_industry_metall", "Metall"),
        short: t("checklist_industry_metall_short", "Met"),
      },
      {
        key: "buero",
        label: t("checklist_industry_buero", "Büro"),
        short: t("checklist_industry_buero_short", "Bü"),
      },
      {
        key: "forst",
        label: t("checklist_industry_forst", "Forst"),
        short: t("checklist_industry_forst_short", "For"),
      },
      {
        key: "holz",
        label: t("checklist_industry_holz", "Holz"),
        short: t("checklist_industry_holz_short", "Holz"),
      },
      {
        key: "bau",
        label: t("checklist_industry_bau", "Bau- und Installationsgewerbe"),
        short: t("checklist_industry_bau_short", "Bau"),
      },
      {
        key: "uebrige",
        label: t("checklist_industry_uebrige", "Übrige"),
        short: t("checklist_industry_uebrige_short", "Übr"),
      },
    ]);

    const getChecklistCategories = () => ([
      {
        key: "vital",
        label: t("checklist_filter_vital", "Lebenswichtige Regeln"),
      },
      {
        key: "ekas",
        label: t("checklist_filter_ekas", "EKAS"),
      },
    ]);

    const formatChecklistTitle = (item) => {
      const title = String(item?.title || "").trim();
      if (!title) return "";
      if (getChecklistCategory(item) !== "ekas") return title;
      const normalized = normalizeChecklistText(title);
      const prefixes = [
        "checkliste:",
        "liste de controle:",
        "lista di controllo:",
        "checklist:",
      ];
      const isChecklist = prefixes.some((prefix) => normalized.startsWith(prefix));
      if (!isChecklist) return title;
      const colonIndex = title.indexOf(":");
      if (colonIndex === -1) return `${t("checklist_filter_ekas", "EKAS")} ${title}`;
      const remainder = title.slice(colonIndex + 1).trim();
      if (!remainder) return `${t("checklist_filter_ekas", "EKAS")} ${title}`;
      return `${t("checklist_filter_ekas", "EKAS")} ${remainder}`;
    };

    const formatChecklistMarkdown = (item) => {
      const prefix = t("checklist_see_also", "Siehe auch");
      const title = formatChecklistTitle(item);
      const rawUrl = String(item.url || "").trim();
      const linkTarget = normalizeChecklistUrl(rawUrl);
      const link = rawUrl && linkTarget ? `[${rawUrl}](${linkTarget})` : "";
      if (title && link) return `${prefix}: ${title} ${link}`;
      if (title) return `${prefix}: ${title}`;
      if (link) return `${prefix}: ${link}`;
      return "";
    };

    const copyToClipboard = async (value) => {
      if (!value) return false;
      if (navigator.clipboard?.writeText) {
        try {
          await navigator.clipboard.writeText(value);
          return true;
        } catch (err) {
          // fallback below
        }
      }
      const textarea = document.createElement("textarea");
      textarea.value = value;
      textarea.setAttribute("readonly", "");
      textarea.style.position = "fixed";
      textarea.style.opacity = "0";
      textarea.style.left = "-9999px";
      textarea.style.top = "0";
      document.body.appendChild(textarea);
      textarea.focus();
      textarea.select();
      let ok = false;
      try {
        ok = document.execCommand("copy");
      } catch (err) {
        ok = false;
      }
      document.body.removeChild(textarea);
      return ok;
    };

    const copyChecklistItem = async (item, button) => {
      const text = formatChecklistMarkdown(item);
      const ok = await copyToClipboard(text);
      if (ok) {
        setStatus(`Copied: ${formatChecklistTitle(item)}`);
        if (button) {
          button.classList.add("is-copied");
          window.setTimeout(() => {
            button.classList.remove("is-copied");
          }, 1200);
        }
      } else {
        setStatus("Copy failed.");
      }
    };

    const renderChecklistOverlay = async () => {
      if (!checklistOverlayEl || !checklistListEl) return;
      let source;
      try {
        source = await loadChecklistData();
      } catch (err) {
        setStatus(err.message || "Checklist load failed.");
        debug.logLine("error", `Checklist load failed: ${err.message || err}`);
        return;
      }
      const industries = getChecklistIndustries();
      const categories = getChecklistCategories();
      state.checklistOverlay = { ...state.checklistOverlay, ...source };
      const activeIndustry = state.checklistOverlay.industryFilter || "";
      const activeCategory = state.checklistOverlay.categoryFilter || "";
      const searchQuery = normalizeChecklistText(state.checklistOverlay.searchQuery || "");
      const items = Array.isArray(source.items) ? source.items : [];
      const filteredItems = items.filter((item) => {
        if (activeCategory) {
          if (getChecklistCategory(item) !== activeCategory) return false;
        }
        if (activeIndustry) {
          if (!item.industries?.includes(activeIndustry)) return false;
        }
        if (searchQuery) {
          const title = normalizeChecklistText(formatChecklistTitle(item));
          if (!title.includes(searchQuery)) return false;
        }
        return true;
      });

      if (checklistOverlayTitle) {
        const baseTitle = t("checklist_overlay_title", "Suva-Checklisten");
        const localeLabel = formatChecklistLocaleLabel(source);
        checklistOverlayTitle.textContent = `${baseTitle} (${localeLabel})`;
      }
      if (checklistOverlayHint) {
        const hint = t("checklist_overlay_hint", "Click ⧉ to copy “Siehe auch: … [link]”.");
        checklistOverlayHint.textContent = hint;
      }
      if (checklistSearchEl) {
        if (!checklistSearchEl.dataset.bound) {
          checklistSearchEl.addEventListener("input", () => {
            state.checklistOverlay.searchQuery = checklistSearchEl.value || "";
            renderChecklistOverlay();
          });
          checklistSearchEl.addEventListener("keydown", (event) => {
            if (event.key !== "Escape") return;
            if (!checklistSearchEl.value) return;
            checklistSearchEl.value = "";
            state.checklistOverlay.searchQuery = "";
            renderChecklistOverlay();
          });
          checklistSearchEl.dataset.bound = "1";
        }
        checklistSearchEl.value = state.checklistOverlay.searchQuery || "";
      }
      if (checklistFiltersEl) {
        checklistFiltersEl.innerHTML = "";
        const allBtn = document.createElement("button");
        allBtn.type = "button";
        allBtn.className = "checklist-filter";
        allBtn.textContent = t("checklist_filter_all", "All");
        if (!activeIndustry && !activeCategory) allBtn.classList.add("is-active");
        allBtn.addEventListener("click", () => {
          state.checklistOverlay.industryFilter = "";
          state.checklistOverlay.categoryFilter = "";
          renderChecklistOverlay();
        });
        checklistFiltersEl.appendChild(allBtn);
        industries.forEach((industry) => {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "checklist-filter";
          btn.textContent = industry.label;
          btn.title = industry.label;
          if (activeIndustry === industry.key) btn.classList.add("is-active");
          btn.addEventListener("click", () => {
            const isActive = activeIndustry === industry.key;
            state.checklistOverlay.industryFilter = isActive ? "" : industry.key;
            if (!isActive) state.checklistOverlay.categoryFilter = "";
            renderChecklistOverlay();
          });
          checklistFiltersEl.appendChild(btn);
        });
        categories.forEach((category) => {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "checklist-filter";
          btn.textContent = category.label;
          btn.title = category.label;
          if (activeCategory === category.key) btn.classList.add("is-active");
          btn.addEventListener("click", () => {
            const isActive = activeCategory === category.key;
            state.checklistOverlay.categoryFilter = isActive ? "" : category.key;
            if (!isActive) state.checklistOverlay.industryFilter = "";
            renderChecklistOverlay();
          });
          checklistFiltersEl.appendChild(btn);
        });
      }

      checklistListEl.innerHTML = "";
      const fragment = document.createDocumentFragment();

      const header = document.createElement("div");
      header.className = "checklist-row checklist-row--header";
      const headerIndustries = document.createElement("div");
      headerIndustries.className = "checklist-row__industries";
      industries.forEach((industry) => {
        const cell = document.createElement("div");
        cell.className = "checklist-industry-header";
        cell.textContent = industry.short;
        headerIndustries.appendChild(cell);
      });
      const headerTitle = document.createElement("div");
      headerTitle.className = "checklist-row__title";
      headerTitle.textContent = t("checklist_header_title", "Titel");
      const headerLink = document.createElement("div");
      headerLink.className = "checklist-row__link";
      headerLink.textContent = t("checklist_header_link", "Link");
      const headerCopy = document.createElement("div");
      headerCopy.className = "checklist-row__copy";
      headerCopy.textContent = t("checklist_header_copy", "Copy");
      header.appendChild(headerIndustries);
      header.appendChild(headerTitle);
      header.appendChild(headerLink);
      header.appendChild(headerCopy);
      fragment.appendChild(header);

      let lastSection = "";
      filteredItems.forEach((item) => {
        const section = String(item.section || "").trim();
        if (section && section !== lastSection) {
          const sectionRow = document.createElement("div");
          sectionRow.className = "checklist-row checklist-row--section";
          const sectionTitle = document.createElement("div");
          sectionTitle.className = "checklist-section__title";
          sectionTitle.textContent = section;
          sectionRow.appendChild(sectionTitle);
          fragment.appendChild(sectionRow);
          lastSection = section;
        } else if (!section) {
          lastSection = "";
        }
        const row = document.createElement("div");
        row.className = "checklist-row";
        const itemIndustries = new Set(item.industries || []);
        const industriesEl = document.createElement("div");
        industriesEl.className = "checklist-row__industries";
        industries.forEach((industry) => {
          const cell = document.createElement("div");
          cell.className = "checklist-industry-cell";
          if (itemIndustries.has(industry.key)) {
            cell.classList.add("is-on");
            cell.textContent = "x";
          }
          industriesEl.appendChild(cell);
        });

        const title = document.createElement("div");
        title.className = "checklist-row__title";
        title.textContent = formatChecklistTitle(item);
        const link = document.createElement("a");
        link.className = "checklist-row__link";
        const rawUrl = String(item.url || "").trim();
        link.textContent = rawUrl;
        link.href = normalizeChecklistUrl(rawUrl);
        link.target = "_blank";
        link.rel = "noopener";
        const copyBtn = document.createElement("button");
        copyBtn.type = "button";
        copyBtn.className = "checklist-copy";
        const copyLabel = t("checklist_copy_title", "Copy checklist link");
        const copyText = t("checklist_copy_button", "Copy");
        copyBtn.textContent = copyText;
        copyBtn.title = copyLabel;
        copyBtn.setAttribute("aria-label", copyLabel);
        copyBtn.addEventListener("click", () => {
          copyChecklistItem(item, copyBtn);
        });

        row.appendChild(industriesEl);
        row.appendChild(title);
        row.appendChild(link);
        row.appendChild(copyBtn);
        fragment.appendChild(row);
      });
      checklistListEl.appendChild(fragment);
      checklistListEl.scrollTop = 0;
    };

    const openChecklistOverlay = async () => {
      if (!checklistOverlayEl) return;
      state.checklistOverlay.industryFilter = "";
      state.checklistOverlay.categoryFilter = "";
      await renderChecklistOverlay();
      checklistOverlayEl.classList.add("is-open");
      checklistOverlayEl.setAttribute("aria-hidden", "false");
    };

    const closeChecklistOverlay = () => {
      if (!checklistOverlayEl) return;
      checklistOverlayEl.classList.remove("is-open");
      checklistOverlayEl.setAttribute("aria-hidden", "true");
    };

    const OBS_FILTER_OPTIONS = [
      { key: "all", label: t("obs_filter_all", "All") },
      { key: "with-photos", label: t("obs_filter_with_photos", "With Photos") },
      { key: "no-photos", label: t("obs_filter_no_photos", "No Photos") },
    ];

    const OBS_SORT_OPTIONS = [
      { key: "manual", label: t("obs_sort_manual", "Manual") },
      { key: "count-desc", label: t("obs_sort_count_desc", "Count ↓") },
      { key: "count-asc", label: t("obs_sort_count_asc", "Count ↑") },
      { key: "alpha-asc", label: t("obs_sort_alpha_asc", "A-Z") },
      { key: "alpha-desc", label: t("obs_sort_alpha_desc", "Z-A") },
    ];

    const observationsOrganizer = {
      chapterId: "4.8",
      baseOrder: [],
      draftOrder: [],
      filterMode: "all",
      sortMode: "manual",
      searchQuery: "",
      draggingId: "",
    };

    const getChapterById = (chapterId) => (
      (state.project?.chapters || []).find((chapter) => String(chapter?.id || "") === String(chapterId || ""))
    );

    const getObservationChapter = () => getChapterById("4.8");

    const getObservationRows = (chapter) => {
      if (!chapter) return [];
      const ordered = normalizeHelpers.orderObservationRows(chapter);
      return ordered.filter((row) => String(row?.kind || "").toLowerCase() !== "section");
    };

    const getObservationTagLabel = (row) => (
      String(row?.tag || row?.titleOverride || row?.id || "").trim()
    );

    const getObservationPhotoCount = (row) => {
      const tag = getObservationTagLabel(row);
      if (!tag) return 0;
      return getPhotosForTag(tag, "observations").length;
    };

    const normalizeOrganizerOrder = (order, rowMap) => {
      const seen = new Set();
      const normalized = [];
      order.forEach((rowId) => {
        const key = String(rowId || "");
        if (!key || seen.has(key) || !rowMap.has(key)) return;
        seen.add(key);
        normalized.push(key);
      });
      rowMap.forEach((_, rowId) => {
        if (seen.has(rowId)) return;
        seen.add(rowId);
        normalized.push(rowId);
      });
      return normalized;
    };

    const getObservationRowMap = () => {
      const chapter = getObservationChapter();
      const rowMap = new Map();
      getObservationRows(chapter).forEach((row) => {
        rowMap.set(String(row.id || ""), row);
      });
      return rowMap;
    };

    const compareText = (left, right) => String(left || "").localeCompare(String(right || ""), "de", { numeric: true, sensitivity: "base" });

    const getSortedObservationIds = (sortMode = observationsOrganizer.sortMode) => {
      const rowMap = getObservationRowMap();
      const manualOrder = normalizeOrganizerOrder(observationsOrganizer.draftOrder, rowMap);
      if (sortMode === "manual") return manualOrder;
      const entries = manualOrder.map((rowId) => {
        const row = rowMap.get(rowId);
        return {
          rowId,
          title: getObservationTagLabel(row),
          count: getObservationPhotoCount(row),
        };
      });
      entries.sort((a, b) => {
        if (sortMode === "count-desc") {
          if (b.count !== a.count) return b.count - a.count;
          const byTitle = compareText(a.title, b.title);
          if (byTitle !== 0) return byTitle;
          return compareText(a.rowId, b.rowId);
        }
        if (sortMode === "count-asc") {
          if (a.count !== b.count) return a.count - b.count;
          const byTitle = compareText(a.title, b.title);
          if (byTitle !== 0) return byTitle;
          return compareText(a.rowId, b.rowId);
        }
        if (sortMode === "alpha-desc") {
          const byTitle = compareText(b.title, a.title);
          if (byTitle !== 0) return byTitle;
          return compareText(b.rowId, a.rowId);
        }
        const byTitle = compareText(a.title, b.title);
        if (byTitle !== 0) return byTitle;
        return compareText(a.rowId, b.rowId);
      });
      return entries.map((entry) => entry.rowId);
    };

    const getOrganizerApplyOrder = () => getSortedObservationIds(observationsOrganizer.sortMode);

    const arraysEqual = (left, right) => {
      if (left.length !== right.length) return false;
      for (let i = 0; i < left.length; i += 1) {
        if (left[i] !== right[i]) return false;
      }
      return true;
    };

    const hasObservationOrganizerChanges = () => {
      const nextOrder = getOrganizerApplyOrder();
      return !arraysEqual(nextOrder, observationsOrganizer.baseOrder);
    };

    const canManualReorderInOrganizer = () => (
      observationsOrganizer.sortMode === "manual"
      && observationsOrganizer.filterMode === "all"
      && !String(observationsOrganizer.searchQuery || "").trim()
    );

    const moveObservationDraftRelative = (rowId, delta) => {
      const key = String(rowId || "");
      if (!key) return false;
      const current = [...observationsOrganizer.draftOrder];
      const index = current.indexOf(key);
      if (index === -1) return false;
      const next = index + delta;
      if (next < 0 || next >= current.length) return false;
      [current[index], current[next]] = [current[next], current[index]];
      observationsOrganizer.draftOrder = current;
      return true;
    };

    const moveObservationDraftBefore = (sourceId, targetId) => {
      const source = String(sourceId || "");
      const target = String(targetId || "");
      if (!source || !target || source === target) return false;
      const current = [...observationsOrganizer.draftOrder];
      const sourceIdx = current.indexOf(source);
      const targetIdx = current.indexOf(target);
      if (sourceIdx === -1 || targetIdx === -1) return false;
      current.splice(sourceIdx, 1);
      const nextTargetIdx = current.indexOf(target);
      current.splice(nextTargetIdx, 0, source);
      observationsOrganizer.draftOrder = current;
      return true;
    };

    const renderObservationsOrganizer = () => {
      if (!observationsOrganizeModal || !observationsOrganizeListEl || !observationsOrganizeStatusEl) return;
      const chapter = getObservationChapter();
      if (!chapter) return;
      const rowMap = getObservationRowMap();
      observationsOrganizer.draftOrder = normalizeOrganizerOrder(observationsOrganizer.draftOrder, rowMap);
      const orderedIds = getSortedObservationIds(observationsOrganizer.sortMode);
      const rows = orderedIds.map((rowId) => ({
        rowId,
        row: rowMap.get(rowId),
      })).filter((item) => !!item.row);

      const filterMode = observationsOrganizer.filterMode || "all";
      const searchNorm = String(observationsOrganizer.searchQuery || "").trim().toLowerCase();

      const totalCount = rows.length;
      const withPhotosCount = rows.filter((item) => getObservationPhotoCount(item.row) > 0).length;
      const noPhotosCount = totalCount - withPhotosCount;

      if (observationsOrganizeFiltersEl) {
        observationsOrganizeFiltersEl.innerHTML = "";
        OBS_FILTER_OPTIONS.forEach((option) => {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "organizer-pill";
          if (filterMode === option.key) btn.classList.add("is-active");
          const count = option.key === "with-photos"
            ? withPhotosCount
            : option.key === "no-photos"
              ? noPhotosCount
              : totalCount;
          btn.textContent = `${option.label} ${count}`;
          btn.addEventListener("click", () => {
            observationsOrganizer.filterMode = option.key;
            renderObservationsOrganizer();
          });
          observationsOrganizeFiltersEl.appendChild(btn);
        });
      }

      if (observationsOrganizeSortsEl) {
        observationsOrganizeSortsEl.innerHTML = "";
        OBS_SORT_OPTIONS.forEach((option) => {
          const btn = document.createElement("button");
          btn.type = "button";
          btn.className = "organizer-pill";
          if (observationsOrganizer.sortMode === option.key) btn.classList.add("is-active");
          btn.textContent = option.label;
          btn.addEventListener("click", () => {
            observationsOrganizer.sortMode = option.key;
            renderObservationsOrganizer();
          });
          observationsOrganizeSortsEl.appendChild(btn);
        });
      }

      const filteredRows = rows.filter((item) => {
        const count = getObservationPhotoCount(item.row);
        if (filterMode === "with-photos" && count <= 0) return false;
        if (filterMode === "no-photos" && count > 0) return false;
        if (searchNorm) {
          const hay = getObservationTagLabel(item.row).toLowerCase();
          if (!hay.includes(searchNorm)) return false;
        }
        return true;
      });

      const canReorder = canManualReorderInOrganizer();
      observationsOrganizeListEl.innerHTML = "";
      const listFragment = document.createDocumentFragment();

      filteredRows.forEach((item) => {
        const { rowId, row } = item;
        const label = getObservationTagLabel(row) || rowId;
        const photoCount = getObservationPhotoCount(row);
        const draftIndex = observationsOrganizer.draftOrder.indexOf(rowId);

        const entry = document.createElement("div");
        entry.className = "organizer-row";
        if (canReorder) {
          entry.draggable = true;
          entry.addEventListener("dragstart", () => {
            observationsOrganizer.draggingId = rowId;
            entry.classList.add("is-dragging");
          });
          entry.addEventListener("dragend", () => {
            observationsOrganizer.draggingId = "";
            entry.classList.remove("is-dragging");
          });
          entry.addEventListener("dragover", (event) => {
            event.preventDefault();
            entry.classList.add("is-drag-over");
          });
          entry.addEventListener("dragleave", () => {
            entry.classList.remove("is-drag-over");
          });
          entry.addEventListener("drop", (event) => {
            event.preventDefault();
            entry.classList.remove("is-drag-over");
            if (!observationsOrganizer.draggingId) return;
            if (moveObservationDraftBefore(observationsOrganizer.draggingId, rowId)) {
              observationsOrganizer.draggingId = "";
              renderObservationsOrganizer();
            }
          });
        }

        const dragHandle = document.createElement("button");
        dragHandle.type = "button";
        dragHandle.className = "organizer-drag";
        dragHandle.textContent = "↕";
        dragHandle.disabled = !canReorder;
        dragHandle.title = canReorder
          ? t("obs_drag_title", "Drag to reorder")
          : t("obs_drag_disabled_title", "Enable Manual + All + empty search to reorder");

        const title = document.createElement("div");
        title.className = "organizer-title";
        title.textContent = label;

        const count = document.createElement("span");
        count.className = "organizer-count";
        count.textContent = `${photoCount}`;
        count.title = `${photoCount} photos`;

        const actions = document.createElement("div");
        actions.className = "organizer-actions";
        const up = document.createElement("button");
        up.type = "button";
        up.className = "ghost";
        up.textContent = "▲";
        up.disabled = !canReorder || draftIndex <= 0;
        up.addEventListener("click", () => {
          if (moveObservationDraftRelative(rowId, -1)) renderObservationsOrganizer();
        });
        const down = document.createElement("button");
        down.type = "button";
        down.className = "ghost";
        down.textContent = "▼";
        down.disabled = !canReorder || draftIndex === -1 || draftIndex >= observationsOrganizer.draftOrder.length - 1;
        down.addEventListener("click", () => {
          if (moveObservationDraftRelative(rowId, 1)) renderObservationsOrganizer();
        });
        actions.appendChild(up);
        actions.appendChild(down);

        entry.appendChild(dragHandle);
        entry.appendChild(title);
        entry.appendChild(count);
        entry.appendChild(actions);
        listFragment.appendChild(entry);
      });

      observationsOrganizeListEl.appendChild(listFragment);

      const pending = hasObservationOrganizerChanges();
      const visibilityText = `${t("obs_showing", "Showing")} ${filteredRows.length} / ${totalCount}`;
      const pendingText = pending
        ? t("obs_pending", "Pending changes")
        : t("obs_no_pending", "No pending changes");
      const reorderHint = canReorder
        ? t("obs_reorder_ready", "Reorder enabled")
        : t("obs_reorder_limited", "Reorder requires Manual + All + empty search");
      observationsOrganizeStatusEl.textContent = `${visibilityText} • ${pendingText} • ${reorderHint}`;
      if (observationsOrganizeApplyBtn) {
        observationsOrganizeApplyBtn.disabled = !pending;
      }
      if (observationsOrganizeSearchEl && observationsOrganizeSearchEl.value !== observationsOrganizer.searchQuery) {
        observationsOrganizeSearchEl.value = observationsOrganizer.searchQuery;
      }
    };

    const openObservationsOrganizer = (chapter) => {
      const targetChapter = chapter && chapter.id === "4.8" ? chapter : getObservationChapter();
      if (!observationsOrganizeModal || !targetChapter) return;
      const rows = getObservationRows(targetChapter);
      const baseOrder = rows.map((row) => String(row.id || ""));
      observationsOrganizer.baseOrder = [...baseOrder];
      observationsOrganizer.draftOrder = [...baseOrder];
      observationsOrganizer.filterMode = "all";
      observationsOrganizer.sortMode = "manual";
      observationsOrganizer.searchQuery = "";
      observationsOrganizer.draggingId = "";
      renderObservationsOrganizer();
      observationsOrganizeModal.classList.add("is-open");
      observationsOrganizeModal.setAttribute("aria-hidden", "false");
      if (observationsOrganizeSearchEl) observationsOrganizeSearchEl.focus();
    };

    const closeObservationsOrganizer = () => {
      if (!observationsOrganizeModal) return;
      observationsOrganizeModal.classList.remove("is-open");
      observationsOrganizeModal.setAttribute("aria-hidden", "true");
    };

    const resetObservationsOrganizer = () => {
      observationsOrganizer.draftOrder = [...observationsOrganizer.baseOrder];
      observationsOrganizer.filterMode = "all";
      observationsOrganizer.sortMode = "manual";
      observationsOrganizer.searchQuery = "";
      observationsOrganizer.draggingId = "";
      renderObservationsOrganizer();
    };

    const setObservationsOrganizerSearch = (value) => {
      observationsOrganizer.searchQuery = String(value || "");
      renderObservationsOrganizer();
    };

    const applyObservationsOrganizer = () => {
      const chapter = getObservationChapter();
      if (!chapter) return;
      const nextOrder = getOrganizerApplyOrder();
      chapter.meta = chapter.meta || {};
      chapter.meta.order = [...nextOrder];
      scheduleAutosave();
      renderRows();
      closeObservationsOrganizer();
    };

    const PREVIEW_COL1_WIDTH = 35;
    const PREVIEW_COL2_WIDTH = 58;
    const PREVIEW_COL3_WIDTH = 7;

    const isSectionRow = (row) => String(row?.kind || "").toLowerCase() === "section";

    const rowToText = (value) => {
      if (typeof stateHelpers.toText === "function") {
        return stateHelpers.toText(value);
      }
      if (Array.isArray(value)) return value.join("\n");
      if (value == null) return "";
      return String(value);
    };

    const isIncludedRow = (row) => {
      const ws = row?.workstate;
      if (!ws || ws.includeFinding == null) return true;
      return ws.includeFinding === true;
    };

    const resolveSectionId = (row, chapterId) => {
      const sectionId = String(row?.sectionId || "").trim();
      if (sectionId) return sectionId;
      const rawId = String(row?.id || "");
      const parts = rawId.split(".");
      if (parts.length >= 2) return `${parts[0]}.${parts[1]}`;
      if (chapterId) return `${chapterId}.1`;
      return "";
    };

    const stripLeadingNumber = (value) => (
      String(value || "")
        .replace(/^\s*\d+(?:\.\d+)*(?:\s|[.:-]\s*)?/, "")
        .trim()
    );

    const resolveSectionTitle = (row) => {
      const rawTitle = String(row?.title || row?.id || "");
      const cleaned = stripLeadingNumber(rawTitle);
      return cleaned || rawTitle;
    };

    const resolveFindingText = (row) => {
      const ws = row?.workstate;
      if (ws && Object.prototype.hasOwnProperty.call(ws, "findingText")) {
        return rowToText(ws.findingText);
      }
      return rowToText(row?.master?.finding);
    };

    const resolveRecommendationText = (row) => {
      const ws = row?.workstate || {};
      if (ws.includeRecommendation === false) return "";
      if (Object.prototype.hasOwnProperty.call(ws, "recommendationText")) {
        return rowToText(ws.recommendationText);
      }
      return rowToText(row?.master?.recommendation);
    };

    const isFieldObservationChapter = (chapterId) => String(chapterId || "").includes(".");

    const buildIncludedSections = (rows) => {
      const included = new Set();
      rows.forEach((row) => {
        if (isSectionRow(row) || !isIncludedRow(row)) return;
        const sectionId = String(row?.sectionId || "").trim();
        if (sectionId) included.add(sectionId);
      });
      return included;
    };

    const buildRenumberMap = (rows, chapterId) => {
      const rowMap = new Map();
      const sectionMap = new Map();
      const sectionCounts = new Map();
      let itemCount = 0;
      rows.forEach((row) => {
        if (isSectionRow(row) || !isIncludedRow(row)) return;
        const rowId = String(row?.id || "").trim();
        if (!rowId) return;
        if (isFieldObservationChapter(chapterId)) {
          itemCount += 1;
          rowMap.set(rowId, `${chapterId}.${itemCount}`);
          return;
        }
        const sectionId = resolveSectionId(row, chapterId) || `${chapterId}.1`;
        if (!sectionMap.has(sectionId)) {
          sectionMap.set(sectionId, sectionMap.size + 1);
        }
        const count = (sectionCounts.get(sectionId) || 0) + 1;
        sectionCounts.set(sectionId, count);
        rowMap.set(rowId, `${chapterId}.${sectionMap.get(sectionId)}.${count}`);
      });
      return { rowMap, sectionMap };
    };

    const buildChapterPreviewRows = (chapter) => {
      let rows = Array.isArray(chapter?.rows) ? chapter.rows : [];
      if (chapter?.id === "4.8" || chapter?.id === "0") {
        rows = normalizeHelpers.orderObservationRows(chapter);
      }
      const includedSections = buildIncludedSections(rows);
      const { rowMap } = buildRenumberMap(rows, chapter?.id || "");
      const output = [];
      rows.forEach((row) => {
        if (isSectionRow(row)) {
          const sectionId = String(row?.id || "").trim();
          if (!sectionId || !includedSections.has(sectionId)) return;
          output.push({
            kind: "section",
            title: resolveSectionTitle(row),
          });
          return;
        }
        if (!isIncludedRow(row)) return;
        const rowId = String(row?.id || "").trim();
        output.push({
          kind: "finding",
          id: rowMap.get(rowId) || rowId,
          finding: resolveFindingText(row),
          recommendation: resolveRecommendationText(row),
        });
      });
      return output;
    };

    const createChapterPreviewTable = (rows) => {
      const table = document.createElement("table");
      table.className = "chapter-preview-table";
      const colgroup = document.createElement("colgroup");
      [PREVIEW_COL1_WIDTH, PREVIEW_COL2_WIDTH, PREVIEW_COL3_WIDTH].forEach((width) => {
        const col = document.createElement("col");
        col.style.width = `${width}%`;
        colgroup.appendChild(col);
      });
      table.appendChild(colgroup);

      const body = document.createElement("tbody");

      const topHeader = document.createElement("tr");
      topHeader.className = "chapter-preview-table__check";
      const topBlank = document.createElement("td");
      topBlank.colSpan = 2;
      topBlank.textContent = "";
      const topCheck = document.createElement("td");
      topCheck.className = "chapter-preview-table__prio";
      topCheck.textContent = "✓";
      topHeader.appendChild(topBlank);
      topHeader.appendChild(topCheck);
      body.appendChild(topHeader);

      const titleHeader = document.createElement("tr");
      titleHeader.className = "chapter-preview-table__title";
      const titleMain = document.createElement("td");
      titleMain.colSpan = 2;
      titleMain.textContent = "Systempunkte mit Verbesserungspotenzial";
      const titleSpacer = document.createElement("td");
      titleSpacer.textContent = "";
      titleHeader.appendChild(titleMain);
      titleHeader.appendChild(titleSpacer);
      body.appendChild(titleHeader);

      const labelHeader = document.createElement("tr");
      labelHeader.className = "chapter-preview-table__labels";
      ["Ist-Zustand", "Lösungsansätze", "Prio"].forEach((label, index) => {
        const th = document.createElement("th");
        th.scope = "col";
        th.textContent = label;
        if (index === 2) th.className = "chapter-preview-table__prio";
        labelHeader.appendChild(th);
      });
      body.appendChild(labelHeader);

      rows.forEach((entry) => {
        if (entry.kind === "section") {
          const row = document.createElement("tr");
          row.className = "chapter-preview-table__section";
          const cell = document.createElement("td");
          cell.colSpan = 3;
          cell.textContent = entry.title || "";
          row.appendChild(cell);
          body.appendChild(row);
          return;
        }

        const row = document.createElement("tr");
        row.className = "chapter-preview-table__item";

        const finding = document.createElement("td");
        const id = document.createElement("div");
        id.className = "chapter-preview-id";
        id.textContent = entry.id || "";
        const findingText = document.createElement("div");
        findingText.className = "chapter-preview-text";
        findingText.innerHTML = markdownToHtml(entry.finding || "");
        finding.appendChild(id);
        finding.appendChild(findingText);

        const recommendation = document.createElement("td");
        recommendation.className = "chapter-preview-text";
        recommendation.innerHTML = markdownToHtml(entry.recommendation || "");

        const prio = document.createElement("td");
        prio.className = "chapter-preview-table__prio";
        prio.textContent = "";

        row.appendChild(finding);
        row.appendChild(recommendation);
        row.appendChild(prio);
        body.appendChild(row);
      });

      table.appendChild(body);

      const wrapper = document.createElement("div");
      wrapper.className = "chapter-preview-table-wrap";
      wrapper.appendChild(table);
      return wrapper;
    };

    const renderChapterPreview = (chapter) => {
      if (!chapterPreviewBody || !chapterPreviewTitle) return;
      chapterPreviewBody.innerHTML = "";
      const chapterLabel = stateHelpers.formatChapterLabel(chapter);
      chapterPreviewTitle.textContent = t("chapter_preview_title", "Word Export Preview");
      const subtitle = document.createElement("div");
      subtitle.className = "chapter-preview-subtitle";
      subtitle.textContent = chapterLabel;
      chapterPreviewBody.appendChild(subtitle);

      const rows = buildChapterPreviewRows(chapter);
      if (!rows.length) {
        const empty = document.createElement("div");
        empty.className = "chapter-preview-empty";
        empty.textContent = t("chapter_preview_empty", "No included findings in this chapter.");
        chapterPreviewBody.appendChild(empty);
        return;
      }
      chapterPreviewBody.appendChild(createChapterPreviewTable(rows));
    };

    const openChapterPreview = (chapter) => {
      if (!chapterPreviewModal || !chapter) return;
      renderChapterPreview(chapter);
      chapterPreviewModal.classList.add("is-open");
      chapterPreviewModal.setAttribute("aria-hidden", "false");
    };

    const closeChapterPreview = () => {
      if (!chapterPreviewModal) return;
      chapterPreviewModal.classList.remove("is-open");
      chapterPreviewModal.setAttribute("aria-hidden", "true");
    };

    const renderChapterList = () => {
      chapterListEl.innerHTML = "";
      const orderedChapters = [...state.project.chapters].sort((a, b) =>
        stateHelpers.compareIdSegments(a.id, b.id),
      );
      orderedChapters.forEach((chapter) => {
        const button = document.createElement("button");
        const buttonLabel = stateHelpers.formatChapterLabel(chapter);
        button.textContent = buttonLabel;
        button.title = buttonLabel;
        button.className = chapter.id === state.selectedChapterId ? "active" : "";
        button.addEventListener("click", () => {
          state.selectedChapterId = chapter.id;
          render();
        });
        chapterListEl.appendChild(button);
      });
    };

    const createPhotoPill = (tag, kind = "report") => {
      const items = getPhotosForTag(tag, kind);
      if (!items.length) return null;
      const pill = document.createElement("button");
      pill.type = "button";
      pill.className = "photo-pill";
      pill.textContent = `Photos ${items.length}`;
      pill.title = `${items.length} photos tagged ${tag}`;
      pill.addEventListener("click", () => {
        openPhotoOverlay(tag, kind);
      });
      return pill;
    };

    const renderChapterTitle = (chapter) => {
      chapterTitleEl.innerHTML = "";
      const title = document.createElement("span");
      title.textContent = stateHelpers.formatChapterLabel(chapter);
      const pill = createPhotoPill(chapter.id);
      if (pill) chapterTitleEl.appendChild(pill);
      chapterTitleEl.appendChild(title);
      const chapterPreviewBtn = document.createElement("button");
      chapterPreviewBtn.type = "button";
      chapterPreviewBtn.className = "checklist-pill";
      chapterPreviewBtn.textContent = t("chapter_preview_button", "Preview chapter");
      chapterPreviewBtn.addEventListener("click", () => {
        openChapterPreview(chapter);
      });
      chapterTitleEl.appendChild(chapterPreviewBtn);
      if (chapter.id === "4.8") {
        const organizeBtn = document.createElement("button");
        organizeBtn.type = "button";
        organizeBtn.className = "checklist-pill";
        const organizeLabel = t("obs_organize_button", "Organize 4.8");
        organizeBtn.textContent = `↕ ${organizeLabel}`;
        organizeBtn.title = organizeLabel;
        organizeBtn.setAttribute("aria-label", organizeLabel);
        organizeBtn.addEventListener("click", () => {
          openObservationsOrganizer(chapter);
        });
        chapterTitleEl.appendChild(organizeBtn);

        const checklistBtn = document.createElement("button");
        checklistBtn.type = "button";
        checklistBtn.className = "checklist-pill";
        const label = t("checklist_button_label", "Checklisten");
        checklistBtn.textContent = `☑ ${label}`;
        const title = t("checklist_button_title", label);
        checklistBtn.title = title;
        checklistBtn.setAttribute("aria-label", title);
        checklistBtn.addEventListener("click", () => {
          openChecklistOverlay();
        });
        chapterTitleEl.appendChild(checklistBtn);
      }
    };

    const createSectionRow = (row) => {
      const sectionRow = document.createElement("div");
      sectionRow.className = "section-row";
      const topLevel = String(row.id || "").split(".")[0];
      const hideSectionPill = ["11", "12", "13", "14"].includes(topLevel);
      if (!hideSectionPill) {
        const pill = createPhotoPill(row.id);
        if (pill) sectionRow.appendChild(pill);
      }
      const label = document.createElement("span");
      label.textContent = row.title;
      sectionRow.appendChild(label);
      return sectionRow;
    };

    const createAnswerBadge = (row) => {
      const badge = document.createElement("span");
      badge.className = "badge";
      const answerState = stateHelpers.getAnswerState(row);
      if (answerState === 1) {
        badge.textContent = "Answer: Yes";
      } else if (answerState === 0) {
        badge.textContent = "Answer: No";
      } else if (answerState === "mixed") {
        badge.textContent = "Answer: Mixed";
      } else {
        badge.textContent = "Answer: —";
      }
      return badge;
    };

    const createRowHeader = (row, score, options = {}) => {
      const header = document.createElement("div");
      header.className = "row-header";
      const meta = document.createElement("div");
      meta.className = "row-meta";
      const displayId = options.displayId || row.id;
      meta.textContent = `${displayId} ${row.titleOverride || ""}`.trim();
      const previewBtn = document.createElement("button");
      previewBtn.className = "preview-btn";
      previewBtn.type = "button";
      previewBtn.textContent = "Preview";
      const headerRight = document.createElement("div");
      headerRight.className = "row-header__right";
      headerRight.appendChild(createAnswerBadge(row));
      if (score !== null && score !== undefined) {
        const scoreBadge = document.createElement("span");
        scoreBadge.className = "score-badge";
        scoreBadge.textContent = `${score}%`;
        headerRight.appendChild(scoreBadge);
      }
      if (options.photoPill) {
        headerRight.appendChild(options.photoPill);
      }
      if (options.reorderControls) {
        headerRight.appendChild(options.reorderControls);
      }
      const comments = stateHelpers.getAnswerComments(row);
      if (comments.length) {
        headerRight.appendChild(createBadgeTooltip("Comment", comments.join("\n"), "comment-badge"));
      }
      const evidence = stateHelpers.getAnswerEvidence(row);
      if (evidence.length) {
        headerRight.appendChild(createBadgeTooltip("Evidence", evidence.join("\n"), "evidence-badge"));
      }
      headerRight.appendChild(previewBtn);
      header.appendChild(meta);
      header.appendChild(headerRight);
      return { header, previewBtn };
    };

    const createSelfAssessmentDetails = (row) => {
      if (row.type === "summary") return null;
      const selfItems = row.customer?.items || [];
      if (selfItems.length <= 1) return null;
      const formatAnswer = (value) => {
        if (value === 1 || value === "1" || value === true) return "Yes";
        if (value === 0 || value === "0" || value === false) return "No";
        return "—";
      };
      const details = document.createElement("details");
      details.className = "self-details";
      const summary = document.createElement("summary");
      summary.textContent = `Selbstbeurteilung (${selfItems.length})`;
      details.appendChild(summary);
      const list = document.createElement("ul");
      list.className = "self-list";
      selfItems.forEach((item) => {
        const li = document.createElement("li");
        const answer = formatAnswer(item.answer);
        const label = `${item.id} [${answer}] — ${item.question || ""}`.trim();
        li.textContent = label;
        list.appendChild(li);
      });
      details.appendChild(list);
      return details;
    };

    const createReorderControls = (chapter, row) => {
      const wrapper = document.createElement("div");
      wrapper.className = "reorder-controls";
      const up = document.createElement("button");
      up.type = "button";
      up.className = "ghost";
      up.textContent = "▲";
      up.addEventListener("click", () => {
        const moved = normalizeHelpers.moveObservationRow(chapter, row.id, -1);
        if (moved) {
          scheduleAutosave();
          renderRows();
        }
      });
      const down = document.createElement("button");
      down.type = "button";
      down.className = "ghost";
      down.textContent = "▼";
      down.addEventListener("click", () => {
        const moved = normalizeHelpers.moveObservationRow(chapter, row.id, 1);
        if (moved) {
          scheduleAutosave();
          renderRows();
        }
      });
      wrapper.appendChild(up);
      wrapper.appendChild(down);
      return wrapper;
    };

    const createToggle = (labelText, checked, onChange) => {
      const label = document.createElement("label");
      label.className = "field-toggle";
      const input = document.createElement("input");
      input.type = "checkbox";
      input.checked = checked;
      input.addEventListener("change", () => onChange(input.checked));
      label.appendChild(input);
      label.appendChild(document.createTextNode(labelText));
      return label;
    };

    const createMarkdownHint = () => {
      const hint = document.createElement("span");
      hint.className = "hint";
      const icon = document.createElement("span");
      icon.className = "hint__icon";
      icon.textContent = "?";
      const tooltip = document.createElement("span");
      tooltip.className = "hint__tooltip";
      tooltip.innerHTML = [
        t("markdown_hint_title", "Markdown:"),
        `<strong>${t("markdown_hint_bold", "**bold**")}</strong>,`,
        `<em>${t("markdown_hint_italic", "*italic*")}</em>,`,
        `<span class="hint__link">${t("markdown_hint_link", "[text](url)")}</span>`,
        `<br>${t("markdown_hint_list", "- List item")}`,
        `<br>${t("markdown_hint_paragraph", "Blank line = new paragraph")}`,
        `<br>${t("markdown_hint_linebreak", "Line breaks kept")}`,
      ].join(" ");
      hint.appendChild(icon);
      hint.appendChild(tooltip);
      return hint;
    };

    const createBadgeTooltip = (label, tooltipText, className) => {
      const wrapper = document.createElement("span");
      wrapper.className = "badge-tooltip";
      const badge = document.createElement("span");
      badge.className = className;
      badge.textContent = label;
      const tooltip = document.createElement("span");
      tooltip.className = "hint__tooltip";
      tooltip.textContent = tooltipText;
      wrapper.appendChild(badge);
      wrapper.appendChild(tooltip);
      return wrapper;
    };

    const createFindingField = (row, ws, onChange) => {
      const findingField = document.createElement("div");
      findingField.className = "field";
      const findingHeader = document.createElement("div");
      findingHeader.className = "field-header";
      const findingLabel = document.createElement("label");
      findingLabel.textContent = "Finding";
      findingLabel.appendChild(createMarkdownHint());

      const includeToggle = createToggle("Include", ws.includeFinding, (checked) => {
        ws.includeFinding = checked;
        scheduleAutosave();
        renderRows();
      });
      const doneToggle = createToggle("Done", ws.done, (checked) => {
        ws.done = checked;
        scheduleAutosave();
        renderRows();
      });

      const findingControls = document.createElement("div");
      findingControls.className = "field-controls";
      findingControls.appendChild(includeToggle);
      findingControls.appendChild(doneToggle);

      const findingArea = document.createElement("textarea");
      findingArea.value = stateHelpers.getFindingText(row);
      findingArea.addEventListener("input", () => {
        ws.findingText = findingArea.value;
        scheduleAutosave();
        autosizeTextarea(findingArea);
        if (onChange) onChange();
      });
      requestAnimationFrame(() => autosizeTextarea(findingArea));

      findingHeader.appendChild(findingLabel);
      findingHeader.appendChild(findingControls);
      findingField.appendChild(findingHeader);
      findingField.appendChild(findingArea);
      return findingField;
    };

    const createRecommendationField = (row, ws, onChange) => {
      const recField = document.createElement("div");
      recField.className = "field";
      const recHeader = document.createElement("div");
      recHeader.className = "field-header";
      const recLabel = document.createElement("label");
      recLabel.textContent = "Recommendation";
      recLabel.appendChild(createMarkdownHint());

      const levelGroup = document.createElement("div");
      levelGroup.className = "level-group";
      [1, 2, 3, 4].forEach((level) => {
        const label = document.createElement("label");
        label.className = "level-radio";
        const input = document.createElement("input");
        input.type = "radio";
        input.name = `level-${row.id}`;
        input.value = String(level);
        input.checked = ws.selectedLevel === level;
        input.addEventListener("change", () => {
          ws.selectedLevel = level;
          scheduleAutosave();
          renderRows();
        });
        label.appendChild(input);
        const pct = typeof stateHelpers.levelToPct === "function"
          ? stateHelpers.levelToPct(level)
          : (level - 1) * 25;
        label.appendChild(document.createTextNode(`${pct}%`));
        levelGroup.appendChild(label);
      });

      const recControls = document.createElement("div");
      recControls.className = "field-controls";
      recControls.appendChild(levelGroup);

      const libraryGroup = document.createElement("div");
      libraryGroup.className = "library-group";
      const libraryLabel = document.createElement("span");
      libraryLabel.textContent = "Library";
      libraryGroup.appendChild(libraryLabel);
      const actionButtons = [
        { key: "off", label: "Off" },
        { key: "append", label: "Append" },
        { key: "replace", label: "Replace" },
      ];
      actionButtons.forEach((option) => {
        const button = document.createElement("button");
        button.type = "button";
        button.textContent = option.label;
        if ((ws.libraryAction || "off") === option.key) {
          button.classList.add("is-active");
        }
        button.addEventListener("click", () => {
          ws.libraryAction = option.key;
          scheduleAutosave();
          renderRows();
        });
        libraryGroup.appendChild(button);
      });

      const recommendationInput = document.createElement("textarea");
      const recommendationText = stateHelpers.getRecommendationText(row);
      recommendationInput.value = recommendationText;
      recommendationInput.addEventListener("input", () => {
        ws.recommendationText = recommendationInput.value;
        scheduleAutosave();
        autosizeTextarea(recommendationInput);
        if (onChange) onChange();
      });
      requestAnimationFrame(() => autosizeTextarea(recommendationInput));

      recHeader.appendChild(recLabel);
      recHeader.appendChild(recControls);
      recHeader.appendChild(libraryGroup);
      recField.appendChild(recHeader);
      recField.appendChild(recommendationInput);
      return recField;
    };

    const createPreviewPanel = (row, previewBtn) => {
      const preview = document.createElement("div");
      preview.className = "row-preview";

      const findingCol = document.createElement("div");
      findingCol.className = "row-preview__col";

      const recCol = document.createElement("div");
      recCol.className = "row-preview__col";

      preview.appendChild(findingCol);
      preview.appendChild(recCol);

      const renderPreview = () => {
        const findingText = stateHelpers.getFindingText(row);
        const recommendation = stateHelpers.getRecommendationText(row);
        findingCol.innerHTML = [
          `<strong>${escapeHtml(t("preview_finding", "Finding"))}</strong>`,
          `${markdownToHtml(findingText || "")}`,
        ].join("\n");
        recCol.innerHTML = [
          `<strong>${escapeHtml(t("preview_recommendation", "Recommendation"))}</strong>`,
          `${markdownToHtml(recommendation || "")}`,
        ].join("\n");
      };

      const showPreview = () => {
        renderPreview();
        preview.classList.add("is-visible");
      };
      const hidePreview = () => {
        preview.classList.remove("is-visible");
      };
      previewBtn.addEventListener("click", (event) => {
        event.preventDefault();
        if (preview.classList.contains("is-visible")) {
          hidePreview();
        } else {
          showPreview();
        }
      });
      renderPreview();
      return {
        preview,
        renderPreview,
        isVisible: () => preview.classList.contains("is-visible"),
      };
    };

    const shouldFilterRow = (row, ws) => {
      if (row.kind === "section") return false;
      const answerMode = state.filters.answer || "all";
      const includeMode = state.filters.include || "all";
      const doneMode = state.filters.done || "all";
      // `shouldFilterRow` returns true when the row should be hidden.
      if (answerMode === "hide-yes") {
        if (stateHelpers.getAnswerState(row) === 1) return true;
      }
      if (includeMode === "included" && ws.includeFinding !== true) return true;
      if (includeMode === "not-included" && ws.includeFinding === true) return true;
      if (doneMode === "done" && ws.done !== true) return true;
      if (doneMode === "not-done" && ws.done === true) return true;
      return false;
    };

    const createRowCard = (row, options = {}) => {
      normalizeHelpers.ensureWorkstateDefaults(row);
      const ws = row.workstate;
      if (shouldFilterRow(row, ws)) return null;
      const card = document.createElement("div");
      card.className = "row-card";
      const score = stateHelpers.calculateScore(row);
      const isObservationRow = row.type === "field_observation";
      const photoKind = isObservationRow ? "observations" : "report";
      const photoTag = isObservationRow
        ? (row.tag || row.titleOverride || row.id)
        : (row.sectionId || row.id);
      const hasParentSectionPill = !isObservationRow
        && String(row.sectionId || "") !== String(row.id || "")
        && Array.isArray(options.chapter?.rows)
        && options.chapter.rows.some((candidate) => (
          candidate
          && candidate.kind === "section"
          && String(candidate.id || "") === String(row.sectionId || "")
        ));
      const showPhotoPill = row.type !== "summary"
        && !options.hidePhotoPill
        && (isObservationRow || !hasParentSectionPill);
      const headerPayload = createRowHeader(row, score, {
        photoPill: showPhotoPill ? createPhotoPill(photoTag, photoKind) : null,
        reorderControls: options.reorderControls,
        displayId: options.displayId,
      });
      const selfDetails = createSelfAssessmentDetails(row);
      if (headerPayload) {
        card.appendChild(headerPayload.header);
      }
      if (selfDetails) {
        card.appendChild(selfDetails);
      }
      const previewPanel = createPreviewPanel(row, headerPayload.previewBtn);
      const onChange = () => {
        if (previewPanel.isVisible()) {
          previewPanel.renderPreview();
        }
      };
      const rowBody = document.createElement("div");
      rowBody.className = "row-body";
      rowBody.appendChild(createFindingField(row, ws, onChange));
      rowBody.appendChild(createRecommendationField(row, ws, onChange));
      card.appendChild(rowBody);
      card.appendChild(previewPanel.preview);
      return card;
    };

    const renderRows = () => {
      rowsEl.innerHTML = "";
      const chapter = state.project.chapters.find((c) => c.id === state.selectedChapterId);
      if (!chapter) return;
      renderChapterTitle(chapter);
      let rows = chapter.rows || [];
      let displayIds = new Map();
      if (chapter.id === "4.8" || chapter.id === "0") {
        rows = normalizeHelpers.orderObservationRows(chapter);
        if (chapter.id === "4.8") {
          rows.forEach((row, index) => {
            displayIds.set(row.id, `4.8.${index + 1}`);
          });
        }
      }

      rows.forEach((row) => {
        if (row.kind === "section") {
          rowsEl.appendChild(createSectionRow(row));
          return;
        }
        const displayId = displayIds.get(row.id);
        const allowReorder = chapter.id === "4.8" || chapter.id === "0";
        const reorderControls = allowReorder ? createReorderControls(chapter, row) : null;
        const card = createRowCard(row, { chapter, displayId, reorderControls });
        if (card) rowsEl.appendChild(card);
      });

      if (observationsOrganizeModal?.classList.contains("is-open")) {
        renderObservationsOrganizer();
      }
    };

    const render = () => {
      renderChapterList();
      renderRows();
    };

    return {
      buildPhotoIndex,
      getPhotosForTag,
      openPhotoOverlay,
      closePhotoOverlay,
      stepPhotoOverlay,
      openChecklistOverlay,
      closeChecklistOverlay,
      openObservationsOrganizer,
      closeObservationsOrganizer,
      resetObservationsOrganizer,
      applyObservationsOrganizer,
      setObservationsOrganizerSearch,
      openChapterPreview,
      closeChapterPreview,
      renderChapterList,
      renderChapterTitle,
      renderRows,
      render,
      setScheduleAutosave,
      renderPhotoOverlay,
    };
  };

  window.AutoBerichtRender = { init };
})();
