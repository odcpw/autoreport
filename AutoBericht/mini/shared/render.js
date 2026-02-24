(() => {
  const init = (ctx, deps) => {
    const { stateHelpers, normalizeHelpers } = deps;
    const { elements, state, runtime, debug, setStatus } = ctx;
    const { t } = ctx.i18n;
    const { escapeHtml = (value) => value, formatInlineMarkdown = (value) => value } = ctx.markdown || {};
    const reportRows = window.AutoBerichtReportRows || {};

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
    let seedBootstrapHandler = null;
    let libraryExcelExportHandler = null;
    let sidecarMigrationHandler = null;

    const setScheduleAutosave = (fn) => {
      scheduleAutosave = typeof fn === "function" ? fn : () => {};
    };

    const setSeedBootstrapHandler = (fn) => {
      seedBootstrapHandler = typeof fn === "function" ? fn : null;
    };

    const setLibraryExcelExportHandler = (fn) => {
      libraryExcelExportHandler = typeof fn === "function" ? fn : null;
    };

    const setSidecarMigrationHandler = (fn) => {
      sidecarMigrationHandler = typeof fn === "function" ? fn : null;
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

    const applyEditorLocale = (inputEl) => {
      if (!inputEl) return;
      const locale = String(state.project?.meta?.locale || "de-CH");
      const spellLang = ctx.i18n?.resolveSpellcheckLang
        ? ctx.i18n.resolveSpellcheckLang(locale)
        : String(locale).toLowerCase().split("-")[0] || "en";
      inputEl.setAttribute("lang", spellLang);
      inputEl.setAttribute("spellcheck", "true");
      inputEl.spellcheck = true;
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
      { key: "prio-asc", label: t("obs_sort_prio_asc", "Prio 1-4") },
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

    const getObservationPriorityValue = (row) => {
      const raw = Number(row?.workstate?.priority);
      if (!Number.isFinite(raw)) return 0;
      return Math.max(0, Math.min(4, Math.round(raw)));
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
          priority: getObservationPriorityValue(row),
        };
      });
      entries.sort((a, b) => {
        if (sortMode === "prio-asc") {
          // Keep manual order inside each priority bucket.
          const aBucket = a.priority === 0 ? 99 : a.priority;
          const bBucket = b.priority === 0 ? 99 : b.priority;
          if (aBucket !== bBucket) return aBucket - bBucket;
          return manualOrder.indexOf(a.rowId) - manualOrder.indexOf(b.rowId);
        }
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
        const priorityValue = getObservationPriorityValue(row);
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

        const priority = document.createElement("span");
        priority.className = "organizer-priority";
        priority.textContent = `P${priorityValue}`;
        priority.title = `Priority ${priorityValue}`;

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
        entry.appendChild(priority);
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

    const isSectionRow = typeof reportRows.isSectionRow === "function"
      ? reportRows.isSectionRow
      : (row) => String(row?.kind || "").toLowerCase() === "section";

    const rowToText = (value) => {
      if (typeof reportRows.rowToText === "function") {
        return reportRows.rowToText(value, stateHelpers.toText);
      }
      if (typeof stateHelpers.toText === "function") {
        return stateHelpers.toText(value);
      }
      if (Array.isArray(value)) return value.join("\n");
      if (value == null) return "";
      return String(value);
    };

    const isIncludedRow = typeof reportRows.isReportReadyRow === "function"
      ? reportRows.isReportReadyRow
      : (row) => {
        const ws = row?.workstate;
        if (!ws || ws.includeFinding == null) return false;
        return ws.includeFinding === true && ws.done === true;
      };

    const resolveSectionTitle = typeof reportRows.resolveSectionTitle === "function"
      ? reportRows.resolveSectionTitle
      : (row) => String(row?.title || row?.id || "");

    const stripLeadingNumber = typeof reportRows.stripLeadingNumber === "function"
      ? reportRows.stripLeadingNumber
      : (value) => String(value || "").replace(/^\s*\d+(?:\.\d+)*(?:\s|[.:-]\s*)?/, "").trim();

    const resolveFindingText = typeof reportRows.resolveFindingText === "function"
      ? (row) => reportRows.resolveFindingText(row, stateHelpers.toText)
      : (row) => rowToText(row?.workstate?.findingText ?? row?.master?.finding);

    const resolveRecommendationText = typeof reportRows.resolveRecommendationText === "function"
      ? (row) => reportRows.resolveRecommendationText(row, stateHelpers.toText)
      : (row) => rowToText(row?.workstate?.recommendationText ?? row?.master?.recommendation);

    const resolvePriorityText = typeof reportRows.resolvePriorityText === "function"
      ? reportRows.resolvePriorityText
      : (row) => {
        const raw = Number(row?.workstate?.priority);
        if (!Number.isFinite(raw)) return "";
        const value = Math.round(raw);
        if (value < 1 || value > 4) return "";
        return String(value);
      };

    const buildChapterPreviewRows = (chapter) => {
      let rows = Array.isArray(chapter?.rows) ? chapter.rows : [];
      if (chapter?.id === "4.8" || chapter?.id === "0") {
        rows = normalizeHelpers.orderObservationRows(chapter);
      }
      if (typeof reportRows.buildChapterRows === "function") {
        return reportRows.buildChapterRows(chapter, {
          rows,
          includeRow: isIncludedRow,
          toText: stateHelpers.toText,
          titleForFinding: (row, chapterId) => (
            chapterId === "4.8" ? String(row?.titleOverride || "").trim() : ""
          ),
        });
      }
      const output = [];
      rows.forEach((row) => {
        if (isSectionRow(row)) {
          output.push({ kind: "section", title: resolveSectionTitle(row) });
          return;
        }
        if (!isIncludedRow(row)) return;
        output.push({
          kind: "finding",
          id: String(row?.id || ""),
          title: chapter?.id === "4.8" ? String(row?.titleOverride || "").trim() : "",
          finding: resolveFindingText(row),
          recommendation: resolveRecommendationText(row),
          priority: resolvePriorityText(row),
        });
      });
      return output;
    };

    const alphaLabel = (index) => {
      let n = Number(index) + 1;
      if (!Number.isFinite(n) || n < 1) return "";
      let out = "";
      while (n > 0) {
        const rem = (n - 1) % 26;
        out = String.fromCharCode(65 + rem) + out;
        n = Math.floor((n - 1) / 26);
      }
      return out;
    };

    const collapseSingleLine = (text) => String(text || "")
      .replace(/\r?\n+/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    const previewWordLikeHtml = (text) => {
      const lines = String(text || "").split(/\r?\n/);
      if (!lines.length) return "";
      const parts = [];
      let inList = false;
      const closeList = () => {
        if (!inList) return;
        parts.push("</ul>");
        inList = false;
      };
      lines.forEach((line) => {
        const trimmed = line.trim();
        if (trimmed.startsWith("- ")) {
          if (!inList) {
            parts.push("<ul>");
            inList = true;
          }
          const item = trimmed.slice(2);
          parts.push(`<li>${formatInlineMarkdown(escapeHtml(item))}</li>`);
          return;
        }
        closeList();
        if (trimmed === "") {
          parts.push("<p></p>");
          return;
        }
        parts.push(`<p>${formatInlineMarkdown(escapeHtml(line))}</p>`);
      });
      closeList();
      return parts.join("");
    };

    const createChapterPreviewTable = (rows, options = {}) => {
      const table = document.createElement("table");
      table.className = "chapter-preview-table";
      const positivesText = String(options?.positivesText || "").trim();
      const chapterId = String(options?.chapterId || "");
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
      topBlank.className = "chapter-preview-text";
      if (positivesText) {
        topBlank.innerHTML = previewWordLikeHtml(positivesText);
      } else {
        topBlank.textContent = "";
      }
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
          const sectionLabel = `${entry.id || ""}${entry.title ? ` ${entry.title}` : ""}`.trim();
          cell.textContent = sectionLabel || entry.title || "";
          row.appendChild(cell);
          body.appendChild(row);
          return;
        }

        const row = document.createElement("tr");
        row.className = "chapter-preview-table__item";

        const finding = document.createElement("td");
        const findingText = document.createElement("div");
        findingText.className = "chapter-preview-text";
        if (chapterId === "4.8") {
          const id = document.createElement("div");
          id.className = "chapter-preview-id";
          id.textContent = `${entry.id || ""}${entry.title ? ` ${entry.title}` : ""}`.trim();
          finding.appendChild(id);
          findingText.innerHTML = previewWordLikeHtml(entry.finding || "");
        } else {
          const collapsedFinding = collapseSingleLine(entry.finding || "");
          const normalizedFinding = stripLeadingNumber(collapsedFinding) || collapsedFinding;
          const inlineFinding = [
            `${entry.id || ""}`.trim(),
            normalizedFinding,
          ].filter(Boolean).join(" ");
          findingText.innerHTML = previewWordLikeHtml(inlineFinding);
        }
        finding.appendChild(findingText);

        const recommendation = document.createElement("td");
        recommendation.className = "chapter-preview-text";
        recommendation.innerHTML = previewWordLikeHtml(entry.recommendation || "");

        const prio = document.createElement("td");
        prio.className = "chapter-preview-table__prio";
        prio.textContent = entry.priority || "";

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

    const createChapter0PreviewList = (rows) => {
      const findings = (rows || []).filter((entry) => entry.kind === "finding");
      if (!findings.length) return null;
      const wrapper = document.createElement("div");
      wrapper.className = "chapter-preview-table-wrap chapter-preview-summary-wrap";
      const list = document.createElement("ol");
      list.className = "chapter-preview-summary-list";
      findings.forEach((entry, index) => {
        const item = document.createElement("li");
        item.setAttribute("data-alpha", alphaLabel(index));
        item.textContent = collapseSingleLine(entry.recommendation || "");
        list.appendChild(item);
      });
      wrapper.appendChild(list);
      return wrapper;
    };

    const renderChapterPreview = (chapter) => {
      if (!chapterPreviewBody || !chapterPreviewTitle) return;
      chapterPreviewBody.innerHTML = "";
      const chapterLabel = stateHelpers.formatChapterLabel(chapter, state.project?.meta?.locale);
      chapterPreviewTitle.textContent = t("chapter_preview_title", "Word Export Preview");
      const subtitle = document.createElement("div");
      subtitle.className = "chapter-preview-subtitle";
      subtitle.textContent = chapterLabel;
      chapterPreviewBody.appendChild(subtitle);

      const rows = buildChapterPreviewRows(chapter);
      if (String(chapter?.id || "") === "0") {
        const summaryList = createChapter0PreviewList(rows);
        if (!summaryList) {
          const empty = document.createElement("div");
          empty.className = "chapter-preview-empty";
          empty.textContent = t("chapter_preview_empty", "No included findings in this chapter.");
          chapterPreviewBody.appendChild(empty);
          return;
        }
        chapterPreviewBody.appendChild(summaryList);
        return;
      }
      const positivesText = stateHelpers.getChapterPositivesExportText
        ? stateHelpers.getChapterPositivesExportText(chapter)
        : "";
      if (!rows.length && !positivesText) {
        const empty = document.createElement("div");
        empty.className = "chapter-preview-empty";
        empty.textContent = t("chapter_preview_empty", "No included findings in this chapter.");
        chapterPreviewBody.appendChild(empty);
        return;
      }
      chapterPreviewBody.appendChild(createChapterPreviewTable(rows, {
        positivesText,
        chapterId: String(chapter?.id || ""),
      }));
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

    const PROJECT_VIEW_ID = "__project__";

    const formatIsoDate = (value) => {
      const raw = String(value || "");
      if (raw.length < 10) return "";
      const y = raw.slice(0, 4);
      const m = raw.slice(5, 7);
      const d = raw.slice(8, 10);
      if (!/^\d{4}$/.test(y) || !/^\d{2}$/.test(m) || !/^\d{2}$/.test(d)) return raw;
      return `${d}.${m}.${y}`;
    };

    const parseSpiderRowOverrides = (tableBody) => {
      const next = {};
      if (!tableBody) return next;
      tableBody.querySelectorAll("tr").forEach((tr) => {
        const id = tr.querySelector("input[data-id]")?.dataset.id;
        if (!id) return;
        const companyVal = tr.querySelector('input[data-field="company"]')?.value;
        const consultantVal = tr.querySelector('input[data-field="consultant"]')?.value;
        const useCompany = tr.querySelector('input[data-field="useCompany"]')?.checked || false;
        const useConsultant = tr.querySelector('input[data-field="useConsultant"]')?.checked || false;
        const company = companyVal === "" ? null : Number(companyVal);
        const consultant = consultantVal === "" ? null : Number(consultantVal);
        next[id] = {
          useCompany,
          useConsultant,
          company: Number.isFinite(company) ? company : null,
          consultant: Number.isFinite(consultant) ? consultant : null,
        };
      });
      return next;
    };

    const drawSpiderCanvas = (canvas, rows, companyLabel = "Company") => {
      if (!canvas) return;
      const spiderChart = window.AutoBerichtSpiderChart || {};
      if (typeof spiderChart.drawToCanvas !== "function") return;
      spiderChart.drawToCanvas(canvas, rows, {
        width: 760,
        height: 500,
        dpr: Math.max(1, Math.min(2, window.devicePixelRatio || 1)),
        setCssSize: true,
        companyLabel,
        suvaLabel: "Suva",
      });
    };

    const renderProjectPage = () => {
      rowsEl.innerHTML = "";
      if (runtime.awaitingLocaleBootstrap) {
        if (!state.project.meta || typeof state.project.meta !== "object") state.project.meta = {};
        const meta = state.project.meta;
        if (!meta.createdAt) meta.createdAt = new Date().toISOString();
        if (meta.moderator == null) meta.moderator = "";
        if (meta.moderatorInitials == null) meta.moderatorInitials = "";
        if (meta.coModerator == null) meta.coModerator = "";
        if (meta.coModeratorInitials == null) meta.coModeratorInitials = "";
        if (meta.company == null) meta.company = "";
        if (meta.companyId == null) meta.companyId = "";
        if (meta.address == null) meta.address = "";
        if (meta.plz == null) meta.plz = "";
        if (meta.city == null) meta.city = "";
        if (meta.autobackupMinutes == null) meta.autobackupMinutes = 30;
        meta.author = meta.moderator;
        meta.initials = meta.moderatorInitials;
      } else {
        normalizeHelpers.ensureProjectMeta(state.project, ctx.i18n.setLocale);
      }
      const meta = state.project.meta || {};
      const page = document.createElement("section");
      page.className = "project-page";
      const urlParams = new URLSearchParams(window.location.search || "");
      const migrationFlag = String(urlParams.get("migration") || "").toLowerCase();
      const migrationUiEnabled = ["1", "true", "yes", "on"].includes(migrationFlag);

      const createToolCard = ({
        title,
        hint,
        buttonLabel,
        buttonClass = "ghost",
        onClick,
      }) => {
        const card = document.createElement("div");
        card.className = "project-card";
        const cardTitle = document.createElement("h4");
        cardTitle.textContent = title;
        card.appendChild(cardTitle);
        if (hint) {
          const hintEl = document.createElement("p");
          hintEl.className = "project-card__hint";
          hintEl.textContent = hint;
          card.appendChild(hintEl);
        }
        const actions = document.createElement("div");
        actions.className = "project-page__actions";
        const button = document.createElement("button");
        button.type = "button";
        if (buttonClass) button.className = buttonClass;
        button.textContent = buttonLabel;
        button.addEventListener("click", async () => {
          await onClick(button);
        });
        actions.appendChild(button);
        card.appendChild(actions);
        return { card, button };
      };

      const triggerImportSelf = () => {
        if (elements.importSelfBtn) {
          elements.importSelfBtn.click();
        } else {
          setStatus("Import tool is not available.");
        }
      };

      const triggerGenerateLibrary = () => {
        if (elements.generateLibraryBtn) {
          elements.generateLibraryBtn.click();
        } else {
          setStatus("Library tool is not available.");
        }
      };

      const triggerSaveLog = () => {
        if (elements.saveLogBtn) {
          elements.saveLogBtn.click();
        } else {
          setStatus("Log tool is not available.");
        }
      };

      const formCard = document.createElement("div");
      formCard.className = "project-card";
      const formTitle = document.createElement("h4");
      formTitle.textContent = t("project_meta_title", "Project Metadata");
      formCard.appendChild(formTitle);
      if (runtime.awaitingLocaleBootstrap) {
        const bootstrapHint = document.createElement("p");
        bootstrapHint.className = "project-card__hint";
        bootstrapHint.textContent = t(
          "project_bootstrap_hint",
          "This folder is empty. Select report language to load seed content."
        );
        formCard.appendChild(bootstrapHint);
      }
      const formGrid = document.createElement("div");
      formGrid.className = "project-meta-grid";
      const scheduleProjectAutosave = () => {
        if (runtime.awaitingLocaleBootstrap) return;
        scheduleAutosave();
      };

      const createMetaField = (labelText, value, onInput, options = {}) => {
        const label = document.createElement("label");
        label.className = "project-meta-field";
        const span = document.createElement("span");
        span.textContent = labelText;
        label.appendChild(span);
        const input = options.select ? document.createElement("select") : document.createElement("input");
        if (options.select) {
          (options.items || []).forEach((item) => {
            const option = document.createElement("option");
            option.value = item.value;
            option.textContent = item.label;
            input.appendChild(option);
          });
          input.value = value;
        } else {
          input.type = options.type || "text";
          input.value = value;
          if (options.placeholder) input.placeholder = options.placeholder;
          if (options.min != null) input.min = String(options.min);
          if (options.step != null) input.step = String(options.step);
        }
        if (options.select) {
          input.addEventListener("change", () => onInput(input.value));
        } else {
          input.addEventListener("input", () => onInput(input.value));
        }
        label.appendChild(input);
        if (options.colSpan) {
          const span = Math.max(1, Math.floor(Number(options.colSpan) || 1));
          label.style.gridColumn = `span ${span}`;
        }
        return label;
      };

      formGrid.appendChild(createMetaField(t("project_meta_moderator", "Moderator"), meta.moderator || "", (value) => {
        meta.moderator = String(value || "").trim();
        meta.author = meta.moderator;
        scheduleProjectAutosave();
      }, { colSpan: 5 }));
      formGrid.appendChild(createMetaField(t("project_meta_moderator_initials", "Moderator initials"), meta.moderatorInitials || "", (value) => {
        meta.moderatorInitials = String(value || "").trim();
        meta.initials = meta.moderatorInitials;
        scheduleProjectAutosave();
      }, { colSpan: 5 }));
      formGrid.appendChild(createMetaField(t("project_meta_co_moderator", "Co-moderator"), meta.coModerator || "", (value) => {
        meta.coModerator = String(value || "").trim();
        scheduleProjectAutosave();
      }, { colSpan: 5 }));
      formGrid.appendChild(createMetaField(t("project_meta_co_initials", "Co-moderator initials"), meta.coModeratorInitials || "", (value) => {
        meta.coModeratorInitials = String(value || "").trim();
        scheduleProjectAutosave();
      }, { colSpan: 5 }));
      formGrid.appendChild(createMetaField(t("project_meta_company", "Company"), meta.company || "", (value) => {
        meta.company = String(value || "").trim();
        scheduleProjectAutosave();
      }, { colSpan: 3 }));
      formGrid.appendChild(createMetaField(t("project_meta_companyid", "Company ID"), meta.companyId || "", (value) => {
        meta.companyId = String(value || "").trim();
        scheduleProjectAutosave();
      }, { colSpan: 2 }));
      formGrid.appendChild(createMetaField(t("project_meta_address", "Address"), meta.address || "", (value) => {
        meta.address = String(value || "").trim();
        scheduleProjectAutosave();
      }, { colSpan: 2 }));
      formGrid.appendChild(createMetaField(t("project_meta_postal", "PLZ"), meta.plz || "", (value) => {
        meta.plz = String(value || "").trim();
        scheduleProjectAutosave();
      }, { colSpan: 1 }));
      formGrid.appendChild(createMetaField(t("project_meta_city", "City"), meta.city || "", (value) => {
        meta.city = String(value || "").trim();
        scheduleProjectAutosave();
      }, { colSpan: 2 }));
      formGrid.appendChild(createMetaField(t("project_meta_locale", "Locale"), meta.locale || "", (value) => {
        const nextLocale = String(value || "").trim();
        meta.locale = nextLocale;
        if (nextLocale) {
          ctx.i18n.setLocale(nextLocale);
        }
        const shouldBootstrap = runtime.awaitingLocaleBootstrap
          && Array.isArray(state.project?.chapters)
          && state.project.chapters.length === 0
          && typeof seedBootstrapHandler === "function"
          && !!nextLocale;
        if (shouldBootstrap) {
          seedBootstrapHandler(nextLocale, { deferSave: true }).catch((err) => {
            setStatus(`Seed bootstrap failed: ${err.message || err}`);
            debug.logLine("error", `Seed bootstrap failed: ${err.message || err}`);
          });
          return;
        }
        scheduleProjectAutosave();
      }, {
        colSpan: 5,
        select: true,
        items: [
          { value: "", label: t("project_meta_locale_select", "Select language") },
          { value: "de-CH", label: "DE" },
          { value: "fr-CH", label: "FR" },
          { value: "it-CH", label: "IT" },
        ],
      }));
      formGrid.appendChild(createMetaField(t("project_meta_backup", "Auto-backup (minutes)"), String(meta.autobackupMinutes ?? 30), (value) => {
        const raw = Number(value);
        meta.autobackupMinutes = Number.isFinite(raw) ? Math.max(0, Math.floor(raw)) : 30;
        scheduleProjectAutosave();
      }, {
        colSpan: 5,
        type: "number",
        min: 0,
        step: 1,
      }));
      formCard.appendChild(formGrid);

      const created = document.createElement("div");
      created.className = "project-meta-created";
      created.textContent = `${t("project_meta_created", "Created")}: ${formatIsoDate(meta.createdAt) || "—"}`;
      formCard.appendChild(created);
      page.appendChild(formCard);

      const importCard = createToolCard({
        title: t("project_tool_import_title", "Import Selbstbeurteilung"),
        hint: t("project_tool_import_hint", "Pick the Selbstbeurteilung file. It will be imported and copied into the project inputs folder."),
        buttonLabel: t("project_tool_import_self", "Import Selbstbeurteilung"),
        buttonClass: "",
        onClick: async () => {
          triggerImportSelf();
        },
      });
      page.appendChild(importCard.card);

      const exportCard = createToolCard({
        title: t("project_export_card_title", "Word Export"),
        hint: t("project_export_card_hint", "Pick the DOCX template and create a report in outputs using the current sidecar data."),
        buttonLabel: t("project_export_no_vba", "Word Export (No VBA)"),
        buttonClass: "ghost",
        onClick: async (button) => {
          if (!runtime.dirHandle) {
            setStatus(t("project_open_folder_first", "Open project folder first."));
            return;
          }
          const exporter = window.AutoBerichtWordExport || {};
          if (!exporter.exportReportDocx) {
            setStatus(t("project_export_missing", "No-VBA exporter not available."));
            return;
          }
          button.disabled = true;
          const previous = button.textContent;
          button.textContent = t("project_export_running", "Exporting ...");
          try {
            const result = await exporter.exportReportDocx({
              project: state.project,
              sidecarDoc: runtime.sidecarDoc,
              projectHandle: runtime.dirHandle,
              spiderOverrides: state.spiderOverrides || {},
              computeSpider: window.AutoBerichtSpider?.computeSpider,
              compareIdSegments: stateHelpers.compareIdSegments,
              toText: stateHelpers.toText,
            });
            if (result?.savedAs) {
              setStatus(`Exported ${result.savedAs}`);
            } else {
              setStatus(t("project_export_done", "Export complete."));
            }
          } catch (err) {
            setStatus(`Export failed: ${err.message || err}`);
            debug.logLine("error", `No-VBA export failed: ${err.message || err}`);
          } finally {
            button.disabled = false;
            button.textContent = previous;
          }
        },
      });
      page.appendChild(exportCard.card);

      const libraryCard = createToolCard({
        title: t("project_tool_library_title", "Library Export"),
        hint: t("project_tool_library_hint", "Generate or update the library from the current project content."),
        buttonLabel: t("project_tool_library", "Generate / Update Library"),
        buttonClass: "ghost",
        onClick: async () => {
          triggerGenerateLibrary();
        },
      });
      page.appendChild(libraryCard.card);

      const libraryExcelCard = createToolCard({
        title: t("project_tool_library_excel_title", "Library Excel Export"),
        hint: t("project_tool_library_excel_hint", "Export the current library JSON into an Excel workbook (human-readable and machine-ingestible)."),
        buttonLabel: t("project_tool_library_excel", "Export Library Excel"),
        buttonClass: "ghost",
        onClick: async () => {
          if (typeof libraryExcelExportHandler !== "function") {
            setStatus(t("project_tool_library_excel_missing", "Library Excel exporter is not available."));
            return;
          }
          await libraryExcelExportHandler();
        },
      });
      page.appendChild(libraryExcelCard.card);

      if (migrationUiEnabled) {
        const migrationCard = createToolCard({
          title: t("project_tool_migration_title", "Sidecar Migration"),
          hint: t("project_tool_migration_hint", "Run one-time cleanup for legacy sidecar schema and save a backup in backup/."),
          buttonLabel: t("project_tool_migration", "Migrate Legacy Sidecar"),
          buttonClass: "ghost",
          onClick: async (button) => {
            if (typeof sidecarMigrationHandler !== "function") {
              setStatus(t("project_tool_migration_missing", "Migration handler is not available."));
              return;
            }
            button.disabled = true;
            const prev = button.textContent;
            button.textContent = t("project_tool_migration_running", "Migrating ...");
            try {
              await sidecarMigrationHandler();
            } finally {
              button.disabled = false;
              button.textContent = prev;
            }
          },
        });
        page.appendChild(migrationCard.card);
      }

      const logoCard = document.createElement("div");
      logoCard.className = "project-card";
      const logoTitle = document.createElement("h4");
      logoTitle.textContent = t("project_logo_title", "Logo Assets");
      logoCard.appendChild(logoTitle);
      const logoHint = document.createElement("p");
      logoHint.className = "project-card__hint";
      logoHint.textContent = t("project_logo_hint", "Place logo.png/logo.jpg in inputs, then click to pick the logo file manually. The app writes outputs/logo-large.png and outputs/logo-small.png.");
      logoCard.appendChild(logoHint);
      const logoPaths = document.createElement("div");
      logoPaths.className = "project-logo-paths";
      const renderLogoPaths = () => {
        const large = String(meta.logoLargePath || "outputs/logo-large.png");
        const small = String(meta.logoSmallPath || "outputs/logo-small.png");
        logoPaths.textContent = `Large: ${large}  |  Small: ${small}`;
      };
      renderLogoPaths();
      logoCard.appendChild(logoPaths);
      const logoActions = document.createElement("div");
      logoActions.className = "project-page__actions";
      const logoBtn = document.createElement("button");
      logoBtn.type = "button";
      logoBtn.className = "ghost";
      logoBtn.textContent = t("project_logo_import", "Import + Resize Logo");
      logoBtn.addEventListener("click", async () => {
        if (!runtime.dirHandle) {
          setStatus(t("project_open_folder_first", "Open project folder first."));
          return;
        }
        const exporter = window.AutoBerichtWordExport || {};
        if (!exporter.prepareLogosForProject) {
          setStatus(t("project_logo_missing", "Logo pipeline not available."));
          return;
        }
        logoBtn.disabled = true;
        const prev = logoBtn.textContent;
        logoBtn.textContent = t("project_logo_processing", "Processing ...");
        try {
          const result = await exporter.prepareLogosForProject({
            projectHandle: runtime.dirHandle,
            meta,
          });
          meta.logoLargePath = result.logoLargePath;
          meta.logoSmallPath = result.logoSmallPath;
          renderLogoPaths();
          scheduleAutosave();
          setStatus(t("project_logo_done", "Logo assets prepared."));
        } catch (err) {
          setStatus(`Logo processing failed: ${err.message || err}`);
          debug.logLine("error", `Logo processing failed: ${err.message || err}`);
        } finally {
          logoBtn.disabled = false;
          logoBtn.textContent = prev;
        }
      });
      logoActions.appendChild(logoBtn);
      logoCard.appendChild(logoActions);
      page.appendChild(logoCard);

      const spiderCard = document.createElement("div");
      spiderCard.className = "project-card";
      const spiderTitle = document.createElement("h4");
      spiderTitle.textContent = t("project_spider_title", "Spider Editor");
      spiderCard.appendChild(spiderTitle);
      const spiderHint = document.createElement("p");
      spiderHint.className = "project-card__hint";
      spiderHint.textContent = t("project_spider_hint", "Adjust chapter overrides and preview the live spider chart.");
      spiderCard.appendChild(spiderHint);
      const canvas = document.createElement("canvas");
      canvas.className = "project-spider-canvas";
      canvas.width = 980;
      canvas.height = 560;
      spiderCard.appendChild(canvas);
      const spiderTableWrap = document.createElement("div");
      spiderTableWrap.className = "spider-table-wrapper";
      const spiderTable = document.createElement("table");
      spiderTable.className = "spider-table";
      const getCompanyLegendLabel = () => String(meta.company || "").trim() || "Company";
      const getSpiderDisplayLabel = (row) => {
        const chapterId = String(row?.id || "");
        const chapter = (state.project?.chapters || []).find((item) => String(item?.id || "") === chapterId);
        if (!chapter) return chapterId;
        return stateHelpers.formatChapterLabel(chapter, state.project?.meta?.locale);
      };
      spiderTable.innerHTML = [
        "<thead><tr>",
        "<th>Kapitel</th>",
        `<th>${escapeHtml(getCompanyLegendLabel())} %</th>`,
        "<th>Override</th>",
        "<th>Use</th>",
        "<th>Suva %</th>",
        "<th>Override</th>",
        "<th>Use</th>",
        "</tr></thead>",
      ].join("");
      const spiderBody = document.createElement("tbody");
      spiderTable.appendChild(spiderBody);
      spiderTableWrap.appendChild(spiderTable);
      spiderCard.appendChild(spiderTableWrap);
      const spiderActions = document.createElement("div");
      spiderActions.className = "project-page__actions";
      const applySpiderBtn = document.createElement("button");
      applySpiderBtn.type = "button";
      applySpiderBtn.textContent = t("project_spider_apply", "Apply Spider Overrides");
      const recalcSpiderBtn = document.createElement("button");
      recalcSpiderBtn.type = "button";
      recalcSpiderBtn.className = "ghost";
      recalcSpiderBtn.textContent = t("project_spider_recalc", "Recalculate");
      spiderActions.appendChild(applySpiderBtn);
      spiderActions.appendChild(recalcSpiderBtn);
      spiderCard.appendChild(spiderActions);

      const renderSpider = async () => {
        const spider = window.AutoBerichtSpider || {};
        if (!spider.computeSpider) {
          const row = document.createElement("tr");
          row.innerHTML = "<td colspan=\"7\">Spider module unavailable.</td>";
          spiderBody.innerHTML = "";
          spiderBody.appendChild(row);
          return;
        }
        spiderBody.innerHTML = "";
        try {
          const spiderData = await spider.computeSpider({
            project: state.project,
            overrides: state.spiderOverrides || {},
            dirHandle: runtime.dirHandle,
          });
          const rows = spiderData.effective?.chapters_1_11 || [];
          const baseline = spiderData.baseline?.chapters_1_11 || [];
          const baselineMap = new Map(baseline.map((item) => [item.id, item]));
          const displayRows = rows.map((row) => ({
            ...row,
            displayLabel: getSpiderDisplayLabel(row),
          }));
          displayRows.forEach((row) => {
            const base = baselineMap.get(row.id) || row;
            const ov = state.spiderOverrides?.[row.id] || {};
            const tr = document.createElement("tr");
            tr.innerHTML = [
              `<td>${escapeHtml(row.displayLabel || row.id)}</td>`,
              `<td>${Number(base.company || 0).toFixed(1)}%</td>`,
              `<td><input type="number" step="0.1" min="0" max="100" data-field="company" data-id="${row.id}" value="${Number.isFinite(ov.company) ? ov.company : ""}"></td>`,
              `<td><input type="checkbox" data-field="useCompany" data-id="${row.id}" ${ov.useCompany ? "checked" : ""}></td>`,
              `<td>${Number(base.consultant || 0).toFixed(1)}%</td>`,
              `<td><input type="number" step="0.1" min="0" max="100" data-field="consultant" data-id="${row.id}" value="${Number.isFinite(ov.consultant) ? ov.consultant : ""}"></td>`,
              `<td><input type="checkbox" data-field="useConsultant" data-id="${row.id}" ${ov.useConsultant ? "checked" : ""}></td>`,
            ].join("");
            spiderBody.appendChild(tr);
          });
          drawSpiderCanvas(canvas, displayRows, getCompanyLegendLabel());
        } catch (err) {
          const row = document.createElement("tr");
          row.innerHTML = `<td colspan="7">Spider failed: ${escapeHtml(err.message || String(err))}</td>`;
          spiderBody.appendChild(row);
          debug.logLine("error", `Project spider failed: ${err.message || err}`);
        }
      };

      applySpiderBtn.addEventListener("click", async () => {
        state.spiderOverrides = parseSpiderRowOverrides(spiderBody);
        scheduleAutosave();
        setStatus(t("project_spider_saved", "Spider overrides saved."));
        await renderSpider();
      });
      recalcSpiderBtn.addEventListener("click", async () => {
        await renderSpider();
      });

      page.appendChild(spiderCard);

      const debugCard = createToolCard({
        title: t("project_tool_log_title", "Debug Log"),
        hint: t("project_tool_log_hint", "Save the current debug log to share diagnostics when troubleshooting."),
        buttonLabel: t("project_tool_log", "Save debug log"),
        buttonClass: "ghost",
        onClick: async () => {
          triggerSaveLog();
        },
      });
      page.appendChild(debugCard.card);

      rowsEl.appendChild(page);
      renderSpider().catch((err) => {
        debug.logLine("error", `Project spider initial render failed: ${err.message || err}`);
      });
    };

    const renderChapterList = () => {
      chapterListEl.innerHTML = "";
      if (elements.sidebarProjectBtn) {
        if (!elements.sidebarProjectBtn.dataset.bound) {
          elements.sidebarProjectBtn.addEventListener("click", () => {
            state.selectedChapterId = PROJECT_VIEW_ID;
            render();
          });
          elements.sidebarProjectBtn.dataset.bound = "1";
        }
        elements.sidebarProjectBtn.classList.toggle("active", state.selectedChapterId === PROJECT_VIEW_ID);
      }
      const orderedChapters = [...state.project.chapters].sort((a, b) =>
        stateHelpers.compareIdSegments(a.id, b.id),
      );
      orderedChapters.forEach((chapter) => {
        const button = document.createElement("button");
        const buttonLabel = stateHelpers.formatChapterLabel(chapter, state.project?.meta?.locale);
        button.textContent = buttonLabel;
        button.title = buttonLabel;
        button.className = chapter.id === state.selectedChapterId ? "active" : "";
        button.addEventListener("click", () => {
          const leavingProject = state.selectedChapterId === PROJECT_VIEW_ID && chapter.id !== PROJECT_VIEW_ID;
          if (leavingProject && runtime.pendingBootstrapWrite) {
            scheduleAutosave();
          }
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
      const isProjectView = chapter.id === PROJECT_VIEW_ID;
      const title = document.createElement("span");
      title.textContent = isProjectView
        ? t("project_page_nav", "Project")
        : stateHelpers.formatChapterLabel(chapter, state.project?.meta?.locale);
      const pill = createPhotoPill(chapter.id);
      if (pill) chapterTitleEl.appendChild(pill);
      chapterTitleEl.appendChild(title);
      if (isProjectView) return;
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
      const headerAux = document.createElement("div");
      headerAux.className = "row-header__aux";
      const headerStable = document.createElement("div");
      headerStable.className = "row-header__stable";
      if (score !== null && score !== undefined) {
        const scoreBadge = document.createElement("span");
        scoreBadge.className = "score-badge";
        scoreBadge.textContent = `${score}%`;
        headerAux.appendChild(scoreBadge);
      }
      if (options.photoPill) {
        headerAux.appendChild(options.photoPill);
      }
      if (options.reorderControls) {
        headerAux.appendChild(options.reorderControls);
      }
      const comments = stateHelpers.getAnswerComments(row);
      if (comments.length) {
        headerAux.appendChild(createBadgeTooltip("Comment", comments.join("\n"), "comment-badge"));
      }
      const evidence = stateHelpers.getAnswerEvidence(row);
      if (evidence.length) {
        headerAux.appendChild(createBadgeTooltip("Evidence", evidence.join("\n"), "evidence-badge"));
      }
      headerStable.appendChild(createAnswerBadge(row));
      headerStable.appendChild(previewBtn);
      if (headerAux.childElementCount > 0) {
        headerRight.appendChild(headerAux);
      }
      headerRight.appendChild(headerStable);
      header.appendChild(meta);
      header.appendChild(headerRight);
      return { header, previewBtn };
    };

    const createSelfAssessmentDetails = (row, options = {}) => {
      if (options.isSummary === true || row.type === "summary") return null;
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

    const createPriorityControls = (row, ws) => {
      const group = document.createElement("div");
      group.className = "priority-group";
      const title = document.createElement("span");
      title.className = "priority-group__label";
      title.textContent = "Prio";
      group.appendChild(title);
      const current = Number(ws.priority);
      const hasPriority = Number.isFinite(current) && current >= 0 && current <= 4;
      const options = [1, 2, 3, 4, 0];
      options.forEach((value) => {
        const label = document.createElement("label");
        label.className = "level-radio priority-radio";
        const input = document.createElement("input");
        input.type = "radio";
        input.name = `priority-${row.id}`;
        input.value = String(value);
        input.checked = hasPriority ? current === value : value === 0;
        input.addEventListener("change", () => {
          ws.priority = value;
          scheduleAutosave();
        });
        label.appendChild(input);
        label.appendChild(document.createTextNode(String(value)));
        group.appendChild(label);
      });
      return group;
    };

    const createLibraryGroup = (currentAction, onSelect, labelText = "Library") => {
      const libraryGroup = document.createElement("div");
      libraryGroup.className = "library-group";
      const libraryLabel = document.createElement("span");
      libraryLabel.textContent = labelText;
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
        if ((currentAction || "off") === option.key) {
          button.classList.add("is-active");
        }
        button.addEventListener("click", () => onSelect(option.key));
        libraryGroup.appendChild(button);
      });
      return libraryGroup;
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
      findingControls.className = "field-controls finding-controls";
      findingControls.appendChild(includeToggle);
      findingControls.appendChild(doneToggle);
      findingControls.appendChild(createPriorityControls(row, ws));

      const findingArea = document.createElement("textarea");
      applyEditorLocale(findingArea);
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

      const libraryGroup = createLibraryGroup(
        ws.libraryAction || "off",
        (nextAction) => {
          ws.libraryAction = nextAction;
          scheduleAutosave();
          renderRows();
        },
        "Library",
      );

      const recommendationInput = document.createElement("textarea");
      applyEditorLocale(recommendationInput);
      const recommendationText = stateHelpers.getRecommendationText(row);
      recommendationInput.value = recommendationText;
      recommendationInput.addEventListener("input", () => {
        ws.recommendationText = recommendationInput.value;
        scheduleAutosave();
        autosizeTextarea(recommendationInput);
        if (onChange) onChange();
      });
      requestAnimationFrame(() => autosizeTextarea(recommendationInput));

      const recommendationControls = document.createElement("div");
      recommendationControls.className = "recommendation-controls";
      recommendationControls.appendChild(recControls);
      recommendationControls.appendChild(libraryGroup);

      recHeader.appendChild(recLabel);
      recHeader.appendChild(recommendationControls);
      recField.appendChild(recHeader);
      recField.appendChild(recommendationInput);
      return recField;
    };

    const createSummaryField = (row, ws, onChange) => {
      const summaryField = document.createElement("div");
      summaryField.className = "field";
      const summaryHeader = document.createElement("div");
      summaryHeader.className = "field-header";
      const summaryLabel = document.createElement("label");
      summaryLabel.textContent = "Summary";
      summaryLabel.appendChild(createMarkdownHint());

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

      const summaryControls = document.createElement("div");
      summaryControls.className = "field-controls";
      summaryControls.appendChild(includeToggle);
      summaryControls.appendChild(doneToggle);
      summaryControls.appendChild(createLibraryGroup(
        ws.libraryAction || "off",
        (nextAction) => {
          ws.libraryAction = nextAction;
          scheduleAutosave();
          renderRows();
        },
        "Library",
      ));

      const summaryInput = document.createElement("textarea");
      applyEditorLocale(summaryInput);
      summaryInput.value = stateHelpers.getRecommendationText(row);
      summaryInput.addEventListener("input", () => {
        ws.recommendationText = summaryInput.value;
        scheduleAutosave();
        autosizeTextarea(summaryInput);
        if (onChange) onChange();
      });
      requestAnimationFrame(() => autosizeTextarea(summaryInput));

      summaryHeader.appendChild(summaryLabel);
      summaryHeader.appendChild(summaryControls);
      summaryField.appendChild(summaryHeader);
      summaryField.appendChild(summaryInput);
      return summaryField;
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
          `${previewWordLikeHtml(findingText || "")}`,
        ].join("\n");
        recCol.innerHTML = [
          `<strong>${escapeHtml(t("preview_recommendation", "Recommendation"))}</strong>`,
          `${previewWordLikeHtml(recommendation || "")}`,
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
      const isSummaryRow = String(options.chapter?.id || "") === "0";
      if (shouldFilterRow(row, ws)) return null;
      const card = document.createElement("div");
      card.className = "row-card";
      const score = isSummaryRow ? null : stateHelpers.calculateScore(row);
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
      const showPhotoPill = !isSummaryRow
        && !options.hidePhotoPill
        && (isObservationRow || !hasParentSectionPill);
      const headerPayload = createRowHeader(row, score, {
        photoPill: showPhotoPill ? createPhotoPill(photoTag, photoKind) : null,
        reorderControls: options.reorderControls,
        displayId: options.displayId,
      });
      const selfDetails = createSelfAssessmentDetails(row, { isSummary: isSummaryRow });
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
      rowBody.className = isSummaryRow ? "row-body row-body--single" : "row-body";
      if (isSummaryRow) {
        rowBody.appendChild(createSummaryField(row, ws, onChange));
      } else {
        rowBody.appendChild(createFindingField(row, ws, onChange));
        rowBody.appendChild(createRecommendationField(row, ws, onChange));
      }
      card.appendChild(rowBody);
      card.appendChild(previewPanel.preview);
      return card;
    };

    const createChapterPositivesCard = (chapter) => {
      if (!chapter || String(chapter.id || "") === "0") return null;
      if (typeof normalizeHelpers.ensureChapterMetaDefaults === "function") {
        normalizeHelpers.ensureChapterMetaDefaults(chapter);
      } else {
        chapter.meta = chapter.meta || {};
        if (chapter.meta.positivesText == null) chapter.meta.positivesText = "";
        if (chapter.meta.positivesInclude == null) chapter.meta.positivesInclude = false;
        if (chapter.meta.positivesDone == null) chapter.meta.positivesDone = false;
        if (!chapter.meta.positivesLibraryAction) chapter.meta.positivesLibraryAction = "off";
        if (chapter.meta.positivesLibraryHash == null) chapter.meta.positivesLibraryHash = "";
      }
      const meta = chapter.meta;
      const card = document.createElement("div");
      card.className = "row-card chapter-positives-card";

      const header = document.createElement("div");
      header.className = "field-header";
      const label = document.createElement("label");
      label.textContent = "Positives";
      label.appendChild(createMarkdownHint());

      const includeToggle = createToggle("Include", meta.positivesInclude === true, (checked) => {
        meta.positivesInclude = checked;
        scheduleAutosave();
        renderRows();
      });
      const doneToggle = createToggle("Done", meta.positivesDone === true, (checked) => {
        meta.positivesDone = checked;
        scheduleAutosave();
        renderRows();
      });

      const controls = document.createElement("div");
      controls.className = "field-controls";
      controls.appendChild(includeToggle);
      controls.appendChild(doneToggle);
      controls.appendChild(createLibraryGroup(
        meta.positivesLibraryAction || "off",
        (nextAction) => {
          meta.positivesLibraryAction = nextAction;
          scheduleAutosave();
          renderRows();
        },
        "Library",
      ));

      const input = document.createElement("textarea");
      applyEditorLocale(input);
      input.value = stateHelpers.getChapterPositivesText
        ? stateHelpers.getChapterPositivesText(chapter)
        : String(meta.positivesText || "");
      input.addEventListener("input", () => {
        meta.positivesText = input.value;
        scheduleAutosave();
        autosizeTextarea(input);
      });
      requestAnimationFrame(() => autosizeTextarea(input));

      header.appendChild(label);
      header.appendChild(controls);
      card.appendChild(header);
      card.appendChild(input);
      return card;
    };

    const renderRows = () => {
      rowsEl.innerHTML = "";
      if (state.selectedChapterId === PROJECT_VIEW_ID) {
        renderChapterTitle({
          id: PROJECT_VIEW_ID,
          title: { de: t("project_page_nav", "Project") },
        });
        renderProjectPage();
        return;
      }
      const chapter = state.project.chapters.find((c) => c.id === state.selectedChapterId);
      if (!chapter) return;
      renderChapterTitle(chapter);
      const chapterPositivesCard = createChapterPositivesCard(chapter);
      if (chapterPositivesCard) rowsEl.appendChild(chapterPositivesCard);
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
      setSeedBootstrapHandler,
      setLibraryExcelExportHandler,
      setSidecarMigrationHandler,
      renderPhotoOverlay,
    };
  };

  window.AutoBerichtRender = { init };
})();
