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
    } = elements;

    let scheduleAutosave = () => {};

    const setScheduleAutosave = (fn) => {
      scheduleAutosave = typeof fn === "function" ? fn : () => {};
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
      const items = Array.isArray(source.items) ? source.items : [];
      const filteredItems = items.filter((item) => {
        if (activeCategory) {
          if (getChecklistCategory(item) !== activeCategory) return false;
        }
        if (activeIndustry) {
          if (!item.industries?.includes(activeIndustry)) return false;
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
      if (chapter.id === "4.8") {
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

    const createFindingField = (row, ws) => {
      const findingField = document.createElement("div");
      findingField.className = "field";
      const findingHeader = document.createElement("div");
      findingHeader.className = "field-header";
      const findingLabel = document.createElement("label");
      findingLabel.textContent = "Finding";
      findingLabel.appendChild(createMarkdownHint());

      const findingToggle = createToggle("Override", ws.useFindingOverride, (checked) => {
        ws.useFindingOverride = checked;
        if (ws.useFindingOverride && !ws.findingOverride) {
          ws.findingOverride = stateHelpers.toText(row.master?.finding);
        }
        scheduleAutosave();
        renderRows();
      });
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
      findingControls.appendChild(findingToggle);
      findingControls.appendChild(includeToggle);
      findingControls.appendChild(doneToggle);

      const findingArea = document.createElement("textarea");
      findingArea.value = stateHelpers.getFindingText(row);
      findingArea.disabled = !ws.useFindingOverride;
      findingArea.addEventListener("input", () => {
        ws.findingOverride = findingArea.value;
        scheduleAutosave();
      });

      findingHeader.appendChild(findingLabel);
      findingHeader.appendChild(findingControls);
      findingField.appendChild(findingHeader);
      findingField.appendChild(findingArea);
      return findingField;
    };

    const createRecommendationField = (row, ws) => {
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
        label.appendChild(document.createTextNode(String(level)));
        levelGroup.appendChild(label);
      });

      const levelKey = String(ws.selectedLevel);
      const overrideToggle = createToggle("Override", !!ws.useLevelOverride?.[levelKey], (checked) => {
        ws.useLevelOverride[levelKey] = checked;
        if (ws.useLevelOverride[levelKey] && !ws.levelOverrides[levelKey]) {
          ws.levelOverrides[levelKey] = stateHelpers.toText(row.master?.levels?.[levelKey]);
        }
        scheduleAutosave();
        renderRows();
      });

      const recControls = document.createElement("div");
      recControls.className = "field-controls";
      recControls.appendChild(levelGroup);
      recControls.appendChild(overrideToggle);

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
        if ((ws.libraryActions?.[levelKey] || "off") === option.key) {
          button.classList.add("is-active");
        }
        button.addEventListener("click", () => {
          ws.libraryActions[levelKey] = option.key;
          scheduleAutosave();
          renderRows();
        });
        libraryGroup.appendChild(button);
      });

      const levelOverrideInput = document.createElement("textarea");
      const levelText = stateHelpers.getRecommendationText(row, levelKey);
      levelOverrideInput.value = levelText;
      levelOverrideInput.disabled = !ws.useLevelOverride[levelKey];
      levelOverrideInput.addEventListener("input", () => {
        ws.levelOverrides[levelKey] = levelOverrideInput.value;
        scheduleAutosave();
      });

      recHeader.appendChild(recLabel);
      recHeader.appendChild(recControls);
      recHeader.appendChild(libraryGroup);
      recField.appendChild(recHeader);
      recField.appendChild(levelOverrideInput);
      return recField;
    };

    const createPreviewPanel = (row, ws, previewBtn) => {
      const preview = document.createElement("div");
      preview.className = "row-preview";
      const findingText = ws.useFindingOverride ? ws.findingOverride : row.master?.finding;
      const levelKey = String(ws.selectedLevel);
      const recommendation = ws.useLevelOverride?.[levelKey]
        ? ws.levelOverrides?.[levelKey]
        : row.master?.levels?.[levelKey];

      const findingCol = document.createElement("div");
      findingCol.className = "row-preview__col";
      findingCol.innerHTML = [
        `<strong>${escapeHtml(t("preview_finding", "Finding"))}</strong>`,
        `<p>${markdownToHtml(findingText || "")}</p>`,
      ].join("\n");

      const recCol = document.createElement("div");
      recCol.className = "row-preview__col";
      recCol.innerHTML = [
        `<strong>${escapeHtml(t("preview_recommendation", "Recommendation"))}</strong>`,
        `<p>${markdownToHtml(recommendation || "")}</p>`,
      ].join("\n");

      preview.appendChild(findingCol);
      preview.appendChild(recCol);

      const showPreview = () => {
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
      return preview;
    };

    const shouldFilterRow = (row, ws) => {
      if (row.kind === "section") return false;
      const mode = state.filters.mode;
      if (mode === "include-only") return ws.includeFinding === true;
      if (mode === "done-only") return ws.done === true;
      if (mode === "hide-done") return ws.done !== true;
      if (mode === "hide-yes") {
        return stateHelpers.getAnswerState(row) === 1;
      }
      return false;
    };

    const createRowCard = (row, options = {}) => {
      normalizeHelpers.ensureWorkstateDefaults(row);
      const ws = row.workstate;
      if (shouldFilterRow(row, ws)) return null;
      const card = document.createElement("div");
      card.className = "row-card";
      const score = stateHelpers.calculateScore(row);
      const photoTag = row.sectionId || row.id;
      const showPhotoPill = row.type !== "summary" && !options.hidePhotoPill;
      const headerPayload = createRowHeader(row, score, {
        photoPill: showPhotoPill ? createPhotoPill(photoTag) : null,
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
      const rowBody = document.createElement("div");
      rowBody.className = "row-body";
      rowBody.appendChild(createFindingField(row, ws));
      rowBody.appendChild(createRecommendationField(row, ws));
      card.appendChild(rowBody);
      card.appendChild(createPreviewPanel(row, ws, headerPayload.previewBtn));
      return card;
    };

    const renderRows = () => {
      rowsEl.innerHTML = "";
      const chapter = state.project.chapters.find((c) => c.id === state.selectedChapterId);
      if (!chapter) return;
      renderChapterTitle(chapter);
      let rows = chapter.rows || [];
      let displayIds = new Map();
      if (chapter.id === "4.8") {
        rows = normalizeHelpers.orderObservationRows(chapter);
        rows.forEach((row, index) => {
          displayIds.set(row.id, `4.8.${index + 1}`);
        });
      }

      rows.forEach((row) => {
        if (row.kind === "section") {
          rowsEl.appendChild(createSectionRow(row));
          return;
        }
        const displayId = displayIds.get(row.id);
        const reorderControls = chapter.id === "4.8" ? createReorderControls(chapter, row) : null;
        const card = createRowCard(row, { chapter, displayId, reorderControls });
        if (card) rowsEl.appendChild(card);
      });
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
