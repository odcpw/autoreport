(() => {
  const init = (ctx, deps) => {
    const { state, runtime, setStatus } = ctx;
    const { elements, tagsApi, photosApi, actions } = deps;

    const updateStatusVisibility = (isHidden) => {
      elements.statusEl?.classList.toggle("is-hidden", isHidden);
    };

    const setStatusHidden = (hidden) => {
      if (window.localStorage) {
        window.localStorage.setItem("photosorterStatusHidden", hidden ? "1" : "0");
      }
      updateStatusVisibility(hidden);
    };

    const updateMeta = () => {
      const filtered = photosApi.getFilteredPhotos();
      const total = state.photos.length;
      const unsorted = state.photos.filter(photosApi.isPhotoUnsorted).length;
      const current = filtered.length ? state.currentIndex + 1 : 0;
      if (elements.photoMetaEl) {
        elements.photoMetaEl.textContent = `Image ${current} of ${filtered.length} • Total ${total} • Unsorted ${unsorted}`;
      }
    };

    const buildTagCounts = () => {
      const counts = {
        report: new Map(),
        observations: new Map(),
        training: new Map(),
      };
      state.photos.forEach((photo) => {
        ["report", "observations", "training"].forEach((group) => {
          (photo.tags?.[group] || []).forEach((tag) => {
            const key = String(tag || "").trim();
            if (!key) return;
            counts[group].set(key, (counts[group].get(key) || 0) + 1);
          });
        });
      });
      return counts;
    };

    const hasAnyActiveFilters = () => (
      state.filterMode !== "all" || !!photosApi.hasActiveTagFilters?.()
    );

    const createFilterIcon = (withSlash = false) => {
      const ns = "http://www.w3.org/2000/svg";
      const svg = document.createElementNS(ns, "svg");
      svg.setAttribute("viewBox", "0 0 24 24");
      svg.setAttribute("aria-hidden", "true");
      svg.classList.add("filter-icon");

      const funnel = document.createElementNS(ns, "path");
      funnel.setAttribute("d", "M3 4h18l-7.2 8.4v6.5l-3.6-2.1v-4.4z");
      svg.appendChild(funnel);

      if (withSlash) {
        const slash = document.createElementNS(ns, "path");
        slash.setAttribute("d", "M4 20L20 4");
        slash.setAttribute("class", "filter-icon__slash");
        svg.appendChild(slash);
      }

      return svg;
    };

    const createSplitTagButton = ({
      group,
      option,
      labelText,
      selected,
      filtered,
      canTag = true,
      chapter = false,
    }) => {
      const wrapper = document.createElement("div");
      wrapper.className = `tag-split${chapter ? " tag-split--chapter" : ""}`;

      const filterBtn = document.createElement("button");
      filterBtn.type = "button";
      filterBtn.className = "tag-split__filter";
      filterBtn.setAttribute("aria-label", `Filter ${option.label}`);
      filterBtn.setAttribute("aria-pressed", String(filtered));
      filterBtn.title = filtered ? `Filter active: ${option.label}` : `Filter by ${option.label}`;
      if (filtered) filterBtn.classList.add("is-active");
      filterBtn.appendChild(createFilterIcon(false));
      filterBtn.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        actions.toggleFilterTag(group, option.value);
      });

      const tagBtn = document.createElement("button");
      tagBtn.type = "button";
      tagBtn.className = "tag-split__tag";
      tagBtn.textContent = labelText;
      tagBtn.title = labelText;
      tagBtn.setAttribute("aria-pressed", String(selected));
      if (selected) tagBtn.classList.add("active");
      tagBtn.disabled = !canTag;
      tagBtn.addEventListener("click", () => {
        actions.toggleTag(group, option.value);
      });

      wrapper.append(filterBtn, tagBtn);
      return wrapper;
    };

    const renderViewer = () => {
      const current = photosApi.getCurrentPhoto();
      if (!current) {
        if (elements.photoImageEl) {
          elements.photoImageEl.removeAttribute("src");
          elements.photoImageEl.alt = "No photo loaded";
        }
        if (elements.photoFilenameEl) elements.photoFilenameEl.textContent = "";
        if (elements.notesEl) {
          elements.notesEl.value = "";
          elements.notesEl.disabled = true;
        }
        if (elements.photoUnsortedBtn) {
          elements.photoUnsortedBtn.classList.remove("active");
          elements.photoUnsortedBtn.disabled = true;
        }
        if (elements.photoClearFiltersBtn) {
          const hasFilters = hasAnyActiveFilters();
          elements.photoClearFiltersBtn.classList.toggle("active", hasFilters);
          elements.photoClearFiltersBtn.disabled = !hasFilters;
        }
        updateMeta();
        return;
      }
      if (elements.photoFilenameEl) {
        elements.photoFilenameEl.textContent = current.path.split("/").pop();
      }
      if (elements.notesEl) {
        const locale = document.documentElement?.getAttribute("lang") || "de-CH";
        const spellLang = ctx.i18n?.resolveSpellcheckLang
          ? ctx.i18n.resolveSpellcheckLang(locale)
          : String(locale).toLowerCase().split("-")[0] || "en";
        elements.notesEl.setAttribute("lang", spellLang);
        elements.notesEl.setAttribute("spellcheck", "true");
        elements.notesEl.spellcheck = true;
        elements.notesEl.value = current.notes || "";
        elements.notesEl.disabled = false;
      }
      if (elements.photoUnsortedBtn) {
        const isUnsorted = photosApi.isPhotoUnsorted(current);
        elements.photoUnsortedBtn.classList.toggle("active", isUnsorted);
        elements.photoUnsortedBtn.disabled = false;
      }
      if (elements.photoClearFiltersBtn) {
        const hasFilters = hasAnyActiveFilters();
        elements.photoClearFiltersBtn.classList.toggle("active", hasFilters);
        elements.photoClearFiltersBtn.disabled = !hasFilters;
      }
      updateMeta();
      photosApi.loadPhotoUrl(current).then((url) => {
        if (!url) return;
        if (elements.photoImageEl) {
          elements.photoImageEl.src = url;
          elements.photoImageEl.alt = current.path;
        }
      }).catch((err) => {
        setStatus(`Photo load failed: ${err.message}`);
      });
    };

    const renderTagPanel = (group, config) => {
      const container = elements.panels?.[group];
      if (!container) return;

      container.innerHTML = "";
      const title = document.createElement("h3");
      title.textContent = config.title;
      const description = document.createElement("p");
      description.textContent = config.description;

      const controls = document.createElement("div");
      controls.className = "panel__controls";
      const filterInput = document.createElement("input");
      filterInput.type = "text";
      filterInput.placeholder = "Filter tags";
      filterInput.value = config.filter || "";
      filterInput.addEventListener("input", () => {
        state.tagFilters[group] = filterInput.value;
        renderPanels();
      });

      controls.append(filterInput);
      if (config.allowAdd) {
        const addInput = document.createElement("input");
        addInput.type = "text";
        addInput.placeholder = "Add new tag";
        addInput.addEventListener("keydown", (event) => {
          if (event.key !== "Enter") return;
          event.preventDefault();
          const value = addInput.value.trim();
          if (!value) return;
          const existing = state.tagOptions[group] || [];
          if (!existing.some((option) => option.value === value)) {
            const next = tagsApi.sortOptionsForGroup(group, [...existing, { value, label: value }]);
            state.tagOptions[group] = next;
            actions.persistTagOptions?.();
          }
          addInput.value = "";
          renderPanels();
        });
        controls.append(addInput);
      }

      const current = photosApi.getCurrentPhoto();
      const selected = new Set(current?.tags?.[group] || []);
      const filteredSet = new Set(state.activeTagFilters?.[group] || []);
      const canTag = !!current;
      const filteredOptions = (state.tagOptions?.[group] || [])
        .filter((option) => option.label.toLowerCase().includes((config.filter || "").toLowerCase()));

      const { chapters, rest } = config.splitChapters
        ? tagsApi.splitChapterOptions(filteredOptions)
        : { chapters: [], rest: filteredOptions };
      const counts = config.counts || new Map();
      const formatLabel = (option) => {
        if (!state.showTagCounts) return option.label;
        const count = counts.get(option.value) || 0;
        return `(${count}) ${option.label}`;
      };

      if (config.splitChapters) {
        const chapterRow = document.createElement("div");
        chapterRow.className = "panel__chapters";
        chapters.forEach((option) => {
          const labelText = formatLabel(option);
          chapterRow.appendChild(createSplitTagButton({
            group,
            option,
            labelText,
            selected: selected.has(option.value),
            filtered: filteredSet.has(option.value),
            chapter: true,
            canTag,
          }));
        });
        container.append(chapterRow);
      }

      const tagsEl = document.createElement("div");
      tagsEl.className = "panel__tags";

      rest.forEach((option) => {
        const labelText = formatLabel(option);
        tagsEl.appendChild(createSplitTagButton({
          group,
          option,
          labelText,
          selected: selected.has(option.value),
          filtered: filteredSet.has(option.value),
          chapter: false,
          canTag,
        }));
      });

      container.append(title, description, controls, tagsEl);
    };

    const renderPanels = () => {
      const counts = buildTagCounts();
      renderTagPanel("report", {
        title: "Bericht",
        description: "Kapitel & Unterkapitel (1.x / 1.2 / 4.8 etc.)",
        filter: state.tagFilters.report,
        splitChapters: true,
        allowAdd: false,
        counts: counts.report,
      });
      renderTagPanel("observations", {
        title: "4.8 Beobachtungen",
        description: "Diese Tags werden als Kapitel 4.8 im Bericht verwendet.",
        filter: state.tagFilters.observations,
        allowAdd: true,
        counts: counts.observations,
      });
      renderTagPanel("training", {
        title: "Training",
        description: "Seminar-/Schulungskategorien.",
        filter: state.tagFilters.training,
        allowAdd: false,
        counts: counts.training,
      });
    };

    const renderAllNow = () => {
      const filtered = photosApi.getFilteredPhotos();
      if (state.currentIndex >= filtered.length) {
        state.currentIndex = Math.max(0, filtered.length - 1);
      }
      if (elements.filterToggleBtn) {
        elements.filterToggleBtn.textContent = state.filterMode === "all" ? "Show Unsorted" : "Show All";
      }
      if (elements.countToggleBtn) {
        elements.countToggleBtn.textContent = state.showTagCounts ? "Hide counts" : "Show counts";
        elements.countToggleBtn.classList.toggle("active", state.showTagCounts);
        elements.countToggleBtn.setAttribute("aria-pressed", String(state.showTagCounts));
      }
      renderViewer();
      renderPanels();
      actions.enableActions();
    };

    const scheduleRenderAll = () => {
      if (runtime.renderTimer) clearTimeout(runtime.renderTimer);
      runtime.renderTimer = setTimeout(() => {
        runtime.renderTimer = null;
        renderAllNow();
      }, 32);
    };

    const renderAll = () => {
      scheduleRenderAll();
    };

    const renderObservationTagList = () => {
      if (!elements.obsTagList) return;
      elements.obsTagList.innerHTML = "";
      const options = state.tagOptions?.observations || [];
      const counts = new Map();
      state.photos.forEach((photo) => {
        (photo.tags?.observations || []).forEach((tag) => {
          counts.set(tag, (counts.get(tag) || 0) + 1);
        });
      });

      options.forEach((option) => {
        const row = document.createElement("div");
        row.className = "settings-tags__row";
        const label = document.createElement("span");
        label.textContent = option.label;
        const count = document.createElement("span");
        count.className = "settings-tags__count";
        count.textContent = String(counts.get(option.value) || 0);
        const removeBtn = document.createElement("button");
        removeBtn.type = "button";
        removeBtn.textContent = "Remove";
        removeBtn.addEventListener("click", () => {
          const confirmed = window.confirm(
            `Remove "${option.label}"? This clears it from all photos and Chapter 4.8 in the report.`
          );
          if (!confirmed) return;
          actions.removeObservationTag(option.value);
        });
        row.append(label, count, removeBtn);
        elements.obsTagList.appendChild(row);
      });
    };

    return {
      updateStatusVisibility,
      setStatusHidden,
      renderViewer,
      renderPanels,
      renderAll,
      renderAllNow,
      renderObservationTagList,
    };
  };

  window.AutoBerichtPhotoSorterRender = { init };
})();
