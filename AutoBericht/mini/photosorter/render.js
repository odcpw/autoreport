(() => {
  const init = (ctx, deps) => {
    const { state, runtime, setStatus } = ctx;
    const { elements, tagsApi, photosApi, actions } = deps;

    const applyLayoutMode = () => {
      document.body.classList.toggle("layout-tabs", state.layoutMode === "tabs");
      document.body.classList.toggle("layout-stacked", state.layoutMode === "stacked");
    };

    const placeObservationsPanel = () => {
      const panel = elements.panels?.observations;
      if (!panel) return;
      const target = state.layoutMode === "tabs" ? elements.observationsSlot : elements.viewerObservations;
      if (!target || panel.parentElement === target) return;
      target.appendChild(panel);
    };

    const updateLayoutToggle = () => {
      elements.layoutToggleButtons?.forEach((button) => {
        const isActive = button.dataset.layout === state.layoutMode;
        button.classList.toggle("active", isActive);
        button.setAttribute("aria-pressed", String(isActive));
      });
    };

    const setLayoutMode = (mode) => {
      if (mode !== "tabs" && mode !== "stacked") return;
      state.layoutMode = mode;
      if (window.localStorage) {
        window.localStorage.setItem("photosorterLayout", mode);
      }
      applyLayoutMode();
      placeObservationsPanel();
      updateLayoutToggle();
      renderPanels();
    };

    const setActivePanel = (group) => {
      if (!elements.panels?.[group]) return;
      state.activePanel = group;
      elements.panelTabButtons?.forEach((btn) => {
        const isActive = btn.dataset.panelTab === group;
        btn.classList.toggle("active", isActive);
        btn.setAttribute("aria-pressed", String(isActive));
      });
      renderPanels();
    };

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

      container.classList.toggle("is-active", state.activePanel === group);
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
          }
          addInput.value = "";
          renderPanels();
        });
        controls.append(addInput);
      }

      const current = photosApi.getCurrentPhoto();
      const selected = new Set(current?.tags?.[group] || []);
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
          const button = document.createElement("button");
          button.type = "button";
          button.className = "tag-button tag-button--chapter";
          const labelText = formatLabel(option);
          button.textContent = labelText;
          button.title = labelText;
          if (selected.has(option.value)) {
            button.classList.add("active");
          }
          button.addEventListener("click", () => {
            actions.toggleTag(group, option.value);
          });
          chapterRow.appendChild(button);
        });
        container.append(chapterRow);
      }

      const tagsEl = document.createElement("div");
      tagsEl.className = "panel__tags";

      rest.forEach((option) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "tag-button";
        const labelText = formatLabel(option);
        button.textContent = labelText;
        button.title = labelText;
        if (selected.has(option.value)) {
          button.classList.add("active");
        }
        button.addEventListener("click", () => {
          actions.toggleTag(group, option.value);
        });
        tagsEl.appendChild(button);
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
      applyLayoutMode,
      placeObservationsPanel,
      updateLayoutToggle,
      setLayoutMode,
      setActivePanel,
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
