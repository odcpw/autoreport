(() => {
  const bind = (ctx, deps) => {
    const {
      renderApi,
      ioApi,
      seeds,
      stateHelpers,
      normalizeHelpers,
      importSelfHandler,
      ensureFsAccess,
      enableActions,
      persistHandle,
      flushAutosave,
    } = deps;
    const { elements, state, runtime, debug, setStatus, i18n } = ctx;

    if (elements.pickFolderBtn) {
      elements.pickFolderBtn.addEventListener("click", async () => {
        if (!ensureFsAccess()) return;
        try {
          runtime.dirHandle = await window.showDirectoryPicker({ mode: "readwrite", id: "autobericht-project" });
          await persistHandle(runtime.dirHandle);
          enableActions();
          setStatus(`Selected folder: ${runtime.dirHandle.name}`);
          debug.logLine("info", `Selected folder: ${runtime.dirHandle.name}`);
          await ioApi.loadProjectFromFolder();
        } catch (err) {
          setStatus(`Folder pick canceled or failed: ${err.message}`);
          debug.logLine("warn", `Folder pick canceled or failed: ${err.message}`);
        }
      });
    }

    if (elements.loadSidecarBtn) {
      elements.loadSidecarBtn.addEventListener("click", async () => {
        if (!runtime.dirHandle) return;
        await ioApi.loadProjectFromFolder();
      });
    }

    if (elements.saveSidecarBtn) {
      elements.saveSidecarBtn.addEventListener("click", async () => {
        if (!runtime.dirHandle) return;
        try {
          await ioApi.saveSidecar();
          setStatus("Saved project_sidecar.json");
          debug.logLine("info", "Saved project_sidecar.json");
        } catch (err) {
          setStatus(`Save failed: ${err.message}`);
          debug.logLine("error", `Save failed: ${err.message}`);
        }
      });
    }

    if (elements.loadSeedsBtn) {
      elements.loadSeedsBtn.addEventListener("click", async () => {
        if (!runtime.dirHandle) return;
        try {
          state.project = await seeds.loadSeedsForProject({
            dirHandle: runtime.dirHandle,
            getLibraryFileName: () => stateHelpers.getLibraryFileName(state.project.meta || {}),
          });
          normalizeHelpers.ensureProjectMeta(state.project, i18n.setLocale);
          state.selectedChapterId = state.project.chapters[0]?.id || "";
          renderApi.render();
          setStatus("Loaded seed data.");
          debug.logLine("info", "Loaded seed data.");
        } catch (err) {
          setStatus(`Seed load failed: ${err.message}`);
          debug.logLine("error", `Seed load failed: ${err.message}`);
        }
      });
    }

    if (elements.saveLogBtn) {
      elements.saveLogBtn.addEventListener("click", async () => {
        try {
          const result = await debug.saveLog({
            suggestedName: `mini-editor-log-${new Date().toISOString().replace(/[:.]/g, "-")}.txt`,
            dirHandle: runtime.dirHandle,
          });
          setStatus(`Saved log (${result.location}): ${result.filename}`);
        } catch (err) {
          setStatus(`Log save failed: ${err.message}`);
        }
      });
    }

    if (elements.filterModeEls) {
      elements.filterModeEls.forEach((input) => {
        input.addEventListener("change", () => {
          if (!input.checked) return;
          state.filters.mode = input.value;
          renderApi.renderRows();
        });
      });
    }

    const openSettings = () => {
      if (!elements.settingsModal) return;
      elements.settingsModeratorEl.value = state.project.meta?.moderator || state.project.meta?.author || "";
      elements.settingsModeratorInitialsEl.value = state.project.meta?.moderatorInitials || state.project.meta?.initials || "";
      elements.settingsCoModeratorEl.value = state.project.meta?.coModerator || "";
      elements.settingsCoInitialsEl.value = state.project.meta?.coModeratorInitials || "";
      elements.settingsCompanyEl.value = state.project.meta?.company || "";
      elements.settingsCompanyIdEl.value = state.project.meta?.companyId || "";
      elements.settingsLocaleEl.value = state.project.meta?.locale || "de-CH";
      const libraryName = stateHelpers.getLibraryFileName(state.project.meta || {});
      if (elements.settingsLibraryHintEl) {
        elements.settingsLibraryHintEl.textContent = `Library file: ${libraryName} (timestamped backup on generate).`;
      }
      elements.settingsModal.classList.add("is-open");
      elements.settingsModal.setAttribute("aria-hidden", "false");
    };

    const closeSettings = () => {
      if (!elements.settingsModal) return;
      elements.settingsModal.classList.remove("is-open");
      elements.settingsModal.setAttribute("aria-hidden", "true");
    };

    const saveSettings = () => {
      normalizeHelpers.ensureProjectMeta(state.project, i18n.setLocale);
      state.project.meta.moderator = elements.settingsModeratorEl.value.trim();
      state.project.meta.moderatorInitials = elements.settingsModeratorInitialsEl.value.trim();
      state.project.meta.coModerator = elements.settingsCoModeratorEl.value.trim();
      state.project.meta.coModeratorInitials = elements.settingsCoInitialsEl.value.trim();
      state.project.meta.author = state.project.meta.moderator; // legacy
      state.project.meta.initials = state.project.meta.moderatorInitials; // legacy
      state.project.meta.company = elements.settingsCompanyEl.value.trim();
      state.project.meta.companyId = elements.settingsCompanyIdEl.value.trim();
      state.project.meta.locale = elements.settingsLocaleEl.value || "de-CH";
      if (elements.settingsLibraryHintEl) {
        elements.settingsLibraryHintEl.textContent = `Library file: ${stateHelpers.getLibraryFileName(state.project.meta)} (timestamped backup on generate).`;
      }
      setStatus("Settings saved (remember to save sidecar).");
    };

    if (elements.openSettingsBtn) {
      elements.openSettingsBtn.addEventListener("click", openSettings);
    }
    if (elements.closeSettingsBtn) {
      elements.closeSettingsBtn.addEventListener("click", closeSettings);
    }
    if (elements.settingsBackdrop) {
      elements.settingsBackdrop.addEventListener("click", closeSettings);
    }
    if (elements.saveSettingsBtn) {
      elements.saveSettingsBtn.addEventListener("click", saveSettings);
    }
    if (elements.generateLibraryBtn) {
      elements.generateLibraryBtn.addEventListener("click", async () => {
        try {
          await ioApi.generateLibrary();
        } catch (err) {
          setStatus(`Library update failed: ${err.message}`);
          debug.logLine("error", `Library update failed: ${err.message || err}`);
        }
      });
    }

    if (elements.photoOverlayClose) {
      elements.photoOverlayClose.addEventListener("click", () => {
        renderApi.closePhotoOverlay();
      });
    }
    if (elements.photoOverlayCloseBtn) {
      elements.photoOverlayCloseBtn.addEventListener("click", () => {
        renderApi.closePhotoOverlay();
      });
    }
    if (elements.photoOverlayPrevBtn) {
      elements.photoOverlayPrevBtn.addEventListener("click", () => {
        renderApi.stepPhotoOverlay(-1);
      });
    }
    if (elements.photoOverlayNextBtn) {
      elements.photoOverlayNextBtn.addEventListener("click", () => {
        renderApi.stepPhotoOverlay(1);
      });
    }
    if (elements.checklistOverlayClose) {
      elements.checklistOverlayClose.addEventListener("click", () => {
        renderApi.closeChecklistOverlay();
      });
    }
    if (elements.checklistOverlayCloseBtn) {
      elements.checklistOverlayCloseBtn.addEventListener("click", () => {
        renderApi.closeChecklistOverlay();
      });
    }
    document.addEventListener("keydown", (event) => {
      if (!elements.photoOverlayEl || !elements.photoOverlayEl.classList.contains("is-open")) return;
      if (event.target && ["INPUT", "TEXTAREA"].includes(event.target.tagName)) return;
      if (event.key === "Escape") renderApi.closePhotoOverlay();
      if (event.key === "a" || event.key === "A") renderApi.stepPhotoOverlay(-1);
      if (event.key === "d" || event.key === "D") renderApi.stepPhotoOverlay(1);
    });
    document.addEventListener("keydown", (event) => {
      if (!elements.checklistOverlayEl || !elements.checklistOverlayEl.classList.contains("is-open")) return;
      if (event.key === "Escape") renderApi.closeChecklistOverlay();
    });

    window.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "hidden") {
        flushAutosave();
      }
    });
    window.addEventListener("pagehide", () => {
      flushAutosave();
    });

    if (elements.importSelfBtn) {
      elements.importSelfBtn.addEventListener("click", importSelfHandler);
    }
  };

  window.AutoBerichtBindEvents = { bind };
})();
