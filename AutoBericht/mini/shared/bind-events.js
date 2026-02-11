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
      flushAutosave,
      setFirstRunVisible,
      applyAutoBackup,
    } = deps;
    const { elements, state, runtime, debug, setStatus, i18n } = ctx;

    const maybeOpenSettings = (result) => {
      if (result && result.source && result.source !== "sidecar") {
        openSettings();
      }
    };

    const pickProjectFolder = async () => {
      if (!ensureFsAccess()) return;
      try {
        runtime.dirHandle = await window.showDirectoryPicker({ mode: "readwrite", id: "autobericht-project" });
        if (ctx.fs?.saveHandle) {
          await ctx.fs.saveHandle(runtime.dirHandle);
        }
        enableActions();
        setStatus(`Selected folder: ${runtime.dirHandle.name}`);
        debug.logLine("info", `Selected folder: ${runtime.dirHandle.name}`);
        if (setFirstRunVisible) setFirstRunVisible(false);
        const result = await ioApi.loadProjectFromFolder();
        if (result?.ok) {
          maybeOpenSettings(result);
          if (applyAutoBackup) applyAutoBackup();
        } else if (setFirstRunVisible) {
          setFirstRunVisible(true);
        }
      } catch (err) {
        setStatus(`Folder pick canceled or failed: ${err.message}`);
        debug.logLine("warn", `Folder pick canceled or failed: ${err.message}`);
      }
    };

    if (elements.pickFolderBtn) {
      elements.pickFolderBtn.addEventListener("click", pickProjectFolder);
    }
    if (elements.firstRunPickBtn) {
      elements.firstRunPickBtn.addEventListener("click", pickProjectFolder);
    }

    if (elements.loadSidecarBtn) {
      elements.loadSidecarBtn.addEventListener("click", async () => {
        if (!runtime.dirHandle) return;
        const result = await ioApi.loadProjectFromFolder();
        if (result?.ok) {
          maybeOpenSettings(result);
          if (applyAutoBackup) applyAutoBackup();
        }
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

    const bindFilterGroup = (inputs, key) => {
      if (!inputs || !inputs.length) return;
      inputs.forEach((input) => {
        input.addEventListener("change", () => {
          if (!input.checked) return;
          state.filters[key] = input.value;
          renderApi.renderRows();
        });
      });
    };

    bindFilterGroup(elements.filterAnswerModeEls, "answer");
    bindFilterGroup(elements.filterIncludeModeEls, "include");
    bindFilterGroup(elements.filterDoneModeEls, "done");

    const openSettings = () => {
      if (!elements.settingsModal) return;
      elements.settingsModeratorEl.value = state.project.meta?.moderator || state.project.meta?.author || "";
      elements.settingsModeratorInitialsEl.value = state.project.meta?.moderatorInitials || state.project.meta?.initials || "";
      elements.settingsCoModeratorEl.value = state.project.meta?.coModerator || "";
      elements.settingsCoInitialsEl.value = state.project.meta?.coModeratorInitials || "";
      elements.settingsCompanyEl.value = state.project.meta?.company || "";
      elements.settingsCompanyIdEl.value = state.project.meta?.companyId || "";
      elements.settingsLocaleEl.value = state.project.meta?.locale || "de-CH";
      if (elements.settingsBackupMinutesEl) {
        const minutes = Number(state.project.meta?.autobackupMinutes);
        elements.settingsBackupMinutesEl.value = Number.isFinite(minutes) ? String(minutes) : "30";
      }
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

    const saveSettings = async () => {
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
      if (elements.settingsBackupMinutesEl) {
        const raw = Number(elements.settingsBackupMinutesEl.value);
        const minutes = Number.isFinite(raw) ? Math.max(0, Math.floor(raw)) : 30;
        state.project.meta.autobackupMinutes = minutes;
      }
      if (elements.settingsLibraryHintEl) {
        elements.settingsLibraryHintEl.textContent = `Library file: ${stateHelpers.getLibraryFileName(state.project.meta)} (timestamped backup on generate).`;
      }
      if (applyAutoBackup) applyAutoBackup();
      if (runtime.dirHandle) {
        try {
          await ioApi.saveSidecar();
          setStatus("Settings saved.");
        } catch (err) {
          setStatus(`Settings saved, but sidecar save failed: ${err.message}`);
          debug.logLine("error", `Sidecar save failed: ${err.message || err}`);
        }
      } else {
        setStatus("Settings saved (remember to save sidecar).");
      }
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
    if (elements.chapterPreviewBackdrop) {
      elements.chapterPreviewBackdrop.addEventListener("click", () => {
        renderApi.closeChapterPreview();
      });
    }
    if (elements.chapterPreviewCloseBtn) {
      elements.chapterPreviewCloseBtn.addEventListener("click", () => {
        renderApi.closeChapterPreview();
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
    document.addEventListener("keydown", (event) => {
      if (!elements.chapterPreviewModal || !elements.chapterPreviewModal.classList.contains("is-open")) return;
      if (event.key === "Escape") renderApi.closeChapterPreview();
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
