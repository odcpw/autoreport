(() => {
  const bind = (ctx, deps) => {
    const { state, runtime, setStatus, debug } = ctx;
    const { elements, tagsApi, photosApi, ioApi, renderApi, actions } = deps;

    const ensureFsAccess = () => {
      if (!window.showDirectoryPicker) {
        setStatus("File System Access API is not available. Open via http://localhost in Edge/Chrome to enable file access.");
        if (elements.pickProjectBtn) {
          elements.pickProjectBtn.disabled = true;
        }
        return false;
      }
      return true;
    };

    const enableActions = () => {
      const hasProject = !!state.projectHandle;
      const hasVisiblePhotos = photosApi.getFilteredPhotos().length > 0;
      if (elements.loadSidecarBtn) elements.loadSidecarBtn.disabled = !hasProject;
      if (elements.pickPhotosBtn) elements.pickPhotosBtn.disabled = !hasProject;
      if (elements.saveSidecarBtn) elements.saveSidecarBtn.disabled = !hasProject;
      if (elements.filterToggleBtn) elements.filterToggleBtn.disabled = state.photos.length === 0;
      if (elements.prevBtn) elements.prevBtn.disabled = !hasVisiblePhotos;
      if (elements.nextBtn) elements.nextBtn.disabled = !hasVisiblePhotos;
      if (elements.importActionBtn) elements.importActionBtn.disabled = !hasProject;
      if (elements.rescanPhotosBtn) elements.rescanPhotosBtn.disabled = !hasProject;
      if (elements.exportActionBtn) elements.exportActionBtn.disabled = !hasProject;
    };

    const openSettings = () => {
      if (!elements.settingsModal) return;
      elements.settingsModal.classList.add("is-open");
      elements.settingsModal.setAttribute("aria-hidden", "false");
      renderApi.renderObservationTagList();
    };

    const closeSettings = () => {
      if (!elements.settingsModal) return;
      elements.settingsModal.classList.remove("is-open");
      elements.settingsModal.setAttribute("aria-hidden", "true");
    };

    const openImportModal = () => {
      if (!elements.importModal) return;
      elements.importModal.classList.add("is-open");
      elements.importModal.setAttribute("aria-hidden", "false");
    };

    const closeImportModal = () => {
      if (!elements.importModal) return;
      elements.importModal.classList.remove("is-open");
      elements.importModal.setAttribute("aria-hidden", "true");
    };

    const addObservationTag = () => {
      if (!elements.obsTagInput) return;
      const value = elements.obsTagInput.value.trim();
      if (!value) return;
      const existing = state.tagOptions.observations || [];
      if (!existing.some((option) => option.value === value)) {
        state.tagOptions.observations = tagsApi.sortOptionsForGroup("observations", [
          ...existing,
          { value, label: value },
        ]);
        ioApi.scheduleAutosave();
        renderApi.renderAll();
        renderApi.renderObservationTagList();
      }
      elements.obsTagInput.value = "";
    };

    if (elements.statusCloseBtn) {
      elements.statusCloseBtn.addEventListener("click", () => {
        renderApi.setStatusHidden(true);
      });
    }

    elements.panelTabButtons?.forEach((button) => {
      button.addEventListener("click", () => {
        renderApi.setActivePanel(button.dataset.panelTab);
      });
    });

    elements.layoutToggleButtons?.forEach((button) => {
      button.addEventListener("click", () => {
        renderApi.setLayoutMode(button.dataset.layout);
      });
    });

    if (elements.pickProjectBtn) {
      elements.pickProjectBtn.addEventListener("click", async () => {
        if (!ensureFsAccess()) return;
        try {
          state.projectHandle = await window.showDirectoryPicker({ mode: "readwrite", id: "autobericht-project" });
          await ioApi.persistHandle(state.projectHandle);
          setStatus(`Project folder: ${state.projectHandle.name}`);
          await ioApi.loadProjectSidecar();
          setStatus(`Project folder: ${state.projectHandle.name}`);
        } catch (err) {
          setStatus(`Project pick canceled: ${err.message}`);
        }
        enableActions();
      });
    }

    if (elements.loadSidecarBtn) {
      elements.loadSidecarBtn.addEventListener("click", async () => {
        await ioApi.loadProjectSidecar();
      });
    }

    if (elements.pickPhotosBtn) {
      elements.pickPhotosBtn.addEventListener("click", () => {
        openImportModal();
      });
    }

    if (elements.importActionBtn) {
      elements.importActionBtn.addEventListener("click", async () => {
        try {
          if (!state.projectHandle) {
            setStatus("Open project folder first.");
            return;
          }
          if (!ctx.photoImport?.importRawPhotos) {
            setStatus("Photo import module not available.");
            return;
          }
          const result = await ctx.photoImport.importRawPhotos({
            projectHandle: state.projectHandle,
            getNestedDirectory: ioApi.getNestedDirectory,
            isImageFile: photosApi.isImageFile,
            setStatus,
            resizeMax: ctx.constants.RESIZE_MAX,
            resizeQuality: ctx.constants.RESIZE_QUALITY,
          });
          if (!result?.resizedHandle) return;
          state.photoHandle = result.resizedHandle;
          state.photoRootName = result.photoRootName || "";
          if (state.projectDoc) {
            state.projectDoc.photoRoot = state.photoRootName;
          }
          await ioApi.saveProjectSidecar();
          await photosApi.scanPhotos();
        } catch (err) {
          setStatus(`Import failed: ${err.message}`);
        }
      });
    }

    if (elements.exportActionBtn) {
      elements.exportActionBtn.addEventListener("click", async () => {
        try {
          if (!state.projectHandle) {
            setStatus("Open project folder first.");
            return;
          }
          if (!ctx.photoImport?.exportTaggedPhotos) {
            setStatus("Photo export module not available.");
            return;
          }
          if (!state.photoHandle) {
            await ioApi.setDefaultPhotoHandle();
          }
          if (!state.photoHandle) {
            setStatus("No photo folder found. Import or scan photos first.");
            return;
          }
          if (!state.photos.length) {
            await photosApi.scanPhotos();
          }
          const result = await ctx.photoImport.exportTaggedPhotos({
            projectHandle: state.projectHandle,
            getNestedDirectory: ioApi.getNestedDirectory,
            setStatus,
            photos: state.photos,
          });
          if (result?.count) {
            setStatus(`Exported ${result.count} photos to ${result.exportRootName}`);
          }
        } catch (err) {
          setStatus(`Export failed: ${err.message}`);
        }
      });
    }

    if (elements.rescanPhotosBtn) {
      elements.rescanPhotosBtn.addEventListener("click", async () => {
        if (!state.projectHandle) {
          setStatus("Open project folder first.");
          return;
        }
        if (!state.photoHandle) {
          await ioApi.setDefaultPhotoHandle();
        }
        if (!state.photoHandle) {
          setStatus("No photo folder found. Import photos first.");
          return;
        }
        await photosApi.scanPhotos();
      });
    }

    if (elements.saveSidecarBtn) {
      elements.saveSidecarBtn.addEventListener("click", async () => {
        await ioApi.saveProjectSidecar();
      });
    }

    if (elements.openSettingsBtn) {
      elements.openSettingsBtn.addEventListener("click", openSettings);
    }
    if (elements.settingsCloseBtn) {
      elements.settingsCloseBtn.addEventListener("click", closeSettings);
    }
    if (elements.settingsBackdrop) {
      elements.settingsBackdrop.addEventListener("click", closeSettings);
    }
    if (elements.importCloseBtn) {
      elements.importCloseBtn.addEventListener("click", closeImportModal);
    }
    if (elements.importBackdrop) {
      elements.importBackdrop.addEventListener("click", closeImportModal);
    }
    if (elements.obsTagAddBtn) {
      elements.obsTagAddBtn.addEventListener("click", addObservationTag);
    }
    if (elements.obsTagInput) {
      elements.obsTagInput.addEventListener("keydown", (event) => {
        if (event.key !== "Enter") return;
        event.preventDefault();
        addObservationTag();
      });
    }

    if (elements.photoUnsortedBtn) {
      elements.photoUnsortedBtn.addEventListener("click", () => {
        actions.setPhotoUnsorted();
      });
    }

    if (elements.filterToggleBtn) {
      elements.filterToggleBtn.addEventListener("click", () => {
        state.filterMode = state.filterMode === "all" ? "unsorted" : "all";
        elements.filterToggleBtn.textContent = state.filterMode === "all" ? "Show Unsorted" : "Show All";
        state.currentIndex = 0;
        renderApi.renderAll();
      });
    }

    if (elements.prevBtn) {
      elements.prevBtn.addEventListener("click", () => {
        const filtered = photosApi.getFilteredPhotos();
        if (!filtered.length) return;
        state.currentIndex = (state.currentIndex - 1 + filtered.length) % filtered.length;
        renderApi.renderAll();
      });
    }

    if (elements.nextBtn) {
      elements.nextBtn.addEventListener("click", () => {
        const filtered = photosApi.getFilteredPhotos();
        if (!filtered.length) return;
        state.currentIndex = (state.currentIndex + 1) % filtered.length;
        renderApi.renderAll();
      });
    }

    document.addEventListener("keydown", (event) => {
      if (state.layoutMode !== "tabs") return;
      const target = event.target;
      const isInput = target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA");
      if (isInput) return;
      if (event.key === "1") renderApi.setActivePanel("report");
      if (event.key === "2") renderApi.setActivePanel("observations");
      if (event.key === "3") renderApi.setActivePanel("training");
    });

    if (elements.notesEl) {
      elements.notesEl.addEventListener("input", () => {
        const current = photosApi.getCurrentPhoto();
        if (!current) return;
        current.notes = elements.notesEl.value;
        ioApi.scheduleAutosave();
      });
    }

    window.addEventListener("keydown", (event) => {
      const target = event.target;
      const isInput = target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement;
      if (isInput) return;
      if (event.key.toLowerCase() === "a") {
        event.preventDefault();
        elements.prevBtn?.click();
      } else if (event.key.toLowerCase() === "d") {
        event.preventDefault();
        elements.nextBtn?.click();
      }
    });

    if (elements.saveLogBtn) {
      elements.saveLogBtn.addEventListener("click", async () => {
        try {
          const result = await debug.saveLog({
            suggestedName: `photosorter-log-${new Date().toISOString().replace(/[:.]/g, "-")}.txt`,
            dirHandle: state.projectHandle || null,
          });
          setStatus(`Saved log (${result.location}): ${result.filename}`);
        } catch (err) {
          setStatus(`Log save failed: ${err.message}`);
        }
      });
    }

    window.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "hidden") {
        ioApi.flushAutosave();
      }
    });
    window.addEventListener("pagehide", () => {
      ioApi.flushAutosave();
    });

    return {
      ensureFsAccess,
      enableActions,
    };
  };

  window.AutoBerichtPhotoSorterBindEvents = { bind };
})();
