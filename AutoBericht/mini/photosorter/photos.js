(() => {
  const init = (ctx, deps) => {
    const { state, runtime, setStatus } = ctx;
    const { tagsApi, notifyChange } = deps;
    const constants = window.AutoBerichtPhotoSorterState || {};

    const isImageFile = (name) => {
      const lower = name.toLowerCase();
      const dot = lower.lastIndexOf(".");
      if (dot === -1) return false;
      return constants.IMAGE_EXTENSIONS?.has(lower.slice(dot));
    };

    const buildPhotoEntry = (path, fileHandle, file) => {
      const previous = state.projectDoc?.photos?.[path];
      const tags = tagsApi.normalizePhotoTags(previous?.tags);
      return {
        path,
        fileHandle: fileHandle || null,
        file: file || null,
        notes: previous?.notes || "",
        tags,
      };
    };

    const getPhotoFile = async (photo) => {
      if (!photo) return null;
      if (photo.file) return photo.file;
      if (photo.fileHandle) {
        return photo.fileHandle.getFile();
      }
      return null;
    };

    const loadPhotoUrl = async (photo) => {
      if (!photo) return null;
      state.currentPhotoToken += 1;
      const token = state.currentPhotoToken;
      const file = await getPhotoFile(photo);
      if (!file) return null;
      if (token !== state.currentPhotoToken) return null;
      if (state.currentPhotoUrl) {
        URL.revokeObjectURL(state.currentPhotoUrl);
      }
      const url = URL.createObjectURL(file);
      state.currentPhotoUrl = url;
      return url;
    };

    const clearPhotoUrl = () => {
      if (state.currentPhotoUrl) {
        URL.revokeObjectURL(state.currentPhotoUrl);
        state.currentPhotoUrl = "";
      }
    };

    const collectImages = async (handle, prefix, collection) => {
      for await (const entry of handle.values()) {
        if (entry.kind === "file") {
          if (!isImageFile(entry.name)) continue;
          const path = `${prefix}${entry.name}`;
          collection.push(buildPhotoEntry(path, entry, null));
        } else if (entry.kind === "directory") {
          const lower = entry.name.toLowerCase();
          if (lower === "raw") continue; // skip raw tree for display; resized is source of truth
          await collectImages(entry, `${prefix}${entry.name}/`, collection);
        }
      }
    };

    const scanPhotos = async () => {
      if (!state.photoHandle) return;
      clearPhotoUrl();
      setStatus("Scanning photos...");
      const collection = [];
      const prefix = state.photoRootName ? `${state.photoRootName}/` : "";
      await collectImages(state.photoHandle, prefix, collection);
      collection.sort((a, b) => a.path.localeCompare(b.path));
      state.photos = collection;
      state.filterMode = "all";
      state.currentIndex = 0;
      setStatus(`Loaded ${state.photos.length} photos from ${state.photoRootName}`);
      if (notifyChange) notifyChange();
    };

    const isPhotoUnsorted = (photo) => {
      if (!photo?.tags) return true;
      const { report, observations, training } = photo.tags;
      return !((report || []).length || (observations || []).length || (training || []).length);
    };

    const hasActiveTagFilters = () => (
      ["report", "observations", "training"].some(
        (group) => (state.activeTagFilters?.[group] || []).length > 0
      )
    );

    const photoMatchesActiveTagFilters = (photo) => (
      ["report", "observations", "training"].every((group) => {
        const active = state.activeTagFilters?.[group] || [];
        if (!active.length) return true;
        const assigned = new Set(photo?.tags?.[group] || []);
        return active.every((tag) => assigned.has(tag));
      })
    );

    const getFilteredPhotos = () => {
      let filtered = state.photos;
      if (state.filterMode === "unsorted") {
        filtered = filtered.filter(isPhotoUnsorted);
      }
      if (hasActiveTagFilters()) {
        filtered = filtered.filter(photoMatchesActiveTagFilters);
      }
      return filtered;
    };

    const getCurrentPhoto = () => {
      const filtered = getFilteredPhotos();
      if (!filtered.length) return null;
      return filtered[state.currentIndex] || null;
    };

    const maybeAutoScan = async () => {
      if (!state.photoHandle) return false;
      if (state.photos.length > 0) return false;
      const savedPhotos = state.projectDoc?.photos || {};
      const hasSaved = Object.keys(savedPhotos).length > 0;
      const hasRoot = !!(state.projectDoc?.photoRoot || state.photoRootName);
      if (!hasSaved && !hasRoot) return false;
      await scanPhotos();
      return true;
    };

    const loadDemoPhotos = async () => {
      if (!constants.DEMO_PHOTO_URLS?.length) return;
      if (!state.projectDoc) {
        state.projectDoc = deps.ioApi?.createEmptyProjectDoc
          ? deps.ioApi.createEmptyProjectDoc()
          : { photos: {}, photoTagOptions: tagsApi.SEED_TAG_OPTIONS, photoRoot: "" };
      }
      setStatus("Loading demo photos...");
      const collection = [];
      for (let i = 0; i < constants.DEMO_PHOTO_URLS.length; i += 1) {
        const url = constants.DEMO_PHOTO_URLS[i];
        const response = await fetch(url);
        if (!response.ok) {
          throw new Error(`Failed to fetch ${url}`);
        }
        const blob = await response.blob();
        const file = new File([blob], `${i + 1}.jpg`, { type: blob.type || "image/jpeg" });
        collection.push(buildPhotoEntry(`demo-photos/${i + 1}.jpg`, null, file));
      }
      state.photoRootName = "demo-photos";
      state.photos = collection;
      state.filterMode = "all";
      state.currentIndex = 0;
      setStatus(`Loaded ${collection.length} demo photos`);
      if (notifyChange) notifyChange();
    };

    const serializePhotos = () => {
      const output = {};
      state.photos.forEach((photo) => {
        output[photo.path] = {
          notes: photo.notes || "",
          tags: photo.tags,
        };
      });
      return output;
    };

    return {
      isImageFile,
      buildPhotoEntry,
      getPhotoFile,
      loadPhotoUrl,
      clearPhotoUrl,
      collectImages,
      scanPhotos,
      maybeAutoScan,
      loadDemoPhotos,
      isPhotoUnsorted,
      hasActiveTagFilters,
      photoMatchesActiveTagFilters,
      getFilteredPhotos,
      getCurrentPhoto,
      serializePhotos,
    };
  };

  window.AutoBerichtPhotoSorterPhotos = { init };
})();
