(() => {
  const init = (ctx, deps) => {
    const { state, runtime, setStatus } = ctx;
    const { t, tf } = ctx.i18n;
    const { tagsApi, notifyChange } = deps;
    const constants = window.AutoBerichtPhotoSorterState || {};

    const isImageFile = (name) => {
      const lower = name.toLowerCase();
      const dot = lower.lastIndexOf(".");
      if (dot === -1) return false;
      return constants.IMAGE_EXTENSIONS?.has(lower.slice(dot));
    };

    const buildPhotoEntry = (path, fileHandle, file) => {
      const previous = state.photos.find((photo) => photo.path === path) || state.projectDoc?.photos?.[path];
      const tags = tagsApi.normalizePhotoTags(previous?.tags);
      return {
        path,
        photoNumber: previous?.photoNumber,
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

    // Numbers identify photos within this project, independently of the current
    // list/filter. Keep a high-water mark even when the highest photo is removed.
    const assignPhotoNumbers = (collection) => {
      const saved = state.projectDoc?.photos || {};
      const byNumber = new Map();
      let last = state.projectDoc?.meta?.lastPhotoNumber ?? 0;
      if (!Number.isSafeInteger(last) || last < 0) {
        throw new Error("Invalid saved photo counter. Restore a valid sidecar before assigning photo numbers.");
      }
      const reserve = (path, number) => {
        if (number == null) return;
        if (!Number.isSafeInteger(number) || number < 1) {
          throw new Error(`Invalid photo number for ${path}.`);
        }
        if (byNumber.has(number) && byNumber.get(number) !== path) {
          throw new Error(`Photo ${number} is assigned to both ${byNumber.get(number)} and ${path}. Correct the sidecar before continuing.`);
        }
        byNumber.set(number, path);
        last = Math.max(last, number);
      };
      Object.entries(saved).forEach(([path, photo]) => reserve(path, photo.photoNumber));
      state.photos.forEach((photo) => reserve(photo.path, photo.photoNumber));
      collection.forEach((photo) => reserve(photo.path, photo.photoNumber));
      // scanPhotos supplies the same unfiltered path order as the existing UI.
      const numbers = collection.map((photo) => {
        if (photo.photoNumber != null) return photo.photoNumber;
        if (last >= Number.MAX_SAFE_INTEGER) throw new Error("No photo numbers available.");
        last += 1;
        return last;
      });
      collection.forEach((photo, index) => { photo.photoNumber = numbers[index]; });
      if (!state.projectDoc) state.projectDoc = {};
      if (!state.projectDoc.meta) state.projectDoc.meta = {};
      state.projectDoc.meta.lastPhotoNumber = last;
    };

    const scanPhotos = async () => {
      if (!state.photoHandle) return;
      clearPhotoUrl();
      setStatus(t("status_scanning_photos"));
      const collection = [];
      const prefix = state.photoRootName ? `${state.photoRootName}/` : "";
      await collectImages(state.photoHandle, prefix, collection);
      collection.sort((a, b) => a.path.localeCompare(b.path));
      assignPhotoNumbers(collection);
      state.photos = collection;
      state.filterMode = "all";
      state.keepPath = "";
      state.currentIndex = 0;
      setStatus(tf("status_loaded_photos", "Loaded {count} photos from {folder}.", { count: state.photos.length, folder: state.photoRootName }));
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
        filtered = filtered.filter((photo) => isPhotoUnsorted(photo) || photo.path === state.keepPath);
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
      setStatus(t("status_loading_demo_photos"));
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
      assignPhotoNumbers(collection);
      state.photos = collection;
      state.filterMode = "all";
      state.currentIndex = 0;
      setStatus(tf("status_loaded_demo_photos", "Loaded {count} demo photos.", { count: collection.length }));
      if (notifyChange) notifyChange();
    };

    const serializePhotos = () => {
      const output = {};
      state.photos.forEach((photo) => {
        output[photo.path] = {
          ...state.projectDoc?.photos?.[photo.path],
          photoNumber: photo.photoNumber,
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
