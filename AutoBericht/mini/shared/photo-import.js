(() => {
  const getAscii = (view, offset, length) => {
    let out = "";
    for (let i = 0; i < length; i += 1) {
      const code = view.getUint8(offset + i);
      if (code === 0) break;
      out += String.fromCharCode(code);
    }
    return out;
  };

  const readExifTagString = (view, tiffStart, ifdOffset, tagId, littleEndian) => {
    const entryCount = view.getUint16(ifdOffset, littleEndian);
    for (let i = 0; i < entryCount; i += 1) {
      const entryOffset = ifdOffset + 2 + i * 12;
      const tag = view.getUint16(entryOffset, littleEndian);
      if (tag !== tagId) continue;
      const type = view.getUint16(entryOffset + 2, littleEndian);
      const count = view.getUint32(entryOffset + 4, littleEndian);
      const valueOffset = view.getUint32(entryOffset + 8, littleEndian);
      if (type !== 2 || count === 0) return null;
      if (count <= 4) {
        return getAscii(view, entryOffset + 8, count);
      }
      const absolute = tiffStart + valueOffset;
      if (absolute + count > view.byteLength) return null;
      return getAscii(view, absolute, count);
    }
    return null;
  };

  const readExifTagOffset = (view, ifdOffset, tagId, littleEndian) => {
    const entryCount = view.getUint16(ifdOffset, littleEndian);
    for (let i = 0; i < entryCount; i += 1) {
      const entryOffset = ifdOffset + 2 + i * 12;
      const tag = view.getUint16(entryOffset, littleEndian);
      if (tag !== tagId) continue;
      const type = view.getUint16(entryOffset + 2, littleEndian);
      const valueOffset = view.getUint32(entryOffset + 8, littleEndian);
      if (type !== 4) return null;
      return valueOffset;
    }
    return null;
  };

  const readExifDateFromBuffer = (buffer) => {
    const view = new DataView(buffer);
    if (view.byteLength < 12) return null;
    if (view.getUint16(0, false) !== 0xffd8) return null;
    let offset = 2;
    while (offset + 4 < view.byteLength) {
      if (view.getUint8(offset) !== 0xff) {
        offset += 1;
        continue;
      }
      const marker = view.getUint16(offset, false);
      if (marker === 0xffe1) {
        const length = view.getUint16(offset + 2, false);
        const start = offset + 4;
        if (start + 6 > view.byteLength) return null;
        const header = getAscii(view, start, 6);
        if (header !== "Exif\u0000\u0000") return null;
        const tiffStart = start + 6;
        const endian = getAscii(view, tiffStart, 2);
        const littleEndian = endian === "II";
        const firstIfdOffset = view.getUint32(tiffStart + 4, littleEndian);
        const ifd0 = tiffStart + firstIfdOffset;
        if (ifd0 + 2 > view.byteLength) return null;
        const exifOffset = readExifTagOffset(view, ifd0, 0x8769, littleEndian);
        let dateString = null;
        if (exifOffset !== null) {
          const exifIfd = tiffStart + Number(exifOffset);
          dateString = readExifTagString(view, tiffStart, exifIfd, 0x9003, littleEndian);
        }
        if (!dateString) {
          dateString = readExifTagString(view, tiffStart, ifd0, 0x0132, littleEndian);
        }
        if (!dateString) return null;
        const match = /^(\d{4}):(\d{2}):(\d{2}) (\d{2}):(\d{2})(?::(\d{2}))?/.exec(dateString);
        if (!match) return null;
        const [, y, m, d, hh, mm, ss] = match;
        return new Date(
          Number(y),
          Number(m) - 1,
          Number(d),
          Number(hh),
          Number(mm),
          ss ? Number(ss) : 0,
        );
      }
      if (marker === 0xffd9 || marker === 0xffda) break;
      const size = view.getUint16(offset + 2, false);
      if (!size) break;
      offset += 2 + size;
    }
    return null;
  };

  const getPhotoTimestamp = async (file) => {
    try {
      const slice = file.slice(0, 256 * 1024);
      const buffer = await slice.arrayBuffer();
      const exifDate = readExifDateFromBuffer(buffer);
      if (exifDate) return exifDate;
    } catch (err) {
      // ignore
    }
    return new Date(file.lastModified);
  };

  const formatTimestamp = (date) => {
    const pad = (value) => String(value).padStart(2, "0");
    return [
      date.getFullYear(),
      pad(date.getMonth() + 1),
      pad(date.getDate()),
    ].join("-") + `-${pad(date.getHours())}-${pad(date.getMinutes())}`;
  };

  const padSequence = (value) => String(value).padStart(4, "0");

  const hasExistingResizedFiles = async (resizedHandle, isImageFile) => {
    if (!resizedHandle?.values) return false;
    for await (const entry of resizedHandle.values()) {
      if (entry.kind !== "file") continue;
      if (!isImageFile || isImageFile(entry.name)) return true;
    }
    return false;
  };

  const sortTasksByOwnerAndName = (tasks) => {
    return tasks.sort((a, b) => {
      const ownerCmp = String(a.owner).localeCompare(String(b.owner), "de", { numeric: true });
      if (ownerCmp !== 0) return ownerCmp;
      return String(a.file?.name || "").localeCompare(String(b.file?.name || ""), "de", { numeric: true });
    });
  };

  const resizePhoto = async (file, maxSize, quality) => {
    const image = await createImageBitmap(file);
    const longSide = Math.max(image.width, image.height);
    const scale = longSide > maxSize ? maxSize / longSide : 1;
    const targetWidth = Math.max(1, Math.round(image.width * scale));
    const targetHeight = Math.max(1, Math.round(image.height * scale));
    const canvas = document.createElement("canvas");
    canvas.width = targetWidth;
    canvas.height = targetHeight;
    const ctx = canvas.getContext("2d");
    ctx.drawImage(image, 0, 0, targetWidth, targetHeight);
    return new Promise((resolve) => {
      canvas.toBlob((blob) => resolve(blob), "image/jpeg", quality);
    });
  };

  const findRawFolder = async (projectHandle, getNestedDirectory) => {
    const candidates = [
      ["Photos", "raw"],
      ["Photos", "Raw"],
      ["photos", "raw"],
      ["photos", "Raw"],
    ];
    for (const parts of candidates) {
      try {
        const handle = await getNestedDirectory(projectHandle, parts);
        return { handle, label: parts.join("/") };
      } catch (err) {
        // try next
      }
    }
    return null;
  };

  const collectRawTasks = async (rawHandle, isImageFile) => {
    const invalidFolders = [];
    const folders = [];
    for await (const entry of rawHandle.values()) {
      if (entry.kind !== "directory") continue;
      if (entry.name.length !== 3) {
        invalidFolders.push(entry.name);
        continue;
      }
      folders.push(entry);
    }
    if (invalidFolders.length) {
      throw new Error(`Raw folder names must be 3 chars: ${invalidFolders.join(", ")}`);
    }
    const tasks = [];
    for (const folder of folders) {
      for await (const entry of folder.values()) {
        if (entry.kind !== "file") continue;
        if (!isImageFile(entry.name)) continue;
        const file = await entry.getFile();
        tasks.push({
          owner: folder.name,
          file,
        });
      }
    }
    return tasks;
  };

  const ensureResizedFolder = async (projectHandle, getNestedDirectory) => {
    const photosHandle =
      (await getNestedDirectory(projectHandle, ["photos"], { create: true }).catch(() => null))
      || (await getNestedDirectory(projectHandle, ["Photos"], { create: true }));
    return getNestedDirectory(photosHandle, ["resized"], { create: true }).catch(async () =>
      getNestedDirectory(photosHandle, ["Resized"], { create: true }));
  };

  const fileExists = async (dirHandle, name) => {
    try {
      await dirHandle.getFileHandle(name);
      return true;
    } catch (err) {
      return false;
    }
  };

  const sanitizeFolderName = (value) => {
    const cleaned = String(value || "")
      .trim()
      .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_")
      .replace(/[.\s]+$/g, "");
    return cleaned || "tag";
  };

  const ensureUniqueFileName = async (dirHandle, fileName) => {
    if (!(await fileExists(dirHandle, fileName))) return fileName;
    const dot = fileName.lastIndexOf(".");
    const stem = dot === -1 ? fileName : fileName.slice(0, dot);
    const ext = dot === -1 ? "" : fileName.slice(dot);
    let counter = 1;
    while (counter < 1000) {
      const nextName = `${stem}-${counter}${ext}`;
      if (!(await fileExists(dirHandle, nextName))) return nextName;
      counter += 1;
    }
    return `${stem}-${Date.now()}${ext}`;
  };

  const getFileHandleFromPath = async (projectHandle, path) => {
    const parts = String(path || "")
      .split("/")
      .map((part) => part.trim())
      .filter(Boolean);
    if (!parts.length) return null;
    let current = projectHandle;
    for (let i = 0; i < parts.length - 1; i += 1) {
      current = await current.getDirectoryHandle(parts[i]);
    }
    return current.getFileHandle(parts[parts.length - 1]);
  };

  const getPhotoFile = async (photo, projectHandle) => {
    if (!photo) return null;
    try {
      if (photo.file) return photo.file;
      if (photo.fileHandle) return await photo.fileHandle.getFile();
    } catch (err) {
      // fall through to path lookup
    }
    if (projectHandle && photo.path) {
      try {
        const handle = await getFileHandleFromPath(projectHandle, photo.path);
        if (handle) return await handle.getFile();
      } catch (err) {
        // ignore
      }
    }
    return null;
  };

  const getPhotoFileName = (photo) => {
    if (!photo) return "photo.jpg";
    if (photo.file?.name) return photo.file.name;
    if (photo.fileHandle?.name) return photo.fileHandle.name;
    if (photo.path) {
      const parts = String(photo.path).split("/");
      return parts[parts.length - 1] || "photo.jpg";
    }
    return "photo.jpg";
  };

  const findPhotosRoot = async (projectHandle) => {
    try {
      const handle = await projectHandle.getDirectoryHandle("photos");
      return { handle, name: "photos" };
    } catch (err) {
      // ignore
    }
    try {
      const handle = await projectHandle.getDirectoryHandle("Photos");
      return { handle, name: "Photos" };
    } catch (err) {
      // ignore
    }
    const handle = await projectHandle.getDirectoryHandle("photos", { create: true });
    return { handle, name: "photos" };
  };

  const findExportRoot = async (photosHandle) => {
    try {
      const handle = await photosHandle.getDirectoryHandle("export");
      return { handle, name: "export" };
    } catch (err) {
      // ignore
    }
    try {
      const handle = await photosHandle.getDirectoryHandle("Export");
      return { handle, name: "Export" };
    } catch (err) {
      // ignore
    }
    const handle = await photosHandle.getDirectoryHandle("export", { create: true });
    return { handle, name: "export" };
  };

  const importRawPhotos = async (context) => {
    const {
      projectHandle,
      getNestedDirectory,
      isImageFile,
      setStatus,
      resizeMax = 1200,
      resizeQuality = 0.85,
    } = context || {};

    if (!projectHandle) {
      setStatus?.("Open project folder first.");
      return null;
    }

    const raw = await findRawFolder(projectHandle, getNestedDirectory);
    if (!raw) {
      setStatus?.("Missing Photos/raw. Expected Photos/raw/ABC folders.");
      return null;
    }
    const tasks = await collectRawTasks(raw.handle, isImageFile);
    if (!tasks.length) {
      setStatus?.("No raw images found.");
      return null;
    }
    const resizedHandle = await ensureResizedFolder(projectHandle, getNestedDirectory);
    if (await hasExistingResizedFiles(resizedHandle, isImageFile)) {
      setStatus?.("Photos/resized is not empty. Clear it before importing to avoid breaking tags.");
      return null;
    }
    const counters = new Map();
    let completed = 0;

    const orderedTasks = sortTasksByOwnerAndName(tasks);
    for (const task of orderedTasks) {
      const timestamp = formatTimestamp(await getPhotoTimestamp(task.file));
      let counter = (counters.get(task.owner) || 0) + 1;
      let filename = "";
      while (true) {
        filename = `${timestamp}_${task.owner}_${padSequence(counter)}.jpg`;
        if (!(await fileExists(resizedHandle, filename))) break;
        counter += 1;
      }
      counters.set(task.owner, counter);
      const blob = await resizePhoto(task.file, resizeMax, resizeQuality);
      if (!blob) {
        throw new Error(`Failed to resize ${task.file.name}`);
      }
      const handle = await resizedHandle.getFileHandle(filename, { create: true });
      const writable = await handle.createWritable();
      await writable.write(blob);
      await writable.close();

      completed += 1;
      const pct = Math.round((completed / tasks.length) * 100);
      setStatus?.(`Importing photos ${completed}/${tasks.length} (${pct}%)`);
    }

    return {
      resizedHandle,
      photoRootName: "photos/resized",
      count: tasks.length,
    };
  };

  const exportTaggedPhotos = async (context) => {
    const {
      projectHandle,
      setStatus,
      photos,
      tagOptions,
    } = context || {};
    if (!projectHandle) throw new Error("Project folder is missing.");
    if (!Array.isArray(photos) || !photos.length) {
      return { count: 0, exportRootName: "Photos/export" };
    }

    setStatus?.("Preparing export...");
    const labelMap = new Map();
    const addOptionsToLabelMap = (options = []) => {
      options.forEach((opt) => {
        if (!opt) return;
        const value = String(opt.value || opt.label || "").trim();
        const label = String(opt.label || opt.value || "").trim();
        if (!value || !label) return;
        if (!labelMap.has(value)) labelMap.set(value, label);
      });
    };
    addOptionsToLabelMap(tagOptions?.report);
    addOptionsToLabelMap(tagOptions?.observations);
    addOptionsToLabelMap(tagOptions?.training);

    const tagMap = new Map();
    const addPhotoToTag = (tag, photo) => {
      const raw = String(tag || "").trim();
      if (!raw) return;
      const key = labelMap.get(raw) || raw;
      if (!key) return;
      if (!tagMap.has(key)) tagMap.set(key, []);
      tagMap.get(key).push(photo);
    };

    const isExportPath = (path) => {
      const parts = String(path || "")
        .split("/")
        .map((part) => part.trim().toLowerCase())
        .filter(Boolean);
      return parts.includes("export");
    };

    photos.forEach((photo) => {
      if (isExportPath(photo?.path)) return;
      const tags = photo?.tags || {};
      const list = new Set([
        ...(tags.report || []),
        ...(tags.observations || []),
        ...(tags.training || []),
      ].map((tag) => String(tag || "").trim()).filter(Boolean));
      if (!list.size) {
        addPhotoToTag("unsorted", photo);
      } else {
        list.forEach((tag) => addPhotoToTag(tag, photo));
      }
    });

    const { handle: photosHandle, name: photosName } = await findPhotosRoot(projectHandle);
    const { handle: exportHandle, name: exportName } = await findExportRoot(photosHandle);
    const exportRootName = `${photosName}/${exportName}`;

    const uniquePhotos = new Set();
    let total = 0;
    for (const items of tagMap.values()) {
      total += items.length;
      items.forEach((photo) => {
        if (photo?.path) {
          uniquePhotos.add(photo.path);
        } else if (photo?.file?.name) {
          uniquePhotos.add(photo.file.name);
        } else if (photo?.fileHandle?.name) {
          uniquePhotos.add(photo.fileHandle.name);
        }
      });
    }
    let completed = 0;

    for (const [tag, items] of tagMap.entries()) {
      const folderName = sanitizeFolderName(tag);
      const tagHandle = await exportHandle.getDirectoryHandle(folderName, { create: true });
      for (const photo of items) {
        const file = await getPhotoFile(photo, projectHandle);
        if (!file) continue;
        const rawName = getPhotoFileName(photo);
        const safeName = await ensureUniqueFileName(tagHandle, rawName);
        const target = await tagHandle.getFileHandle(safeName, { create: true });
        const writable = await target.createWritable();
        await writable.write(file);
        await writable.close();
        completed += 1;
        setStatus?.(`Exporting photos ${completed}/${total}`);
      }
    }

    return { count: uniquePhotos.size, copyCount: completed, exportRootName };
  };

  window.AutoBerichtPhotoImport = {
    importRawPhotos,
    exportTaggedPhotos,
  };
})();
