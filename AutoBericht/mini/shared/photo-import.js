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
    const counters = new Map();
    let completed = 0;

    for (const task of tasks) {
      const timestamp = formatTimestamp(await getPhotoTimestamp(task.file));
      const counterKey = `${timestamp}_${task.owner}`;
      let counter = counters.get(counterKey) || 1;
      let filename = "";
      while (true) {
        filename = `${timestamp}_${task.owner}_${padSequence(counter)}.jpg`;
        if (!(await fileExists(resizedHandle, filename))) break;
        counter += 1;
      }
      counters.set(counterKey, counter + 1);
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

  window.AutoBerichtPhotoImport = {
    importRawPhotos,
  };
})();
