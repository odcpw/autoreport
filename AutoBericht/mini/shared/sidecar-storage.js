/*
 * Shared persistence for project_sidecar.json.
 *
 * AutoBericht and PhotoSorter edit different branches of the same document.
 * Every save takes a cross-tab lock, reads the latest file, lets the caller
 * replace only its own branch, and writes the result back. Two open tabs can
 * therefore never overwrite each other's branch.
 */
(() => {
  const SIDECAR_FILENAME = "project_sidecar.json";
  const LOCK_NAME = "autobericht-project-sidecar";

  const isPlainObject = (value) => (
    !!value
    && typeof value === "object"
    && !Array.isArray(value)
  );

  const isNotFoundError = (err) => String(err?.name || "") === "NotFoundError";

  const createStorageError = (message, name = "SidecarStorageError", cause = null) => {
    const error = new Error(message);
    error.name = name;
    if (cause) error.cause = cause;
    return error;
  };

  const parseSidecarText = (text) => {
    try {
      const parsed = JSON.parse(String(text || ""));
      if (!isPlainObject(parsed)) {
        throw new Error("The root value must be a JSON object.");
      }
      return parsed;
    } catch (err) {
      throw createStorageError(
        `project_sidecar.json is corrupted or invalid JSON: ${err.message || err}`,
        "SidecarParseError",
        err,
      );
    }
  };

  const readSidecar = async (dirHandle, options = {}) => {
    if (!dirHandle) throw createStorageError("Project folder is not open.");
    try {
      const handle = await dirHandle.getFileHandle(SIDECAR_FILENAME);
      const file = await handle.getFile();
      return parseSidecarText(await file.text());
    } catch (err) {
      if (isNotFoundError(err) && options.allowMissing !== false) return null;
      if (err?.name === "SidecarParseError") throw err;
      throw createStorageError(
        `Could not read project_sidecar.json: ${err.message || err}`,
        "SidecarReadError",
        err,
      );
    }
  };

  const withSidecarLock = async (task) => {
    const locks = globalThis.navigator?.locks;
    if (locks?.request) return locks.request(LOCK_NAME, { mode: "exclusive" }, task);
    return task();
  };

  const writeSidecar = async (dirHandle, payload) => {
    const handle = await dirHandle.getFileHandle(SIDECAR_FILENAME, { create: true });
    let writable = null;
    try {
      writable = await handle.createWritable();
      await writable.write(JSON.stringify(payload, null, 2));
      await writable.close();
      writable = null;
    } catch (err) {
      if (writable?.abort) {
        try {
          await writable.abort();
        } catch (abortErr) {
          // Preserve the original write error.
        }
      }
      throw createStorageError(
        `Could not write project_sidecar.json: ${err.message || err}`,
        "SidecarWriteError",
        err,
      );
    }
    return payload;
  };

  // merge(latest) receives the document currently on disk (or null) and must
  // return the full document to write, with only the caller's branch replaced.
  const saveSidecar = async ({ dirHandle, merge }) => {
    if (typeof merge !== "function") throw createStorageError("A sidecar merge function is required.");
    return withSidecarLock(async () => {
      const latest = await readSidecar(dirHandle, { allowMissing: true });
      const merged = merge(latest);
      if (!isPlainObject(merged)) {
        throw createStorageError("The sidecar merge did not return a JSON object.");
      }
      merged.meta = isPlainObject(merged.meta) ? merged.meta : {};
      merged.meta.updatedAt = new Date().toISOString();
      return writeSidecar(dirHandle, merged);
    });
  };

  const enqueue = (runtime, operation) => {
    if (!runtime || typeof operation !== "function") {
      return Promise.reject(createStorageError("A save runtime and operation are required."));
    }
    const task = Promise.resolve(runtime.saveQueue)
      .catch(() => undefined)
      .then(operation);
    runtime.saveQueue = task.catch(() => undefined);
    return task;
  };

  window.AutoBerichtSidecarStorage = {
    SIDECAR_FILENAME,
    parseSidecarText,
    readSidecar,
    saveSidecar,
    enqueue,
    withSidecarLock,
  };
})();
