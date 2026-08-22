/*
 * Shared, conflict-aware persistence for project_sidecar.json.
 *
 * AutoBericht and PhotoSorter edit different branches of the same document.
 * This module serializes cross-tab writes, rejects same-branch stale writes,
 * and reads the saved payload back before reporting success.
 */
(() => {
  const SIDECAR_FILENAME = "project_sidecar.json";
  const LOCK_NAME = "autobericht-project-sidecar";
  const FALLBACK_LOCK_KEY = "autobericht-project-sidecar-lock";
  const FALLBACK_LOCK_TTL_MS = 15000;
  const FALLBACK_LOCK_WAIT_MS = 40;
  const FALLBACK_LOCK_TIMEOUT_MS = 20000;

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

  const branchFingerprint = (doc, branch) => JSON.stringify(
    isPlainObject(doc) && Object.prototype.hasOwnProperty.call(doc, branch)
      ? doc[branch]
      : null,
  );

  const readRevision = (doc) => {
    const value = Number(doc?.meta?.revision);
    return Number.isSafeInteger(value) && value >= 0 ? value : 0;
  };

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const createLockToken = () => {
    if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID();
    return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  };

  const withLocalStorageLock = async (task) => {
    let storage = null;
    try {
      storage = globalThis.localStorage;
    } catch (err) {
      return task();
    }
    if (!storage) return task();
    const token = createLockToken();
    const deadline = Date.now() + FALLBACK_LOCK_TIMEOUT_MS;

    while (Date.now() < deadline) {
      let current = null;
      try {
        current = JSON.parse(storage.getItem(FALLBACK_LOCK_KEY) || "null");
      } catch (err) {
        current = null;
      }
      const now = Date.now();
      if (!current || Number(current.expiresAt) <= now) {
        const candidate = JSON.stringify({ token, expiresAt: now + FALLBACK_LOCK_TTL_MS });
        try {
          storage.setItem(FALLBACK_LOCK_KEY, candidate);
        } catch (err) {
          return task();
        }
        let confirmed = null;
        try {
          confirmed = JSON.parse(storage.getItem(FALLBACK_LOCK_KEY) || "null");
        } catch (err) {
          confirmed = null;
        }
        if (confirmed?.token === token) {
          try {
            return await task();
          } finally {
            try {
              const owner = JSON.parse(storage.getItem(FALLBACK_LOCK_KEY) || "null");
              if (owner?.token === token) storage.removeItem(FALLBACK_LOCK_KEY);
            } catch (err) {
              // A malformed fallback lock expires automatically.
            }
          }
        }
      }
      await sleep(FALLBACK_LOCK_WAIT_MS);
    }
    throw createStorageError(
      "Timed out waiting for another AutoBericht window to finish saving.",
      "SidecarLockTimeoutError",
    );
  };

  const withSidecarLock = async (task) => {
    let lockManager = null;
    try {
      lockManager = globalThis.navigator?.locks;
    } catch (err) {
      lockManager = null;
    }
    if (lockManager?.request) {
      return lockManager.request(LOCK_NAME, { mode: "exclusive" }, task);
    }
    return withLocalStorageLock(task);
  };

  const writeAndVerify = async (dirHandle, payload) => {
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

    const verified = await readSidecar(dirHandle, { allowMissing: false });
    if (JSON.stringify(verified) !== JSON.stringify(payload)) {
      throw createStorageError(
        "The saved sidecar did not read back exactly as written.",
        "SidecarVerificationError",
      );
    }
    return verified;
  };

  const saveBranch = async ({
    dirHandle,
    baseDoc,
    branch,
    merge,
    writerId = "",
  }) => {
    if (!branch) throw createStorageError("A sidecar branch name is required.");
    if (typeof merge !== "function") throw createStorageError("A sidecar merge function is required.");

    return withSidecarLock(async () => {
      const latest = await readSidecar(dirHandle, { allowMissing: true });
      const expected = branchFingerprint(baseDoc, branch);
      const actual = branchFingerprint(latest, branch);
      if (expected !== actual) {
        throw createStorageError(
          `project_sidecar.json changed in another window (${branch} branch). Reload before saving; your unsaved edits remain open.`,
          "SidecarConflictError",
        );
      }

      const merged = merge(latest);
      if (!isPlainObject(merged)) {
        throw createStorageError("The sidecar merge did not return a JSON object.");
      }
      merged.meta = isPlainObject(merged.meta) ? merged.meta : {};
      merged.meta.updatedAt = new Date().toISOString();
      merged.meta.revision = Math.max(readRevision(latest), readRevision(baseDoc)) + 1;
      if (writerId) merged.meta.lastWriter = String(writerId);
      return writeAndVerify(dirHandle, merged);
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
    branchFingerprint,
    readRevision,
    saveBranch,
    enqueue,
    withSidecarLock,
  };
})();
