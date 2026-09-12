(() => {
  const HANDLE_DB_NAME = "autobericht";
  const HANDLE_STORE = "handles";
  const HANDLE_KEY = "projectDir";

  const openHandleDb = () => new Promise((resolve, reject) => {
    if (!window.indexedDB) {
      resolve(null);
      return;
    }
    const request = window.indexedDB.open(HANDLE_DB_NAME, 1);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(HANDLE_STORE)) {
        db.createObjectStore(HANDLE_STORE);
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });

  const saveHandle = async (handle) => {
    const db = await openHandleDb();
    if (!db) return;
    await new Promise((resolve, reject) => {
      const tx = db.transaction(HANDLE_STORE, "readwrite");
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error || new Error("Could not save the project-folder handle."));
      tx.onabort = () => reject(tx.error || new Error("Saving the project-folder handle was aborted."));
      tx.objectStore(HANDLE_STORE).put(handle, HANDLE_KEY);
    });
  };

  const loadHandle = async () => {
    const db = await openHandleDb();
    if (!db) return null;
    const tx = db.transaction(HANDLE_STORE, "readonly");
    return await new Promise((resolve) => {
      const req = tx.objectStore(HANDLE_STORE).get(HANDLE_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  };

  const requestHandlePermission = async (handle) => {
    const opts = { mode: "readwrite" };
    try {
      if (typeof handle?.queryPermission === "function" && (await handle.queryPermission(opts)) === "granted") return true;
    } catch (err) {
      // Older implementations may not support queryPermission options; try an explicit request.
    }
    try {
      if (typeof handle?.requestPermission === "function" && (await handle.requestPermission(opts)) === "granted") return true;
    } catch (err) {
      return false;
    }
    return false;
  };

  window.AutoBerichtFsHandle = {
    openHandleDb,
    saveHandle,
    loadHandle,
    requestHandlePermission,
    HANDLE_DB_NAME,
    HANDLE_STORE,
    HANDLE_KEY,
  };
})();
