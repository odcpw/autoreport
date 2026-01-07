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
    const tx = db.transaction(HANDLE_STORE, "readwrite");
    tx.objectStore(HANDLE_STORE).put(handle, HANDLE_KEY);
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
    try {
      const opts = { mode: "readwrite" };
      if ((await handle.queryPermission(opts)) === "granted") return true;
      if ((await handle.requestPermission(opts)) === "granted") return true;
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
