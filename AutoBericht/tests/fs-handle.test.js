const test = require("node:test");
const assert = require("node:assert/strict");
const { loadBrowserScripts } = require("./helpers");

const createIndexedDb = ({ fail = false } = {}) => ({
  open() {
    const request = {};
    const db = {
      objectStoreNames: { contains: () => true },
      transaction() {
        const tx = {
          error: fail ? new Error("handle persistence failed") : null,
          objectStore: () => ({
            put() {
              setTimeout(() => {
                if (fail) tx.onerror?.();
                else tx.oncomplete?.();
              }, 15);
            },
          }),
        };
        return tx;
      },
    };
    setTimeout(() => {
      request.result = db;
      request.onsuccess?.();
    }, 0);
    return request;
  },
});

test("folder-handle persistence waits for the IndexedDB transaction commit", async () => {
  const indexedDB = createIndexedDb();
  const api = loadBrowserScripts(["mini/shared/fs-handle.js"], { indexedDB }).AutoBerichtFsHandle;
  const started = Date.now();
  await api.saveHandle({ kind: "directory" });
  assert.equal(Date.now() - started >= 10, true);
});

test("folder-handle persistence surfaces IndexedDB transaction failure", async () => {
  const indexedDB = createIndexedDb({ fail: true });
  const api = loadBrowserScripts(["mini/shared/fs-handle.js"], { indexedDB }).AutoBerichtFsHandle;
  await assert.rejects(api.saveHandle({ kind: "directory" }), /handle persistence failed/);
});
