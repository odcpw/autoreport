const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const { webcrypto } = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");

const loadBrowserScripts = (relativePaths, additions = {}) => {
  const context = {
    console,
    crypto: webcrypto,
    structuredClone,
    TextEncoder,
    TextDecoder,
    URL,
    setTimeout,
    clearTimeout,
    ...additions,
  };
  context.window = context;
  context.globalThis = context;
  vm.createContext(context);
  relativePaths.forEach((relativePath) => {
    const filename = path.join(ROOT, relativePath);
    vm.runInContext(fs.readFileSync(filename, "utf8"), context, { filename });
  });
  return context;
};

const createMemoryDirectory = (initial = {}, options = {}) => {
  const files = new Map(Object.entries(initial).map(([name, value]) => [name, String(value)]));
  const directory = {
    name: "project",
    async getFileHandle(name, fileOptions = {}) {
      if (!files.has(name) && !fileOptions.create) {
        const error = new Error(`${name} not found`);
        error.name = "NotFoundError";
        throw error;
      }
      if (!files.has(name)) files.set(name, "");
      return {
        kind: "file",
        name,
        async getFile() {
          const text = files.get(name);
          return {
            size: Buffer.byteLength(text),
            async text() { return text; },
          };
        },
        async createWritable() {
          let pending = "";
          return {
            async write(value) {
              if (options.failWrite) throw new Error("disk full");
              pending = typeof value === "string" ? value : Buffer.from(value).toString("utf8");
            },
            async close() {
              if (options.failClose) throw new Error("close failed");
              files.set(name, options.mutateAfterWrite ? options.mutateAfterWrite(pending) : pending);
            },
            async abort() {},
          };
        },
      };
    },
    async *entries() {
      for (const name of files.keys()) {
        yield [name, await directory.getFileHandle(name)];
      }
    },
    read(name) { return files.get(name); },
  };
  return directory;
};

module.exports = { ROOT, loadBrowserScripts, createMemoryDirectory };
