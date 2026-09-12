(() => {
  const fail = (message, options = {}) => {
    const appName = String(options.appName || "AutoBericht");
    const detail = `${appName} cannot start: ${message}`;
    if (typeof document !== "undefined") {
      const statusEl = document.getElementById(String(options.statusElementId || "status"));
      if (statusEl) statusEl.textContent = detail;
    }
    throw new Error(detail);
  };

  const requireModules = (requirements, options = {}) => {
    const resolved = {};
    Object.entries(requirements || {}).forEach(([globalName, functionNames]) => {
      const module = window[globalName];
      if (!module || (typeof module !== "object" && typeof module !== "function")) {
        fail(`required module ${globalName} did not load.`, options);
      }
      (functionNames || []).forEach((functionName) => {
        if (typeof module[functionName] !== "function") {
          fail(`required function ${globalName}.${functionName} did not load.`, options);
        }
      });
      resolved[globalName] = module;
    });
    return resolved;
  };

  window.AutoBerichtDependencies = { requireModules };
})();
