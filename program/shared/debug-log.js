(() => {
  if (window.AutoReportDebug) return;

  const logLines = [];

  const logLine = (level, message) => {
    const timestamp = new Date().toISOString();
    const line = `[${timestamp}] ${level.toUpperCase()}: ${message}`;
    logLines.push(line);
  };

  const formatArgs = (args) => args.map((item) => {
    if (typeof item === "string") return item;
    try {
      return JSON.stringify(item);
    } catch (err) {
      return String(item);
    }
  }).join(" ");

  const captureConsole = () => {
    const original = {
      log: console.log,
      warn: console.warn,
      error: console.error,
      info: console.info,
      debug: console.debug,
    };

    Object.keys(original).forEach((key) => {
      console[key] = (...args) => {
        logLine(key, formatArgs(args));
        original[key](...args);
      };
    });

    window.addEventListener("error", (event) => {
      logLine("error", `${event.message} @ ${event.filename}:${event.lineno}:${event.colno}`);
    });

    window.addEventListener("unhandledrejection", (event) => {
      logLine("error", `Unhandled rejection: ${event.reason}`);
    });
  };

  const saveLog = async ({ suggestedName, dirHandle } = {}) => {
    const content = logLines.join("\n") || "No log entries yet.";
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const filename = suggestedName || `debug-log-${timestamp}.txt`;

    const writeToHandle = async (handle) => {
      const writable = await handle.createWritable();
      await writable.write(content);
      await writable.close();
    };

    if (dirHandle) {
      try {
        const handle = await dirHandle.getFileHandle(filename, { create: true });
        await writeToHandle(handle);
        return { location: "folder", filename };
      } catch (err) {
        logLine("warn", `Folder log save failed: ${err.message || err}`);
      }
    }

    if (window.showSaveFilePicker) {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName: filename,
          types: [{ description: "Text", accept: { "text/plain": [".txt"] } }],
        });
        await writeToHandle(handle);
        return { location: "picker", filename };
      } catch (err) {
        logLine("warn", `Save picker canceled or failed: ${err.message || err}`);
      }
    }

    const blob = new Blob([content], { type: "text/plain" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
    URL.revokeObjectURL(link.href);
    return { location: "download", filename };
  };

  captureConsole();

  window.AutoReportDebug = {
    logLines,
    logLine,
    saveLog,
  };
})();
