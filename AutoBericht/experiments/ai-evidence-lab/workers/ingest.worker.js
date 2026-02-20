function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

self.__canceledRuns = self.__canceledRuns || new Set();

self.onmessage = async (event) => {
  const msg = event.data || {};
  if (msg.type === "cancel") {
    const runId = String(msg.runId || "");
    if (runId) self.__canceledRuns.add(runId);
    return;
  }
  if (msg.type !== "start") return;

  const runId = String(msg.runId || "");
  const items = Array.isArray(msg.items) ? msg.items : [];
  const batchSize = Math.max(1, Math.min(256, Number(msg.batchSize || 16)));
  const batchDelay = Math.max(1, Math.round(20 / Math.sqrt(batchSize)));

  self.postMessage({
    type: "start",
    runId,
    total: items.length,
  });

  let completed = 0;
  for (const item of items) {
    if (self.__canceledRuns.has(runId)) {
      self.__canceledRuns.delete(runId);
      self.postMessage({
        type: "canceled",
        runId,
        completed,
        total: items.length,
      });
      return;
    }

    self.postMessage({
      type: "progress",
      runId,
      path: String(item.path || ""),
      status: "running",
      completed,
      total: items.length,
    });
    await sleep(batchDelay);
    completed += 1;
    self.postMessage({
      type: "progress",
      runId,
      path: String(item.path || ""),
      status: "done",
      completed,
      total: items.length,
    });
  }

  self.postMessage({
    type: "done",
    runId,
    completed,
    total: items.length,
  });
};
