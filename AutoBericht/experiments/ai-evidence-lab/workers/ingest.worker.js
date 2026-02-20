function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

self.onmessage = async (event) => {
  const msg = event.data || {};
  if (msg.type !== "start") return;

  const runId = String(msg.runId || "");
  const items = Array.isArray(msg.items) ? msg.items : [];

  self.postMessage({
    type: "start",
    runId,
    total: items.length,
  });

  let completed = 0;
  for (const item of items) {
    await sleep(4);
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
