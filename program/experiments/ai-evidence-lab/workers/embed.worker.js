self.onmessage = (event) => {
  const msg = event.data || {};
  if (msg.type === "ping") {
    self.postMessage({ type: "pong", worker: "embed" });
    return;
  }

  if (msg.type === "start") {
    self.postMessage({
      type: "done",
      worker: "embed",
      runId: String(msg.runId || ""),
      completed: true,
    });
  }
};
