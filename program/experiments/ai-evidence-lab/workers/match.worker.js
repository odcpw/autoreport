self.onmessage = (event) => {
  const msg = event.data || {};
  if (msg.type === "ping") {
    self.postMessage({ type: "pong", worker: "match" });
    return;
  }

  if (msg.type === "start") {
    self.postMessage({
      type: "done",
      worker: "match",
      runId: String(msg.runId || ""),
      completed: true,
    });
  }
};
