(() => {
  const init = (ctx, deps) => {
    const { elements, state, runtime, setStatus, debug } = ctx;
    const { spider, ioApi } = deps;

    const closeModal = () => {
      if (!elements.spiderModal) return;
      elements.spiderModal.classList.remove("is-open");
      elements.spiderModal.setAttribute("aria-hidden", "true");
    };

    const openModal = async () => {
      if (!elements.spiderModal) return;
      try {
        setStatus("Recalculating spider values...");
        const spiderData = await spider.computeSpider({
          project: state.project,
          overrides: state.spiderOverrides || {},
          dirHandle: runtime.dirHandle,
        });
        renderTable(spiderData);
        setStatus("Spider values ready.");
        elements.spiderModal.classList.add("is-open");
        elements.spiderModal.setAttribute("aria-hidden", "false");
      } catch (err) {
        setStatus(`Spider calc failed: ${err.message || err}`);
        debug.logLine("error", `Spider calc failed: ${err.message || err}`);
      }
    };

    const renderTable = (spiderData) => {
      if (!elements.spiderTableBody) return;
      const rows = spiderData.effective.chapters_1_14;
      const overrides = spiderData.overrides || {};
      const baseline = spiderData.baseline.chapters_1_14;
      const baselineMap = new Map(baseline.map((r) => [r.id, r]));
      elements.spiderTableBody.innerHTML = "";
      rows.forEach((row) => {
        const base = baselineMap.get(row.id) || row;
        const ov = overrides[row.id] || {};
        const tr = document.createElement("tr");
        const cells = [
          `<td>${row.id}</td>`,
          `<td>${(base.company ?? 0).toFixed(1)}%</td>`,
          `<td><input type="number" step="0.1" min="0" max="100" data-field="company" data-id="${row.id}" value="${Number.isFinite(ov.company) ? ov.company : ""}"></td>`,
          `<td><input type="checkbox" data-field="useCompany" data-id="${row.id}" ${ov.useCompany ? "checked" : ""}></td>`,
          `<td>${(base.consultant ?? 0).toFixed(1)}%</td>`,
          `<td><input type="number" step="0.1" min="0" max="100" data-field="consultant" data-id="${row.id}" value="${Number.isFinite(ov.consultant) ? ov.consultant : ""}"></td>`,
          `<td><input type="checkbox" data-field="useConsultant" data-id="${row.id}" ${ov.useConsultant ? "checked" : ""}></td>`,
        ];
        tr.innerHTML = cells.join("");
        elements.spiderTableBody.appendChild(tr);
      });
    };

    const saveOverrides = () => {
      const body = elements.spiderTableBody;
      if (!body) return;
      const next = {};
      body.querySelectorAll("tr").forEach((tr) => {
        const id = tr.querySelector("input[data-id]")?.dataset.id;
        if (!id) return;
        const companyVal = tr.querySelector('input[data-field="company"]')?.value;
        const consultantVal = tr.querySelector('input[data-field="consultant"]')?.value;
        const useCompany = tr.querySelector('input[data-field="useCompany"]')?.checked || false;
        const useConsultant = tr.querySelector('input[data-field="useConsultant"]')?.checked || false;
        const company = companyVal === "" ? null : Number(companyVal);
        const consultant = consultantVal === "" ? null : Number(consultantVal);
        next[id] = {
          useCompany,
          useConsultant,
          company: Number.isFinite(company) ? company : null,
          consultant: Number.isFinite(consultant) ? consultant : null,
        };
      });
      state.spiderOverrides = next;
      setStatus("Spider overrides saved (autosave will persist).");
    };

    if (elements.openSpiderBtn) {
      elements.openSpiderBtn.addEventListener("click", openModal);
    }
    if (elements.spiderCloseBtn) {
      elements.spiderCloseBtn.addEventListener("click", closeModal);
    }
    if (elements.spiderBackdrop) {
      elements.spiderBackdrop.addEventListener("click", closeModal);
    }
    if (elements.spiderCancelBtn) {
      elements.spiderCancelBtn.addEventListener("click", closeModal);
    }
    if (elements.spiderSaveBtn) {
      elements.spiderSaveBtn.addEventListener("click", () => {
        saveOverrides();
        if (ioApi?.saveSidecar) {
          ioApi.saveSidecar().catch(() => {});
        }
        closeModal();
      });
    }

    return {
      openModal,
      saveOverrides,
    };
  };

  window.AutoBerichtSpiderUi = { init };
})();
