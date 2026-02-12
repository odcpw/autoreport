/*
 * Shared spider chart renderer used by:
 * - Project page UI preview (`drawToCanvas`)
 * - Word export image generation (`drawToBlob`)
 *
 * Keep style/layout logic centralized here to avoid UI/export drift.
 */
(() => {
  const FONT_FAMILY = "system-ui, -apple-system, Segoe UI, sans-serif";

  const wrapLabel = (ctx, text, maxWidth, maxLines) => {
    const words = String(text || "").trim().split(/\s+/).filter(Boolean);
    if (!words.length) return [""];

    const lines = [];
    let current = "";
    words.forEach((word) => {
      const candidate = current ? `${current} ${word}` : word;
      if (!current || ctx.measureText(candidate).width <= maxWidth) {
        current = candidate;
      } else {
        lines.push(current);
        current = word;
      }
    });
    if (current) lines.push(current);
    if (lines.length <= maxLines) return lines;

    const trimmed = lines.slice(0, maxLines);
    let last = trimmed[maxLines - 1];
    while (last.length > 1 && ctx.measureText(`${last}…`).width > maxWidth) {
      last = last.slice(0, -1);
    }
    trimmed[maxLines - 1] = `${last}…`;
    return trimmed;
  };

  const drawSpider = (ctx, rows, options = {}) => {
    const width = Number(options.width || 760);
    const height = Number(options.height || 500);
    const companyLabel = String(options.companyLabel || "Company");
    const suvaLabel = String(options.suvaLabel || "Suva");
    const spiderRows = Array.isArray(rows) ? rows : [];
    const labels = spiderRows.map((row) => String(row?.displayLabel || row?.label || row?.id || ""));
    const companyValues = spiderRows.map((row) => Number(row?.company || 0));
    const consultantValues = spiderRows.map((row) => Number(row?.consultant || 0));

    const cx = Math.round(width * 0.5);
    const cy = Math.round(height * 0.46);
    const radius = Math.round(Math.min(width, height) * 0.31);
    const count = labels.length || 1;
    const steps = 5;
    const maxValue = 100;

    ctx.clearRect(0, 0, width, height);
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, width, height);

    ctx.strokeStyle = "#d9d5ce";
    ctx.lineWidth = 1;
    for (let ring = 1; ring <= steps; ring += 1) {
      const r = (radius * ring) / steps;
      ctx.beginPath();
      for (let i = 0; i < count; i += 1) {
        const angle = (-Math.PI / 2) + (Math.PI * 2 * i) / count;
        const x = cx + Math.cos(angle) * r;
        const y = cy + Math.sin(angle) * r;
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      }
      ctx.closePath();
      ctx.stroke();
    }

    ctx.strokeStyle = "#c4beb4";
    labels.forEach((label, index) => {
      const angle = (-Math.PI / 2) + (Math.PI * 2 * index) / count;
      const x = cx + Math.cos(angle) * radius;
      const y = cy + Math.sin(angle) * radius;
      ctx.beginPath();
      ctx.moveTo(cx, cy);
      ctx.lineTo(x, y);
      ctx.stroke();

      const textX = cx + Math.cos(angle) * (radius + 16);
      const textY = cy + Math.sin(angle) * (radius + 16);
      ctx.fillStyle = "#4b5563";
      ctx.font = `14px ${FONT_FAMILY}`;
      const horizontal = Math.cos(angle);
      const align = horizontal > 0.2 ? "left" : horizontal < -0.2 ? "right" : "center";
      ctx.textAlign = align;
      ctx.textBaseline = "top";
      const maxWidth = align === "center" ? 260 : 220;
      const lines = wrapLabel(ctx, label, maxWidth, 2);
      let xLabel = textX;
      if (align === "left") xLabel = Math.min(xLabel, width - maxWidth - 10);
      if (align === "right") xLabel = Math.max(xLabel, maxWidth + 10);
      if (align === "center") xLabel = Math.max((maxWidth * 0.5) + 10, Math.min(width - (maxWidth * 0.5) - 10, xLabel));
      if (/^\s*4\./.test(label)) xLabel += 10;
      const lineHeight = 16;
      const topNudge = index === 0 ? -14 : 0;
      const startY = textY - ((lines.length - 1) * lineHeight * 0.5) + topNudge;
      lines.forEach((line, lineIndex) => {
        ctx.fillText(line, xLabel, startY + (lineIndex * lineHeight));
      });
    });

    const drawPolygon = (values, stroke, fill) => {
      ctx.beginPath();
      values.forEach((value, index) => {
        const pct = Math.max(0, Math.min(maxValue, Number(value || 0))) / maxValue;
        const angle = (-Math.PI / 2) + (Math.PI * 2 * index) / count;
        const x = cx + Math.cos(angle) * radius * pct;
        const y = cy + Math.sin(angle) * radius * pct;
        if (index === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.closePath();
      ctx.fillStyle = fill;
      ctx.strokeStyle = stroke;
      ctx.lineWidth = 2;
      ctx.fill();
      ctx.stroke();
    };

    drawPolygon(companyValues, "#2563eb", "rgba(37,99,235,0.14)");
    drawPolygon(consultantValues, "#dc2626", "rgba(220,38,38,0.11)");

    const legendTopY = height - 52;
    const legendLineHeight = 18;
    const legendColumnMaxWidth = 220;
    const markerWidth = 16;
    const markerTextGap = 8;
    const legendGap = 36;

    ctx.textAlign = "left";
    ctx.textBaseline = "top";
    ctx.font = `17px ${FONT_FAMILY}`;

    const leftLegendLines = wrapLabel(ctx, companyLabel, legendColumnMaxWidth, 2);
    const rightLegendLines = wrapLabel(ctx, suvaLabel, legendColumnMaxWidth, 2);
    const leftTextWidth = leftLegendLines.reduce((max, line) => Math.max(max, ctx.measureText(line).width), 0);
    const rightTextWidth = rightLegendLines.reduce((max, line) => Math.max(max, ctx.measureText(line).width), 0);
    const leftBlockWidth = markerWidth + markerTextGap + leftTextWidth;
    const rightBlockWidth = markerWidth + markerTextGap + rightTextWidth;
    const totalLegendWidth = leftBlockWidth + legendGap + rightBlockWidth;
    const leftMarkerX = Math.round((width - totalLegendWidth) * 0.5);
    const rightMarkerX = leftMarkerX + leftBlockWidth + legendGap;
    const markerY = legendTopY + 8;
    const leftTextX = leftMarkerX + markerWidth + markerTextGap;
    const rightTextX = rightMarkerX + markerWidth + markerTextGap;

    ctx.fillStyle = "#2563eb";
    ctx.fillRect(leftMarkerX, markerY, markerWidth, 3);
    leftLegendLines.forEach((line, lineIndex) => {
      ctx.fillText(line, leftTextX, legendTopY + (lineIndex * legendLineHeight));
    });

    ctx.fillStyle = "#dc2626";
    ctx.fillRect(rightMarkerX, markerY, markerWidth, 3);
    rightLegendLines.forEach((line, lineIndex) => {
      ctx.fillText(line, rightTextX, legendTopY + (lineIndex * legendLineHeight));
    });
  };

  const drawToCanvas = (canvas, rows, options = {}) => {
    if (!canvas) return;
    const width = Number(options.width || 760);
    const height = Number(options.height || 500);
    const rawDpr = Number(options.dpr);
    const dpr = Number.isFinite(rawDpr) && rawDpr > 0 ? rawDpr : 1;
    const setCssSize = options.setCssSize !== false;
    if (setCssSize) {
      canvas.style.width = `${width}px`;
      canvas.style.height = `${height}px`;
    }
    canvas.width = Math.round(width * dpr);
    canvas.height = Math.round(height * dpr);
    const ctx = canvas.getContext("2d");
    if (!ctx) return;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    drawSpider(ctx, rows, { ...options, width, height });
  };

  const drawToBlob = async (rows, options = {}) => {
    const width = Number(options.width || 760);
    const height = Number(options.height || 500);
    const dpr = Number.isFinite(Number(options.dpr)) && Number(options.dpr) > 0
      ? Number(options.dpr)
      : 2;
    const type = String(options.type || "image/png");
    const quality = Number.isFinite(Number(options.quality)) ? Number(options.quality) : 0.95;

    const canvas = document.createElement("canvas");
    drawToCanvas(canvas, rows, {
      ...options,
      width,
      height,
      dpr,
      setCssSize: false,
    });

    const blob = await new Promise((resolve) => {
      canvas.toBlob((result) => resolve(result), type, quality);
    });
    if (!blob) throw new Error("Failed to render spider image");
    return blob;
  };

  window.AutoBerichtSpiderChart = {
    drawToCanvas,
    drawToBlob,
  };
})();
