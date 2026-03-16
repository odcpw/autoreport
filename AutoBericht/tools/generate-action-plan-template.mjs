import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import XLSX from "../libs/sheetjs/xlsx.full.min.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..", "..");

const outputArg = process.argv[2] || "sources/Template_Aktionsplan_IS_de.xlsx";
const outputPath = path.isAbsolute(outputArg) ? outputArg : path.resolve(repoRoot, outputArg);

const now = new Date();

const nextQuarterStart = (date) => {
  const year = date.getFullYear();
  const month = date.getMonth();
  const currentQuarter = Math.floor(month / 3) + 1;
  if (currentQuarter === 4) {
    return { year: year + 1, quarter: 1 };
  }
  return { year, quarter: currentQuarter + 1 };
};

const buildQuarterWindow = (date) => {
  const start = nextQuarterStart(date);
  const currentQuarter = Math.floor(date.getMonth() / 3) + 1;
  const remainingCurrentYear = Math.max(0, 4 - currentQuarter);
  const futureExtra = 8;
  const total = remainingCurrentYear + futureExtra;
  const quarters = [];
  let year = start.year;
  let quarter = start.quarter;
  for (let i = 0; i < total; i += 1) {
    quarters.push({ year, quarter, label: `Q${quarter}` });
    quarter += 1;
    if (quarter === 5) {
      quarter = 1;
      year += 1;
    }
  }
  return quarters;
};

const quarterWindow = buildQuarterWindow(now);

const headerFill = "D97A00";
const headerText = "FFFFFF";
const suvaFill = "FCE4D6";
const companyFill = "F5F5F5";
const bodyWhite = "FFFFFF";
const accentFill = "FFF2CC";
const borderColor = "D0D0D0";

const mkStyle = (overrides = {}) => ({
  font: { name: "Calibri", sz: 11, color: { rgb: "1F1F1F" } },
  alignment: { vertical: "top", wrapText: true },
  border: {
    top: { style: "thin", color: { rgb: borderColor } },
    bottom: { style: "thin", color: { rgb: borderColor } },
    left: { style: "thin", color: { rgb: borderColor } },
    right: { style: "thin", color: { rgb: borderColor } },
  },
  ...overrides,
});

const titleStyle = mkStyle({
  font: { name: "Calibri", sz: 15, bold: true, color: { rgb: headerText } },
  fill: { patternType: "solid", fgColor: { rgb: headerFill } },
  alignment: { horizontal: "left", vertical: "center" },
});

const noteStyle = mkStyle({
  font: { name: "Calibri", sz: 10, italic: true, color: { rgb: "555555" } },
  fill: { patternType: "solid", fgColor: { rgb: bodyWhite } },
});

const sectionStyle = mkStyle({
  font: { name: "Calibri", sz: 11, bold: true, color: { rgb: headerText } },
  fill: { patternType: "solid", fgColor: { rgb: headerFill } },
  alignment: { horizontal: "center", vertical: "center" },
});

const suvaHeaderStyle = mkStyle({
  font: { name: "Calibri", sz: 10, bold: true, color: { rgb: "7A3F00" } },
  fill: { patternType: "solid", fgColor: { rgb: suvaFill } },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
});

const companyHeaderStyle = mkStyle({
  font: { name: "Calibri", sz: 10, bold: true, color: { rgb: "333333" } },
  fill: { patternType: "solid", fgColor: { rgb: companyFill } },
  alignment: { horizontal: "center", vertical: "center", wrapText: true },
});

const quarterHeaderStyle = mkStyle({
  font: { name: "Calibri", sz: 10, bold: true, color: { rgb: "333333" } },
  fill: { patternType: "solid", fgColor: { rgb: accentFill } },
  alignment: { horizontal: "center", vertical: "center" },
});

const bodyStyle = mkStyle({
  fill: { patternType: "solid", fgColor: { rgb: bodyWhite } },
});

const suvaBodyStyle = mkStyle({
  fill: { patternType: "solid", fgColor: { rgb: "FFF8F2" } },
});

const companyBodyStyle = mkStyle({
  fill: { patternType: "solid", fgColor: { rgb: bodyWhite } },
});

const setCell = (ws, addr, value, style, opts = {}) => {
  ws[addr] = { v: value, t: opts.type || "s" };
  if (opts.formula) {
    ws[addr] = { f: opts.formula, t: opts.type || "s" };
  }
  if (style) ws[addr].s = style;
};

const rangeRef = (endCol, endRow) => `A1:${endCol}${endRow}`;

const colLetter = (index) => {
  let n = index;
  let out = "";
  while (n >= 0) {
    out = String.fromCharCode((n % 26) + 65) + out;
    n = Math.floor(n / 26) - 1;
  }
  return out;
};

const buildActionPlanSheet = () => {
  const ws = {};
  const baseColumns = [
    "Code",
    "Massnahmenpaket",
    "Ziel",
    "Inhalt / Umfang",
    "Verantwortlich",
    "Start",
    "Termin",
    "Status",
    "Strategisches Feld",
  ];
  const quarterStartCol = baseColumns.length;
  const remarksCol = quarterStartCol + quarterWindow.length;
  const headers = [...baseColumns, ...quarterWindow.map((q) => q.label), "Bemerkung"];

  setCell(ws, "A1", "AKTIONSPLAN - BASISPROJEKT INTEGRIERTE SICHERHEIT", titleStyle);
  setCell(
    ws,
    "A2",
    "Eine Zeile pro Massnahmenpaket. Die Berichtsbasis ordnet einzelne Feststellungen diesen Paketen zu.",
    noteStyle,
  );

  const yearGroups = [];
  quarterWindow.forEach((q, index) => {
    const last = yearGroups[yearGroups.length - 1];
    if (last && last.year === q.year) {
      last.end = index;
    } else {
      yearGroups.push({ year: q.year, start: index, end: index });
    }
  });

  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: remarksCol } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: remarksCol } },
  ];

  yearGroups.forEach((group) => {
    const startCol = quarterStartCol + group.start;
    const endCol = quarterStartCol + group.end;
    const addr = `${colLetter(startCol)}4`;
    setCell(ws, addr, String(group.year), sectionStyle);
    ws["!merges"].push({
      s: { r: 3, c: startCol },
      e: { r: 3, c: endCol },
    });
  });

  headers.forEach((header, index) => {
    const cell = `${colLetter(index)}5`;
    const style = index >= quarterStartCol && index < remarksCol
      ? quarterHeaderStyle
      : companyHeaderStyle;
    setCell(ws, cell, header, style);
  });

  const startRow = 6;
  const blankRows = 40;
  for (let r = 0; r < blankRows; r += 1) {
    const rowNo = startRow + r;
    headers.forEach((_, index) => {
      const style = index >= quarterStartCol && index < remarksCol ? companyBodyStyle : bodyStyle;
      setCell(ws, `${colLetter(index)}${rowNo}`, "", style);
    });
  }

  ws["!cols"] = [
    { wch: 14 },
    { wch: 26 },
    { wch: 26 },
    { wch: 34 },
    { wch: 18 },
    { wch: 12 },
    { wch: 12 },
    { wch: 14 },
    { wch: 18 },
    ...quarterWindow.map(() => ({ wch: 10 })),
    { wch: 24 },
  ];
  ws["!ref"] = rangeRef(colLetter(remarksCol), startRow + blankRows - 1);
  ws["!autofilter"] = { ref: `A5:${colLetter(remarksCol)}5` };
  ws["!rows"] = [{ hpt: 24 }, { hpt: 22 }, {}, { hpt: 22 }, { hpt: 36 }];
  return ws;
};

const buildReportBaseSheet = () => {
  const ws = {};
  const headers = [
    "Berichts-Nr.",
    "Kapitel",
    "Thema",
    "Feststellung",
    "Empfehlung Suva",
    "Prioritaet",
    "Entscheid",
    "Massnahmenpaket-Code",
    "Massnahmenpaket",
    "Kommentar",
  ];

  ws["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
    { s: { r: 0, c: 6 }, e: { r: 0, c: headers.length - 1 } },
  ];

  setCell(ws, "A1", "SUVA / BERICHT", sectionStyle);
  setCell(ws, "G1", "UNTERNEHMEN", sectionStyle);

  headers.forEach((header, index) => {
    const style = index <= 5 ? suvaHeaderStyle : companyHeaderStyle;
    setCell(ws, `${colLetter(index)}2`, header, style);
  });

  const startRow = 3;
  const blankRows = 120;
  for (let r = 0; r < blankRows; r += 1) {
    const rowNo = startRow + r;
    headers.forEach((_, index) => {
      const cell = `${colLetter(index)}${rowNo}`;
      const style = index <= 5 ? suvaBodyStyle : companyBodyStyle;
      if (index === 8) {
        setCell(ws, cell, "", style, {
          formula: `IF($H${rowNo}="","",IFERROR(INDEX(Aktionsplan!$B:$B,MATCH($H${rowNo},Aktionsplan!$A:$A,0)),""))`,
        });
      } else {
        setCell(ws, cell, "", style);
      }
    });
  }

  ws["!cols"] = [
    { wch: 12 },
    { wch: 12 },
    { wch: 20 },
    { wch: 42 },
    { wch: 42 },
    { wch: 10 },
    { wch: 18 },
    { wch: 20 },
    { wch: 26 },
    { wch: 24 },
  ];
  ws["!ref"] = rangeRef(colLetter(headers.length - 1), startRow + blankRows - 1);
  ws["!autofilter"] = { ref: `A2:${colLetter(headers.length - 1)}2` };
  ws["!rows"] = [{ hpt: 22 }, { hpt: 36 }];
  return ws;
};

const buildInstructionsSheet = () => {
  const ws = {};
  const lines = [
    "HINWEISE",
    "",
    "1. Im Tab 'Aktionsplan' zunaechst die Massnahmenpakete definieren.",
    "2. Im Tab 'Berichtsbasis' jede relevante Feststellung einem Entscheid zuordnen und anschliessend ueber den Code einem Massnahmenpaket zuweisen.",
    "3. Mehrere Berichts-Punkte duerfen demselben Massnahmenpaket zugeordnet werden.",
    "4. Der Tab 'Aktionsplan' ist die Fuehrungssicht; der Tab 'Berichtsbasis' bleibt die Rueckverfolgbarkeit zum Bericht.",
    `5. Planungshorizont ab naechstem Quartal: ${quarterWindow[0].year}-Q${quarterWindow[0].quarter} bis ${quarterWindow[quarterWindow.length - 1].year}-Q${quarterWindow[quarterWindow.length - 1].quarter}.`,
    "",
    "Farblogik:",
    "- Links orange getoent = Bericht / Suva",
    "- Rechts neutral = Verantwortung des Unternehmens",
  ];

  lines.forEach((line, index) => {
    const row = index + 1;
    const style = row === 1 ? titleStyle : noteStyle;
    setCell(ws, `A${row}`, line, style);
  });
  ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }];
  ws["!cols"] = [{ wch: 110 }];
  ws["!ref"] = "A1:A11";
  return ws;
};

const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, buildActionPlanSheet(), "Aktionsplan");
XLSX.utils.book_append_sheet(workbook, buildReportBaseSheet(), "Berichtsbasis");
XLSX.utils.book_append_sheet(workbook, buildInstructionsSheet(), "Hinweise");

workbook.Props = {
  Title: "Template Aktionsplan Basisprojekt Integrierte Sicherheit",
  Subject: "Draft template for IS action planning",
  Author: "Codex",
  Company: "Suva",
};

const buffer = XLSX.write(workbook, {
  bookType: "xlsx",
  type: "buffer",
  cellStyles: true,
});
fs.mkdirSync(path.dirname(outputPath), { recursive: true });
fs.writeFileSync(outputPath, buffer);

const startQuarter = quarterWindow[0];
const endQuarter = quarterWindow[quarterWindow.length - 1];
console.log(`Wrote ${outputPath}`);
console.log(`Quarter window: ${startQuarter.year}-Q${startQuarter.quarter} -> ${endQuarter.year}-Q${endQuarter.quarter}`);
