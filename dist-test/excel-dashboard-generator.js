"use strict";
var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// src/lib/excel-dashboard-generator.ts
var excel_dashboard_generator_exports = {};
__export(excel_dashboard_generator_exports, {
  generateDashboardExcel: () => generateDashboardExcel
});
module.exports = __toCommonJS(excel_dashboard_generator_exports);
var XLSX = __toESM(require("xlsx-js-style"));
var FMT_ACCT_DOLLAR = '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"??_);_(@_)';
var FMT_ACCT_COUNT = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)';
var FMT_PCT_WHOLE = "0%";
var FMT_INT_COMMA = "#,##0";
var FMT_DATE_MMM_YY = "[$-409]mmm\\-yy;@";
var NAVY = "0F2044";
var WHITE = "FFFFFF";
var BORDER_COL = "D1D5DB";
var THIN_BORDER = {
  top: { style: "thin", color: { rgb: BORDER_COL } },
  bottom: { style: "thin", color: { rgb: BORDER_COL } },
  left: { style: "thin", color: { rgb: BORDER_COL } },
  right: { style: "thin", color: { rgb: BORDER_COL } }
};
function navyHeader(text) {
  return {
    v: text ?? "",
    t: "s",
    s: {
      fill: { fgColor: { rgb: NAVY } },
      font: { bold: true, color: { rgb: WHITE }, sz: 11 },
      border: THIN_BORDER,
      alignment: { horizontal: "center", vertical: "center", wrapText: true }
    }
  };
}
function goldTitle(text) {
  return {
    v: text,
    t: "s",
    s: {
      font: { bold: true, sz: 14 },
      alignment: { horizontal: "center", vertical: "center" }
    }
  };
}
function sectionHeader(text) {
  return {
    v: text,
    t: "s",
    s: {
      fill: { fgColor: { rgb: NAVY } },
      font: { bold: true, color: { rgb: WHITE }, sz: 14 },
      alignment: { horizontal: "center", vertical: "center" }
    }
  };
}
function noFillFormulaCell(formula, numFmt, cachedVal) {
  return { f: formula, v: cachedVal ?? 0, t: "n", s: { ...numFmt ? { numFmt } : {} } };
}
function noFillBoldFormulaCell(formula, numFmt, cachedVal) {
  return { f: formula, v: cachedVal ?? 0, t: "n", s: { font: { bold: true }, ...numFmt ? { numFmt } : {} } };
}
function str(v) {
  return v == null ? "" : String(v).trim();
}
function num(v) {
  if (v == null || v === "") return null;
  const s = String(v).replace(/[$,%\s]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function isExcelSerial(v) {
  return typeof v === "number" && v > 4e4 && v < 6e4;
}
function serialToMonthCode(serial) {
  const d = XLSX.SSF.parse_date_code(serial);
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const yr = String(d.y).slice(-2);
  return `${months[d.m - 1]}-${yr}`;
}
function colLetter(idx) {
  let result = "";
  let n = idx;
  do {
    result = String.fromCharCode(65 + n % 26) + result;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return result;
}
function sheetRef(sheetName, addr) {
  return `'${sheetName}'!${addr}`;
}
function parseSalesByMonth(wb) {
  const sheetName = wb.SheetNames.find((n) => /sales.?by.?month/i.test(n));
  if (!sheetName) throw new Error("Cannot find 'Sales By Month' sheet in workbook");
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
  let headerRowIdx = -1;
  const monthColIndices = [];
  const monthSerials = [];
  for (let i = 0; i < rows.length; i++) {
    const serials = rows[i].filter(isExcelSerial);
    if (serials.length >= 3) {
      headerRowIdx = i;
      rows[i].forEach((cell, j) => {
        if (isExcelSerial(cell)) {
          monthSerials.push(cell);
          monthColIndices.push(j);
        }
      });
      break;
    }
  }
  if (headerRowIdx < 0) throw new Error("Cannot find monthly header row in Sales By Month sheet");
  const monthCodes = monthSerials.map(serialToMonthCode);
  const lookupByLabel = /* @__PURE__ */ new Map();
  const lookupRowIdx = /* @__PURE__ */ new Map();
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const label = str(row[0]);
    if (!label) continue;
    const values = monthColIndices.map((j) => num(row[j]));
    const key = label.toLowerCase().trim();
    lookupByLabel.set(key, values);
    lookupRowIdx.set(key, i + 1);
  }
  let currentMonthIdx = 0;
  const incomeLine = lookupByLabel.get("total income") ?? lookupByLabel.get("total revenue") ?? lookupByLabel.get("flex income");
  if (incomeLine) {
    const nonZero = incomeLine.filter((v) => v != null && v > 0);
    const peak = nonZero.length > 0 ? Math.max(...nonZero) : 0;
    const threshold = peak * 0.15;
    for (let i = incomeLine.length - 1; i >= 0; i--) {
      if ((incomeLine[i] ?? 0) >= threshold) {
        currentMonthIdx = i;
        break;
      }
    }
  }
  return {
    sheetName,
    rows,
    headerRowIdx,
    monthCodes,
    monthSerials,
    monthColIndices,
    currentMonthIdx,
    lookupByLabel,
    lookupRowIdx
  };
}
function parseOccupancyStats(wb) {
  const sheetName = wb.SheetNames.find((n) => n.startsWith("P.3"));
  if (!sheetName) throw new Error("Cannot find P.3 Occupancy sheet");
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });
  let headerRowIdx = -1;
  let statusCol = 3, spaceTypeCol = 4, sqFtCol = 6;
  for (let i = 0; i < rows.length; i++) {
    const rowLower = rows[i].map((c) => str(c).toLowerCase());
    if (rowLower.includes("status")) {
      headerRowIdx = i;
      const si = rowLower.findIndex((c) => c === "status");
      const ti = rowLower.findIndex((c) => c.includes("type"));
      const qi = rowLower.findIndex((c) => c.includes("sq"));
      if (si >= 0) statusCol = si;
      if (ti >= 0) spaceTypeCol = ti;
      if (qi >= 0) sqFtCol = qi;
      break;
    }
  }
  if (headerRowIdx < 0) throw new Error("Cannot find header row in P.3 Occupancy");
  let location = "";
  const dd = { total: 0, occupied: 0, totalSqFt: 0, occupiedSqFt: 0 };
  const po = { total: 0, occupied: 0, totalSqFt: 0, occupiedSqFt: 0 };
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const spaceName = str(row[1]);
    if (!spaceName) continue;
    if (!location) {
      const propVal = str(row[0]);
      const m = propVal.match(/coworking\s+(.+?)(?:\s*\(|$)/i);
      if (m) location = m[1].trim();
    }
    const isOccupied = str(row[statusCol]).toLowerCase().includes("occup");
    const spaceType = str(row[spaceTypeCol]).toLowerCase();
    const sqFt = num(row[sqFtCol]) ?? 0;
    if (spaceType.includes("private office")) {
      po.total++;
      po.totalSqFt += sqFt;
      if (isOccupied) {
        po.occupied++;
        po.occupiedSqFt += sqFt;
      }
    } else if (spaceType.includes("dedicated")) {
      dd.total++;
      dd.totalSqFt += sqFt;
      if (isOccupied) {
        dd.occupied++;
        dd.occupiedSqFt += sqFt;
      }
    }
  }
  return {
    sheetName,
    location: location || "Waco",
    statusCol,
    typeCol: spaceTypeCol,
    sqFtCol,
    headerRowIdx,
    dedicatedDesk: dd,
    privateOffice: po
  };
}
function buildP1Sheet(occ, sales, p2SheetName, wb) {
  const { currentMonthIdx, monthCodes, monthSerials, lookupByLabel, lookupRowIdx } = sales;
  const reportMonthSerial = monthSerials[currentMonthIdx];
  const statusLetter = colLetter(occ.statusCol);
  const typeLetter = colLetter(occ.typeCol);
  const sqftLetter = colLetter(occ.sqFtCol);
  const p3 = occ.sheetName;
  const currentMonthP2ColLetter = colLetter(sales.monthColIndices[currentMonthIdx] + 1);
  const revFormula = `+${sheetRef(p2SheetName, `${currentMonthP2ColLetter}16`)}`;
  const dd = occ.dedicatedDesk;
  const po = occ.privateOffice;
  function makeKpiRow(label, sheetRow, isTotal) {
    const bRef = `B${sheetRow}`;
    let cTotal, dOcc, fSqFt, gOccSqFt;
    if (isTotal) {
      const a1 = sheetRow - 2;
      const a2 = sheetRow - 1;
      cTotal = noFillBoldFormulaCell(`+C${a1}+C${a2}`, FMT_INT_COMMA, dd.total + po.total);
      dOcc = noFillBoldFormulaCell(`+D${a1}+D${a2}`, FMT_INT_COMMA, dd.occupied + po.occupied);
      fSqFt = noFillBoldFormulaCell(`+F${a1}+F${a2}`, FMT_INT_COMMA, dd.totalSqFt + po.totalSqFt);
      gOccSqFt = noFillBoldFormulaCell(`+G${a1}+G${a2}`, FMT_INT_COMMA, dd.occupiedSqFt + po.occupiedSqFt);
    } else {
      const isDd = label === "Dedicated Desk";
      cTotal = noFillFormulaCell(
        `COUNTIF(${sheetRef(p3, `${typeLetter}:${typeLetter}`)},${bRef})`,
        FMT_INT_COMMA,
        isDd ? dd.total : po.total
      );
      dOcc = noFillFormulaCell(
        `COUNTIFS(${sheetRef(p3, `${statusLetter}:${statusLetter}`)},"Occupied",${sheetRef(p3, `${typeLetter}:${typeLetter}`)},${bRef})`,
        FMT_INT_COMMA,
        isDd ? dd.occupied : po.occupied
      );
      fSqFt = noFillFormulaCell(
        `+SUMIF(${sheetRef(p3, `${typeLetter}:${typeLetter}`)},${bRef},${sheetRef(p3, `${sqftLetter}:${sqftLetter}`)})`,
        FMT_INT_COMMA,
        isDd ? dd.totalSqFt : po.totalSqFt
      );
      gOccSqFt = noFillFormulaCell(
        `SUMIFS(${sheetRef(p3, `$${sqftLetter}:$${sqftLetter}`)},${sheetRef(p3, `${statusLetter}:${statusLetter}`)},"Occupied",${sheetRef(p3, `${typeLetter}:${typeLetter}`)},${bRef})`,
        FMT_INT_COMMA,
        isDd ? dd.occupiedSqFt : po.occupiedSqFt
      );
    }
    const eSpacePct = isTotal ? noFillBoldFormulaCell(`+D${sheetRow}/C${sheetRow}`, FMT_PCT_WHOLE) : noFillFormulaCell(`+D${sheetRow}/C${sheetRow}`, FMT_PCT_WHOLE);
    const hSqFtPct = isTotal ? noFillBoldFormulaCell(`+G${sheetRow}/F${sheetRow}`, FMT_PCT_WHOLE) : noFillFormulaCell(`+G${sheetRow}/F${sheetRow}`, FMT_PCT_WHOLE);
    const lbl = isTotal ? { v: label, t: "s", s: { font: { bold: true }, alignment: { horizontal: "right" } } } : { v: label, t: "s" };
    return [null, lbl, cTotal, dOcc, eSpacePct, fSqFt, gOccSqFt, hSqFtPct];
  }
  const origP1SheetName = wb.SheetNames.find((n) => n.startsWith("P.1"));
  const allSnapRows = [];
  const snapRowHeights = [];
  let copiedAvgRow = null;
  if (origP1SheetName) {
    const origWs = wb.Sheets[origP1SheetName];
    const origRowInfo = origWs["!rows"] ?? [];
    const origRange = XLSX.utils.decode_range(origWs["!ref"] ?? "A1:A1");
    const snapStartRowIdx = 17;
    for (let rowIdx = snapStartRowIdx; rowIdx <= origRange.e.r; rowIdx++) {
      const bAddr = XLSX.utils.encode_cell({ r: rowIdx, c: 1 });
      const bCell = origWs[bAddr];
      if (!bCell) continue;
      const bVal = bCell.v;
      const bStr = typeof bVal === "string" ? bVal.trim() : "";
      if (!isExcelSerial(bVal) && /average/i.test(bStr)) {
        const RED_8 = "FFFF0000";
        const LGRAY = "E7E6E6";
        const avgStyleBase = (fmt, align) => ({
          fill: { fgColor: { rgb: LGRAY } },
          font: { bold: true, color: { rgb: RED_8 } },
          alignment: { horizontal: align },
          numFmt: fmt
        });
        const avgCells = [null];
        const bSrc = origWs[XLSX.utils.encode_cell({ r: rowIdx, c: 1 })];
        avgCells.push({ v: bSrc?.v ?? "Average", t: "s", s: avgStyleBase("mm-dd-yy", "right"), z: "mm-dd-yy" });
        for (let col = 2; col <= 8; col++) {
          const addr = XLSX.utils.encode_cell({ r: rowIdx, c: col });
          const c = origWs[addr];
          const fmt = c?.z ?? FMT_ACCT_COUNT;
          const cell = { v: c?.v ?? 0, t: "n", s: avgStyleBase(fmt, "center"), z: fmt };
          if (c?.f) cell.f = c.f;
          avgCells.push(cell);
        }
        copiedAvgRow = avgCells;
        break;
      }
      const dataCells = [null];
      for (let col = 1; col <= 8; col++) {
        const addr = XLSX.utils.encode_cell({ r: rowIdx, c: col });
        const c = origWs[addr];
        if (!c) {
          dataCells.push(null);
          continue;
        }
        const cell = { v: c.v, t: c.t ?? "s" };
        if (c.f) cell.f = c.f;
        if (c.s) cell.s = c.s;
        const fmt = c.z ?? c.s?.numFmt;
        if (fmt) cell.z = fmt;
        dataCells.push(cell);
      }
      allSnapRows.push(dataCells);
      const ri = origRowInfo[rowIdx];
      if (ri?.hidden) {
        snapRowHeights.push({ hidden: true });
      } else if (ri?.hpt != null) {
        snapRowHeights.push({ hpt: ri.hpt });
      } else {
        snapRowHeights.push(void 0);
      }
    }
  }
  const avgRow = copiedAvgRow ?? [
    null,
    { v: "Average", t: "s", s: { font: { bold: true, color: { rgb: "FF0000" } }, alignment: { horizontal: "right" }, numFmt: "mm-dd-yy" } },
    { f: "AVERAGE(C$36:C$47)", v: 0, t: "n", s: { font: { bold: true, color: { rgb: "FF0000" } }, numFmt: FMT_PCT_WHOLE } },
    { f: "AVERAGE(D$36:D$47)", v: 0, t: "n", s: { font: { bold: true, color: { rgb: "FF0000" } }, numFmt: FMT_PCT_WHOLE } },
    { f: "AVERAGE(E$36:E$47)", v: 0, t: "n", s: { font: { bold: true, color: { rgb: "FF0000" } }, numFmt: FMT_ACCT_COUNT } },
    { f: "AVERAGE(F$36:F$47)", v: 0, t: "n", s: { font: { bold: true, color: { rgb: "FF0000" } }, numFmt: FMT_ACCT_COUNT } },
    { f: "AVERAGE(G$36:G$47)", v: 0, t: "n", s: { font: { bold: true, color: { rgb: "FF0000" } }, numFmt: FMT_ACCT_COUNT } },
    { f: "AVERAGE(H$36:H$47)", v: 0, t: "n", s: { font: { bold: true, color: { rgb: "FF0000" } }, numFmt: FMT_ACCT_COUNT } },
    { f: "AVERAGE(I$36:I$47)", v: 0, t: "n", s: { font: { bold: true, color: { rgb: "FF0000" } }, numFmt: FMT_ACCT_COUNT } }
  ];
  const rowHeights = [
    { hpt: 18.75 },
    // row 1
    { hpt: 15.75 },
    // row 2
    { hpt: 6 },
    // row 3
    { hpt: 15.75 },
    // row 4
    { hpt: 15.75 },
    // row 5
    { hpt: 19.5 },
    // row 6
    { hpt: 15.75 },
    // row 7
    { hpt: 18.75 },
    // row 8
    { hpt: 45 },
    // row 9
    void 0,
    // row 10
    void 0,
    // row 11
    void 0,
    // row 12
    { hpt: 15.75 },
    // row 13
    { hpt: 9 },
    // row 14
    { hpt: 8.25 },
    // row 15
    { hpt: 19.5 },
    // row 16
    { hpt: 75.75 }
    // row 17 (snapshot column headers)
  ];
  for (const rh of snapRowHeights) {
    rowHeights.push(rh);
  }
  rowHeights.push({ hpt: 15.75 });
  const merges = [
    { s: { r: 0, c: 1 }, e: { r: 0, c: 7 } },
    // title
    { s: { r: 5, c: 1 }, e: { r: 5, c: 7 } },
    // KPI section header
    { s: { r: 7, c: 1 }, e: { r: 7, c: 7 } },
    // Current Occupancy header
    { s: { r: 15, c: 1 }, e: { r: 15, c: 8 } }
    // 12 Week Snapshot header
  ];
  const colWidths = [2.21, 31.64, 9.21, 9.21, 9.21, 9.21, 9.21, 9.21, 17.07, 21.21];
  const revenueDisplayCell = {
    f: revFormula,
    v: 0,
    t: "n",
    s: {
      font: { bold: true, sz: 11, color: { rgb: "FFFF0000" } },
      numFmt: FMT_ACCT_DOLLAR
    },
    z: FMT_ACCT_DOLLAR
  };
  const monthCodeCell = {
    v: reportMonthSerial,
    t: "n",
    s: { font: { bold: true, sz: 11, color: { rgb: NAVY } }, numFmt: FMT_DATE_MMM_YY },
    z: FMT_DATE_MMM_YY
  };
  const data = [
    // Row 1: Title merged B1:H1
    [null, goldTitle(`25N COWORKING ${occ.location.toUpperCase()} DASHBOARD`)],
    // Row 2: Month code (date serial)
    [null, monthCodeCell],
    // Row 3: thin spacer blank
    [],
    // Row 4: Revenue label + value
    // Historical: B4 is plain text (not bold, not navy)
    [
      null,
      { v: "Sales (For the month) ", t: "s" },
      null,
      null,
      null,
      null,
      null,
      revenueDisplayCell
    ],
    // Row 5: blank
    [],
    // Row 6: KPI section header merged B6:H6
    [null, sectionHeader("KEY PERFORMANCE INDICATORS - 25N COWORKING")],
    // Row 7: blank
    [],
    // Row 8: CURRENT OCCUPANCY header merged B8:H8
    [null, sectionHeader("CURRENT OCCUPANCY")],
    // Row 9: KPI column headers (tall row)
    // Historical: B9 "Workspace" is right-aligned; all other headers are centered
    [
      null,
      {
        v: "Workspace",
        t: "s",
        s: {
          fill: { fgColor: { rgb: NAVY } },
          font: { bold: true, color: { rgb: WHITE }, sz: 11 },
          border: THIN_BORDER,
          alignment: { horizontal: "right", vertical: "center", wrapText: true }
        }
      },
      navyHeader("# of Spaces"),
      navyHeader("Occupied"),
      navyHeader("% of Spaces"),
      navyHeader("Total Square Feet"),
      navyHeader("Occupied"),
      navyHeader("% of Square Feet")
    ],
    // Row 10: Dedicated Desk
    makeKpiRow("Dedicated Desk", 10, false),
    // Row 11: Private Office
    makeKpiRow("Private Office", 11, false),
    // Row 12: Total Spaces
    makeKpiRow("Total Spaces", 12, true),
    // Rows 13-15: blank rows with preserved number formatting (matches historical cols C-H)
    [
      null,
      null,
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_PCT_WHOLE } },
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_PCT_WHOLE } }
    ],
    [
      null,
      null,
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_PCT_WHOLE } },
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_PCT_WHOLE } }
    ],
    [
      null,
      null,
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_PCT_WHOLE } },
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_INT_COMMA } },
      { v: null, t: "s", s: { numFmt: FMT_PCT_WHOLE } }
    ],
    // Row 16: 12 Week Snapshot header merged B16:I16 (trailing space matches historical)
    [null, sectionHeader("12 Week Snapshot ")],
    // Row 17: Snapshot column headers (NO blank row between header and col headers)
    [
      null,
      navyHeader("Month"),
      navyHeader("Occupied % of Total Spaces"),
      navyHeader("Occupied % of Total Square Feet"),
      navyHeader("# of Contracts"),
      navyHeader("# of Flex Members"),
      navyHeader("# of Mail Members"),
      navyHeader("Recurring & Promo Meeting Hours Sold"),
      navyHeader("Meeting Hours Sold $")
    ],
    // Rows 18+: snapshot data (copied verbatim from source P.1)
    ...allSnapRows,
    // Last: average row (copied from source or fallback)
    avgRow
  ];
  return { data, merges, rowHeights, colWidths };
}
function aoaToStyledSheet(aoa, merges, rowHeights) {
  const ws = {};
  let maxR = 0;
  let maxC = 0;
  for (let r = 0; r < aoa.length; r++) {
    const row = aoa[r];
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (cell == null) continue;
      const addr = XLSX.utils.encode_cell({ r, c });
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
      if (typeof cell === "string") {
        ws[addr] = { t: "s", v: cell };
      } else if (typeof cell === "number") {
        ws[addr] = { t: "n", v: cell };
      } else {
        const sc = cell;
        if (sc.f) {
          ws[addr] = { t: sc.t ?? "n", f: sc.f, v: sc.v ?? 0 };
        } else {
          ws[addr] = { t: sc.t ?? "s", v: sc.v ?? "" };
        }
        if (sc.s) ws[addr].s = sc.s;
        const fmt = sc.z ?? sc.s?.numFmt;
        if (fmt) {
          ws[addr].z = fmt;
          if (ws[addr].s) ws[addr].s.numFmt = fmt;
          else ws[addr].s = { numFmt: fmt };
        }
      }
    }
  }
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
  if (merges && merges.length > 0) {
    ws["!merges"] = merges;
    for (const m of merges) {
      const anchorAddr = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
      const anchorFmt = ws[anchorAddr]?.s?.numFmt ?? "General";
      for (let r = m.s.r; r <= m.e.r; r++) {
        for (let c = m.s.c; c <= m.e.c; c++) {
          if (r === m.s.r && c === m.s.c) continue;
          const addr = XLSX.utils.encode_cell({ r, c });
          if (!ws[addr]) {
            ws[addr] = { t: "s", v: "", z: anchorFmt, s: { numFmt: anchorFmt } };
          }
        }
      }
    }
  }
  if (rowHeights && rowHeights.length > 0) {
    ws["!rows"] = rowHeights.map((r) => r ?? {});
  }
  return ws;
}
function setColWidthsPx(ws, widthsPx) {
  ws["!cols"] = widthsPx.map((w) => w > 0 ? { wpx: w } : {});
}
function copyColWidths(ws, src) {
  if (src && src["!cols"]) ws["!cols"] = src["!cols"];
}
function copyEmptyCellFormats(ws, src) {
  let added = false;
  for (const addr of Object.keys(src)) {
    if (addr.startsWith("!")) continue;
    if (ws[addr]) continue;
    const srcCell = src[addr];
    const srcVal = srcCell.v;
    const isEmpty = srcVal === void 0 || srcVal === null || srcVal === "";
    if (isEmpty && srcCell.z) {
      ws[addr] = { t: "s", v: "", z: srcCell.z, s: { numFmt: srcCell.z } };
      added = true;
    }
  }
  if (added) {
    let maxR = -1, maxC = -1, minR = Infinity, minC = Infinity;
    for (const addr of Object.keys(ws)) {
      if (addr.startsWith("!")) continue;
      const { r, c } = XLSX.utils.decode_cell(addr);
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
      minR = Math.min(minR, r);
      minC = Math.min(minC, c);
    }
    if (maxR >= 0) {
      ws["!ref"] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
    }
  }
}
function generateDashboardExcel(fileBuffer) {
  const wb = XLSX.read(fileBuffer, { type: "array", cellDates: false, cellStyles: true, cellNF: true });
  const sales = parseSalesByMonth(wb);
  const occupancyStats = parseOccupancyStats(wb);
  const dashNames = {
    p1: wb.SheetNames.find((n) => n.startsWith("P.1")) ?? "P.1 Dashboard NEW",
    p2: wb.SheetNames.find((n) => n.startsWith("P2")) ?? "P2. Monthly Income",
    p4: wb.SheetNames.find((n) => n.startsWith("P.4")) ?? "P.4 Forecast & Budget to Actual"
  };
  const p1Build = buildP1Sheet(occupancyStats, sales, dashNames.p2, wb);
  const newWb = XLSX.utils.book_new();
  const dashSet = new Set(Object.values(dashNames));
  const sheetOrder = [
    dashNames.p1,
    dashNames.p2,
    dashNames.p4,
    ...wb.SheetNames.filter((n) => !dashSet.has(n))
  ];
  for (const name of sheetOrder) {
    if (name === dashNames.p1) {
      const ws = aoaToStyledSheet(p1Build.data, p1Build.merges, p1Build.rowHeights);
      const srcP1 = wb.Sheets[dashNames.p1];
      if (srcP1 && srcP1["!cols"]) {
        copyColWidths(ws, srcP1);
      } else {
        setColWidthsPx(ws, [36, 448, 134, 134, 134, 134, 134, 134, 244, 302]);
      }
      if (srcP1) copyEmptyCellFormats(ws, srcP1);
      XLSX.utils.book_append_sheet(newWb, ws, name);
    } else if (name === dashNames.p2) {
      const srcP2 = wb.Sheets[dashNames.p2];
      if (srcP2) {
        const ws = JSON.parse(JSON.stringify(srcP2));
        for (const addr of Object.keys(ws)) {
          if (addr.startsWith("!")) continue;
          const cell = ws[addr];
          if (cell && cell.t === "z") {
            const fmt = cell.z ?? cell.s?.numFmt;
            ws[addr] = fmt ? { t: "s", v: "", z: fmt, s: { numFmt: fmt } } : { t: "s", v: "" };
          }
        }
        XLSX.utils.book_append_sheet(newWb, ws, name);
      }
    } else if (name === dashNames.p4) {
      const srcP4 = wb.Sheets[dashNames.p4];
      if (srcP4) {
        const ws = JSON.parse(JSON.stringify(srcP4));
        for (const addr of Object.keys(ws)) {
          if (addr.startsWith("!")) continue;
          const cell = ws[addr];
          if (cell && cell.t === "z") {
            const fmt = cell.z ?? cell.s?.numFmt;
            ws[addr] = fmt ? { t: "s", v: "", z: fmt, s: { numFmt: fmt } } : { t: "s", v: "" };
          }
        }
        XLSX.utils.book_append_sheet(newWb, ws, name);
      }
    } else if (wb.Sheets[name]) {
      XLSX.utils.book_append_sheet(newWb, wb.Sheets[name], name);
    }
  }
  return Buffer.from(XLSX.write(newWb, { type: "buffer", bookType: "xlsx" }));
}
// Annotate the CommonJS export names for ESM import in node:
0 && (module.exports = {
  generateDashboardExcel
});
