import * as XLSX from "xlsx-js-style";

// ─── Types ────────────────────────────────────────────────────────────────────

type CellValue = string | number | null;

interface StyledCell {
  v?: CellValue;
  f?: string;           // formula (no leading =)
  t?: string;           // type: 'n'|'s'|'b'
  z?: string;           // number format
  s?: CellStyle;
}

interface CellStyle {
  fill?: { fgColor: { rgb: string } };
  font?: { bold?: boolean; color?: { rgb: string }; sz?: number };
  border?: {
    top?: { style: string; color: { rgb: string } };
    bottom?: { style: string; color: { rgb: string } };
    left?: { style: string; color: { rgb: string } };
    right?: { style: string; color: { rgb: string } };
  };
  alignment?: { horizontal?: string; vertical?: string; wrapText?: boolean };
  numFmt?: string;
}

type AoaCell = CellValue | StyledCell;

// ─── Style Constants ─────────────────────────────────────────────────────────

const NAVY = "0F2044";
const GOLD = "C9A84C";
const WHITE = "FFFFFF";
const LIGHT_GRAY = "F3F4F6";
const MED_GRAY = "E5E7EB";
const BORDER_COLOR = "D1D5DB";

const THIN_BORDER = {
  top: { style: "thin", color: { rgb: BORDER_COLOR } },
  bottom: { style: "thin", color: { rgb: BORDER_COLOR } },
  left: { style: "thin", color: { rgb: BORDER_COLOR } },
  right: { style: "thin", color: { rgb: BORDER_COLOR } },
};

function navyHeader(text: string | null): StyledCell {
  return {
    v: text ?? "",
    t: "s",
    s: {
      fill: { fgColor: { rgb: NAVY } },
      font: { bold: true, color: { rgb: WHITE }, sz: 11 },
      border: THIN_BORDER,
      alignment: { horizontal: "center", vertical: "center" },
    },
  };
}

function goldTitle(text: string): StyledCell {
  return {
    v: text,
    t: "s",
    s: {
      fill: { fgColor: { rgb: GOLD } },
      font: { bold: true, color: { rgb: WHITE }, sz: 13 },
      alignment: { horizontal: "left", vertical: "center" },
    },
  };
}

function sectionHeader(text: string): StyledCell {
  return {
    v: text,
    t: "s",
    s: {
      fill: { fgColor: { rgb: NAVY } },
      font: { bold: true, color: { rgb: WHITE }, sz: 11 },
      alignment: { horizontal: "left", vertical: "center" },
    },
  };
}

function boldCell(value: CellValue, numFmt?: string): StyledCell {
  return {
    v: value ?? "",
    t: value == null ? "s" : typeof value === "number" ? "n" : "s",
    s: {
      fill: { fgColor: { rgb: MED_GRAY } },
      font: { bold: true, color: { rgb: "111827" } },
      border: THIN_BORDER,
      ...(numFmt ? { numFmt } : {}),
    },
  };
}

function dataCell(value: CellValue, numFmt?: string, rowIdx?: number): StyledCell {
  const isAlt = rowIdx != null && rowIdx % 2 === 1;
  return {
    v: value ?? "",
    t: value == null ? "s" : typeof value === "number" ? "n" : "s",
    s: {
      fill: { fgColor: { rgb: isAlt ? LIGHT_GRAY : WHITE } },
      font: { color: { rgb: "374151" } },
      border: THIN_BORDER,
      ...(numFmt ? { numFmt } : {}),
    },
  };
}

// cachedVal: optional display value (stored as v so Excel shows it before calculating)
function formulaCell(formula: string, numFmt?: string, rowIdx?: number, cachedVal?: number): StyledCell {
  const isAlt = rowIdx != null && rowIdx % 2 === 1;
  return {
    f: formula,
    v: cachedVal ?? 0,
    t: "n",
    s: {
      fill: { fgColor: { rgb: isAlt ? LIGHT_GRAY : WHITE } },
      font: { color: { rgb: "374151" } },
      border: THIN_BORDER,
      ...(numFmt ? { numFmt } : {}),
    },
  };
}

function boldFormulaCell(formula: string, numFmt?: string, cachedVal?: number): StyledCell {
  return {
    f: formula,
    v: cachedVal ?? 0,
    t: "n",
    s: {
      fill: { fgColor: { rgb: MED_GRAY } },
      font: { bold: true, color: { rgb: "111827" } },
      border: THIN_BORDER,
      ...(numFmt ? { numFmt } : {}),
    },
  };
}

function labelCell(text: string, rowIdx?: number): StyledCell {
  const isAlt = rowIdx != null && rowIdx % 2 === 1;
  return {
    v: text,
    t: "s",
    s: {
      fill: { fgColor: { rgb: isAlt ? LIGHT_GRAY : WHITE } },
      font: { bold: false, color: { rgb: "374151" } },
      border: THIN_BORDER,
      alignment: { horizontal: "left" },
    },
  };
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function str(v: unknown): string {
  return v == null ? "" : String(v).trim();
}

function num(v: unknown): number | null {
  if (v == null || v === "") return null;
  const s = String(v).replace(/[$,%\s]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function isExcelSerial(v: unknown): v is number {
  return typeof v === "number" && v > 40000 && v < 60000;
}

function serialToMonthCode(serial: number): string {
  const d = XLSX.SSF.parse_date_code(serial);
  const months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const yr = String(d.y).slice(-2);
  return `${months[d.m - 1]}-${yr}`;
}

/** Convert 0-based column index to Excel column letter(s): 0→A, 25→Z, 26→AA */
function colLetter(idx: number): string {
  let result = "";
  let n = idx;
  do {
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return result;
}

/** 1-indexed Excel cell address */
function cellAddr(col: number, row: number): string {
  return `${colLetter(col)}${row}`;
}

/** Quote a sheet name for use in a formula reference */
function sheetRef(sheetName: string, addr: string): string {
  return `'${sheetName}'!${addr}`;
}

// ─── Source Parsers ───────────────────────────────────────────────────────────

interface SalesByMonthData {
  sheetName: string;
  rows: (CellValue[])[];
  headerRowIdx: number;
  monthCodes: string[];
  monthColIndices: number[];   // 0-based source column indices for each month
  currentMonthIdx: number;     // index into monthCodes / monthColIndices
  lookupByLabel: Map<string, (number | null)[]>;
  lookupRowIdx: Map<string, number>;  // label (lowercase) → 1-indexed row in source sheet
}

function parseSalesByMonth(wb: XLSX.WorkBook): SalesByMonthData {
  const sheetName = wb.SheetNames.find((n) => /sales.?by.?month/i.test(n));
  if (!sheetName) throw new Error("Cannot find 'Sales By Month' sheet in workbook");

  const ws = wb.Sheets[sheetName];
  const rows: (CellValue[])[] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true }) as (CellValue[])[];

  let headerRowIdx = -1;
  const monthColIndices: number[] = [];
  const monthSerials: number[] = [];

  for (let i = 0; i < rows.length; i++) {
    const serials = rows[i].filter(isExcelSerial);
    if (serials.length >= 3) {
      headerRowIdx = i;
      rows[i].forEach((cell, j) => {
        if (isExcelSerial(cell)) {
          monthSerials.push(cell as number);
          monthColIndices.push(j);
        }
      });
      break;
    }
  }

  if (headerRowIdx < 0) throw new Error("Cannot find monthly header row in Sales By Month sheet");

  const monthCodes = monthSerials.map(serialToMonthCode);
  const lookupByLabel = new Map<string, (number | null)[]>();
  const lookupRowIdx = new Map<string, number>();

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const label = str(row[0]);
    if (!label) continue;
    const values = monthColIndices.map((j) => num(row[j]));
    const key = label.toLowerCase().trim();
    lookupByLabel.set(key, values);
    lookupRowIdx.set(key, i + 1); // 1-indexed
  }

  let currentMonthIdx = 0;
  const incomeLine =
    lookupByLabel.get("total income") ??
    lookupByLabel.get("total revenue") ??
    lookupByLabel.get("flex income");
  if (incomeLine) {
    const nonZero = incomeLine.filter((v): v is number => v != null && v > 0);
    const peak = nonZero.length > 0 ? Math.max(...nonZero) : 0;
    const threshold = peak * 0.15;
    for (let i = incomeLine.length - 1; i >= 0; i--) {
      if ((incomeLine[i] ?? 0) >= threshold) {
        currentMonthIdx = i;
        break;
      }
    }
  }

  return { sheetName, rows, headerRowIdx, monthCodes, monthColIndices, currentMonthIdx, lookupByLabel, lookupRowIdx };
}

interface OccupancyStats {
  sheetName: string;
  location: string;
  statusCol: number;  // 0-based
  typeCol: number;
  sqFtCol: number;
  headerRowIdx: number;
  dedicatedDesk: { total: number; occupied: number; totalSqFt: number; occupiedSqFt: number };
  privateOffice: { total: number; occupied: number; totalSqFt: number; occupiedSqFt: number };
}

function parseOccupancyStats(wb: XLSX.WorkBook): OccupancyStats {
  const sheetName = wb.SheetNames.find((n) => n.startsWith("P.3"));
  if (!sheetName) throw new Error("Cannot find P.3 Occupancy sheet");

  const ws = wb.Sheets[sheetName];
  const rows: (CellValue[])[] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false }) as (CellValue[])[];

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
      if (isOccupied) { po.occupied++; po.occupiedSqFt += sqFt; }
    } else if (spaceType.includes("dedicated")) {
      dd.total++;
      dd.totalSqFt += sqFt;
      if (isOccupied) { dd.occupied++; dd.occupiedSqFt += sqFt; }
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
    privateOffice: po,
  };
}

// ─── P.1 Dashboard Sheet ──────────────────────────────────────────────────────

function buildP1Sheet(
  occ: OccupancyStats,
  sales: SalesByMonthData,
  p2SheetName: string,
  wb: XLSX.WorkBook,
): AoaCell[][] {
  const { currentMonthIdx, monthCodes, lookupByLabel, lookupRowIdx } = sales;
  const reportMonthCode = monthCodes[currentMonthIdx];

  // P.3 column letters (detected dynamically from header)
  const statusLetter = colLetter(occ.statusCol);   // e.g. "D"
  const typeLetter   = colLetter(occ.typeCol);      // e.g. "E"
  const sqftLetter   = colLetter(occ.sqFtCol);      // e.g. "G"
  const p3 = occ.sheetName;

  // P.2 month column: source month at col j maps to P.2 col j+1 (labels in col B)
  const currentMonthP2ColLetter = colLetter(sales.monthColIndices[currentMonthIdx] + 1);

  // Revenue formula: references P.2 "Total with Relief" row 16
  const revFormula = sheetRef(p2SheetName, `${currentMonthP2ColLetter}16`);

  // Computed fallbacks
  const dd = occ.dedicatedDesk;
  const po = occ.privateOffice;

  // ── KPI table (rows 10, 11, 12 in output sheet) ───────────────────────────
  // Layout: A(blank) | B(label) | C(#spaces) | D(occupied) | E(%spaces) | F(sqft) | G(occ sqft) | H(%sqft)
  // Criterion in formulas is the cell reference B10 / B11 (the workspace name value)

  function makeKpiRow(label: string, sheetRow: number, isTotal: boolean): AoaCell[] {
    const bRef = `B${sheetRow}`;  // cell reference to workspace name

    let cTotal: AoaCell, dOcc: AoaCell, fSqFt: AoaCell, gOccSqFt: AoaCell;

    if (isTotal) {
      const above1 = sheetRow - 2;
      const above2 = sheetRow - 1;
      cTotal   = boldFormulaCell(`C${above1}+C${above2}`, "#,##0", dd.total + po.total);
      dOcc     = boldFormulaCell(`D${above1}+D${above2}`, "#,##0", dd.occupied + po.occupied);
      fSqFt    = boldFormulaCell(`F${above1}+F${above2}`, "#,##0", dd.totalSqFt + po.totalSqFt);
      gOccSqFt = boldFormulaCell(`G${above1}+G${above2}`, "#,##0", dd.occupiedSqFt + po.occupiedSqFt);
    } else {
      // COUNTIF/SUMIF using the workspace name cell (B10 or B11) as criterion — exact match
      cTotal   = formulaCell(`COUNTIF(${sheetRef(p3, `${typeLetter}:${typeLetter}`)},${bRef})`, "#,##0", undefined, label === "Dedicated Desk" ? dd.total : po.total);
      dOcc     = formulaCell(`COUNTIFS(${sheetRef(p3, `${typeLetter}:${typeLetter}`)},${bRef},${sheetRef(p3, `${statusLetter}:${statusLetter}`)},"Occupied")`, "#,##0", undefined, label === "Dedicated Desk" ? dd.occupied : po.occupied);
      fSqFt    = formulaCell(`SUMIF(${sheetRef(p3, `${typeLetter}:${typeLetter}`)},${bRef},${sheetRef(p3, `${sqftLetter}:${sqftLetter}`)})`, "#,##0", undefined, label === "Dedicated Desk" ? dd.totalSqFt : po.totalSqFt);
      gOccSqFt = formulaCell(`SUMIFS(${sheetRef(p3, `${sqftLetter}:${sqftLetter}`)},${sheetRef(p3, `${typeLetter}:${typeLetter}`)},${bRef},${sheetRef(p3, `${statusLetter}:${statusLetter}`)},"Occupied")`, "#,##0", undefined, label === "Dedicated Desk" ? dd.occupiedSqFt : po.occupiedSqFt);
    }

    const eSpacePct = isTotal
      ? boldFormulaCell(`IF(C${sheetRow}>0,D${sheetRow}/C${sheetRow},"")`, "0.0%")
      : formulaCell(`IF(C${sheetRow}>0,D${sheetRow}/C${sheetRow},"")`, "0.0%");
    const hSqFtPct  = isTotal
      ? boldFormulaCell(`IF(F${sheetRow}>0,G${sheetRow}/F${sheetRow},"")`, "0.0%")
      : formulaCell(`IF(F${sheetRow}>0,G${sheetRow}/F${sheetRow},"")`, "0.0%");

    const labelCell_ = isTotal
      ? boldCell(label)
      : { v: label, t: "s", s: labelCell(label).s };

    return [null, labelCell_, cTotal, dOcc, eSpacePct, fSqFt, gOccSqFt, hSqFtPct];
  }

  // ── Parse Monthly Metrics for 2024 per-month counts ─────────────────────
  // Monthly Metrics has count-based data (# contracts, flex, mail, mtg hours) per month
  interface MonthlyMetricsRow { poCount: number; ddCount: number; flexCount: number; mailCount: number; mtgHours: number; totalContracts: number; }
  const mmData = new Map<string, MonthlyMetricsRow>();

  const mmSheetName = wb.SheetNames.find((n) => n === "Monthly Metrics" || /monthly.?metrics/i.test(n));
  if (mmSheetName) {
    const mmWs = wb.Sheets[mmSheetName];
    const mmRows: (CellValue[])[] = XLSX.utils.sheet_to_json(mmWs, { header: 1, defval: null, raw: true }) as (CellValue[])[];
    const mmHeader = mmRows[0] ?? [];
    const mmLookup = new Map<string, CellValue[]>();
    for (let r = 1; r < mmRows.length; r++) {
      const lbl = str(mmRows[r][0]).trim().toLowerCase();
      if (lbl) mmLookup.set(lbl, mmRows[r] as CellValue[]);
    }
    const mailKeysMM = [
      "basic mailbox package", "club 25n mailbox package", "digital mailbox package",
      "virtual office package", "personal mail assistant", "virtual office plus package",
      "club 25n", "virtual mail service",
    ];
    for (let c = 1; c < mmHeader.length; c++) {
      const raw = mmHeader[c];
      const monthCode = isExcelSerial(raw) ? serialToMonthCode(raw as number) : str(raw);
      if (!monthCode) continue;
      const n = (key: string) => Math.round(num((mmLookup.get(key) ?? [])[c]) ?? 0);
      const mailCount = mailKeysMM.reduce((sum, k) => sum + n(k), 0);
      // Total contracts = sum of all rows that represent individual client/service counts
      const contractKeys = [
        "private office", "dedicated plus space", "dedicated space",
        "flex income", "basic mailbox package", "club 25n mailbox package",
        "digital mailbox package", "virtual office package", "personal mail assistant",
        "virtual office plus package", "it services", "phone",
        "club 25n", "virtual mail service",
      ];
      const totalContracts = contractKeys.reduce((sum, k) => sum + n(k), 0);
      mmData.set(monthCode, {
        poCount: n("private office"),
        ddCount: n("dedicated plus space") || n("dedicated space"),
        flexCount: n("flex income"),
        mailCount,
        mtgHours: n("recurring & promo meeting room hours"),
        totalContracts,
      });
    }
  }

  // ── Parse historical snapshot rows from original P.1 ─────────────────────
  // Historical rows (pre-2024) are copied verbatim from original P.1 snapshot
  const origP1SheetName = wb.SheetNames.find((n) => n.startsWith("P.1"));
  const historicalSnapRows: AoaCell[][] = [];

  if (origP1SheetName) {
    const origWs = wb.Sheets[origP1SheetName];
    const origRows: (CellValue[])[] = XLSX.utils.sheet_to_json(origWs, { header: 1, defval: null, raw: false }) as (CellValue[])[];
    let snapStart = -1;
    for (let i = 0; i < origRows.length; i++) {
      if (/12 week|snapshot/i.test(str(origRows[i][0]))) { snapStart = i; break; }
    }
    if (snapStart >= 0) {
      // Skip section header + column header row
      for (let i = snapStart + 2; i < origRows.length; i++) {
        const monthCode = str(origRows[i][0]);
        if (!monthCode || /average/i.test(monthCode)) break;
        // Only include rows NOT in our current year (already in mmData / sales.monthCodes)
        if (!mmData.has(monthCode) && !sales.monthCodes.includes(monthCode)) {
          const r = origRows[i];
          const pOccSpaces = str(r[1]).replace(/%/g, "").trim();
          const pOccSqFt   = str(r[2]).replace(/%/g, "").trim();
          historicalSnapRows.push([
            null,
            dataCell(monthCode),
            dataCell(pOccSpaces ? parseFloat(pOccSpaces) / 100 : null, "0%"),
            dataCell(pOccSqFt   ? parseFloat(pOccSqFt)   / 100 : null, "0%"),
            dataCell(num(r[3]) ?? null, "#,##0"),
            dataCell(num(r[4]) ?? null, "#,##0"),
            dataCell(num(r[5]) ?? null, "#,##0"),
            dataCell(r[6] != null && str(r[6]).trim() !== "" && str(r[6]).trim() !== "-" ? num(r[6]) : null, "#,##0"),
            dataCell(num(r[7]) ?? null, "$#,##0"),
          ]);
        }
      }
    }
  }

  // ── Build 2024 snapshot rows from Monthly Metrics ─────────────────────────
  const SNAPSHOT_FIRST_ROW = 19 + historicalSnapRows.length;
  const snapshotRows: AoaCell[][] = [];
  const totalSpaces = dd.total + po.total;
  const totalSqFt   = dd.totalSqFt + po.totalSqFt;
  const mtgRevVals  = sales.lookupByLabel.get("total meeting space rental") ?? [];

  // Only include months up to currentMonthIdx (skip future months with no MM data)
  for (let i = 0; i <= currentMonthIdx; i++) {
    const monthCode = sales.monthCodes[i];
    const isCurrent = i === currentMonthIdx;
    const rowIdx = historicalSnapRows.length + i;
    const outputRow = 19 + rowIdx;

    const mm = mmData.get(monthCode);
    const poCount  = mm?.poCount  ?? 0;
    const ddCount  = mm?.ddCount  ?? 0;
    const contracts = mm?.totalContracts ?? (poCount + ddCount);
    const occPct = totalSpaces > 0 ? (poCount + ddCount) / totalSpaces : 0;
    const poOccSqFt = po.total > 0 ? (poCount / po.total) * po.totalSqFt : 0;
    const ddOccSqFt = dd.total > 0 ? (ddCount / dd.total) * dd.totalSqFt : 0;
    const sqFtPct = totalSqFt > 0 ? (poOccSqFt + ddOccSqFt) / totalSqFt : 0;
    const flexCount = mm?.flexCount  ?? 0;
    const mailCount = mm?.mailCount  ?? 0;
    const mtgHours  = mm?.mtgHours   ?? 0;
    const mtgRev    = (mtgRevVals[i] ?? 0) as number;

    let bCell: AoaCell, cCell: AoaCell, dCell: AoaCell;
    if (isCurrent) {
      bCell = { f: "B2",  v: monthCode, t: "s", s: dataCell(monthCode, undefined, rowIdx).s };
      cCell = { f: "E12", v: occPct,    t: "n", s: dataCell(occPct, "0%", rowIdx).s, z: "0%" };
      dCell = { f: "H12", v: sqFtPct,   t: "n", s: dataCell(sqFtPct, "0%", rowIdx).s, z: "0%" };
    } else {
      bCell = dataCell(monthCode, undefined, rowIdx);
      cCell = dataCell(occPct,  "0%", rowIdx);
      dCell = dataCell(sqFtPct, "0%", rowIdx);
    }

    snapshotRows.push([
      null,
      bCell,
      cCell,
      dCell,
      dataCell(contracts, "#,##0", rowIdx),
      dataCell(flexCount, "#,##0", rowIdx),
      dataCell(mailCount, "#,##0", rowIdx),
      dataCell(mtgHours || null, "#,##0", rowIdx),
      isCurrent
        ? { f: sheetRef(p2SheetName, `${currentMonthP2ColLetter}16`), v: mtgRev, t: "n", s: dataCell(mtgRev, "$#,##0", rowIdx).s, z: "$#,##0" }
        : dataCell(mtgRev, "$#,##0", rowIdx),
    ]);
  }

  // Averages row — spans all snapshot rows (historical + 2024)
  const allSnapshotCount = historicalSnapRows.length + snapshotRows.length;
  const snapStart2 = 19;
  const snapEnd2 = 19 + allSnapshotCount - 1;
  const avgRow: AoaCell[] = [null, boldCell("Average")];
  const avgCols = ["C","D","E","F","G","H","I"];
  for (const col of avgCols) {
    avgRow.push(boldFormulaCell(`AVERAGE(${col}${snapStart2}:${col}${snapEnd2})`));
  }

  return [
    // Row 1: Title
    [goldTitle(`25N COWORKING ${occ.location.toUpperCase()} DASHBOARD`)],
    // Row 2: Report month
    [null, { v: reportMonthCode, t: "s", s: { font: { bold: true, sz: 11, color: { rgb: NAVY } } } }],
    // Row 3: blank
    [],
    // Row 4: Revenue — label in col B, value formula in col H
    [
      null,
      { v: "Sales (For the month)", t: "s", s: { font: { bold: true, color: { rgb: NAVY } } } },
      null, null, null, null, null,
      { f: revFormula, v: 0, t: "n", s: { font: { bold: true, sz: 14, color: { rgb: NAVY } }, alignment: { horizontal: "left" } }, z: "$#,##0" },
    ],
    // Row 5: blank
    [],
    // Row 6: Section header
    [sectionHeader("KEY PERFORMANCE INDICATORS - 25N COWORKING")],
    // Row 7: blank
    [],
    // Row 8: "CURRENT OCCUPANCY" sub-header (matches original)
    [sectionHeader("CURRENT OCCUPANCY")],
    // Row 9: KPI column headers
    [
      null,
      navyHeader("Workspace"),
      navyHeader("# of Spaces"),
      navyHeader("Occupied"),
      navyHeader("% of Spaces"),
      navyHeader("Total Square Feet"),
      navyHeader("Occupied"),
      navyHeader("% of Square Feet"),
    ],
    // Row 10: Dedicated Desk
    makeKpiRow("Dedicated Desk", 10, false),
    // Row 11: Private Office
    makeKpiRow("Private Office", 11, false),
    // Row 12: Total Spaces
    makeKpiRow("Total Spaces", 12, true),
    // Rows 13-15: blank
    [], [], [],
    // Row 16: Snapshot section header
    [sectionHeader("12 Week Snapshot")],
    // Row 17: blank (to match original "12 Week Snapshot" → blank → headers)
    [],
    // Row 18: Snapshot column headers
    [
      null,
      navyHeader("Month"),
      navyHeader("Occupied % of Total Spaces"),
      navyHeader("Occupied % of Total Square Feet"),
      navyHeader("# of Contracts"),
      navyHeader("# of Flex Members"),
      navyHeader("# of Mail Members"),
      navyHeader("Recurring & Promo Meeting Hours Sold"),
      navyHeader("Meeting Hours Sold $"),
    ],
    // Rows 19+: historical snapshot data then 2024 months
    ...historicalSnapRows,
    ...snapshotRows,
    // Last: averages
    avgRow,
  ];
}

// ─── P.2 Monthly Income Sheet — Condensed 16-row Summary ─────────────────────

function buildP2Sheet(sales: SalesByMonthData): AoaCell[][] {
  const { sheetName: srcName, monthColIndices, monthCodes, lookupRowIdx } = sales;
  const numMonths = monthColIndices.length;

  // Month column letters in OUTPUT sheet:
  // Source month at 0-based col j → P.2 output col j+1
  // (labels in col B, months start at col C)
  const firstMonthOutputCol = 2;  // col C (0-based index 2)
  const lastMonthOutputColLetter = colLetter(firstMonthOutputCol + numMonths - 1);
  const annualOutputColLetter = colLetter(firstMonthOutputCol + numMonths);

  // Month column name lookup
  const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];

  // Helper: formula referencing source sheet at a specific row + month column
  function srcFormula(sourceRow1: number, monthIdx: number): string {
    return sheetRef(srcName, cellAddr(monthColIndices[monthIdx], sourceRow1));
  }

  // Build a single-source category row (one source row → per-month formulas)
  function categoryRow(
    outputRow1: number,
    label: string,
    sourceRow1: number | null,
    isBold = false,
  ): AoaCell[] {
    const cols: AoaCell[] = [
      null,
      isBold ? boldCell(label) : labelCell(label),
    ];
    for (let i = 0; i < numMonths; i++) {
      if (sourceRow1 != null) {
        cols.push(formulaCell(srcFormula(sourceRow1, i), "$#,##0"));
      } else {
        cols.push(dataCell(0, "$#,##0"));
      }
    }
    // Annual SUM
    const firstCol = colLetter(firstMonthOutputCol);
    const lastCol = lastMonthOutputColLetter;
    cols.push(
      isBold
        ? boldFormulaCell(`SUM(${firstCol}${outputRow1}:${lastCol}${outputRow1})`, "$#,##0")
        : formulaCell(`SUM(${firstCol}${outputRow1}:${lastCol}${outputRow1})`, "$#,##0")
    );
    return cols;
  }

  // "Recurring Revenue" = TOTAL INCOME minus the other 9 separately-listed categories
  // This guarantees: SUM(rows 3-12) = TOTAL INCOME (no double-counting)
  function recurringRow(outputRow1: number): AoaCell[] {
    const totalIncomeRow = lookupRowIdx.get("total income") ?? null;
    // The 9 categories that are listed separately in P.2 rows 4-12
    const subtractKeys = [
      "total meeting space rental",  // Meeting Rooms
      "total events revenue",        // Special Events
      "registration set up",         // Registrations
      "daily space",                 // Daily Space
      "printing",                    // Printing
      "furniture enhancements",      // Furniture Enhancements
      "tenant improvement",          // Tenant Improvement
      "sponsorship income",          // Sponsorship
      "miscellaneous income",        // Other Miscellaneous (specific line item)
    ];
    const subtractRows = subtractKeys
      .map((k) => lookupRowIdx.get(k))
      .filter((r): r is number => r != null);

    const cols: AoaCell[] = [null, labelCell("Recurring Revenue")];
    for (let i = 0; i < numMonths; i++) {
      const srcCol = monthColIndices[i];
      if (totalIncomeRow) {
        const totalRef = sheetRef(srcName, cellAddr(srcCol, totalIncomeRow));
        const subParts = subtractRows.map((r) => sheetRef(srcName, cellAddr(srcCol, r)));
        const formula = subParts.length > 0
          ? `${totalRef}-${subParts.join("-")}`
          : totalRef;
        cols.push(formulaCell(formula, "$#,##0"));
      } else {
        // Fallback when TOTAL INCOME row not found
        const parts = ["total membership income", "total workspace rental", "total package revenue"]
          .map((k) => lookupRowIdx.get(k))
          .filter((r): r is number => r != null)
          .map((r) => sheetRef(srcName, cellAddr(srcCol, r)));
        cols.push(parts.length > 0 ? formulaCell(parts.join("+"), "$#,##0") : dataCell(0, "$#,##0"));
      }
    }
    const firstCol = colLetter(firstMonthOutputCol);
    cols.push(formulaCell(`SUM(${firstCol}${outputRow1}:${lastMonthOutputColLetter}${outputRow1})`, "$#,##0"));
    return cols;
  }

  // SUM formula row (Total / Total with Relief)
  function sumRow(outputRow1: number, label: string, sumFrom: number, sumTo: number): AoaCell[] {
    const cols: AoaCell[] = [null, boldCell(label)];
    for (let i = 0; i < numMonths; i++) {
      const col = colLetter(firstMonthOutputCol + i);
      cols.push(boldFormulaCell(`SUM(${col}${sumFrom}:${col}${sumTo})`, "$#,##0"));
    }
    const annualCol = annualOutputColLetter;
    cols.push(boldFormulaCell(`SUM(${colLetter(firstMonthOutputCol)}${outputRow1}:${lastMonthOutputColLetter}${outputRow1})`, "$#,##0"));
    return cols;
  }

  // Category definitions in order (matches original P.2)
  const categories: Array<{ label: string; key?: string; special?: string }> = [
    { label: "Recurring Revenue", special: "recurring" },
    { label: "Meeting Rooms",          key: "total meeting space rental" },
    { label: "Special Events",         key: "total events revenue" },
    { label: "Registrations",          key: "registration set up" },
    { label: "Daily Space",            key: "daily space" },
    { label: "Printing",               key: "printing" },
    { label: "Furniture Enhancements", key: "furniture enhancements" },
    { label: "Tenant Improvement",     key: "tenant improvement" },
    { label: "Sponsorship",            key: "sponsorship income" },
    { label: "Other Miscellaneous",    key: "miscellaneous income" },
  ];

  // Discount rows: look for "discounts given", "discounts", "discounts - other"
  const discountCandidates = ["discounts given", "discounts", "discounts - other"];
  const discountSourceRows = discountCandidates
    .map((k) => lookupRowIdx.get(k))
    .filter((r): r is number => r != null);

  // Build output AOA
  const output: AoaCell[][] = [];

  // Row 1: "MONTHLY TOTAL" header in col A
  output.push([{ v: "MONTHLY TOTAL", t: "s", s: sectionHeader("MONTHLY TOTAL").s }]);

  // Row 2: "SALES COMPARATIVE BY MONTH" + month name headers + "Annual"
  const headerRow: AoaCell[] = [null, navyHeader("SALES COMPARATIVE BY MONTH")];
  for (let i = 0; i < numMonths; i++) {
    // Parse month name from code (e.g. "Aug-24" → "August")
    const code = monthCodes[i]; // "Aug-24"
    const abbr = code.split("-")[0];
    const monthName = MONTH_NAMES.find((m) => m.startsWith(abbr)) ?? code;
    headerRow.push(navyHeader(monthName));
  }
  headerRow.push(navyHeader("Annual"));
  output.push(headerRow);

  // Rows 3-12: 10 summary categories
  for (let ci = 0; ci < categories.length; ci++) {
    const cat = categories[ci];
    const rowNum = 3 + ci;  // 1-indexed output row
    if (cat.special === "recurring") {
      output.push(recurringRow(rowNum));
    } else {
      const sourceRow = cat.key ? (lookupRowIdx.get(cat.key) ?? null) : null;
      output.push(categoryRow(rowNum, cat.label, sourceRow));
    }
  }

  // Row 13: Total = SUM(rows 3-12)
  output.push(sumRow(13, "Total", 3, 12));

  // Row 14: Discounts Given (first discount row or zero)
  output.push(categoryRow(14, "Discounts Given", discountSourceRows[0] ?? null));

  // Row 15: Discounts Given (second discount row or zero)
  output.push(categoryRow(15, "Discounts Given", discountSourceRows[1] ?? null));

  // Row 16: Total with Relief = SUM(rows 13-15)
  output.push(sumRow(16, "Total with Relief", 13, 15));

  return output;
}

// ─── P.4 Budget vs Actual Sheet ───────────────────────────────────────────────

function buildP4Sheet(wb: XLSX.WorkBook, sales: SalesByMonthData, p2SheetName: string): AoaCell[][] {
  const srcSheetName = wb.SheetNames.find((n) => n.startsWith("P.4"));
  if (!srcSheetName) throw new Error("Cannot find P.4 sheet");

  const ws = wb.Sheets[srcSheetName];
  const rows: (CellValue[])[] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true }) as (CellValue[])[];

  const { lookupByLabel, lookupRowIdx, currentMonthIdx, monthCodes, monthColIndices } = sales;
  const reportMonthCode = monthCodes[currentMonthIdx];

  // P.2 column letter for current month (source col j → P.2 col j+1)
  const currentMonthP2ColLetter = colLetter(monthColIndices[currentMonthIdx] + 1);

  let headerRowIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    const rowStr = rows[i].map((c) => str(c)).join(" ").toLowerCase();
    if (rowStr.includes("ptd actual") || (rowStr.includes("actual") && rowStr.includes("budget"))) {
      headerRowIdx = i;
      break;
    }
  }

  const output: AoaCell[][] = [];

  // Pre-header metadata rows
  for (let i = 0; i < Math.max(headerRowIdx, 0); i++) {
    output.push(
      rows[i].map((cell): AoaCell => {
        if (isExcelSerial(cell)) return reportMonthCode;
        const n = num(cell);
        return n !== null ? n : str(cell) || null;
      })
    );
  }

  // Column headers
  if (headerRowIdx >= 0) {
    output.push([
      navyHeader(""),
      navyHeader("PTD Actual"),
      navyHeader("PTD Budget Forecast"),
      navyHeader("Budget Variance $"),
      navyHeader("Budget Variance %"),
    ]);
  }

  const dataStartOutputRow = output.length + 1;
  const dataStart = headerRowIdx >= 0 ? headerRowIdx + 1 : 0;
  let dataIdx = 0;

  for (let i = dataStart; i < rows.length; i++) {
    const row = rows[i];
    const label = str(row[0]);
    if (!label) {
      output.push([]);
      dataIdx++;
      continue;
    }

    const isTotalRow = /total|subtotal/i.test(label);
    const labelLower = label.toLowerCase().trim();
    const outputRow = dataStartOutputRow + dataIdx;
    const rowIdx = dataIdx;

    // Budget: reference source P.4 column C (budget values are in the pass-through sheet)
    const srcDataRow = i + 1;
    const budgetCell: AoaCell = {
      f: sheetRef(srcSheetName, `C${srcDataRow}`),
      v: num(row[2]) ?? 0,
      t: "n",
      s: (isTotalRow ? boldCell(0, "$#,##0") : dataCell(0, "$#,##0", rowIdx)).s,
      z: "$#,##0",
    };

    // Actual: reference P.2 at current month column
    const salesRow = lookupRowIdx.get(labelLower) ?? null;
    let actualCell: AoaCell;
    if (salesRow != null) {
      const actualVal = (lookupByLabel.get(labelLower) ?? [])[currentMonthIdx] ?? 0;
      actualCell = {
        f: sheetRef(p2SheetName, `${currentMonthP2ColLetter}${salesRow}`),
        v: actualVal,
        t: "n",
        s: (isTotalRow ? boldCell(0, "$#,##0") : dataCell(0, "$#,##0", rowIdx)).s,
        z: "$#,##0",
      };
    } else {
      const monthlyVals = lookupByLabel.get(labelLower);
      const actualVal = monthlyVals != null ? (monthlyVals[currentMonthIdx] ?? null) : null;
      actualCell = isTotalRow ? boldCell(actualVal, "$#,##0") : dataCell(actualVal, "$#,##0", rowIdx);
    }

    // Variance $: =B{row}-C{row}
    const varDollarCell: AoaCell = {
      f: `B${outputRow}-C${outputRow}`,
      v: 0,
      t: "n",
      s: (isTotalRow ? boldCell(0, "$#,##0") : dataCell(0, "$#,##0", rowIdx)).s,
      z: "$#,##0",
    };

    // Variance %: =IF(ABS(C{row})>0,(B{row}-C{row})/ABS(C{row}),"")
    const varPctCell: AoaCell = {
      f: `IF(ABS(C${outputRow})>0,(B${outputRow}-C${outputRow})/ABS(C${outputRow}),"")`,
      v: 0,
      t: "n",
      s: (isTotalRow ? boldCell(0, "0.0%") : dataCell(0, "0.0%", rowIdx)).s,
      z: "0.0%",
    };

    output.push([
      isTotalRow ? boldCell(label) : labelCell(label, rowIdx),
      actualCell,
      budgetCell,
      varDollarCell,
      varPctCell,
    ]);
    dataIdx++;
  }

  return output;
}

// ─── Apply Styles to Sheet ────────────────────────────────────────────────────

function aoaToStyledSheet(aoa: AoaCell[][]): XLSX.WorkSheet {
  const ws: XLSX.WorkSheet = {};
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
        const sc = cell as StyledCell;
        if (sc.f) {
          // CRITICAL: must include v (cached value) for xlsx-js-style to write formula cells
          ws[addr] = { t: sc.t ?? "n", f: sc.f, v: sc.v ?? 0 };
        } else {
          ws[addr] = { t: sc.t ?? "s", v: sc.v ?? "" };
        }
        if (sc.s)  ws[addr].s = sc.s;
        if (sc.z)  ws[addr].z = sc.z;
      }
    }
  }

  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
  return ws;
}

// ─── Column Widths ────────────────────────────────────────────────────────────

function setColWidths(ws: XLSX.WorkSheet, widths: number[]) {
  ws["!cols"] = widths.map((w) => ({ wch: w }));
}

// ─── Main Export ─────────────────────────────────────────────────────────────

export function generateDashboardExcel(fileBuffer: ArrayBuffer): Buffer {
  const wb = XLSX.read(fileBuffer, { type: "array", cellDates: false });

  const sales = parseSalesByMonth(wb);
  const occupancyStats = parseOccupancyStats(wb);

  const dashNames = {
    p1: wb.SheetNames.find((n) => n.startsWith("P.1")) ?? "P.1 Dashboard NEW",
    p2: wb.SheetNames.find((n) => n.startsWith("P2")) ?? "P2. Monthly Income",
    p4: wb.SheetNames.find((n) => n.startsWith("P.4")) ?? "P.4 Forecast & Budget to Actual",
  };

  const p2Data = buildP2Sheet(sales);
  const p1Data = buildP1Sheet(occupancyStats, sales, dashNames.p2, wb);
  const p4Data = buildP4Sheet(wb, sales, dashNames.p2);

  const newWb = XLSX.utils.book_new();

  const dashSet = new Set(Object.values(dashNames));
  const sheetOrder = [
    dashNames.p1,
    dashNames.p2,
    dashNames.p4,
    ...wb.SheetNames.filter((n) => !dashSet.has(n)),
  ];

  // P.2 month columns: one per month + label col (B) + annual col + blank col A
  const p2MonthCols = Array(sales.monthColIndices.length + 2).fill(14);
  p2MonthCols[0] = 4;   // col A narrow
  p2MonthCols[1] = 30;  // col B (labels) wider

  for (const name of sheetOrder) {
    if (name === dashNames.p1) {
      const ws = aoaToStyledSheet(p1Data);
      setColWidths(ws, [4, 30, 14, 14, 12, 16, 14, 12, 16]);
      XLSX.utils.book_append_sheet(newWb, ws, name);
    } else if (name === dashNames.p2) {
      const ws = aoaToStyledSheet(p2Data);
      setColWidths(ws, p2MonthCols);
      XLSX.utils.book_append_sheet(newWb, ws, name);
    } else if (name === dashNames.p4) {
      const ws = aoaToStyledSheet(p4Data);
      setColWidths(ws, [36, 16, 18, 18, 16]);
      XLSX.utils.book_append_sheet(newWb, ws, name);
    } else if (wb.Sheets[name]) {
      XLSX.utils.book_append_sheet(newWb, wb.Sheets[name], name);
    }
  }

  return Buffer.from(XLSX.write(newWb, { type: "buffer", bookType: "xlsx" }));
}
