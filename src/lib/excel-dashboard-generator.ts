import * as XLSX from "xlsx-js-style";

// ─── Types ────────────────────────────────────────────────────────────────────

type CellValue = string | number | null;

// xlsx-js-style cell object — value, formula, style, type
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
const GREEN = "16A34A";
const RED = "DC2626";
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

function formulaCell(formula: string, numFmt?: string, rowIdx?: number, colored?: "pos" | "neg"): StyledCell {
  const isAlt = rowIdx != null && rowIdx % 2 === 1;
  const fontColor = colored === "pos" ? GREEN : colored === "neg" ? RED : "374151";
  return {
    f: formula,
    t: "n",
    s: {
      fill: { fgColor: { rgb: isAlt ? LIGHT_GRAY : WHITE } },
      font: { color: { rgb: fontColor }, ...(colored ? { bold: true } : {}) },
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

function avgCell(formula: string, numFmt?: string): StyledCell {
  return {
    f: formula,
    t: "n",
    s: {
      fill: { fgColor: { rgb: MED_GRAY } },
      font: { bold: true, color: { rgb: "111827" } },
      border: THIN_BORDER,
      ...(numFmt ? { numFmt } : {}),
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

/** 1-indexed Excel row address */
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
  monthColIndices: number[];
  currentMonthIdx: number;
  lookupByLabel: Map<string, (number | null)[]>;
  lookupRowIdx: Map<string, number>;  // label → 1-indexed row in source sheet
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

interface MonthlyMetricsData {
  sheetName: string | null;
  metrics: Map<string, (number | null)[]>;
  monthCodes: string[];
  colIndices: number[];
  labelRowIdx: Map<string, number>; // label → 1-indexed row
}

function parseMonthlyMetrics(wb: XLSX.WorkBook): MonthlyMetricsData {
  const sheetName = wb.SheetNames.find((n) => /monthly.?metrics/i.test(n)) ?? null;
  if (!sheetName) return { sheetName: null, metrics: new Map(), monthCodes: [], colIndices: [], labelRowIdx: new Map() };

  const ws = wb.Sheets[sheetName];
  const rows: (CellValue[])[] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true }) as (CellValue[])[];
  if (rows.length === 0) return { sheetName, metrics: new Map(), monthCodes: [], colIndices: [], labelRowIdx: new Map() };

  const headerRow = rows[0];
  const colIndices: number[] = [];
  const monthSerials: number[] = [];
  headerRow.forEach((cell, j) => {
    if (j > 0 && isExcelSerial(cell)) {
      monthSerials.push(cell as number);
      colIndices.push(j);
    }
  });

  const monthCodes = monthSerials.map(serialToMonthCode);
  const metrics = new Map<string, (number | null)[]>();
  const labelRowIdx = new Map<string, number>();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const label = str(row[0]);
    if (!label) continue;
    const values = colIndices.map((j) => num(row[j]));
    metrics.set(label, values);
    metrics.set(label.toLowerCase().trim(), values);
    labelRowIdx.set(label, i + 1); // 1-indexed
    labelRowIdx.set(label.toLowerCase().trim(), i + 1);
  }

  return { sheetName, metrics, monthCodes, colIndices, labelRowIdx };
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
  let statusCol = 3, spaceTypeCol = 4, sqFtCol = 6, propCol = 0;

  for (let i = 0; i < rows.length; i++) {
    const rowLower = rows[i].map((c) => str(c).toLowerCase());
    if (rowLower.includes("status")) {
      headerRowIdx = i;
      const si = rowLower.findIndex((c) => c === "status");
      const ti = rowLower.findIndex((c) => c.includes("type"));
      const qi = rowLower.findIndex((c) => c.includes("sq"));
      const pi = rowLower.findIndex((c) => c.includes("property"));
      if (si >= 0) statusCol = si;
      if (ti >= 0) spaceTypeCol = ti;
      if (qi >= 0) sqFtCol = qi;
      if (pi >= 0) propCol = pi;
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

    if (!location && propCol >= 0) {
      const propVal = str(row[propCol]);
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
  metrics: MonthlyMetricsData,
  p2SheetName: string,
): AoaCell[][] {
  const { currentMonthIdx, monthCodes, lookupByLabel, lookupRowIdx } = sales;
  const reportMonthCode = monthCodes[currentMonthIdx];
  const currentMonthCol = sales.monthColIndices[currentMonthIdx];

  // Detect current revenue row for formula
  const revenueLabels = ["total with relief", "total revenue", "total income"];
  let revenueRowLabel = "";
  for (const l of revenueLabels) {
    if (lookupRowIdx.has(l)) { revenueRowLabel = l; break; }
  }
  const revenueRow = revenueRowLabel ? lookupRowIdx.get(revenueRowLabel)! : null;
  const revenueFormula = revenueRow
    ? sheetRef(p2SheetName, cellAddr(currentMonthCol, revenueRow))
    : null;

  // P.3 column letters for COUNTIFS/SUMIFS
  const statusLetter = colLetter(occ.statusCol);
  const typeLetter = colLetter(occ.typeCol);
  const sqftLetter = colLetter(occ.sqFtCol);
  const p3 = occ.sheetName;

  // Computed fallbacks for display even when formula is used
  const dd = occ.dedicatedDesk;
  const po = occ.privateOffice;
  const totalSpaces = dd.total + po.total;
  const totalOccupied = dd.occupied + po.occupied;
  const totalSqFt = dd.totalSqFt + po.totalSqFt;
  const totalOccSqFt = dd.occupiedSqFt + po.occupiedSqFt;

  // Monthly Metrics lookups
  function getMetric(key: string): (number | null)[] {
    return metrics.metrics.get(key) ?? metrics.metrics.get(key.toLowerCase()) ?? [];
  }
  function getMetricRow(key: string): number | null {
    return metrics.labelRowIdx.get(key) ?? metrics.labelRowIdx.get(key.toLowerCase()) ?? null;
  }

  const poMetrics = getMetric("Private Office");
  const ddMetrics = getMetric("Dedicated Plus Space");
  const flexMetrics = getMetric("Flex Income");
  const hoursMetrics = getMetric("Recurring & Promo Meeting Room Hours");

  const mailKeys = [
    "Basic Mailbox Package",
    "Club 25N Mailbox Package",
    "Digital Mailbox Package",
    "Virtual Office Package",
    "Personal Mail Assistant",
    "Virtual Office Plus Package",
  ];
  const mailArrays = mailKeys.map((k) => getMetric(k));

  const mtgRevenueByMonth = lookupByLabel.get("total meeting space rental");
  const mtgRevenueLabel = "total meeting space rental";
  const mtgRevenueRow = lookupRowIdx.get(mtgRevenueLabel) ?? null;

  // ── Snapshot rows (formula-based if metrics sheet available) ──────────────
  const snapshotRows: AoaCell[][] = [];
  const snapStartDataRow = 18; // 1-indexed row where first snapshot data row will be in P.1

  for (let i = 0; i < metrics.monthCodes.length; i++) {
    const monthCode = metrics.monthCodes[i];
    const poCount = Math.round(poMetrics[i] ?? 0);
    const ddCount = Math.round(ddMetrics[i] ?? 0);
    const occupiedCount = poCount + ddCount;
    const occupancyPct = totalSpaces > 0 ? occupiedCount / totalSpaces : 0;
    const poOccSqFt = po.total > 0 ? (poCount / po.total) * po.totalSqFt : 0;
    const ddOccSqFt = dd.total > 0 ? (ddCount / dd.total) * dd.totalSqFt : 0;
    const sqFtPct = totalSqFt > 0 ? (poOccSqFt + ddOccSqFt) / totalSqFt : 0;
    const flexCount = Math.round(flexMetrics[i] ?? 0);
    const mailCount = mailArrays.reduce((sum, arr) => sum + Math.round(arr[i] ?? 0), 0);
    const hours = hoursMetrics[i] ?? 0;
    const salesIdx = sales.monthCodes.indexOf(monthCode);
    const meetingRev = salesIdx >= 0 ? (mtgRevenueByMonth?.[salesIdx] ?? 0) : 0;

    const rowIdx = i; // for alternating color

    // Use formulas if we have the metrics sheet name; else fall back to values
    const msn = metrics.sheetName;
    const poRow = getMetricRow("Private Office");
    const ddRow = getMetricRow("Dedicated Plus Space");
    const flexRow = getMetricRow("Flex Income");
    const hoursRow = getMetricRow("Recurring & Promo Meeting Room Hours");

    // For the snapshot % and count cells, we reference the P.3 totals via hardcoded formulas
    // rather than metric sheet, because metric sheet only has count by type not total spaces.
    // The occupancy % = (PO occupied + DD occupied) / (PO total + DD total)
    // We express this using numeric references to the KPI table cells (rows 9,10,11 in P.1):
    // Row 9 = Dedicated Desk row, Row 10 = Private Office row, Row 11 = Total row
    // Col D=occupied(3), Col C=total(2) — but P.1 itself uses formulas for those too.
    // Simplest clean approach: use metric sheet references for counts, P.3 formulas for totals.

    let poCell: AoaCell, ddCell: AoaCell, contCell: AoaCell;
    let occPctCell: AoaCell, sqFtPctCell: AoaCell;
    let flexCell: AoaCell, hoursCell: AoaCell;
    let meetRevCell: AoaCell;
    // Mail: sum of multiple rows — use computed value (too complex for clean formula chain)
    const mailCell = dataCell(mailCount, "#,##0", rowIdx);

    if (msn && poRow && ddRow) {
      const poColLtr = colLetter(metrics.colIndices[i] ?? (i + 1));
      poCell = formulaCell(sheetRef(msn, cellAddr(metrics.colIndices[i] ?? i, poRow)), "#,##0", rowIdx);
      ddCell = formulaCell(sheetRef(msn, cellAddr(metrics.colIndices[i] ?? i, ddRow)), "#,##0", rowIdx);
      // contracts = PO + DD
      const p1Row = snapStartDataRow + i;
      contCell = { f: `B${p1Row}+C${p1Row}`, t: "n", s: dataCell(occupiedCount, "#,##0", rowIdx).s };
      // occupancy % = contracts / total spaces (KPI table row 11 col C)
      occPctCell = { f: `D${p1Row}/11`, t: "n", s: dataCell(occupancyPct, "0.0%", rowIdx).s, z: "0.0%" };
      // sqft % = approx via computed value (complex cross-reference, use value)
      sqFtPctCell = dataCell(sqFtPct, "0.0%", rowIdx);
    } else {
      poCell = dataCell(poCount, "#,##0", rowIdx);
      ddCell = dataCell(ddCount, "#,##0", rowIdx);
      contCell = dataCell(occupiedCount, "#,##0", rowIdx);
      occPctCell = dataCell(occupancyPct, "0.0%", rowIdx);
      sqFtPctCell = dataCell(sqFtPct, "0.0%", rowIdx);
    }

    if (msn && flexRow) {
      flexCell = formulaCell(sheetRef(msn, cellAddr(metrics.colIndices[i] ?? i, flexRow)), "#,##0", rowIdx);
    } else {
      flexCell = dataCell(flexCount, "#,##0", rowIdx);
    }
    if (msn && hoursRow) {
      hoursCell = formulaCell(sheetRef(msn, cellAddr(metrics.colIndices[i] ?? i, hoursRow)), "#,##0.0", rowIdx);
    } else {
      hoursCell = dataCell(hours, "#,##0.0", rowIdx);
    }

    // Meeting revenue: reference P.2 (which already references Sales By Month)
    if (mtgRevenueRow && salesIdx >= 0) {
      // P.2 has same row structure as Sales By Month for this label
      meetRevCell = formulaCell(sheetRef(p2SheetName, cellAddr(currentMonthCol, mtgRevenueRow)), "$#,##0", rowIdx);
    } else {
      meetRevCell = dataCell(meetingRev, "$#,##0", rowIdx);
    }

    snapshotRows.push([
      dataCell(monthCode, undefined, rowIdx),
      occPctCell,
      sqFtPctCell,
      contCell,
      flexCell,
      mailCell,
      hoursCell,
      meetRevCell,
    ]);
  }

  // Averages row — use AVERAGE formulas over snapshot range
  const snapStart = snapStartDataRow;
  const snapEnd = snapStart + snapshotRows.length - 1;
  const avgRow: AoaCell[] = [boldCell("Average")];
  const avgCols = ["B","C","D","E","F","G","H","I"];
  for (const c of avgCols) {
    if (snapshotRows.length > 0) {
      avgRow.push(avgCell(`AVERAGE(${c}${snapStart}:${c}${snapEnd})`));
    } else {
      avgRow.push(boldCell(null));
    }
  }

  // ── KPI table rows (P.3 COUNTIFS formulas) ────────────────────────────────
  // Row positions in output sheet (1-indexed): 9=DD, 10=PO, 11=Total
  // Cols: B=workspace label, C=# spaces, D=occupied, E=% spaces, F=total sqft, G=occ sqft, H=% sqft

  function makeKpiRow(label: string, typePattern: string, isTotal: boolean, sheetRow: number): AoaCell[] {
    const p3Range = `${typeLetter}:${typeLetter}`;
    const statusRange = `${statusLetter}:${statusLetter}`;
    const sqftRange = `${sqftLetter}:${sqftLetter}`;

    let totalSpacesF: AoaCell, occupiedF: AoaCell, totalSqFtF: AoaCell, occSqFtF: AoaCell;
    if (isTotal) {
      // Total row sums the two above it
      totalSpacesF = { f: `C${sheetRow-2}+C${sheetRow-1}`, t: "n", s: boldCell(totalSpaces, "#,##0").s, z: "#,##0" };
      occupiedF    = { f: `D${sheetRow-2}+D${sheetRow-1}`, t: "n", s: boldCell(totalOccupied, "#,##0").s, z: "#,##0" };
      totalSqFtF   = { f: `F${sheetRow-2}+F${sheetRow-1}`, t: "n", s: boldCell(totalSqFt, "#,##0").s, z: "#,##0" };
      occSqFtF     = { f: `G${sheetRow-2}+G${sheetRow-1}`, t: "n", s: boldCell(totalOccSqFt, "#,##0").s, z: "#,##0" };
    } else {
      totalSpacesF = { f: `COUNTIF(${sheetRef(p3, p3Range)},"*${typePattern}*")`, t: "n", s: dataCell(0, "#,##0").s, z: "#,##0" };
      occupiedF    = { f: `COUNTIFS(${sheetRef(p3, p3Range)},"*${typePattern}*",${sheetRef(p3, statusRange)},"Occup*")`, t: "n", s: dataCell(0, "#,##0").s, z: "#,##0" };
      totalSqFtF   = { f: `SUMIF(${sheetRef(p3, p3Range)},"*${typePattern}*",${sheetRef(p3, sqftRange)})`, t: "n", s: dataCell(0, "#,##0").s, z: "#,##0" };
      occSqFtF     = { f: `SUMIFS(${sheetRef(p3, sqftRange)},${sheetRef(p3, p3Range)},"*${typePattern}*",${sheetRef(p3, statusRange)},"Occup*")`, t: "n", s: dataCell(0, "#,##0").s, z: "#,##0" };
    }
    const occPctF = { f: `IF(C${sheetRow}>0,D${sheetRow}/C${sheetRow},"")`, t: "n", s: (isTotal ? boldCell(0) : dataCell(0)).s, z: "0.0%" };
    const sqPctF  = { f: `IF(F${sheetRow}>0,G${sheetRow}/F${sheetRow},"")`, t: "n", s: (isTotal ? boldCell(0) : dataCell(0)).s, z: "0.0%" };

    const rowStyle = isTotal ? boldCell(label) : labelCell(label);
    return [rowStyle, navyHeader(null), totalSpacesF, occupiedF, occPctF, totalSqFtF, occSqFtF, sqPctF];
  }

  // Row 9=DD (sheetRow=9), Row 10=PO (sheetRow=10), Row 11=Total (sheetRow=11)
  const ddRow = makeKpiRow("Dedicated Desk", "Dedicated", false, 9);
  const poRow2 = makeKpiRow("Private Office", "Private Office", false, 10);
  const totalRow = makeKpiRow("Total Spaces", "", true, 11);

  // Revenue cell
  const revCell: AoaCell = revenueFormula
    ? { f: revenueFormula, t: "n", s: { font: { bold: true, sz: 14, color: { rgb: NAVY } }, alignment: { horizontal: "left" } }, z: "$#,##0" }
    : { v: 0, t: "n", s: { font: { bold: true, sz: 14 } }, z: "$#,##0" };

  return [
    // Row 1: Title
    [goldTitle(`25N COWORKING ${occ.location.toUpperCase()} DASHBOARD`)],
    // Row 2: Report month
    [{ v: reportMonthCode, t: "s", s: { font: { bold: true, sz: 11, color: { rgb: NAVY } } } }],
    // Row 3: blank
    [],
    // Row 4: Revenue
    [revCell],
    // Row 5: blank
    [],
    // Row 6: Section header
    [sectionHeader("KEY PERFORMANCE INDICATORS - 25N COWORKING")],
    // Row 7: blank
    [],
    // Row 8: KPI column headers
    [
      navyHeader(""),
      navyHeader("Workspace"),
      navyHeader("# of Spaces"),
      navyHeader("Occupied"),
      navyHeader("% of Spaces"),
      navyHeader("Total Square Feet"),
      navyHeader("Occupied"),
      navyHeader("% of Square Feet"),
    ],
    // Row 9: Dedicated Desk
    ddRow,
    // Row 10: Private Office
    poRow2,
    // Row 11: Total
    totalRow,
    // Row 12-13: blank
    [],
    [],
    // Row 14: 12-month section header
    [sectionHeader("12 MONTH SNAPSHOT")],
    // Row 15: blank
    [],
    // Row 16: blank
    [],
    // Row 17: snapshot column headers
    [
      navyHeader("Month"),
      navyHeader("Occupied % of Total Spaces"),
      navyHeader("Occupied % of Total Square Feet"),
      navyHeader("# of Contracts"),
      navyHeader("# of Flex Members"),
      navyHeader("# of Mail Members"),
      navyHeader("Recurring & Promo Meeting Hours Sold"),
      navyHeader("Meeting Hours Sold $"),
    ],
    // Rows 18+: snapshot data
    ...snapshotRows,
    // Last row: averages
    avgRow,
  ];
}

// ─── P.2 Monthly Income Sheet ─────────────────────────────────────────────────

function buildP2Sheet(sales: SalesByMonthData): AoaCell[][] {
  const { sheetName: srcName, rows, headerRowIdx, monthColIndices, monthCodes } = sales;
  const output: AoaCell[][] = [];

  // Pre-header rows
  for (let i = 0; i < headerRowIdx; i++) {
    output.push(
      rows[i].map((cell): AoaCell => {
        if (isExcelSerial(cell)) return serialToMonthCode(cell as number);
        const n = num(cell);
        return n !== null ? n : str(cell) || null;
      })
    );
  }

  // Header row — month columns as "Mon-YY" strings, styled navy
  const headerRow = rows[headerRowIdx];
  const outHeader: AoaCell[] = [];
  for (let j = 0; j < headerRow.length; j++) {
    const midx = monthColIndices.indexOf(j);
    if (midx >= 0) {
      outHeader.push(navyHeader(monthCodes[midx]));
    } else if (j === 0) {
      outHeader.push(navyHeader(str(headerRow[j]) || "Sales Comparative by Month"));
    } else {
      const cell = headerRow[j];
      outHeader.push(navyHeader(isExcelSerial(cell) ? serialToMonthCode(cell as number) : str(cell) || null));
    }
  }
  output.push(outHeader);

  // Data rows — month value cells become cross-sheet formula references
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const label = str(row[0]);
    const srcRow = i + 1; // 1-indexed row in source Sales By Month sheet
    const rowIdx = i - headerRowIdx - 1; // for alternating color
    const isTotalRow = /total|subtotal/i.test(label);

    const outRow: AoaCell[] = [];
    for (let j = 0; j < headerRow.length; j++) {
      if (j === 0) {
        outRow.push(isTotalRow ? boldCell(label || null) : labelCell(label, rowIdx));
      } else if (monthColIndices.indexOf(j) >= 0) {
        // Formula reference to source sheet
        const f = sheetRef(srcName, cellAddr(j, srcRow));
        outRow.push(isTotalRow
          ? { f, t: "n", s: boldCell(0, "$#,##0").s, z: "$#,##0" }
          : formulaCell(f, "$#,##0", rowIdx));
      } else {
        const n = num(row[j]);
        const cellVal = (n ?? str(row[j])) || null;
        const cellFmt = n != null ? "$#,##0" : undefined;
        outRow.push(isTotalRow ? boldCell(cellVal, cellFmt) : dataCell(cellVal, cellFmt, rowIdx));
      }
    }
    output.push(outRow);
  }

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
  const currentMonthColLetter = colLetter(monthColIndices[currentMonthIdx]);

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

  // P.4 output starts at row (output.length + 1) in the worksheet
  const dataStartOutputRow = output.length + 1; // 1-indexed row of first data row in output

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
    const outputRow = dataStartOutputRow + dataIdx; // 1-indexed row in P.4 output sheet
    const rowIdx = dataIdx;

    // Budget: reference source P.4 column C (budget column)
    const srcDataRow = i + 1; // 1-indexed in source P.4 sheet
    const budgetCell: AoaCell = { f: sheetRef(srcSheetName, `C${srcDataRow}`), t: "n", s: (isTotalRow ? boldCell(0, "$#,##0") : dataCell(0, "$#,##0", rowIdx)).s, z: "$#,##0" };

    // Actual: reference P.2 at current month column for this label's row
    const salesRow = lookupRowIdx.get(labelLower) ?? null;
    let actualCell: AoaCell;
    if (salesRow != null) {
      // P.2 has the same row structure as Sales By Month (P.2 is built from it)
      // The P.2 header row occupies the same relative position, so we reference P.2 directly
      const p2DataRow = salesRow; // P.2 mirrors source row indices
      actualCell = { f: sheetRef(p2SheetName, `${currentMonthColLetter}${p2DataRow}`), t: "n", s: (isTotalRow ? boldCell(0, "$#,##0") : dataCell(0, "$#,##0", rowIdx)).s, z: "$#,##0" };
    } else {
      // Fall back to computed value
      const monthlyVals = lookupByLabel.get(labelLower);
      const actualVal = monthlyVals != null ? (monthlyVals[currentMonthIdx] ?? null) : null;
      actualCell = isTotalRow ? boldCell(actualVal, "$#,##0") : dataCell(actualVal, "$#,##0", rowIdx);
    }

    // Variance $: =B{row}-C{row}
    const varDollarCell: AoaCell = {
      f: `B${outputRow}-C${outputRow}`,
      t: "n",
      s: (isTotalRow ? boldCell(0, "$#,##0") : dataCell(0, "$#,##0", rowIdx)).s,
      z: "$#,##0",
    };

    // Variance %: =IF(ABS(C{row})>0,(B{row}-C{row})/ABS(C{row}),"")
    const varPctCell: AoaCell = {
      f: `IF(ABS(C${outputRow})>0,(B${outputRow}-C${outputRow})/ABS(C${outputRow}),"")`,
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
  // Build a plain value AOA for xlsx-js-style, then overlay style + formula objects
  const ws = XLSX.utils.aoa_to_sheet(
    aoa.map((row) =>
      row.map((cell) => {
        if (cell == null || typeof cell !== "object") return cell;
        const c = cell as StyledCell;
        if (c.f) return null; // placeholder — formula cells get written below
        return c.v ?? null;
      })
    )
  );

  for (let r = 0; r < aoa.length; r++) {
    for (let c = 0; c < aoa[r].length; c++) {
      const cell = aoa[r][c];
      if (cell == null || typeof cell !== "object") continue;
      const sc = cell as StyledCell;
      const addr = XLSX.utils.encode_cell({ r, c });

      if (!ws[addr]) {
        ws[addr] = { t: sc.t ?? "s", v: sc.v ?? "" };
      }
      if (sc.f) {
        ws[addr] = { t: sc.t ?? "n", f: sc.f };
      }
      if (sc.s) {
        ws[addr].s = sc.s;
      }
      if (sc.z) {
        ws[addr].z = sc.z;
      }
      if (sc.t) {
        ws[addr].t = sc.t;
      }
    }
  }

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
  const metrics = parseMonthlyMetrics(wb);
  const occupancyStats = parseOccupancyStats(wb);

  const dashNames = {
    p1: wb.SheetNames.find((n) => n.startsWith("P.1")) ?? "P.1 Dashboard NEW",
    p2: wb.SheetNames.find((n) => n.startsWith("P2")) ?? "P2. Monthly Income",
    p4: wb.SheetNames.find((n) => n.startsWith("P.4")) ?? "P.4 Forecast & Budget to Actual",
  };

  const p2Data = buildP2Sheet(sales);
  const p1Data = buildP1Sheet(occupancyStats, sales, metrics, dashNames.p2);
  const p4Data = buildP4Sheet(wb, sales, dashNames.p2);

  const newWb = XLSX.utils.book_new();

  const dashSet = new Set(Object.values(dashNames));
  const sheetOrder = [
    dashNames.p1,
    dashNames.p2,
    dashNames.p4,
    ...wb.SheetNames.filter((n) => !dashSet.has(n)),
  ];

  for (const name of sheetOrder) {
    if (name === dashNames.p1) {
      const ws = aoaToStyledSheet(p1Data);
      setColWidths(ws, [20, 12, 28, 28, 18, 14, 14, 18, 22]);
      XLSX.utils.book_append_sheet(newWb, ws, name);
    } else if (name === dashNames.p2) {
      const ws = aoaToStyledSheet(p2Data);
      setColWidths(ws, [36, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12]);
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
