import * as XLSX from "xlsx-js-style";
import {
  parseSalesByMonth,
  parse12MonthStatement,
  parseOccupancyStats,
  extractCurrentSnapshot,
  parseHistoricalSheet,
  parseBudgetSource,
  auditBudgetMappings,
  type KpiSnapshot,
  type SalesByMonthData,
  type OccupancyStats,
} from "@/lib/excel-dashboard-generator";
import { applyRecalcFlags } from "@/lib/recalc-engine";
import { injectPivotParts, preserveTemplateStyles, type PivotReport } from "@/lib/pivot-engine";
import { reconcileOutput, type ReconciliationReport } from "@/lib/reconciliation-engine";

// ─── Types ────────────────────────────────────────────────────────────────────

export type HydrationSectionStatus = "hydrated" | "preserved" | "source_missing";

export interface HydrationSectionResult {
  section: string;
  status:  HydrationSectionStatus;
  note?:   string;
}

export interface HydrationResult {
  buffer:            Buffer;
  snapshot:          KpiSnapshot;
  historicalParsed?: KpiSnapshot[];
  tabsRefreshed:     string[];
  tabsPreserved:     string[];
  warnings:          string[];
  reconciliation:    ReconciliationReport;
  pivotReport:       PivotReport;
  sectionResults:    HydrationSectionResult[];
}

// ─── Cell clone ───────────────────────────────────────────────────────────────

function cloneCell(cell: XLSX.CellObject): XLSX.CellObject {
  return JSON.parse(JSON.stringify(cell));
}

function cloneSheet(ws: XLSX.WorkSheet): XLSX.WorkSheet {
  const out: XLSX.WorkSheet = {};
  for (const key of Object.keys(ws)) {
    if (key === "!ref" || key === "!merges" || key === "!cols" || key === "!rows" ||
        key === "!margins" || key === "!pageSetup" || key === "!autofilter" ||
        key === "!condfmt") {
      out[key] = JSON.parse(JSON.stringify(ws[key]));
    } else if (!key.startsWith("!")) {
      out[key] = cloneCell(ws[key] as XLSX.CellObject);
    } else {
      out[key] = ws[key];
    }
  }
  return out;
}

// ─── Sales By Month hydration ─────────────────────────────────────────────────

function hydrateSalesByMonthSheet(
  templateSheet: XLSX.WorkSheet,
  sales: SalesByMonthData,
): { sheet: XLSX.WorkSheet; warnings: string[] } {
  const warnings: string[] = [];
  const ws = cloneSheet(templateSheet);

  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");

  let templateHeaderRow = -1;
  const templateMonthCols: { colIdx: number; serial: number }[] = [];

  for (let r = 0; r <= range.e.r; r++) {
    const rowSerials: { colIdx: number; serial: number }[] = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "number" && cell.v > 40000 && cell.v < 60000) {
        rowSerials.push({ colIdx: c, serial: cell.v });
      }
    }
    if (rowSerials.length >= 3) {
      templateHeaderRow = r;
      templateMonthCols.push(...rowSerials);
      break;
    }
  }

  if (templateHeaderRow < 0) {
    // Fallback: template may use text month headers ("Jan 2026") instead of date serials
    const TM_RE  = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})$/i;
    const TM_IDX: Record<string, number> = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
    for (let r2 = 0; r2 <= range.e.r; r2++) {
      const hits: { colIdx: number; serial: number }[] = [];
      for (let c = 0; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r: r2, c })];
        if (!cell?.v) continue;
        const m = TM_RE.exec(String(cell.v).trim());
        if (m) {
          const idx  = TM_IDX[m[1].toLowerCase()];
          const year = parseInt(m[2]);
          if (idx !== undefined && year) {
            const serial = Math.floor(Date.UTC(year, idx, 1) / 86400000) + 25569;
            hits.push({ colIdx: c, serial });
          }
        }
      }
      if (hits.length >= 3) { templateHeaderRow = r2; templateMonthCols.push(...hits); break; }
    }
  }

  if (templateHeaderRow < 0) {
    warnings.push("Sales By Month: could not find date-serial or text-month header row in template — sheet not hydrated");
    return { sheet: ws, warnings };
  }

  const serialToSourceIdx = new Map<number, number>();
  for (let i = 0; i < sales.monthSerials.length; i++) {
    serialToSourceIdx.set(sales.monthSerials[i], i);
  }

  const colMap: { templateCol: number; sourceIdx: number }[] = [];
  for (const { colIdx, serial } of templateMonthCols) {
    const sourceIdx = serialToSourceIdx.get(serial);
    if (sourceIdx !== undefined) {
      colMap.push({ templateCol: colIdx, sourceIdx });
    } else {
      warnings.push(`Sales By Month: template column serial ${serial} not found in source workbook`);
    }
  }

  if (colMap.length === 0) {
    warnings.push("Sales By Month: no month columns matched between template and source — sheet not hydrated");
    return { sheet: ws, warnings };
  }

  for (let r = templateHeaderRow + 1; r <= range.e.r; r++) {
    const labelCell = ws[XLSX.utils.encode_cell({ r, c: 0 })];
    if (!labelCell || labelCell.v == null || String(labelCell.v).trim() === "") continue;

    const key = String(labelCell.v).trim().toLowerCase();
    let sourceValues = sales.lookupByLabel.get(key);
    if (!sourceValues) {
      for (const [k, v] of sales.lookupByLabel) {
        if (k.length >= 5 && (k.includes(key) || (key.length >= 5 && key.includes(k)))) {
          sourceValues = v;
          break;
        }
      }
    }

    if (!sourceValues) continue;

    for (const { templateCol, sourceIdx } of colMap) {
      const addr = XLSX.utils.encode_cell({ r, c: templateCol });
      const existingCell = ws[addr];
      const newValue = sourceValues[sourceIdx] ?? null;
      if (newValue === null) continue;

      if (existingCell) {
        existingCell.v = newValue;
        existingCell.t = "n";
        delete (existingCell as Record<string, unknown>).f;
      } else {
        ws[addr] = { v: newValue, t: "n" };
      }
    }
  }

  return { sheet: ws, warnings };
}

// ─── P.1 snapshot section hydration ──────────────────────────────────────────

const FMT_DATE_MMM_YY = '[$-409]mmm\\-yy;@';
const FMT_PCT_WHOLE   = "0%";
const FMT_INT_COMMA   = "#,##0";
const FMT_ACCT_2DEC   = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';
const FMT_ACCT_DOLLAR = '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"??_);_(@_)';
const FMT_ACCT_COUNT  = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)';

function buildSnapshotCells(snap: KpiSnapshot, isCurrentMonth: boolean): (XLSX.CellObject | null)[] {
  const dateStyle = isCurrentMonth
    ? { font: { bold: true, color: { rgb: "FFFF0000" }, sz: 11 }, fill: { fgColor: { rgb: "E7E6E6" } }, numFmt: FMT_DATE_MMM_YY }
    : { numFmt: FMT_DATE_MMM_YY };

  return [
    null,
    { v: snap.monthSerial,          t: "n", z: FMT_DATE_MMM_YY,  s: dateStyle },
    { v: snap.occupancyPctSpaces,   t: "n", z: FMT_PCT_WHOLE,    s: { numFmt: FMT_PCT_WHOLE   } },
    { v: snap.occupancyPctSqFt,     t: "n", z: FMT_PCT_WHOLE,    s: { numFmt: FMT_PCT_WHOLE   } },
    { v: snap.contracts,            t: "n", z: FMT_INT_COMMA,    s: { numFmt: FMT_INT_COMMA   } },
    { v: snap.flexMembers,          t: "n", z: FMT_INT_COMMA,    s: { numFmt: FMT_INT_COMMA   } },
    { v: snap.mailMembers,          t: "n", z: FMT_INT_COMMA,    s: { numFmt: FMT_INT_COMMA   } },
    { v: snap.meetingHours,         t: "n", z: FMT_ACCT_2DEC,   s: { numFmt: FMT_ACCT_2DEC   } },
    { v: snap.meetingRevenueDollar, t: "n", z: FMT_ACCT_DOLLAR, s: { numFmt: FMT_ACCT_DOLLAR } },
  ] as (XLSX.CellObject | null)[];
}

function hydrateP1Sheet(
  templateSheet: XLSX.WorkSheet,
  snapshots: KpiSnapshot[],
  currentSnapshot: KpiSnapshot,
): { sheet: XLSX.WorkSheet; warnings: string[] } {
  const warnings: string[] = [];
  const ws = cloneSheet(templateSheet);

  const b2Addr = "B2";
  if (ws[b2Addr]) {
    ws[b2Addr].v = currentSnapshot.monthSerial;
    ws[b2Addr].t = "n";
    delete (ws[b2Addr] as Record<string, unknown>).f;
  } else {
    ws[b2Addr] = { v: currentSnapshot.monthSerial, t: "n", z: FMT_DATE_MMM_YY };
  }

  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  const SNAP_START_ROW_0 = 17;
  for (const addr of Object.keys(ws)) {
    if (addr.startsWith("!")) continue;
    const cell = XLSX.utils.decode_cell(addr);
    if (cell.r >= SNAP_START_ROW_0) delete ws[addr];
  }

  const allSnaps: KpiSnapshot[] = [...snapshots];
  const currentAlready = allSnaps.some((s) => s.monthCode === currentSnapshot.monthCode);
  if (!currentAlready && currentSnapshot.monthSerial > 0) allSnaps.push(currentSnapshot);

  const cols = [1, 2, 3, 4, 5, 6, 7, 8];

  for (let i = 0; i < allSnaps.length; i++) {
    const snap = allSnaps[i];
    const isCurrentMonth = snap.monthCode === currentSnapshot.monthCode;
    const cells = buildSnapshotCells(snap, isCurrentMonth);
    const excelRow = SNAP_START_ROW_0 + i;

    for (let c = 0; c < cols.length; c++) {
      const cell = cells[c];
      if (cell == null) continue;
      ws[XLSX.utils.encode_cell({ r: excelRow, c: cols[c] })] = cell;
    }
  }

  const snapEndRow1   = 18 + allSnaps.length - 1;
  const snapStartRow1 = 18;
  const avgFmt = (col: string) => `AVERAGE(${col}${snapStartRow1}:${col}${snapEndRow1})`;
  const avgRow0 = SNAP_START_ROW_0 + allSnaps.length;

  const avgStyle = (numFmt: string, align?: string): XLSX.CellObject["s"] => ({
    font: { bold: true, color: { rgb: "FF0000" } },
    numFmt,
    ...(align ? { alignment: { horizontal: align } } : {}),
  });

  const avgCells: (XLSX.CellObject | null)[] = [
    null,
    { v: "Average", t: "s", s: avgStyle("mm-dd-yy", "right") },
    { f: avgFmt("C"), v: 0, t: "n", s: avgStyle(FMT_PCT_WHOLE,  "center"), z: FMT_PCT_WHOLE   },
    { f: avgFmt("D"), v: 0, t: "n", s: avgStyle(FMT_PCT_WHOLE,  "center"), z: FMT_PCT_WHOLE   },
    { f: avgFmt("E"), v: 0, t: "n", s: avgStyle(FMT_ACCT_COUNT, "center"), z: FMT_ACCT_COUNT  },
    { f: avgFmt("F"), v: 0, t: "n", s: avgStyle(FMT_ACCT_COUNT, "center"), z: FMT_ACCT_COUNT  },
    { f: avgFmt("G"), v: 0, t: "n", s: avgStyle(FMT_ACCT_COUNT, "center"), z: FMT_ACCT_COUNT  },
    { f: avgFmt("H"), v: 0, t: "n", s: avgStyle(FMT_ACCT_COUNT, "center"), z: FMT_ACCT_COUNT  },
    { f: avgFmt("I"), v: 0, t: "n", s: avgStyle(FMT_ACCT_COUNT, "center"), z: FMT_ACCT_COUNT  },
  ] as (XLSX.CellObject | null)[];

  for (let c = 0; c < cols.length; c++) {
    const cell = avgCells[c];
    if (cell == null) continue;
    ws[XLSX.utils.encode_cell({ r: avgRow0, c: cols[c] })] = cell;
  }

  const newLastRow = avgRow0;
  const newLastCol = Math.max(range.e.c, 8);
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: newLastRow, c: newLastCol } });

  if (allSnaps.length === 0) {
    warnings.push("P.1: no snapshots available — snapshot section will be empty");
  }

  return { sheet: ws, warnings };
}

// ─── P.1 rolling snapshot updater (new master template format) ───────────────
//
// Preserves ALL template content. Only:
//  1. Updates B2 with current month serial from sales data
//  2. Hides the oldest visible snapshot row (rolling 12-month window)
//  3. Shifts the AVERAGE formula start row by 1 to match the new window
//
// Does NOT wipe rows, does NOT call hydrateP1Sheet.

function updateP1SnapshotRolling(
  templateSheet:   XLSX.WorkSheet,
  sales:           SalesByMonthData | null,
  currentSnapshot: KpiSnapshot | null,
): { sheet: XLSX.WorkSheet; warnings: string[] } {
  const warnings: string[] = [];
  const ws    = cloneSheet(templateSheet);
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");

  // 1. Update B2 with current month serial
  if (sales) {
    const monthSerial = sales.monthSerials[sales.currentMonthIdx];
    if (monthSerial) {
      if (ws["B2"]) {
        ws["B2"].v = monthSerial;
        ws["B2"].t = "n";
        delete (ws["B2"] as Record<string, unknown>).f;
      } else {
        ws["B2"] = { v: monthSerial, t: "n", z: FMT_DATE_MMM_YY };
      }
    }
  }

  // 2. Find the AVERAGE row first (most reliable anchor) — any col with AVERAGE formula in rows 50-100
  //    Then live row = AVERAGE row - 1.
  //    Fallback: scan for col B formula that is "+B2", "B2", or any short ref to B2.
  let avgRow0  = -1;
  let liveRow0 = -1;

  for (let r = 17; r <= Math.min(range.e.r, 120); r++) {
    for (let c = 2; c <= Math.min(range.e.c, 12); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell?.f && /\bAVERAGE\b/i.test(cell.f)) { avgRow0 = r; break; }
    }
    if (avgRow0 >= 0) break;
  }

  if (avgRow0 >= 0) {
    liveRow0 = avgRow0 - 1;
  } else {
    // Fallback: col B formula = "+B2" or "B2"
    for (let r = 17; r <= Math.min(range.e.r, 120); r++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: 1 })];
      if (cell?.f && /^\+?B2$/.test(cell.f.trim())) { liveRow0 = r; break; }
    }
  }

  if (liveRow0 < 0) {
    warnings.push("P.1: could not find snapshot live row or AVERAGE row — snapshot window not rolled");
    return { sheet: ws, warnings };
  }

  // 3. Find the contiguous block of snapshot rows (col B has date serials) just above liveRow0
  let firstSnapRow0 = liveRow0 - 1;
  {
    const anchor = ws[XLSX.utils.encode_cell({ r: firstSnapRow0, c: 1 })];
    if (!anchor || typeof anchor.v !== "number" || anchor.v <= 40000 || anchor.v >= 70000) {
      warnings.push("P.1: row before live formula row has no date serial — cannot determine snapshot range");
      return { sheet: ws, warnings };
    }
    while (firstSnapRow0 > 0) {
      const prev = ws[XLSX.utils.encode_cell({ r: firstSnapRow0 - 1, c: 1 })];
      if (prev && typeof prev.v === "number" && prev.v > 40000 && prev.v < 70000) {
        firstSnapRow0--;
      } else { break; }
    }
  }

  // 4. Ensure !rows metadata array is initialised and large enough
  const rowsMeta: XLSX.RowInfo[] = (ws["!rows"] as XLSX.RowInfo[] | undefined) ?? [];
  while (rowsMeta.length <= liveRow0 + 1) rowsMeta.push({});
  ws["!rows"] = rowsMeta;

  // 5. Write current snapshot data into the live row and ensure it's visible.
  //    The live row may be hidden (was last month's "current" which got rolled out)
  //    or may have stale static values. Either way, overwrite with fresh data.
  if (currentSnapshot && currentSnapshot.monthSerial > 0) {
    const cells    = buildSnapshotCells(currentSnapshot, true);
    const snapCols = [1, 2, 3, 4, 5, 6, 7, 8];
    for (let ci = 0; ci < snapCols.length; ci++) {
      const cell = cells[snapCols[ci]]; // snapCols[ci] is the correct index into the cells array
      if (cell == null) continue;
      const addr = XLSX.utils.encode_cell({ r: liveRow0, c: snapCols[ci] });
      ws[addr] = cell;
    }
  }
  // Un-hide the live row so the current month is visible in the 12-month window.
  if (!rowsMeta[liveRow0]) rowsMeta[liveRow0] = {};
  rowsMeta[liveRow0].hidden = false;

  // 6. Find the first NON-hidden snapshot row (the oldest visible month)
  let firstVisibleRow0 = -1;
  for (let r = firstSnapRow0; r < liveRow0; r++) {
    if (!rowsMeta[r]?.hidden) { firstVisibleRow0 = r; break; }
  }
  if (firstVisibleRow0 < 0) {
    warnings.push("P.1: all snapshot rows are already hidden — nothing to roll");
    return { sheet: ws, warnings };
  }

  // 7. Hide the oldest visible row (keeps window at 12 months)
  if (!rowsMeta[firstVisibleRow0]) rowsMeta[firstVisibleRow0] = {};
  rowsMeta[firstVisibleRow0].hidden = true;

  // 8. Shift AVERAGE formula start row (avgRow0 already known, or derive from liveRow0)
  const resolvedAvgRow0 = avgRow0 >= 0 ? avgRow0 : liveRow0 + 1;
  const oldRowRef  = firstVisibleRow0 + 1; // 1-indexed Excel row
  const newRowRef  = oldRowRef + 1;
  const oldRefPat  = new RegExp(`\\$${oldRowRef}\\b`, "g");

  for (let c = 2; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: resolvedAvgRow0, c });
    const cell = ws[addr];
    if (!cell?.f) continue;
    const updated = cell.f.replace(oldRefPat, `$${newRowRef}`);
    if (updated !== cell.f) cell.f = updated;
  }

  warnings.push(`P.1: rolled snapshot window — hid row ${oldRowRef} (oldest visible), live row ${liveRow0 + 1} shows ${currentSnapshot?.monthCode ?? "current month"}, AVERAGE starts at row ${newRowRef}`);
  return { sheet: ws, warnings };
}

// ─── Income detail row parser ──────────────────────────────────────────────────

interface IncomeDetailRow {
  invoice:    string | number | null;
  property:   string | null;
  company:    string | null;
  item:       string | null;
  desc:       string | null;
  chargeCode: string | null;
  postMonth:  string | null;
  income:     number | null;
}

function parseIncomeDetailRows(wb: XLSX.WorkBook): IncomeDetailRow[] | null {
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws["!ref"]) continue;
    const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(
      ws, { header: 1, defval: null, raw: true },
    ) as (string | number | null)[][];

    let headerRowIdx = -1;
    let idxInvoice  = -1, idxProperty = -1, idxCompany = -1, idxItem = -1;
    let idxDesc     = -1, idxCharge  = -1, idxMonth   = -1, idxIncome = -1;

    for (let r = 0; r < Math.min(rows.length, 15); r++) {
      const texts = rows[r].map((v) => v != null ? String(v).toLowerCase().trim() : "");
      const has = (kw: string) => texts.some((t) => t.includes(kw));
      // Require at minimum a charge code column OR both company + income columns
      if (!has("charge code") && !(has("company") && (has("income") || has("amount")))) continue;

      idxInvoice  = texts.findIndex((t) => t.includes("invoice"));
      idxProperty = texts.findIndex((t) => t.includes("property"));
      idxCompany  = texts.findIndex((t) => t.includes("company"));
      idxItem     = texts.findIndex((t) => t === "item" || t === "charge item");
      idxDesc     = texts.findIndex((t) => t.includes("description") || t === "desc");
      idxCharge   = texts.findIndex((t) => t.includes("charge code"));
      idxMonth    = texts.findIndex((t) => t.includes("post month") || t === "month" || t === "post date");
      idxIncome   = texts.findIndex((t) => t === "income" || t === "amount" || t.includes("income"));
      headerRowIdx = r;
      break;
    }

    // Require at least charge code and company to proceed (invoice is optional)
    if (headerRowIdx < 0 || (idxCharge < 0 && idxCompany < 0)) continue;

    const result: IncomeDetailRow[] = [];
    for (let r = headerRowIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.every((v) => v == null)) continue;

      const getStr = (idx: number) => idx >= 0 && row[idx] != null ? String(row[idx]) : null;
      const incRaw = idxIncome >= 0 ? row[idxIncome] : null;
      const income = incRaw == null ? null : (() => {
        const n = typeof incRaw === "number" ? incRaw : parseFloat(String(incRaw).replace(/[$,\s]/g, ""));
        return isNaN(n) ? null : n;
      })();

      result.push({
        invoice:    idxInvoice  >= 0 ? row[idxInvoice] as (string | number | null) : null,
        property:   getStr(idxProperty),
        company:    getStr(idxCompany),
        item:       getStr(idxItem),
        desc:       getStr(idxDesc),
        chargeCode: getStr(idxCharge),
        postMonth:  getStr(idxMonth),
        income,
      });
    }

    if (result.length > 0) return result;
  }
  return null;
}

function injectIncomeDetailRows(
  stagingSheet: XLSX.WorkSheet,
  rows:         IncomeDetailRow[],
): XLSX.WorkSheet {
  const ws = cloneSheet(stagingSheet);
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");

  // Locate header row in staging tab (contains "Invoice" and "Charge Code")
  let dataStartRow0 = 8; // default: row 9 (1-indexed)
  for (let r = 0; r <= Math.min(range.e.r, 15); r++) {
    const texts: string[] = [];
    for (let c = 0; c <= Math.min(range.e.c, 12); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell?.v) texts.push(String(cell.v).toLowerCase());
    }
    const rowText = texts.join(" ");
    if (/invoice/.test(rowText) && /charge.?code|income|amount/.test(rowText)) {
      dataStartRow0 = r + 1;
      break;
    }
  }

  // Clear existing data rows
  for (const addr of Object.keys(ws)) {
    if (addr.startsWith("!")) continue;
    const { r } = XLSX.utils.decode_cell(addr);
    if (r >= dataStartRow0) delete ws[addr];
  }

  // Write merged rows into staging tab columns A–H
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const r   = dataStartRow0 + i;
    const vals: (string | number | null)[] = [
      row.invoice, row.property, row.company, row.item,
      row.desc, row.chargeCode, row.postMonth, row.income,
    ];
    for (let c = 0; c < vals.length; c++) {
      const v = vals[c];
      if (v == null) continue;
      ws[XLSX.utils.encode_cell({ r, c })] =
        typeof v === "number" ? { v, t: "n" } : { v: String(v), t: "s" };
    }
  }

  const newEndRow = Math.max(range.e.r, dataStartRow0 + Math.max(rows.length - 1, 0));
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: newEndRow, c: Math.max(range.e.c, 7) } });
  return ws;
}

// ─── Per-tab income detail injector ──────────────────────────────────────────
//
// Injects filtered income detail rows directly into tabs like "Flex" and
// "Mail Members".  Each of these tabs has:
//   LEFT  — raw income detail rows (Company / Item / Desc / Charge Code / Month / Income)
//   RIGHT — summary formulas (COUNTIF/SUMIF) that reference the left-side rows
//
// The function discovers which charge codes belong to a tab by reading the
// existing template data rows — no hardcoding needed.  Only left-side columns
// are cleared and rewritten; right-side summary columns are untouched.

function injectIncomeDetailIntoTab(
  tabSheet: XLSX.WorkSheet,
  allRows:  IncomeDetailRow[],
): { sheet: XLSX.WorkSheet; injected: number; chargeCodesFound: number } {
  const ws = cloneSheet(tabSheet);
  const wsRows = XLSX.utils.sheet_to_json<(string | number | null)[]>(
    ws, { header: 1, defval: null, raw: true },
  ) as (string | number | null)[][];

  // Find the data header row: must contain "charge code" column
  let headerRow0    = -1;
  let colCompany    = -1;
  let colItem       = -1;
  let colDesc       = -1;
  let colChargeCode = -1;
  let colPostMonth  = -1;
  let colIncome     = -1;

  for (let r = 0; r < Math.min(wsRows.length, 15); r++) {
    const texts = wsRows[r].map((v) => (v != null ? String(v).toLowerCase().trim() : ""));
    const rowStr = texts.join(" ");
    if (!rowStr.includes("charge code")) continue;
    colCompany    = texts.findIndex((t) => t === "company name" || t === "company");
    colItem       = texts.findIndex((t) => t === "item" || t === "charge item");
    colDesc       = texts.findIndex((t) => t.includes("description") || t === "desc");
    colChargeCode = texts.findIndex((t) => t.includes("charge code"));
    colPostMonth  = texts.findIndex((t) => t.includes("post month") || t === "month");
    colIncome     = texts.findIndex((t) => t === "income" || t === "amount");
    if (colChargeCode >= 0) { headerRow0 = r; break; }
  }

  if (headerRow0 < 0 || colChargeCode < 0) return { sheet: ws, injected: 0, chargeCodesFound: 0 };

  const dataStart  = headerRow0 + 1;
  const maxDataCol = Math.max(colCompany, colItem, colDesc, colChargeCode, colPostMonth, colIncome);

  // Auto-discover charge codes from the template's existing sample rows
  const existingCodes = new Set<string>();
  for (let r = dataStart; r < wsRows.length; r++) {
    const cc = wsRows[r][colChargeCode];
    if (cc != null && String(cc).trim()) existingCodes.add(String(cc).trim());
  }

  if (existingCodes.size === 0) return { sheet: ws, injected: 0, chargeCodesFound: 0 };

  // Filter all income rows to only those whose charge code appears in this tab
  const filteredRows = allRows.filter(
    (row) => row.chargeCode && existingCodes.has(row.chargeCode.trim()),
  );
  if (filteredRows.length === 0) return { sheet: ws, injected: 0, chargeCodesFound: existingCodes.size };

  // Clear only the left-side data area (≤ maxDataCol) below the header
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  for (const addr of Object.keys(ws)) {
    if (addr.startsWith("!")) continue;
    const { r, c } = XLSX.utils.decode_cell(addr);
    if (r >= dataStart && c <= maxDataCol) delete ws[addr];
  }

  // Write filtered rows
  const colMap: [number, (row: IncomeDetailRow) => string | number | null][] = [
    [colCompany,    (row) => row.company],
    [colItem,       (row) => row.item],
    [colDesc,       (row) => row.desc],
    [colChargeCode, (row) => row.chargeCode],
    [colPostMonth,  (row) => row.postMonth],
    [colIncome,     (row) => row.income],
  ];

  for (let i = 0; i < filteredRows.length; i++) {
    const row = filteredRows[i];
    const r   = dataStart + i;
    for (const [c, getValue] of colMap) {
      if (c < 0) continue;
      const v = getValue(row);
      if (v == null) continue;
      ws[XLSX.utils.encode_cell({ r, c })] =
        typeof v === "number" ? { v, t: "n" } : { v: String(v), t: "s" };
    }
  }

  const newEndRow = Math.max(range.e.r, dataStart + filteredRows.length - 1);
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: newEndRow, c: range.e.c } });
  return { sheet: ws, injected: filteredRows.length, chargeCodesFound: existingCodes.size };
}

// ─── P.2 gap column filler ────────────────────────────────────────────────────

function shiftColLetter(colStr: string, offset: number): string {
  let n = 0;
  for (const ch of colStr.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  n += offset;
  if (n <= 0) n = 1;
  let result = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function fixP2GapColumns(p2Sheet: XLSX.WorkSheet): XLSX.WorkSheet {
  const ws    = p2Sheet; // already cloned before call
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");

  // Find P.2 header row — 3+ date serials in the 40000–60000 range
  let headerRow = -1;
  const monthCols: { colIdx: number; serial: number }[] = [];
  for (let r = 0; r <= Math.min(range.e.r, 10); r++) {
    const hits: { colIdx: number; serial: number }[] = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "number" && cell.v > 40000 && cell.v < 60000) {
        hits.push({ colIdx: c, serial: cell.v });
      }
    }
    if (hits.length >= 3) { headerRow = r; monthCols.push(...hits); break; }
  }
  if (headerRow < 0 || monthCols.length === 0) return ws;

  // Count formula cells per month column
  const formulasByCol = new Map<number, Map<number, string>>();
  for (const { colIdx } of monthCols) {
    const fByRow = new Map<number, string>();
    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c: colIdx })];
      if (cell?.f) fByRow.set(r, cell.f);
    }
    formulasByCol.set(colIdx, fByRow);
  }

  const gapCols    = monthCols.filter(({ colIdx }) => (formulasByCol.get(colIdx)?.size ?? 0) === 0);
  const filledCols = monthCols.filter(({ colIdx }) => (formulasByCol.get(colIdx)?.size ?? 0) > 0);
  if (gapCols.length === 0 || filledCols.length === 0) return ws;

  // Regex to match any SBM sheet column reference, e.g. '2026 Sales By Month'!F16
  const SBM_COL_RE = /(['"]['"\w\s\-/]*(Sales By Month)['"\w\s\-/]*['"])!([A-Z]+)(\d+)/gi;

  for (const gap of gapCols) {
    // Nearest filled column by P2 column distance
    const ref = filledCols.reduce((best, fc) =>
      Math.abs(fc.colIdx - gap.colIdx) < Math.abs(best.colIdx - gap.colIdx) ? fc : best,
    );
    const refFormulas = formulasByCol.get(ref.colIdx);
    if (!refFormulas) continue;

    const colOffset = gap.colIdx - ref.colIdx;

    for (const [rowIdx, formula] of refFormulas) {
      const newFormula = formula.replace(
        SBM_COL_RE,
        (_match, sheetRef, _sbm, col, row) => `${sheetRef}!${shiftColLetter(col, colOffset)}${row}`,
      );
      const addr    = XLSX.utils.encode_cell({ r: rowIdx, c: gap.colIdx });
      const refAddr = XLSX.utils.encode_cell({ r: rowIdx, c: ref.colIdx });
      const refCell = ws[refAddr];
      if (ws[addr]) {
        ws[addr].f = newFormula;
        ws[addr].v = 0;
      } else if (refCell) {
        const nc = cloneCell(refCell);
        nc.f = newFormula;
        nc.v = 0;
        ws[addr] = nc;
      } else {
        ws[addr] = { f: newFormula, v: 0, t: "n" };
      }
    }

    // Ensure header cell has the correct serial
    const headerAddr = XLSX.utils.encode_cell({ r: headerRow, c: gap.colIdx });
    if (!ws[headerAddr] || ws[headerAddr].v == null) {
      const refHeaderAddr = XLSX.utils.encode_cell({ r: headerRow, c: ref.colIdx });
      if (ws[refHeaderAddr]) {
        const hc = cloneCell(ws[refHeaderAddr]);
        hc.v = gap.serial;
        ws[headerAddr] = hc;
      }
    }
  }

  return ws;
}

// ─── Occupancy Data tab refresh ───────────────────────────────────────────────
//
// Copies the raw unit-level Kube occupancy sheet from a source workbook into the
// "Occupancy Data" staging tab of the template.  P.1 Dashboard formulas
// (COUNTIF / SUMIF) read from this tab, so replacing it with fresh source data
// makes all P.1 KPIs (occupancy %, contract counts, sq ft) auto-update.

function refreshOccupancyDataSheet(
  templateSheet: XLSX.WorkSheet,
  sourceWbs:     XLSX.WorkBook[],
): { sheet: XLSX.WorkSheet; found: boolean } {
  const STATUS_RE = /\bstatus\b|\boccupied\b|\boccupancy\b/i;
  const SPACE_RE  = /\bsq\.?\s*ft\b|\bsqft\b|\bunit\b|\bspace\b/i;

  for (const wb of sourceWbs) {
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      if (!ws["!ref"]) continue;
      const range = XLSX.utils.decode_range(ws["!ref"]);
      for (let r = 0; r <= Math.min(range.e.r, 6); r++) {
        const texts: string[] = [];
        for (let c = 0; c <= Math.min(range.e.c, 20); c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (cell?.v) texts.push(String(cell.v).toLowerCase());
        }
        const row = texts.join(" ");
        if (STATUS_RE.test(row) && SPACE_RE.test(row)) {
          return { sheet: cloneSheet(ws), found: true };
        }
      }
    }
  }
  return { sheet: cloneSheet(templateSheet), found: false };
}

// ─── Budget Actual Data staging + P.4 value injection ────────────────────────
//
// 1. Finds the Voyager Budget Comparison sheet across source workbooks
//    (detects by PTD Actual / PTD Budget / Variance column headers).
// 2. Copies that raw sheet into the "Budget Actual Data" staging tab.
// 3. Directly writes PTD Actual, PTD Budget, Variance, and % Var values into P.4
//    rows by fuzzy-matching P.4 col-A account labels against the source labels.

function processBudgetData(
  p4Template:            XLSX.WorkSheet,
  sourceWbs:             XLSX.WorkBook[],
  budgetActualTemplate?: XLSX.WorkSheet | null,
): { p4Sheet: XLSX.WorkSheet; budgetActualSheet: XLSX.WorkSheet | null; found: boolean; warnings: string[] } {
  const warnings: string[] = [];

  // ── 1. Locate Budget Comparison source sheet ──────────────────────────────
  let budgetSourceWs: XLSX.WorkSheet | null = null;
  let budgetRows: (string | number | null)[][] = [];
  let colPTDActual  = -1;
  let colPTDBudget  = -1;
  let colVariance   = -1;
  let colPctVar     = -1;

  outer:
  for (const wb of sourceWbs) {
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      if (!ws["!ref"]) continue;
      const rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(
        ws, { header: 1, defval: null, raw: true },
      ) as (string | number | null)[][];

      for (let r = 0; r < Math.min(10, rows.length); r++) {
        const rowStrs = rows[r].map((v) => (v != null ? String(v).toLowerCase().trim() : ""));
        const rowStr  = rowStrs.join(" ");
        if (rowStr.includes("ptd") && rowStr.includes("actual") && rowStr.includes("budget")) {
          // Map column positions
          for (let c = 0; c < rowStrs.length; c++) {
            const h = rowStrs[c];
            if (h.includes("ptd") && h.includes("actual") && colPTDActual < 0) colPTDActual = c;
            else if (h.includes("ptd") && h.includes("budget") && colPTDBudget < 0) colPTDBudget = c;
            else if ((h === "variance" || h.startsWith("var")) && colVariance < 0 && colPTDActual >= 0) colVariance = c;
            else if ((h.includes("%") || h.includes("pct")) && colPctVar < 0) colPctVar = c;
          }
          budgetSourceWs = ws;
          budgetRows     = rows.slice(r + 1);
          break outer;
        }
      }
    }
  }

  if (!budgetSourceWs || colPTDActual < 0) {
    return { p4Sheet: cloneSheet(p4Template), budgetActualSheet: null, found: false, warnings };
  }

  // ── 2. Build account-label → values lookup ───────────────────────────────
  type BudgetEntry = { ptdActual: number; ptdBudget: number; variance: number; pctVar: number };
  const lookup = new Map<string, BudgetEntry>();

  const toNum = (v: string | number | null): number => {
    if (v == null) return 0;
    if (typeof v === "number") return v;
    return parseFloat(String(v).replace(/[$,\s]/g, "")) || 0;
  };

  for (const row of budgetRows) {
    const rawLabel = row[0] != null ? String(row[0]).trim() : "";
    if (!rawLabel) continue;
    const key = rawLabel.toLowerCase().replace(/\s+/g, " ");
    lookup.set(key, {
      ptdActual: toNum(row[colPTDActual]),
      ptdBudget: colPTDBudget >= 0 ? toNum(row[colPTDBudget]) : 0,
      variance:  colVariance  >= 0 ? toNum(row[colVariance])  : toNum(row[colPTDActual]) - (colPTDBudget >= 0 ? toNum(row[colPTDBudget]) : 0),
      pctVar:    colPctVar    >= 0 ? toNum(row[colPctVar])    : 0,
    });
  }

  // ── 3. Detect P.4 output column positions from its header row ────────────
  const p4Rows = XLSX.utils.sheet_to_json<(string | number | null)[]>(
    p4Template, { header: 1, defval: null, raw: true },
  ) as (string | number | null)[][];

  let p4ColActual  = 1; // default: col B
  let p4ColBudget  = 2; // default: col C
  let p4ColVar     = 3; // default: col D
  let p4ColPct     = 4; // default: col E
  let p4DataStart  = 5; // default: row 6

  for (let r = 0; r < Math.min(20, p4Rows.length); r++) {
    const rowStrs = p4Rows[r].map((v) => (v != null ? String(v).toLowerCase().trim() : ""));
    const rowStr  = rowStrs.join(" ");
    if (rowStr.includes("actual") && rowStr.includes("budget")) {
      p4DataStart = r + 1;
      for (let c = 0; c < rowStrs.length; c++) {
        const h = rowStrs[c];
        if (h.includes("actual") && p4ColActual === 1) p4ColActual = c;
        if (h.includes("budget") && p4ColBudget === 2) p4ColBudget = c;
        if ((h === "variance" || h.startsWith("var")) && p4ColVar === 3) p4ColVar = c;
        if ((h.includes("%") || h.includes("pct")) && p4ColPct === 4) p4ColPct = c;
      }
      break;
    }
  }

  // ── 4. Write values into P.4 rows ────────────────────────────────────────
  const p4 = cloneSheet(p4Template);
  let matchedCount = 0;

  const setCell = (ws: XLSX.WorkSheet, r: number, c: number, value: number) => {
    const addr = XLSX.utils.encode_cell({ r, c });
    if (ws[addr]) {
      ws[addr].v = value;
      ws[addr].t = "n";
      delete (ws[addr] as Record<string, unknown>).f;
    } else {
      ws[addr] = { v: value, t: "n" };
    }
  };

  for (let r = p4DataStart; r < p4Rows.length; r++) {
    const rawLabel = p4Rows[r][0] != null ? String(p4Rows[r][0]).trim() : "";
    if (!rawLabel) continue;
    const key = rawLabel.toLowerCase().replace(/\s+/g, " ");

    let entry = lookup.get(key);
    if (!entry && key.length >= 5) {
      for (const [k, v] of lookup) {
        if (k.length >= 5 && (k.includes(key) || key.includes(k))) { entry = v; break; }
      }
    }
    if (!entry) continue;

    matchedCount++;
    setCell(p4, r, p4ColActual, entry.ptdActual);
    setCell(p4, r, p4ColBudget, entry.ptdBudget);
    setCell(p4, r, p4ColVar,    entry.variance);
    setCell(p4, r, p4ColPct,    entry.pctVar);
  }

  if (matchedCount === 0) {
    warnings.push("P.4: budget source found but no account labels matched P.4 rows — values not updated");
  } else {
    warnings.push(`P.4: updated ${matchedCount} account row(s) from Budget Comparison source`);
  }

  // ── 5. Build "Budget Actual Data" sheet preserving template column structure ─
  // If we have the template's "Budget Actual Data" sheet, inject values into it
  // so that P.4 formula references (e.g. ='Budget Actual Data'!E5) remain valid.
  // Fallback to copying the raw Voyager sheet only when no template is available.
  let budgetActualOut: XLSX.WorkSheet | null = null;

  if (budgetActualTemplate) {
    const batRows = XLSX.utils.sheet_to_json<(string | number | null)[]>(
      budgetActualTemplate, { header: 1, defval: null, raw: true },
    ) as (string | number | null)[][];

    // Detect column positions in the template's Budget Actual Data header row
    let batColActual   = -1;
    let batColBudget   = -1;
    let batColVariance = -1;
    let batColPct      = -1;
    let batDataStart   = 1;

    for (let r = 0; r < Math.min(15, batRows.length); r++) {
      const rowStrs = batRows[r].map((v) => (v != null ? String(v).toLowerCase().trim() : ""));
      const rowStr  = rowStrs.join(" ");
      if (rowStr.includes("actual") || rowStr.includes("budget") || rowStr.includes("variance")) {
        batDataStart = r + 1;
        for (let c = 0; c < rowStrs.length; c++) {
          const h = rowStrs[c];
          if      (h.includes("ptd") && h.includes("actual") && batColActual   < 0) batColActual   = c;
          else if (h.includes("actual")                       && batColActual   < 0) batColActual   = c;
          if      (h.includes("ptd") && h.includes("budget") && batColBudget   < 0) batColBudget   = c;
          else if (h.includes("budget")                       && batColBudget   < 0) batColBudget   = c;
          if ((h === "variance" || h.startsWith("var"))       && batColVariance < 0) batColVariance = c;
          if ((h.includes("%")  || h.includes("pct"))         && batColPct      < 0) batColPct      = c;
        }
        break;
      }
    }

    const bat = cloneSheet(budgetActualTemplate);
    let batMatchCount = 0;

    for (let r = batDataStart; r < batRows.length; r++) {
      const rawLabel = batRows[r][0] != null ? String(batRows[r][0]).trim() : "";
      if (!rawLabel) continue;
      const key = rawLabel.toLowerCase().replace(/\s+/g, " ");

      let entry = lookup.get(key);
      if (!entry && key.length >= 5) {
        for (const [k, v] of lookup) {
          if (k.length >= 5 && (k.includes(key) || key.includes(k))) { entry = v; break; }
        }
      }
      if (!entry) continue;

      batMatchCount++;
      if (batColActual   >= 0) setCell(bat, r, batColActual,   entry.ptdActual);
      if (batColBudget   >= 0) setCell(bat, r, batColBudget,   entry.ptdBudget);
      if (batColVariance >= 0) setCell(bat, r, batColVariance, entry.variance);
      if (batColPct      >= 0) setCell(bat, r, batColPct,      entry.pctVar);
    }

    if (batMatchCount === 0) {
      warnings.push("Budget Actual Data: Voyager source found but no labels matched template rows — structure preserved, values not updated");
    } else {
      warnings.push(`Budget Actual Data: updated ${batMatchCount} row(s) from Voyager Budget Comparison`);
    }
    budgetActualOut = bat;
  } else {
    budgetActualOut = cloneSheet(budgetSourceWs);
  }

  return {
    p4Sheet:           p4,
    budgetActualSheet: budgetActualOut,
    found:             true,
    warnings,
  };
}

// ─── Empty reconciliation (used when source data is too partial for a full check) ──

function emptyReconciliation(note: string): ReconciliationReport {
  return {
    isClean:           true,
    criticalCount:     0,
    errorCount:        0,
    warningCount:      0,
    infoCount:         0,
    findings:          [],
    categories:        [],
    summary:           note,
    staleFormulaCount: 0,
    generationMode:    "hydration",
  };
}

// ─── Snapshot builder when not all source data is available ──────────────────

function buildPartialSnapshot(
  sales:   SalesByMonthData | null,
  occ:     OccupancyStats   | null,
): KpiSnapshot {
  if (sales && occ) return extractCurrentSnapshot(sales, occ);

  const base: KpiSnapshot = {
    monthCode: "Unknown", monthSerial: 0,
    occupancyPctSpaces: 0, occupancyPctSqFt: 0,
    contracts: 0, flexMembers: 0, mailMembers: 0,
    meetingHours: 0, meetingRevenueDollar: 0,
    location: occ?.location ?? "Unknown",
  };

  if (sales) {
    const { currentMonthIdx, monthCodes, monthSerials } = sales;
    base.monthCode   = monthCodes[currentMonthIdx]   ?? "Unknown";
    base.monthSerial = monthSerials[currentMonthIdx]  ?? 0;
  }

  if (occ) {
    const total    = occ.dedicatedDesk.total   + occ.privateOffice.total;
    const occupied = occ.dedicatedDesk.occupied + occ.privateOffice.occupied;
    if (total > 0) {
      base.occupancyPctSpaces = occupied / total;
      const totalSqFt    = occ.dedicatedDesk.totalSqFt    + occ.privateOffice.totalSqFt;
      const occupiedSqFt = occ.dedicatedDesk.occupiedSqFt + occ.privateOffice.occupiedSqFt;
      if (totalSqFt > 0) base.occupancyPctSqFt = occupiedSqFt / totalSqFt;
    }
    if (!base.location || base.location === "Unknown") base.location = occ.location || "Unknown";
  }

  return base;
}

// ─── Main hydration orchestrator ──────────────────────────────────────────────

/**
 * Hydrate a template workbook with available data from one or more source workbooks.
 *
 * Sectional strategy — only hydrates sections where source data is present:
 *   Sales By Month  → hydrated when source has SBM with date-serial header
 *   P.1 Dashboard   → rows 1-17 preserved; snapshot rows 18+ written from history
 *   P.2 Revenue     → preserved from template (formulas chain through SBM)
 *   P.3 Occupancy   → hydrated when source has occupancy unit-level data
 *   P.4 Budget      → preserved from template (workflow engine handles column update)
 *   All other tabs  → source tabs are passed through; template tabs fill ALL remaining gaps
 *
 * Missing source sections never throw — they fall back to the template version
 * and emit an advisory in the warnings array.
 */
export function hydrateWorkbook(
  templateBuffer:       Buffer | ArrayBuffer,
  sourceBuffers:        ArrayBuffer[],
  historicalSnapshots?: KpiSnapshot[],
): HydrationResult {
  const warnings:       string[]                = [];
  const tabsRefreshed:  string[]                = [];
  const tabsPreserved:  string[]                = [];
  const sectionResults: HydrationSectionResult[] = [];

  const templateWb = XLSX.read(templateBuffer, { type: "buffer", cellDates: false, cellStyles: true, cellNF: true });

  // Parse all source workbooks
  const sourceWbs: XLSX.WorkBook[] = sourceBuffers.map((buf) =>
    XLSX.read(buf, { type: "array", cellDates: false, cellStyles: true, cellNF: true }),
  );

  // ── Parse source data — try each buffer, stop when data found ──────────────

  let sales: SalesByMonthData | null = null;
  let salesWb: XLSX.WorkBook | null = null;
  // Try classic Sales By Month format first, then 12-Month Statement format
  for (const wb of sourceWbs) {
    try { sales = parseSalesByMonth(wb); salesWb = wb; break; } catch { /* try next */ }
  }
  if (!sales) {
    for (const wb of sourceWbs) {
      const result = parse12MonthStatement(wb);
      if (result) { sales = result; salesWb = wb; break; }
    }
  }
  if (!sales) {
    warnings.push("Sales By Month not found in source — P.2 revenue section will use template data. Upload a Voyager P&L export to populate this section.");
  }

  let occupancyStats: OccupancyStats | null = null;
  for (const wb of sourceWbs) {
    try {
      occupancyStats = parseOccupancyStats(wb);
      break;
    } catch { /* try next */ }
  }
  if (!occupancyStats) {
    warnings.push("Occupancy data not found in source — P.3 will use template data. Upload a Kube Occupancy report to populate this section.");
  }

  const currentSnapshot = buildPartialSnapshot(sales, occupancyStats);

  // ── Historical snapshots (Supabase → tab fallback) ──────────────────────────
  let resolvedSnapshots: KpiSnapshot[] | undefined = historicalSnapshots;
  let historicalParsed:  KpiSnapshot[] | undefined;

  if (!resolvedSnapshots || resolvedSnapshots.length === 0) {
    for (const wb of sourceWbs) {
      const fromTab = parseHistoricalSheet(wb, occupancyStats?.location ?? "");
      if (fromTab && fromTab.length > 0) {
        resolvedSnapshots = fromTab;
        historicalParsed  = fromTab;
        break;
      }
    }
  }

  // ── Budget audit (non-blocking, only when both sources available) ────────────
  let budgetSource: ReturnType<typeof parseBudgetSource> = null;
  for (const wb of sourceWbs) {
    const bs = parseBudgetSource(wb);
    if (bs) { budgetSource = bs; break; }
  }
  if (budgetSource && sales) {
    const issues    = auditBudgetMappings(budgetSource, sales.lookupByLabel, sales.lookupRowIdx);
    const criticals = issues.filter((i) => i.severity === "critical");
    if (criticals.length > 0) {
      const labels = criticals.map((i) => `"${i.normalizedLabel}"`).join(", ");
      warnings.push(`Budget: ${criticals.length} unmapped item(s): ${labels}`);
    }
  }

  // ── Identify template sheet names ─────────────────────────────────────────────
  // Handle both "P1. Dashboard" (number-period) and "P.1 Dashboard" (period-number) naming
  const templateP1  = templateWb.SheetNames.find((n) => /^P\.?1[\s.]/i.test(n));
  const templateP2  = templateWb.SheetNames.find((n) => /^P\.?2[\s.]/i.test(n));
  const templateP3  = templateWb.SheetNames.find((n) => n.startsWith("P.3"));
  const templateP4  = templateWb.SheetNames.find((n) => n.startsWith("P.4"));
  const templateSBM = templateWb.SheetNames.find((n) => /sales.?by.?month/i.test(n));

  // Find P.3 in any source workbook
  let sourceP3: string | undefined;
  let sourceP3Wb: XLSX.WorkBook | null = null;
  for (const wb of sourceWbs) {
    const name = wb.SheetNames.find((n) => n.startsWith("P.3"));
    if (name) { sourceP3 = name; sourceP3Wb = wb; break; }
  }

  if (!templateP1) warnings.push("Template is missing P.1 sheet");
  if (!templateP2) warnings.push("Template is missing P.2 sheet");

  // ── Build output sheet map (keyed by template sheet name) ───────────────────
  // Output is always built in template tab order — no source tabs are ever added.
  const sheetMap = new Map<string, XLSX.WorkSheet>();

  // P.1: preserve template exactly — ZIP-level insertP1SnapshotRow handles row insertion
  // after XLSX.write so that all template formatting is inherited by the new snapshot row.
  if (templateP1 && templateWb.Sheets[templateP1]) {
    sheetMap.set(templateP1, cloneSheet(templateWb.Sheets[templateP1]));
    tabsPreserved.push(templateP1);
    sectionResults.push({ section: "P.1 Dashboard", status: "hydrated", note: "Snapshot row inserted at ZIP level after style preservation" });
  }

  // P.2: preserve from template — formulas chain through Sales By Month
  if (templateP2 && templateWb.Sheets[templateP2]) {
    sheetMap.set(templateP2, cloneSheet(templateWb.Sheets[templateP2]));
    tabsPreserved.push(templateP2);
    sectionResults.push({ section: "P.2 Revenue Summary", status: "preserved", note: "Formulas chain through Sales By Month" });
  }

  // P.3: use source version (renamed to template tab name) or fall back to template
  if (sourceP3 && sourceP3Wb && sourceP3Wb.Sheets[sourceP3] && templateP3) {
    sheetMap.set(templateP3, cloneSheet(sourceP3Wb.Sheets[sourceP3]));
    tabsRefreshed.push(templateP3);
    sectionResults.push({ section: "P.3 Occupancy", status: "hydrated", note: `Updated from source tab '${sourceP3}'` });
  } else if (templateP3 && templateWb.Sheets[templateP3]) {
    sheetMap.set(templateP3, cloneSheet(templateWb.Sheets[templateP3]));
    tabsPreserved.push(templateP3);
    sectionResults.push({
      section: "P.3 Occupancy",
      status:  occupancyStats ? "hydrated" : "source_missing",
      note:    occupancyStats
        ? "Occupancy data available but sheet structure differs from expected"
        : "Upload a Kube Occupancy report to update this section",
    });
    if (!occupancyStats) warnings.push("P.3 not found in source — using template version");
  }

  // "Occupancy Data" staging tab: replace template copy with fresh Kube unit-level data.
  // P.1 Dashboard reads ALL occupancy KPIs (COUNTIF/SUMIF) from this tab, not from P.3.
  const OCC_DATA_TAB = templateWb.SheetNames.find((n) => /^occupancy.?data$/i.test(n));
  if (OCC_DATA_TAB && occupancyStats) {
    const tmpl = templateWb.Sheets[OCC_DATA_TAB] ?? {};
    const { sheet: freshOccData, found: occDataFound } = refreshOccupancyDataSheet(tmpl, sourceWbs);
    if (occDataFound) {
      sheetMap.set(OCC_DATA_TAB, freshOccData);
      // Will be overridden in "fill remaining tabs" if not explicitly placed here first
      tabsRefreshed.push(OCC_DATA_TAB);
      sectionResults.push({ section: "Occupancy Data", status: "hydrated", note: "Fresh unit-level Kube occupancy data — P.1 KPIs will auto-update" });
    }
  }

  // P.4: inject PTD Actual/Budget/Variance values from Budget Comparison source,
  //      and refresh the "Budget Actual Data" staging tab with the raw source data.
  const BUDGET_ACTUAL_TAB = templateWb.SheetNames.find((n) => /budget.?actual/i.test(n));
  if (templateP4 && templateWb.Sheets[templateP4]) {
    const { p4Sheet, budgetActualSheet, found: budgetFound, warnings: bw } = processBudgetData(
      templateWb.Sheets[templateP4],
      sourceWbs,
      BUDGET_ACTUAL_TAB ? templateWb.Sheets[BUDGET_ACTUAL_TAB] : null,
    );
    warnings.push(...bw);
    sheetMap.set(templateP4, p4Sheet);
    if (budgetFound) {
      tabsRefreshed.push(templateP4);
      sectionResults.push({ section: "P.4 Budget & Variance", status: "hydrated", note: "PTD Actual/Budget/Variance values injected from Budget Comparison source" });
    } else {
      tabsPreserved.push(templateP4);
      sectionResults.push({ section: "P.4 Budget & Variance", status: "source_missing", note: "Upload a Voyager Budget Comparison export to populate this section" });
    }
    // Refresh "Budget Actual Data" staging tab with raw source data so P.4 formulas (if any) also see fresh data
    if (budgetActualSheet && BUDGET_ACTUAL_TAB) {
      sheetMap.set(BUDGET_ACTUAL_TAB, budgetActualSheet);
      const idx = tabsPreserved.indexOf(BUDGET_ACTUAL_TAB);
      if (idx >= 0) tabsPreserved.splice(idx, 1);
      if (!tabsRefreshed.includes(BUDGET_ACTUAL_TAB)) tabsRefreshed.push(BUDGET_ACTUAL_TAB);
      const srIdx = sectionResults.findIndex((s) => s.section === BUDGET_ACTUAL_TAB);
      const note  = "Raw Budget Comparison data copied from source";
      if (srIdx >= 0) sectionResults[srIdx] = { section: "Budget Actual Data", status: "hydrated", note };
      else sectionResults.push({ section: "Budget Actual Data", status: "hydrated", note });
    }
  }

  // Sales By Month: hydrate from source (date-serial or text-month format) or preserve template
  if (sales !== null && templateSBM && templateWb.Sheets[templateSBM]) {
    const { sheet, warnings: w } = hydrateSalesByMonthSheet(templateWb.Sheets[templateSBM], sales);
    warnings.push(...w);
    sheetMap.set(templateSBM, sheet);
    tabsRefreshed.push(templateSBM);
    sectionResults.push({
      section: "Sales By Month",
      status:  "hydrated",
      note:    `${sales.monthSerials.length} month column(s), ${sales.lookupByLabel.size} line items`,
    });
  } else if (templateSBM && templateWb.Sheets[templateSBM]) {
    sheetMap.set(templateSBM, cloneSheet(templateWb.Sheets[templateSBM]));
    tabsPreserved.push(templateSBM);
    sectionResults.push({
      section: "Sales By Month",
      status:  "source_missing",
      note:    "Upload a Voyager P&L export to populate revenue data",
    });
  }

  // P.2: repair month columns with no formulas (template build gaps, e.g. missing May column)
  if (templateP2) {
    const p2ws = sheetMap.get(templateP2);
    if (p2ws) sheetMap.set(templateP2, fixP2GapColumns(p2ws));
  }

  // Fill all remaining template tabs — preserves complete workbook structure in template order
  for (const name of templateWb.SheetNames) {
    if (sheetMap.has(name)) continue;
    if (!templateWb.Sheets[name]) continue;
    sheetMap.set(name, cloneSheet(templateWb.Sheets[name]));
    tabsPreserved.push(name);
    sectionResults.push({ section: name, status: "preserved" as HydrationSectionStatus, note: "Template tab preserved" });
  }

  // Income detail staging tab: merge all incomedetailbyaccount source files and inject
  // Flex, Mail Members, and Daily Space tabs use COUNTIF against this staging tab.
  const STAGING_RE    = /WA\s+Yardi\s+Inc\s+Detail|Income\s+Detail.*Account/i;
  const stagingTabName = templateWb.SheetNames.find((n) => STAGING_RE.test(n));
  if (stagingTabName) {
    const allIncomeRows: IncomeDetailRow[] = [];
    for (const wb of sourceWbs) {
      const parsed = parseIncomeDetailRows(wb);
      if (parsed) allIncomeRows.push(...parsed);
    }
    if (allIncomeRows.length > 0) {
      const stagingSheet = sheetMap.get(stagingTabName);
      if (stagingSheet) {
        sheetMap.set(stagingTabName, injectIncomeDetailRows(stagingSheet, allIncomeRows));
        // Remove from preserved list and record as hydrated
        const idx = tabsPreserved.indexOf(stagingTabName);
        if (idx >= 0) tabsPreserved.splice(idx, 1);
        tabsRefreshed.push(stagingTabName);
        const srIdx = sectionResults.findIndex((s) => s.section === stagingTabName);
        const note  = `${allIncomeRows.length} transaction rows merged from ${sourceWbs.length} source file(s)`;
        if (srIdx >= 0) sectionResults[srIdx] = { section: "Income Detail Staging", status: "hydrated", note };
        else sectionResults.push({ section: "Income Detail Staging", status: "hydrated", note });
        warnings.push(`Income detail: ${allIncomeRows.length} rows injected into '${stagingTabName}'`);
      }
    } else {
      warnings.push("Income detail staging tab found in template but no income detail source files were uploaded — Flex/Mail Members/Daily Space tabs will not auto-populate.");
    }
  }

  // Inject filtered income detail rows directly into Flex, Mail Members, and Daily Space tabs.
  // Each tab has raw transaction rows on the left (auto-populated here) and summary formulas
  // on the right (COUNTIF/SUMIF) that calculate from those rows — no manual entry needed.
  {
    const allIncomeRowsForTabs: IncomeDetailRow[] = [];
    for (const wb of sourceWbs) {
      const parsed = parseIncomeDetailRows(wb);
      if (parsed) allIncomeRowsForTabs.push(...parsed);
    }
    if (allIncomeRowsForTabs.length > 0) {
      const MEMBER_TABS = ["Flex", "Mail Members", "Daily Space"];
      for (const tabName of MEMBER_TABS) {
        const templateTabName = templateWb.SheetNames.find((n) => n.trim() === tabName);
        if (!templateTabName) continue;
        const currentSheet = sheetMap.get(templateTabName) ?? cloneSheet(templateWb.Sheets[templateTabName]);
        const { sheet, injected, chargeCodesFound } = injectIncomeDetailIntoTab(currentSheet, allIncomeRowsForTabs);
        if (injected > 0) {
          sheetMap.set(templateTabName, sheet);
          const idx = tabsPreserved.indexOf(templateTabName);
          if (idx >= 0) tabsPreserved.splice(idx, 1);
          if (!tabsRefreshed.includes(templateTabName)) tabsRefreshed.push(templateTabName);
          const srIdx = sectionResults.findIndex((s) => s.section === templateTabName);
          const note  = `${injected} rows injected (${chargeCodesFound} charge code type(s) from template)`;
          if (srIdx >= 0) sectionResults[srIdx] = { section: tabName, status: "hydrated", note };
          else sectionResults.push({ section: tabName, status: "hydrated", note });
        }
      }
    }
  }

  // ── Build output workbook in exact template tab order ────────────────────────
  const outWb = XLSX.utils.book_new();
  for (const name of templateWb.SheetNames) {
    const sheet = sheetMap.get(name);
    if (sheet) XLSX.utils.book_append_sheet(outWb, sheet, name);
  }

  // ── Write output ─────────────────────────────────────────────────────────────
  applyRecalcFlags(outWb);
  let rawBuffer = Buffer.from(XLSX.write(outWb, { type: "buffer", bookType: "xlsx" }));

  const templateBuf = Buffer.isBuffer(templateBuffer)
    ? templateBuffer
    : Buffer.from(templateBuffer as ArrayBuffer);

  const { buffer: pivotBuffer, report: pivotReport } = injectPivotParts(templateBuf, rawBuffer);
  rawBuffer = pivotBuffer;

  // Copy xl/styles.xml and xl/theme/theme1.xml from the template into the output.
  // This restores all conditional formatting, named styles, theme colours, and
  // bolding that xlsx-js-style strips when it rebuilds the style table from scratch.
  rawBuffer = preserveTemplateStyles(templateBuf, rawBuffer);

  // ── Reconciliation (only when full source data is available) ─────────────────
  let reconciliation: ReconciliationReport;
  if (sales && occupancyStats) {
    const allSnapshots = [
      ...(resolvedSnapshots ?? []),
      ...((resolvedSnapshots ?? []).some((s) => s.monthCode === currentSnapshot.monthCode)
        ? [] : [currentSnapshot]),
    ];
    reconciliation = reconcileOutput(
      rawBuffer, sales, occupancyStats, currentSnapshot,
      allSnapshots, budgetSource, "hydration",
    );
  } else {
    reconciliation = emptyReconciliation(
      sales
        ? "Reconciliation skipped — occupancy data not available in source"
        : occupancyStats
        ? "Reconciliation skipped — Sales By Month not available in source"
        : "Reconciliation skipped — source file does not contain Sales By Month or occupancy data",
    );
  }

  return {
    buffer: rawBuffer,
    snapshot: currentSnapshot,
    historicalParsed,
    tabsRefreshed,
    tabsPreserved,
    warnings,
    reconciliation,
    pivotReport,
    sectionResults,
  };
}
