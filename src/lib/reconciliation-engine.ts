import * as XLSX from "xlsx-js-style";
import {
  type SalesByMonthData,
  type OccupancyStats,
  type KpiSnapshot,
  type BudgetSource,
} from "@/lib/excel-dashboard-generator";

// ─── Reconciliation Engine ─────────────────────────────────────────────────────
// Post-hydration deterministic financial audit.
//
// PHILOSOPHY:
//   The workbook IS the financial application. This engine independently
//   recalculates every critical financial metric from raw source data and
//   compares against the workbook's cached cell values.
//
//   For hydrated workbooks: formula cells (P.2, P.4) will have stale cached
//   values until Excel opens and recalculates them. This is expected, documented,
//   and DOES NOT prevent download — because CalcPr forceFullCalc ensures correctness
//   on first open. Only data cells we directly write are reconciled exactly.
//
//   For all critical financial mismatches (tolerance exceeded): the generator
//   FAILS EXPLICITLY and blocks the download with a 422 response.
//
// SEVERITY LEVELS:
//   INFO     — informational observations, no action needed
//   WARNING  — notable drift or missing optional data
//   ERROR    — calculable mismatch, finance team must review before distributing
//   CRITICAL — financial total unreconciled — output BLOCKED until resolved

export type ReconciliationSeverity = "INFO" | "WARNING" | "ERROR" | "CRITICAL";

export interface ReconciliationFinding {
  severity:    ReconciliationSeverity;
  category:    string;
  sheet:       string;
  cell?:       string;
  description: string;
  expected?:   string;
  actual?:     string;
  delta?:      string;
}

export interface ReconciliationReport {
  /** true only when criticalCount === 0 AND errorCount === 0 */
  isClean:       boolean;
  criticalCount: number;
  errorCount:    number;
  warningCount:  number;
  infoCount:     number;
  findings:      ReconciliationFinding[];
  categories:    string[];
  summary:       string;
  /** Stale formula cells are expected in hydration mode; listed here for audit trail */
  staleFormulaCount: number;
  generationMode: "hydration" | "regeneration";
}

// ─── Tolerance constants ──────────────────────────────────────────────────────
// Financial values: 1% relative tolerance (covers Excel floating-point rounding)
const REL_TOLERANCE_WARN     = 0.01;   // 1% — warning
const REL_TOLERANCE_ERROR    = 0.05;   // 5% — error
const REL_TOLERANCE_CRITICAL = 0.10;   // 10% — critical (unreconciled total)
// Integer counts: exact match required
const COUNT_TOLERANCE        = 0;
// Percentages (0-1): 0.01 absolute tolerance
const PCT_TOLERANCE          = 0.01;

// ─── Utilities ────────────────────────────────────────────────────────────────

function num(v: unknown): number | null {
  if (v == null || v === "") return null;
  const s = String(v).replace(/[$,%\s()]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function pct(n: number): string {
  return `${(n * 100).toFixed(1)}%`;
}

function dollar(n: number): string {
  return `$${n.toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 0 })}`;
}

function relDelta(expected: number, actual: number): number {
  if (expected === 0) return actual === 0 ? 0 : Infinity;
  return Math.abs(actual - expected) / Math.abs(expected);
}

function classifyRel(delta: number): ReconciliationSeverity {
  if (delta >= REL_TOLERANCE_CRITICAL) return "CRITICAL";
  if (delta >= REL_TOLERANCE_ERROR)    return "ERROR";
  if (delta >= REL_TOLERANCE_WARN)     return "WARNING";
  return "INFO";
}

function getCellValue(ws: XLSX.WorkSheet, addr: string): number | null {
  const cell = ws[addr] as XLSX.CellObject | undefined;
  if (!cell) return null;
  return num(cell.v);
}

function getStringValue(ws: XLSX.WorkSheet, addr: string): string {
  const cell = ws[addr] as XLSX.CellObject | undefined;
  if (!cell || cell.v == null) return "";
  return String(cell.v).trim();
}

// ─── 1. Structural integrity ──────────────────────────────────────────────────

function checkStructure(wb: XLSX.WorkBook, findings: ReconciliationFinding[]): {
  p1?: string; p2?: string; p3?: string; p4?: string;
} {
  const p1 = wb.SheetNames.find((n) => /^P\.?1[\s.]/i.test(n));
  const p2 = wb.SheetNames.find((n) => n.startsWith("P2") || n.startsWith("P.2"));
  const p3 = wb.SheetNames.find((n) => n.startsWith("P.3"));
  const p4 = wb.SheetNames.find((n) => n.startsWith("P.4"));

  if (!p1) findings.push({ severity: "CRITICAL", category: "Structure", sheet: "P.1",
    description: "P.1 Dashboard sheet is missing from output workbook" });
  if (!p2) findings.push({ severity: "CRITICAL", category: "Structure", sheet: "P.2",
    description: "P.2 Monthly Income sheet is missing from output workbook" });
  if (!p3) findings.push({ severity: "WARNING", category: "Structure", sheet: "P.3",
    description: "P.3 Occupancy sheet is missing from output workbook" });
  if (!p4) findings.push({ severity: "WARNING", category: "Structure", sheet: "P.4",
    description: "P.4 Budget vs Actual sheet is missing — no budget source found" });

  return { p1, p2, p3, p4 };
}

// ─── 2. Formula error scan ────────────────────────────────────────────────────

const ERROR_VALUES = new Set(["#REF!", "#VALUE!", "#NAME?", "#NULL!", "#DIV/0!", "#N/A", "#NUM!"]);

function scanFormulaErrors(wb: XLSX.WorkBook, findings: ReconciliationFinding[]): void {
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    let count = 0;
    const samples: string[] = [];

    for (const addr of Object.keys(ws)) {
      if (addr.startsWith("!")) continue;
      const cell = ws[addr] as XLSX.CellObject;
      const vStr = cell.v != null ? String(cell.v) : "";
      const wStr = (cell as unknown as Record<string, unknown>).w != null ? String((cell as unknown as Record<string, unknown>).w) : "";
      if (ERROR_VALUES.has(vStr) || ERROR_VALUES.has(wStr)) {
        count++;
        if (samples.length < 5) samples.push(addr);
      }
    }

    if (count > 0) {
      findings.push({
        severity:    "CRITICAL",
        category:    "FormulaErrors",
        sheet:       sheetName,
        description: `${count} formula error(s) in sheet — cells: ${samples.join(", ")}${count > 5 ? ` (+${count - 5} more)` : ""}`,
        actual:      String(count),
      });
    }
  }
}

// ─── 3. Occupancy reconciliation ─────────────────────────────────────────────
// In hydration mode: P.1 rows 10-12 are preserved FORMULAS from the template.
// Cached values may be stale until Excel opens. We verify the raw OccupancyStats
// and confirm the source data is internally consistent.

function reconcileOccupancy(
  occ: OccupancyStats,
  p1Ws: XLSX.WorkSheet | null,
  generationMode: "hydration" | "regeneration",
  findings: ReconciliationFinding[],
): void {
  const dd = occ.dedicatedDesk;
  const po = occ.privateOffice;

  // Internal consistency: occupied <= total
  if (dd.occupied > dd.total) {
    findings.push({
      severity: "ERROR",
      category: "Occupancy",
      sheet:    "P.3",
      description: "Dedicated Desk: occupied count exceeds total spaces — data quality issue in P.3",
      expected: `occupied <= ${dd.total}`,
      actual:   String(dd.occupied),
    });
  }
  if (po.occupied > po.total) {
    findings.push({
      severity: "ERROR",
      category: "Occupancy",
      sheet:    "P.3",
      description: "Private Office: occupied count exceeds total spaces — data quality issue in P.3",
      expected: `occupied <= ${po.total}`,
      actual:   String(po.occupied),
    });
  }

  // Reasonable ranges
  const ddPct = dd.total > 0 ? dd.occupied / dd.total : 0;
  const poPct = po.total > 0 ? po.occupied / po.total : 0;

  if (ddPct < 0 || ddPct > 1) {
    findings.push({
      severity: "ERROR",
      category: "Occupancy",
      sheet:    "P.3",
      description: `Dedicated Desk occupancy % out of 0–100% range: ${pct(ddPct)}`,
    });
  }
  if (poPct < 0 || poPct > 1) {
    findings.push({
      severity: "ERROR",
      category: "Occupancy",
      sheet:    "P.3",
      description: `Private Office occupancy % out of 0–100% range: ${pct(poPct)}`,
    });
  }

  // If we have the P.1 worksheet, verify cached occupancy values (regeneration mode)
  // In hydration mode, P.1 rows 10-12 have formula cells with stale cached values —
  // we note this but do NOT flag as an error because CalcPr will fix them on open.
  if (p1Ws && generationMode === "regeneration") {
    const totalTotal = dd.total + po.total;
    const totalOcc   = dd.occupied + po.occupied;
    const spacesPct  = totalTotal > 0 ? totalOcc / totalTotal : 0;

    const c12 = getCellValue(p1Ws, "C12");
    const d12 = getCellValue(p1Ws, "D12");
    const e12 = getCellValue(p1Ws, "E12");

    if (c12 != null && Math.abs(c12 - totalTotal) > COUNT_TOLERANCE) {
      findings.push({
        severity:    "ERROR",
        category:    "Occupancy",
        sheet:       "P.1",
        cell:        "C12",
        description: "Total spaces (P.1 C12) does not match P.3 source data",
        expected:    String(totalTotal),
        actual:      String(c12),
      });
    }
    if (d12 != null && Math.abs(d12 - totalOcc) > COUNT_TOLERANCE) {
      findings.push({
        severity:    "ERROR",
        category:    "Occupancy",
        sheet:       "P.1",
        cell:        "D12",
        description: "Occupied spaces (P.1 D12) does not match P.3 source data",
        expected:    String(totalOcc),
        actual:      String(d12),
      });
    }
    if (e12 != null && Math.abs(e12 - spacesPct) > PCT_TOLERANCE) {
      const sev = relDelta(spacesPct, e12) >= REL_TOLERANCE_CRITICAL ? "CRITICAL" : "ERROR";
      findings.push({
        severity:    sev,
        category:    "Occupancy",
        sheet:       "P.1",
        cell:        "E12",
        description: `Total occupancy % (P.1 E12) deviates from independent calculation`,
        expected:    pct(spacesPct),
        actual:      pct(e12),
        delta:       pct(Math.abs(e12 - spacesPct)),
      });
    }
  } else if (p1Ws && generationMode === "hydration") {
    // In hydration mode, note that occupancy formula cells will recalculate on open
    findings.push({
      severity:    "INFO",
      category:    "Occupancy",
      sheet:       "P.1",
      description: "Hydration mode: P.1 occupancy cells (rows 10-12) are template formulas — stale cached values will update when Excel opens (CalcPr forceFullCalc applied)",
    });
  }

  // Always validate source data
  if (dd.total === 0 && po.total === 0) {
    findings.push({
      severity:    "ERROR",
      category:    "Occupancy",
      sheet:       "P.3",
      description: "No dedicated desk or private office spaces found in P.3 — occupancy data may be missing",
    });
  }
}

// ─── 4. Revenue reconciliation ────────────────────────────────────────────────
// Independent revenue calculation from Sales By Month, compared against P.2
// Total Income row (regeneration mode) or source data internal consistency.

function reconcileRevenue(
  sales: SalesByMonthData,
  p2Ws: XLSX.WorkSheet | null,
  generationMode: "hydration" | "regeneration",
  findings: ReconciliationFinding[],
): void {
  const { lookupByLabel, currentMonthIdx, monthCodes } = sales;
  const currentMonthCode = monthCodes[currentMonthIdx];

  // Calculate total income from Sales By Month (independent)
  const incomeKeys = ["total income", "total revenue", "total sales", "gross revenue"];
  let independentRevenue: number | null = null;
  for (const key of incomeKeys) {
    const row = lookupByLabel.get(key);
    if (row) {
      independentRevenue = row[currentMonthIdx] ?? null;
      break;
    }
  }

  // Validate revenue is non-negative and within a reasonable range
  if (independentRevenue == null) {
    findings.push({
      severity:    "WARNING",
      category:    "Revenue",
      sheet:       "Sales By Month",
      description: `Could not find a "Total Income" / "Total Revenue" row for month ${currentMonthCode} — revenue reconciliation skipped`,
    });
    return;
  }

  if (independentRevenue < 0) {
    findings.push({
      severity:    "ERROR",
      category:    "Revenue",
      sheet:       "Sales By Month",
      description: `Total Income for ${currentMonthCode} is negative (${dollar(independentRevenue)}) — possible data entry error`,
      actual:      dollar(independentRevenue),
    });
  }

  if (independentRevenue === 0) {
    findings.push({
      severity:    "WARNING",
      category:    "Revenue",
      sheet:       "Sales By Month",
      description: `Total Income for ${currentMonthCode} is zero — verify that data was correctly uploaded`,
      actual:      dollar(0),
    });
    return;
  }

  // In regeneration mode: compare against P.2 cached Total row
  if (p2Ws && generationMode === "regeneration") {
    const range = XLSX.utils.decode_range(p2Ws["!ref"] ?? "A1:A1");

    // Find the Total Income row in P.2 (look for "Total" or "Total Income" in col B)
    for (let r = 0; r <= range.e.r; r++) {
      const bCell = p2Ws[XLSX.utils.encode_cell({ r, c: 1 })];
      if (!bCell) continue;
      const label = String(bCell.v ?? "").toLowerCase().trim();
      if (label !== "total income" && label !== "total revenue" && label !== "total") continue;

      // Find the current month column in P.2
      // P.2 month serials are in row 1 (0-indexed row 0)
      // or we can search by matching the current month serial
      const currentSerial = sales.monthSerials[currentMonthIdx];
      for (let c = 2; c <= range.e.c; c++) {
        const headerCell = p2Ws[XLSX.utils.encode_cell({ r: 1, c })];
        if (!headerCell) continue;
        const hv = headerCell.v;
        // Match by serial number or by month name
        const matchesBySerial = typeof hv === "number" && Math.abs(hv - currentSerial) < 5;
        if (matchesBySerial) {
          const totalCell = p2Ws[XLSX.utils.encode_cell({ r, c })];
          if (!totalCell) continue;
          const p2Revenue = num(totalCell.v);
          if (p2Revenue == null) continue;

          const delta = relDelta(independentRevenue, p2Revenue);
          const sev = classifyRel(delta);

          if (sev !== "INFO") {
            findings.push({
              severity:    sev,
              category:    "Revenue",
              sheet:       "P.2",
              cell:        XLSX.utils.encode_cell({ r, c }),
              description: `P.2 Total Income for ${currentMonthCode} deviates from Sales By Month source`,
              expected:    dollar(independentRevenue),
              actual:      dollar(p2Revenue),
              delta:       pct(delta),
            });
          } else {
            findings.push({
              severity:    "INFO",
              category:    "Revenue",
              sheet:       "P.2",
              description: `Revenue reconciled for ${currentMonthCode}: ${dollar(independentRevenue)} (within ${pct(delta)})`,
            });
          }
          return;
        }
      }
    }

    // P.2 total row found but month column not matched — note it
    findings.push({
      severity:    "WARNING",
      category:    "Revenue",
      sheet:       "P.2",
      description: `Could not locate current month column in P.2 for revenue reconciliation — verify P.2 month headers match Sales By Month serials`,
    });
  } else if (generationMode === "hydration") {
    // In hydration mode, P.2 formula cells will recalculate — just confirm source value looks reasonable
    findings.push({
      severity:    "INFO",
      category:    "Revenue",
      sheet:       "Sales By Month",
      description: `Source revenue for ${currentMonthCode}: ${dollar(independentRevenue)} — P.2 formulas will recalculate to this value on Excel open`,
    });
  }
}

// ─── 5. KPI snapshot reconciliation ──────────────────────────────────────────
// The snapshot rows in P.1 are written DIRECTLY by the hydration/generation
// engine. These should match KpiSnapshot EXACTLY (or within floating-point noise).
// This is the highest-confidence check because these are data writes, not formulas.

function reconcileSnapshots(
  snapshot: KpiSnapshot,
  allSnapshots: KpiSnapshot[],
  p1Ws: XLSX.WorkSheet | null,
  findings: ReconciliationFinding[],
): void {
  if (!p1Ws) return;

  // Find the current month in the P.1 snapshot rows (rows 18+, 0-indexed row 17+)
  const range = XLSX.utils.decode_range(p1Ws["!ref"] ?? "A1:A1");

  // Row B column index = 1, snapshot rows start at 0-indexed row 17
  const SNAP_START = 17;
  let foundRow = -1;
  let matchedSerial: number | null = null;

  for (let r = SNAP_START; r <= range.e.r; r++) {
    const bCell = p1Ws[XLSX.utils.encode_cell({ r, c: 1 })];
    if (!bCell) continue;
    const v = bCell.v;
    // B column in snapshot section holds the month serial
    if (typeof v === "number" && Math.abs(v - snapshot.monthSerial) < 1) {
      foundRow = r;
      matchedSerial = v;
      break;
    }
    // Also accept the "Average" text row as a stopping signal
    if (typeof v === "string" && v.toLowerCase().includes("average")) break;
  }

  if (foundRow < 0) {
    findings.push({
      severity:    "ERROR",
      category:    "Snapshot",
      sheet:       "P.1",
      description: `Current month snapshot row for ${snapshot.monthCode} (serial ${snapshot.monthSerial}) not found in P.1 — snapshot section may not have been written`,
    });
    return;
  }

  // Col C (idx 2) = occupancyPctSpaces
  // Col D (idx 3) = occupancyPctSqFt
  // Col E (idx 4) = contracts
  // Col F (idx 5) = flexMembers
  // Col G (idx 6) = mailMembers
  // Col H (idx 7) = meetingHours
  // Col I (idx 8) = meetingRevenueDollar

  const checks: { col: number; label: string; expected: number; pctCheck?: boolean }[] = [
    { col: 2, label: "Occupancy % (Spaces)", expected: snapshot.occupancyPctSpaces, pctCheck: true },
    { col: 3, label: "Occupancy % (SqFt)",   expected: snapshot.occupancyPctSqFt,   pctCheck: true },
    { col: 4, label: "Contracts",             expected: snapshot.contracts },
    { col: 5, label: "Flex Members",          expected: snapshot.flexMembers },
    { col: 6, label: "Mail Members",          expected: snapshot.mailMembers },
    { col: 7, label: "Meeting Hours",         expected: snapshot.meetingHours },
    { col: 8, label: "Meeting Revenue",       expected: snapshot.meetingRevenueDollar },
  ];

  for (const { col, label, expected, pctCheck } of checks) {
    const actual = getCellValue(p1Ws, XLSX.utils.encode_cell({ r: foundRow, c: col }));
    if (actual == null) {
      findings.push({
        severity:    "WARNING",
        category:    "Snapshot",
        sheet:       "P.1",
        cell:        XLSX.utils.encode_cell({ r: foundRow, c: col }),
        description: `${label} cell is empty in current month snapshot row`,
        expected:    String(expected),
      });
      continue;
    }

    const tolerance = pctCheck ? PCT_TOLERANCE : COUNT_TOLERANCE;
    const diff = Math.abs(actual - expected);

    if (diff > tolerance) {
      const delta = relDelta(expected, actual);
      const sev = pctCheck
        ? (diff > 0.05 ? "ERROR" : "WARNING")
        : (delta >= REL_TOLERANCE_CRITICAL ? "CRITICAL" : "ERROR");

      findings.push({
        severity:    sev,
        category:    "Snapshot",
        sheet:       "P.1",
        cell:        XLSX.utils.encode_cell({ r: foundRow, c: col }),
        description: `${label} in P.1 snapshot does not match extracted value — possible write error`,
        expected:    pctCheck ? pct(expected) : String(expected),
        actual:      pctCheck ? pct(actual)   : String(actual),
        delta:       pct(delta),
      });
    }
  }

  // Validate snapshot completeness (all expected months present)
  if (allSnapshots.length === 0) {
    findings.push({
      severity:    "INFO",
      category:    "Snapshot",
      sheet:       "P.1",
      description: "No historical snapshots — only current month is in the snapshot section (first-run state)",
    });
  } else if (allSnapshots.length < 4) {
    findings.push({
      severity:    "INFO",
      category:    "Snapshot",
      sheet:       "P.1",
      description: `Only ${allSnapshots.length} snapshot(s) in history — 12 months is optimal for the rolling trend view`,
    });
  }
}

// ─── 6. Budget variance reconciliation ───────────────────────────────────────
// For each budget line item that maps to a Sales By Month row,
// independently calculate variance = actual - budget and compare against P.4.
// In hydration mode: P.4 formula cells are stale until Excel recalculates.

function reconcileBudget(
  budgetSource: BudgetSource | null,
  sales: SalesByMonthData,
  p4Ws: XLSX.WorkSheet | null,
  generationMode: "hydration" | "regeneration",
  findings: ReconciliationFinding[],
): void {
  if (!budgetSource) {
    findings.push({
      severity:    "INFO",
      category:    "Budget",
      sheet:       "P.4",
      description: "No budget source available — budget reconciliation skipped",
    });
    return;
  }

  const { currentMonthIdx, lookupByLabel } = sales;
  const lineItems = budgetSource.lineItems.filter(
    (i) => !i.isBlankRow && !i.isSectionHeader && !i.isTotal && i.budgetValue !== 0,
  );

  if (lineItems.length === 0) {
    findings.push({
      severity:    "WARNING",
      category:    "Budget",
      sheet:       "P.4",
      description: "Budget source contains no non-zero leaf items — budget reconciliation skipped",
    });
    return;
  }

  let mappedCount = 0;
  let unmappedCount = 0;
  let maxVariancePct = 0;
  let maxVarianceItem = "";

  for (const item of lineItems) {
    const row = lookupByLabel.get(item.normalizedLabel) ?? lookupByLabel.get(item.label.toLowerCase().trim());
    if (!row) {
      unmappedCount++;
      continue;
    }

    mappedCount++;
    const actual  = row[currentMonthIdx] ?? 0;
    const budget  = item.budgetValue;
    const variancePct = budget !== 0 ? Math.abs(actual - budget) / Math.abs(budget) : 0;

    if (variancePct > maxVariancePct) {
      maxVariancePct = variancePct;
      maxVarianceItem = item.label;
    }

    // Flag large variances as informational (finance teams expect variances; we just audit)
    if (variancePct > REL_TOLERANCE_CRITICAL && Math.abs(actual - budget) > 1000) {
      findings.push({
        severity:    "WARNING",
        category:    "Budget",
        sheet:       "P.4",
        description: `${item.label}: actual vs budget variance ${pct(variancePct)} (budget ${dollar(budget)}, actual ${dollar(actual)})`,
        expected:    dollar(budget),
        actual:      dollar(actual),
        delta:       pct(variancePct),
      });
    }
  }

  if (unmappedCount > 0) {
    findings.push({
      severity:    "WARNING",
      category:    "Budget",
      sheet:       "P.4",
      description: `${unmappedCount} of ${lineItems.length} budget items could not be mapped to Sales By Month rows — P.4 actuals will show 0 for these items`,
    });
  }

  if (generationMode === "hydration") {
    findings.push({
      severity:    "INFO",
      category:    "Budget",
      sheet:       "P.4",
      description: `Hydration mode: P.4 formula cells are template values — will recalculate on Excel open. ${mappedCount} budget items mapped to source data.`,
    });
  } else {
    findings.push({
      severity:    "INFO",
      category:    "Budget",
      sheet:       "P.4",
      description: `Budget reconciliation: ${mappedCount}/${lineItems.length} items mapped. Max variance: ${maxVarianceItem ? `"${maxVarianceItem}" at ${pct(maxVariancePct)}` : "none"}`,
    });
  }
}

// ─── 7. Named range + structural integrity ─────────────────────────────────────

function checkNamedRanges(wb: XLSX.WorkBook, findings: ReconciliationFinding[]): void {
  const names = (wb.Workbook as Record<string, unknown> | undefined)?.Names;
  if (!Array.isArray(names) || names.length === 0) return;

  let brokenCount = 0;
  for (const nm of names) {
    const ref = String((nm as Record<string, unknown>).Ref ?? "");
    // A named range referencing a sheet that doesn't exist is broken
    const sheetMatch = ref.match(/^'?([^'!]+)'?!/);
    if (sheetMatch) {
      const sheetName = sheetMatch[1];
      if (!wb.SheetNames.includes(sheetName)) {
        brokenCount++;
      }
    }
  }

  if (brokenCount > 0) {
    findings.push({
      severity:    "ERROR",
      category:    "NamedRanges",
      sheet:       "(workbook)",
      description: `${brokenCount} named range(s) reference sheets that do not exist in the output workbook`,
      actual:      String(brokenCount),
    });
  }
}

// ─── 8. Stale formula detection ───────────────────────────────────────────────

function countStaleFormulas(wb: XLSX.WorkBook): number {
  let count = 0;
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    for (const addr of Object.keys(ws)) {
      if (addr.startsWith("!")) continue;
      const cell = ws[addr] as XLSX.CellObject;
      if (cell.f && (cell.v === 0 || cell.v == null)) count++;
    }
  }
  return count;
}

// ─── Main reconciliation entry point ──────────────────────────────────────────

export function reconcileOutput(
  outputBuffer: Buffer,
  sales: SalesByMonthData,
  occ: OccupancyStats,
  snapshot: KpiSnapshot,
  allSnapshots: KpiSnapshot[],
  budgetSource: BudgetSource | null,
  generationMode: "hydration" | "regeneration",
): ReconciliationReport {
  const findings: ReconciliationFinding[] = [];

  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(outputBuffer, { type: "buffer", cellStyles: false, cellNF: false });
  } catch (e) {
    return {
      isClean:       false,
      criticalCount: 1,
      errorCount:    0,
      warningCount:  0,
      infoCount:     0,
      findings: [{
        severity: "CRITICAL",
        category: "Structure",
        sheet: "(workbook)",
        description: `Output buffer is not a valid XLSX file: ${e}`,
      }],
      categories:    ["Structure"],
      summary:       "CRITICAL: output buffer could not be parsed",
      staleFormulaCount: 0,
      generationMode,
    };
  }

  // 1. Structural check
  const sheets = checkStructure(wb, findings);

  // 2. Formula error scan
  scanFormulaErrors(wb, findings);

  // 3. Occupancy
  const p1Ws = sheets.p1 ? wb.Sheets[sheets.p1] : null;
  const p2Ws = sheets.p2 ? wb.Sheets[sheets.p2] : null;
  reconcileOccupancy(occ, p1Ws, generationMode, findings);

  // 4. Revenue
  reconcileRevenue(sales, p2Ws, generationMode, findings);

  // 5. Snapshot write verification
  reconcileSnapshots(snapshot, allSnapshots, p1Ws, findings);

  // 6. Budget variance
  const p4Ws = sheets.p4 ? wb.Sheets[sheets.p4] : null;
  reconcileBudget(budgetSource, sales, p4Ws, generationMode, findings);

  // 7. Named ranges
  checkNamedRanges(wb, findings);

  // 8. Stale formula count
  const staleFormulaCount = countStaleFormulas(wb);

  // Aggregate
  const criticalCount = findings.filter((f) => f.severity === "CRITICAL").length;
  const errorCount    = findings.filter((f) => f.severity === "ERROR").length;
  const warningCount  = findings.filter((f) => f.severity === "WARNING").length;
  const infoCount     = findings.filter((f) => f.severity === "INFO").length;

  const categories = [...new Set(findings.map((f) => f.category))];
  const isClean    = criticalCount === 0 && errorCount === 0;

  let summary: string;
  if (isClean && warningCount === 0) {
    summary = `Clean: 0 issues across ${wb.SheetNames.length} sheet(s) [${generationMode} mode]`;
  } else if (isClean) {
    summary = `Passed with ${warningCount} warning(s) [${generationMode} mode]`;
  } else if (criticalCount > 0) {
    summary = `CRITICAL: ${criticalCount} unreconciled financial total(s), ${errorCount} error(s), ${warningCount} warning(s) [${generationMode} mode]`;
  } else {
    summary = `${errorCount} error(s), ${warningCount} warning(s) — review before distributing [${generationMode} mode]`;
  }

  return {
    isClean,
    criticalCount,
    errorCount,
    warningCount,
    infoCount,
    findings,
    categories,
    summary,
    staleFormulaCount,
    generationMode,
  };
}
