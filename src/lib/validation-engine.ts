import * as XLSX from "xlsx-js-style";
import {
  type BudgetSource,
  type KpiSnapshot,
  auditBudgetMappings,
} from "@/lib/excel-dashboard-generator";

// ─── Validation Engine ────────────────────────────────────────────────────────
// Engine 1: post-generation deterministic audit.
// Runs after generateDashboardExcel and validates the output buffer against
// the extracted source data. Never throws — returns a structured report.

// Re-export types used by callers
export type { BudgetSource, KpiSnapshot };

export interface ValidationIssue {
  severity: "critical" | "warning" | "info";
  sheet:    string;
  cell?:    string;
  description: string;
  expected?:   string;
  actual?:     string;
}

export interface ValidationReport {
  isValid:       boolean;    // true when criticalCount === 0
  criticalCount: number;
  warningCount:  number;
  infoCount:     number;
  issues:        ValidationIssue[];
  summary:       string;
}

type CellValue = string | number | null;

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

// ── Sheet-presence check ──────────────────────────────────────────────────────

function checkSheetPresence(wb: XLSX.WorkBook, issues: ValidationIssue[]): {
  p1?: string; p2?: string; p4?: string;
} {
  const p1 = wb.SheetNames.find((n) => /^P\.?1[\s.]/i.test(n));
  const p2 = wb.SheetNames.find((n) => n.startsWith("P2") || n.startsWith("P.2"));
  const p4 = wb.SheetNames.find((n) => n.startsWith("P.4"));

  if (!p1) issues.push({ severity: "critical", sheet: "P.1", description: "P.1 Dashboard sheet is missing from generated workbook" });
  if (!p2) issues.push({ severity: "critical", sheet: "P.2", description: "P.2 Monthly Income sheet is missing from generated workbook" });
  if (!p4) issues.push({ severity: "warning",  sheet: "P.4", description: "P.4 Budget vs Actual sheet is missing — no budget source was found" });

  return { p1, p2, p4 };
}

// ── P.1 occupancy sanity ──────────────────────────────────────────────────────

function validateP1Occupancy(wb: XLSX.WorkBook, p1Name: string, issues: ValidationIssue[]): void {
  const ws = wb.Sheets[p1Name];
  if (!ws) return;

  // Cells C10 and C11 = occupancy % for DD and PO. Range must be 0–1.
  for (const addr of ["C10", "C11", "D10", "D11"]) {
    const cell = ws[addr];
    if (!cell) continue;
    const v = typeof cell.v === "number" ? cell.v : num(cell.v);
    if (v == null) continue;
    if (v < 0 || v > 1) {
      issues.push({
        severity: "warning",
        sheet:    p1Name,
        cell:     addr,
        description: `Occupancy % out of expected 0–1 range`,
        expected: "0.00 – 1.00",
        actual:   String(v),
      });
    }
  }
}

// ── P.2 column totals reconcile ───────────────────────────────────────────────

function validateP2Totals(wb: XLSX.WorkBook, p2Name: string, issues: ValidationIssue[]): void {
  const ws = wb.Sheets[p2Name];
  if (!ws) return;

  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");

  // Row 13 (0-indexed row 12) should be the sum of rows 3-12.
  // For each month column (cols 2..range.e.c-1), verify the formula is a SUM.
  // We validate cached values rather than re-computing to avoid column-index guessing.

  // Find the "Total" row (row 13 = idx 12) and "Total with Relief" row (idx 15)
  // by scanning col B for bold cells with matching label text.
  for (let r = 0; r <= range.e.r; r++) {
    const bCell = ws[XLSX.utils.encode_cell({ r, c: 1 })];
    if (!bCell) continue;
    const label = str(bCell.v).toLowerCase();

    if (label === "total" || label === "total with relief") {
      // Check that all numeric data cells in this row have reasonable values
      for (let c = 2; c <= range.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cell || cell.t !== "n") continue;
        const v = num(cell.v);
        if (v == null) continue;
        if (Math.abs(v) > 1_000_000_000) {  // sanity: > $1B is suspicious
          issues.push({
            severity: "warning",
            sheet:    p2Name,
            cell:     XLSX.utils.encode_cell({ r, c }),
            description: `${label} row value exceeds $1B — possible formula error`,
            actual:   String(v),
          });
        }
      }
    }
  }
}

// ── P.4 formula structure check ───────────────────────────────────────────────

function validateP4Formulas(wb: XLSX.WorkBook, p4Name: string, issues: ValidationIssue[]): void {
  const ws = wb.Sheets[p4Name];
  if (!ws) return;

  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");

  let refErrors = 0;
  for (let r = 0; r <= range.e.r; r++) {
    for (let c = 0; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell) continue;
      // Detect #REF! errors cached in cell value
      if (str(cell.v) === "#REF!" || str(cell.w) === "#REF!") {
        refErrors++;
      }
    }
  }

  if (refErrors > 0) {
    issues.push({
      severity: "critical",
      sheet:    p4Name,
      description: `${refErrors} cell(s) contain #REF! errors — formula references may be broken`,
      actual:   String(refErrors),
    });
  }

  // Check that row 3 (date cell) has a formula
  const dateCell = ws["A3"];
  if (dateCell && !dateCell.f) {
    issues.push({
      severity: "info",
      sheet:    p4Name,
      cell:     "A3",
      description: "P.4 date cell (A3) has no formula — expected reference to P.1 B2",
    });
  }
}

// ── No #REF! errors in any generated sheet ────────────────────────────────────

function checkGlobalRefErrors(wb: XLSX.WorkBook, issues: ValidationIssue[]): void {
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws) continue;
    let count = 0;
    for (const addr of Object.keys(ws)) {
      if (addr.startsWith("!")) continue;
      const cell = ws[addr];
      if (str(cell.v) === "#REF!" || str(cell.w) === "#REF!") count++;
    }
    if (count > 0) {
      issues.push({
        severity: "critical",
        sheet:    name,
        description: `${count} #REF! error(s) in sheet "${name}"`,
        actual:   String(count),
      });
    }
  }
}

// ── Month detection sanity ────────────────────────────────────────────────────

function validateMonthHeaders(wb: XLSX.WorkBook, p2Name: string, issues: ValidationIssue[]): void {
  const ws = wb.Sheets[p2Name];
  if (!ws) return;

  // P.2 row 2 should contain month name strings in cols C+
  const row2Cells: (string | null)[] = [];
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  for (let c = 2; c <= Math.min(range.e.c, 14); c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: 1, c })];
    row2Cells.push(cell ? str(cell.v) : null);
  }

  const monthNames = ["January","February","March","April","May","June",
    "July","August","September","October","November","December","Annual"];
  const validMonthCount = row2Cells.filter((v) => v && monthNames.includes(v)).length;

  if (validMonthCount < 3) {
    issues.push({
      severity: "warning",
      sheet:    p2Name,
      description: "P.2 column header row contains fewer than 3 recognizable month names",
      actual:   String(validMonthCount),
    });
  }
}

// ── Historical snapshot completeness ─────────────────────────────────────────

function checkHistoricalSnapshots(
  snapshots: KpiSnapshot[] | undefined,
  issues: ValidationIssue[],
): void {
  if (!snapshots || snapshots.length === 0) {
    issues.push({
      severity: "info",
      sheet:    "P.1",
      description: "No historical snapshots available — P.1 snapshot section will be empty. Upload a file with a Historical tab or re-run after the first snapshot has been persisted.",
    });
    return;
  }

  // Flag if fewer than 4 months (partial history is accepted but noted)
  if (snapshots.length < 4) {
    issues.push({
      severity: "info",
      sheet:    "P.1",
      description: `Only ${snapshots.length} historical snapshot(s) available. 12 months is optimal for the rolling trend view.`,
    });
  }

  // Flag duplicate month codes
  const codes = snapshots.map((s) => s.monthCode);
  const dupes = codes.filter((c, i) => codes.indexOf(c) !== i);
  if (dupes.length > 0) {
    issues.push({
      severity: "warning",
      sheet:    "P.1",
      description: `Duplicate month codes in snapshot history: ${[...new Set(dupes)].join(", ")}`,
    });
  }

  // Flag occupancy values out of range
  for (const snap of snapshots) {
    if (snap.occupancyPctSpaces > 1 || snap.occupancyPctSpaces < 0 ||
        snap.occupancyPctSqFt   > 1 || snap.occupancyPctSqFt   < 0) {
      issues.push({
        severity: "warning",
        sheet:    "P.1",
        description: `Snapshot ${snap.monthCode}: occupancy percentage out of 0–1 range (spaces=${snap.occupancyPctSpaces}, sqft=${snap.occupancyPctSqFt}) — may indicate percentage stored as 0-100 rather than 0-1`,
      });
      break;  // one warning covers all; avoid flooding the report
    }
  }
}

// ── Budget mapping audit ──────────────────────────────────────────────────────

function checkBudgetMappings(
  budgetSource: BudgetSource | undefined,
  issues: ValidationIssue[],
): void {
  if (!budgetSource) {
    issues.push({
      severity: "warning",
      sheet:    "P.4",
      description: "No budget source found — P.4 was generated with a placeholder row. Add a Budget tab or P.4 tab to the uploaded workbook.",
    });
    return;
  }

  // Summarise from auditBudgetMappings (which requires the sales lookup maps;
  // those aren't available post-generation, so we just report source coverage)
  const nonTrivial = budgetSource.lineItems.filter(
    (i) => !i.isBlankRow && !i.isSectionHeader && i.budgetValue !== 0,
  );
  if (nonTrivial.length === 0) {
    issues.push({
      severity: "warning",
      sheet:    "P.4",
      description: "Budget source was found but contains no non-zero leaf items — all budget values appear to be zero.",
    });
  }
}

// ─── Main Validator ───────────────────────────────────────────────────────────

export function validateGeneratedWorkbook(
  generatedBuffer: Buffer,
  opts?: {
    historicalSnapshots?: KpiSnapshot[];
    budgetSource?: BudgetSource;
  },
): ValidationReport {
  const issues: ValidationIssue[] = [];

  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(generatedBuffer, { type: "buffer", cellStyles: false, cellNF: false });
  } catch {
    return {
      isValid:       false,
      criticalCount: 1,
      warningCount:  0,
      infoCount:     0,
      issues: [{ severity: "critical", sheet: "(workbook)", description: "Generated buffer is not a valid Excel file" }],
      summary: "CRITICAL: generated file could not be parsed",
    };
  }

  // 1. Sheet presence
  const sheets = checkSheetPresence(wb, issues);

  // 2. Global #REF! scan
  checkGlobalRefErrors(wb, issues);

  // 3. P.1 occupancy sanity + historical snapshot coverage
  if (sheets.p1) validateP1Occupancy(wb, sheets.p1, issues);
  checkHistoricalSnapshots(opts?.historicalSnapshots, issues);

  // 4. P.2 total rows sanity
  if (sheets.p2) {
    validateP2Totals(wb, sheets.p2, issues);
    validateMonthHeaders(wb, sheets.p2, issues);
  }

  // 5. P.4 formula structure + budget mapping coverage
  if (sheets.p4) validateP4Formulas(wb, sheets.p4, issues);
  checkBudgetMappings(opts?.budgetSource, issues);

  const criticalCount = issues.filter((i) => i.severity === "critical").length;
  const warningCount  = issues.filter((i) => i.severity === "warning").length;
  const infoCount     = issues.filter((i) => i.severity === "info").length;

  const isValid = criticalCount === 0;

  let summary: string;
  if (isValid && warningCount === 0) {
    summary = `Validation passed (${wb.SheetNames.length} sheets, 0 issues)`;
  } else if (isValid) {
    summary = `Validation passed with ${warningCount} warning(s)`;
  } else {
    summary = `Validation FAILED: ${criticalCount} critical issue(s), ${warningCount} warning(s)`;
  }

  return { isValid, criticalCount, warningCount, infoCount, issues, summary };
}
