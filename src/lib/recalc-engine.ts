import * as XLSX from "xlsx-js-style";

// ─── Recalculation Engine ─────────────────────────────────────────────────────
// Ensures every workbook produced by the hydration pipeline forces a full
// recalculation when opened in Excel, Excel 365, or Excel Online.
//
// Root cause without this: xlsx-js-style writes formula strings + last-cached
// values. If Excel opens the file and finds a valid calcChain, it may NOT
// recalculate all cells — so dependent sheets would show stale numbers.
//
// Solution:
//   1. Set <calcPr fullCalcOnLoad="1" forceFullCalc="1" calcMode="auto"/>
//      → Excel recalculates ALL formulas on open, unconditionally.
//   2. Don't write a calcChain.xml (xlsx-js-style never writes one, so
//      Excel also recalculates due to missing chain — belt-and-suspenders).
//   3. Scan every output cell for cached v=0 on a formula cell
//      → signals a possibly-stale calculation worth auditing.
//
// Excel compatibility:
//   - Excel for Windows/Mac:   fullCalcOnLoad + forceFullCalc are honoured
//   - Excel 365 (web):         fullCalcOnLoad is honoured
//   - Excel Online:            recalculates on every edit anyway
//   - Google Sheets import:    ignores calcPr, recalculates on import

export interface StaleFormulaRecord {
  sheet: string;
  addr:  string;
  formula: string;
  cachedValue: unknown;
}

export interface RecalcReport {
  forceFullCalcApplied: boolean;
  staleFormulaCandidates: StaleFormulaRecord[];
  sheetsScan: string[];
  summary: string;
}

/**
 * Set fullCalcOnLoad + forceFullCalc on the workbook.
 * Must be called on the output WorkBook object BEFORE XLSX.write().
 */
export function applyRecalcFlags(wb: XLSX.WorkBook): void {
  if (!wb.Workbook) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    (wb as any).Workbook = {};
  }
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  (wb.Workbook as any).CalcPr = {
    calcId:           191028,   // arbitrary but non-zero — Excel ignores this value
    fullCalcOnLoad:   true,     // recalculate all formulas on open
    forceFullCalc:    true,     // bypass Excel's optimisation that skips unchanged cells
    calcMode:         "auto",   // ensure auto-calc is active (not manual)
  };
}

/**
 * Scan all sheets for formula cells whose cached value is 0 (or null).
 * These are candidates for stale calculations — not necessarily wrong,
 * but worth including in the audit log.
 *
 * A cell is flagged when:
 *   - It has a formula (cell.f is set)
 *   - Its cached value (cell.v) is 0, null, or undefined
 *   - It is NOT in a row/column that legitimately holds 0
 *     (we can't know for sure — so we flag all of them for human review)
 */
export function detectStaleFormulas(wb: XLSX.WorkBook): RecalcReport {
  const stale: StaleFormulaRecord[] = [];
  const sheetsScan: string[] = [];

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;
    sheetsScan.push(sheetName);

    for (const addr of Object.keys(ws)) {
      if (addr.startsWith("!")) continue;
      const cell = ws[addr] as XLSX.CellObject;
      if (!cell.f) continue; // not a formula cell

      const v = cell.v;
      const isZeroOrEmpty = v === 0 || v === null || v === undefined || v === "";

      if (isZeroOrEmpty) {
        stale.push({
          sheet:       sheetName,
          addr,
          formula:     cell.f,
          cachedValue: v,
        });
      }
    }
  }

  const summary = stale.length === 0
    ? `No stale formula candidates detected across ${sheetsScan.length} sheet(s)`
    : `${stale.length} formula cell(s) with zero/null cached values across ${sheetsScan.length} sheet(s) — will recalculate on Excel open`;

  return {
    forceFullCalcApplied: true,
    staleFormulaCandidates: stale,
    sheetsScan,
    summary,
  };
}
