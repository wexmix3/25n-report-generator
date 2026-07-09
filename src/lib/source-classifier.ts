import * as XLSX from "xlsx-js-style";

// ─── Types ────────────────────────────────────────────────────────────────────

export type SourceType =
  | "VOYAGER_PL"           // Voyager P&L export — contains Sales By Month with revenue rows
  | "VOYAGER_BUDGET"       // Voyager budget export — contains Budget tab with line items
  | "KUBE_OCCUPANCY"       // Kube occupancy report — unit-level occupancy/status data
  | "KUBE_INCOME_DETAIL"   // Kube income detail — per-member revenue breakdown
  | "KUBE_BOOKING_SUMMARY" // Kube booking summary — meeting room utilization
  | "CUSTOMER_LIST"        // Member / customer contact list
  | "CONSOLIDATED"         // Legacy all-in-one: Sales By Month + occupancy in one file
  | "UNKNOWN";

/** Capabilities derived from the source — maps to which dashboard sections can be hydrated */
export type SourceCapability =
  | "SALES_BY_MONTH"  // Can hydrate SBM tab and P.2 revenue formulas
  | "OCCUPANCY"       // Can hydrate P.3 occupancy section
  | "BUDGET"          // Can hydrate P.4 budget column
  | "MEETING_HOURS"   // Can update meeting hours / meeting revenue KPIs
  | "CUSTOMER_DATA"   // Can update member count KPIs
  | "HISTORICAL";     // Contains historical snapshot data for P.1 trend section

export interface SourceClassification {
  sourceType:    SourceType;
  capabilities:  SourceCapability[];
  confidence:    "high" | "medium" | "low";
  detectedHints: string[];
}

// ─── Structural helpers ───────────────────────────────────────────────────────

/** Returns true if the given sheet has a row with 3+ text month headers like "Jan 2026". */
function sheetHasTextMonthHeaders(ws: XLSX.WorkSheet): boolean {
  const TEXT_MONTH = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}$/i;
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  for (let r = 0; r <= Math.min(range.e.r, 10); r++) {
    let count = 0;
    for (let c = 0; c <= Math.min(range.e.c, 20); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell?.v && TEXT_MONTH.test(String(cell.v).trim())) count++;
    }
    if (count >= 3) return true;
  }
  return false;
}

/** Returns true if the given sheet has a row where 3+ cells are Excel date serials (40000–60000). */
function sheetHasDateSerialHeader(ws: XLSX.WorkSheet): boolean {
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  for (let r = 0; r <= Math.min(range.e.r, 6); r++) {
    let count = 0;
    for (let c = 0; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "number" && cell.v > 40_000 && cell.v < 60_000) count++;
    }
    if (count >= 3) return true;
  }
  return false;
}

/**
 * Returns true if the given sheet looks like a unit/space occupancy list:
 * has a header row containing both a status-like column and a space/sqft-like column.
 */
function sheetHasOccupancyStructure(ws: XLSX.WorkSheet): boolean {
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  for (let r = 0; r <= Math.min(range.e.r, 6); r++) {
    const rowTexts: string[] = [];
    for (let c = 0; c <= Math.min(range.e.c, 25); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell?.v) rowTexts.push(String(cell.v).toLowerCase());
    }
    const row = rowTexts.join(" ");
    const hasStatusOrOcc = /\bstatus\b|\boccupancy\b|\boccupied\b/.test(row);
    const hasSpaceOrSqFt = /\bsq\.?\s*ft\b|\bsqft\b|\bunit\b|\bspace\b|\btype\b/.test(row);
    if (hasStatusOrOcc && hasSpaceOrSqFt) return true;
  }
  return false;
}

/**
 * Returns true if the given sheet looks like a Voyager Budget Comparison export.
 * Voyager exports these with a "Report1" sheet name and column headers like
 * "PTD Budget", "YTD Budget", or "Annual".
 */
function sheetHasBudgetComparisonStructure(ws: XLSX.WorkSheet): boolean {
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  for (let r = 0; r <= Math.min(range.e.r, 10); r++) {
    const rowTexts: string[] = [];
    for (let c = 0; c <= Math.min(range.e.c, 15); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell?.v) rowTexts.push(String(cell.v).toLowerCase());
    }
    const row = rowTexts.join(" ");
    const hasBudgetCol  = /\bptd.?budget\b|\bytd.?budget\b|\bannual\b/.test(row);
    const hasActualCol  = /\bptd.?actual\b|\bytd.?actual\b|\bactual\b/.test(row);
    const hasVariance   = /\bvariance\b|\bvar\b/.test(row);
    if (hasBudgetCol && hasActualCol && hasVariance) return true;
  }
  return false;
}

/**
 * Returns true if the given sheet looks like a Kube income-detail-by-account export.
 * These have a header row containing both a "Charge Code" column and an "Income" or "Amount"
 * column. Kube exports use generic sheet names ("Report1", "Sheet1") so structural detection
 * is the only reliable way to identify them.
 */
function sheetHasIncomeDetailStructure(ws: XLSX.WorkSheet): boolean {
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  for (let r = 0; r <= Math.min(range.e.r, 10); r++) {
    const rowTexts: string[] = [];
    for (let c = 0; c <= Math.min(range.e.c, 20); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell?.v) rowTexts.push(String(cell.v).toLowerCase().trim());
    }
    const row = rowTexts.join(" ");
    if (/\bcharge.?code\b/.test(row) && (/\bincome\b/.test(row) || /\bamount\b/.test(row))) return true;
  }
  return false;
}

/** Returns true if the given sheet looks like a booking/meeting-room summary. */
function sheetHasBookingStructure(ws: XLSX.WorkSheet): boolean {
  const range = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  for (let r = 0; r <= Math.min(range.e.r, 6); r++) {
    const rowTexts: string[] = [];
    for (let c = 0; c <= Math.min(range.e.c, 20); c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell?.v) rowTexts.push(String(cell.v).toLowerCase());
    }
    const row = rowTexts.join(" ");
    if (/\bhour[s]?\b/.test(row) && /\broom\b|\bmeeting\b|\bbooking\b|\bconference\b/.test(row)) return true;
  }
  return false;
}

// ─── Main classifier ──────────────────────────────────────────────────────────

export function detectSourceType(wb: XLSX.WorkBook): SourceClassification {
  const names  = wb.SheetNames;
  const lnames = names.map((n) => n.toLowerCase());
  const hints: string[] = [];
  const caps:  SourceCapability[] = [];

  // ── Sheet-name signals ──────────────────────────────────────────────────────
  const hasSBMName      = lnames.some((n) => /sales.?by.?month/.test(n));
  const hasBudgetName   = lnames.some((n) => /\bbudget\b/.test(n));
  const hasOccName      = lnames.some((n) => /\boccupancy\b|\boccupan/.test(n));
  const hasP3Name       = lnames.some((n) => /^p\.?3\b/.test(n));
  const hasBookingName  = lnames.some((n) => /\bbooking\b|\bmeeting\b|\bconference\b/.test(n));
  const hasIncomeName   = lnames.some((n) => /income.?detail|member.?income|income_detail/.test(n));
  const hasCustomerName = lnames.some((n) => /\bcustomer\b|\bmember\b|\bcontact\b/.test(n));
  const hasHistName     = lnames.some((n) => /\bhistorical\b|\bhistory\b/.test(n));

  // ── Structural signals ──────────────────────────────────────────────────────
  const sbmName = names.find((n) => /sales.?by.?month/i.test(n));
  const sbmHasDateSerials = sbmName
    ? sheetHasDateSerialHeader(wb.Sheets[sbmName])
    : false;

  // Check any sheet for text month headers (Voyager 12-Month Statement format)
  let textMonthSheet: string | undefined;
  for (const name of names) {
    if (sheetHasTextMonthHeaders(wb.Sheets[name])) { textMonthSheet = name; break; }
  }

  // Check all sheets for occupancy structure (P.3 may be named differently)
  let occupancySheet: string | undefined;
  for (const name of names) {
    if (sheetHasOccupancyStructure(wb.Sheets[name])) {
      occupancySheet = name;
      break;
    }
  }

  // Check for booking structure
  let bookingSheet: string | undefined;
  for (const name of names) {
    if (sheetHasBookingStructure(wb.Sheets[name])) {
      bookingSheet = name;
      break;
    }
  }

  // Check for budget comparison structure (Voyager Budget Comparison exports use "Report1" sheet)
  let budgetCompSheet: string | undefined;
  for (const name of names) {
    if (sheetHasBudgetComparisonStructure(wb.Sheets[name])) {
      budgetCompSheet = name;
      break;
    }
  }

  if (hasHistName) caps.push("HISTORICAL");

  // ── Classification ──────────────────────────────────────────────────────────

  // CONSOLIDATED: has Sales By Month with date serials AND occupancy structure
  if (sbmHasDateSerials && occupancySheet) {
    hints.push(`Sales By Month (${sbmName}) with date-serial header + occupancy structure in '${occupancySheet}'`);
    caps.push("SALES_BY_MONTH", "OCCUPANCY");
    if (hasBudgetName) caps.push("BUDGET");
    return { sourceType: "CONSOLIDATED", capabilities: caps, confidence: "high", detectedHints: hints };
  }

  // VOYAGER_PL: has Sales By Month with date serials, no occupancy structure
  if (sbmHasDateSerials) {
    hints.push(`Sales By Month (${sbmName}) with date-serial header row → Voyager P&L export`);
    caps.push("SALES_BY_MONTH");
    if (hasBudgetName) { caps.push("BUDGET"); hints.push("Budget tab detected"); }
    return { sourceType: "VOYAGER_PL", capabilities: caps, confidence: "high", detectedHints: hints };
  }

  // VOYAGER_PL: text month headers (Voyager 12-Month Statement / P&L export)
  if (!sbmHasDateSerials && textMonthSheet && !occupancySheet) {
    hints.push(`Text month headers (Jan YYYY format) in '${textMonthSheet}' — Voyager 12-Month Statement`);
    caps.push("SALES_BY_MONTH");
    if (hasBudgetName) { caps.push("BUDGET"); hints.push("Budget tab detected"); }
    return { sourceType: "VOYAGER_PL", capabilities: caps, confidence: "high", detectedHints: hints };
  }

  // Has SBM name but no valid date serials — possibly malformed Voyager
  if (hasSBMName && !sbmHasDateSerials) {
    hints.push(`Sheet named 'Sales By Month' found but no date-serial header row — treating as VOYAGER_PL (low confidence)`);
    caps.push("SALES_BY_MONTH");
    return { sourceType: "VOYAGER_PL", capabilities: caps, confidence: "low", detectedHints: hints };
  }

  // KUBE_OCCUPANCY: has occupancy structure, no SBM
  if (occupancySheet) {
    hints.push(`Occupancy structure (Status/Type/Sq Ft columns) in '${occupancySheet}' without Sales By Month → Kube Occupancy`);
    caps.push("OCCUPANCY");
    return { sourceType: "KUBE_OCCUPANCY", capabilities: caps, confidence: "high", detectedHints: hints };
  }

  // Occupancy sheet name present even without structural confirmation
  if ((hasOccName || hasP3Name) && !hasSBMName) {
    hints.push("Sheet named 'occupancy' or 'P.3' without Sales By Month → likely Kube Occupancy");
    caps.push("OCCUPANCY");
    return { sourceType: "KUBE_OCCUPANCY", capabilities: caps, confidence: "medium", detectedHints: hints };
  }

  // KUBE_BOOKING_SUMMARY: has booking/meeting structure
  if (bookingSheet || hasBookingName) {
    hints.push(`Booking/meeting-room structure in '${bookingSheet ?? ""}' → Kube Booking Summary`);
    caps.push("MEETING_HOURS");
    return { sourceType: "KUBE_BOOKING_SUMMARY", capabilities: caps, confidence: "medium", detectedHints: hints };
  }

  // KUBE_INCOME_DETAIL (name-based)
  if (hasIncomeName) {
    hints.push("Income detail sheet detected → Kube Income Detail");
    return { sourceType: "KUBE_INCOME_DETAIL", capabilities: [], confidence: "medium", detectedHints: hints };
  }

  // KUBE_INCOME_DETAIL (structural) — Kube exports use generic sheet names ("Report1", "Sheet1")
  // so name-based detection misses them. Detect by column headers: Charge Code + Income/Amount.
  {
    let incomeDetailSheet: string | undefined;
    for (const name of names) {
      if (sheetHasIncomeDetailStructure(wb.Sheets[name])) { incomeDetailSheet = name; break; }
    }
    if (incomeDetailSheet) {
      hints.push(`Income detail structure (Charge Code + Income columns) in '${incomeDetailSheet}' → Kube Income Detail`);
      caps.push("CUSTOMER_DATA");
      return { sourceType: "KUBE_INCOME_DETAIL", capabilities: caps, confidence: "high", detectedHints: hints };
    }
  }

  // VOYAGER_BUDGET: structural match (PTD Actual + PTD Budget + Variance columns)
  // Voyager Budget Comparison exports use "Report1" as the sheet name, so name-based
  // detection misses them — structural detection is required.
  if (budgetCompSheet) {
    hints.push(`Budget Comparison structure (PTD Actual/Budget/Variance columns) in '${budgetCompSheet}' → Voyager Budget export`);
    caps.push("BUDGET");
    return { sourceType: "VOYAGER_BUDGET", capabilities: caps, confidence: "high", detectedHints: hints };
  }

  // VOYAGER_BUDGET: budget tab name (fallback for files with explicit "Budget" sheet name)
  if (hasBudgetName) {
    hints.push("Budget sheet without Sales By Month → Voyager Budget export");
    caps.push("BUDGET");
    return { sourceType: "VOYAGER_BUDGET", capabilities: caps, confidence: "medium", detectedHints: hints };
  }

  // CUSTOMER_LIST
  if (hasCustomerName) {
    hints.push("Customer/member sheet without financial data → Customer List");
    caps.push("CUSTOMER_DATA");
    return { sourceType: "CUSTOMER_LIST", capabilities: caps, confidence: "low", detectedHints: hints };
  }

  hints.push(`Unrecognized structure — ${names.length} sheet(s): ${names.slice(0, 5).join(", ")}`);
  return { sourceType: "UNKNOWN", capabilities: [], confidence: "low", detectedHints: hints };
}

/** Human-readable label for a SourceType. */
export const SOURCE_TYPE_LABELS: Record<SourceType, string> = {
  VOYAGER_PL:           "Voyager P&L Export",
  VOYAGER_BUDGET:       "Voyager Budget Export",
  KUBE_OCCUPANCY:       "Kube Occupancy Report",
  KUBE_INCOME_DETAIL:   "Kube Income Detail",
  KUBE_BOOKING_SUMMARY: "Kube Booking Summary",
  CUSTOMER_LIST:        "Customer List",
  CONSOLIDATED:         "Consolidated Source Workbook",
  UNKNOWN:              "Unknown Source",
};

/** Human-readable label for a SourceCapability. */
export const CAPABILITY_LABELS: Record<SourceCapability, string> = {
  SALES_BY_MONTH: "P.2 Revenue Data (Sales By Month)",
  OCCUPANCY:      "P.3 Occupancy Data",
  BUDGET:         "P.4 Budget Column",
  MEETING_HOURS:  "Meeting Hours / Revenue KPIs",
  CUSTOMER_DATA:  "Member Count KPIs",
  HISTORICAL:     "P.1 Historical Trend Data",
};
