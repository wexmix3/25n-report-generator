import * as XLSX from "xlsx-js-style";
import {
  parseSalesByMonth,
  parse12MonthStatement,
  parseOccupancyStats,
  parseIncomeDetailMembers,
  extractCurrentSnapshot,
  parseBudgetSource,
  parseBookingSummary,
  parseBookingSummaryDetail,
  parseDailySpaceRaw,
  parseFlexRaw,
  parseMailRaw,
  parseYardiAdjOccAllRaw,
  parseYardiIncDetailAllRaw,
  type SalesByMonthData,
  type OccupancyStats,
  type KpiSnapshot,
  type BudgetSource,
  type BookingSummaryDetail,
  type DailySpaceRawData,
  type MembershipRawData,
} from "@/lib/excel-dashboard-generator";
import { hydrateWorkbook, type HydrationSectionResult } from "@/lib/workbook-hydrator";
import { reconcileOutput, type ReconciliationReport } from "@/lib/reconciliation-engine";
import { applyRecalcFlags } from "@/lib/recalc-engine";
import { injectPivotParts, preserveTemplateStyles, insertP1SnapshotRow } from "@/lib/pivot-engine";
import {
  detectSourceType,
  SOURCE_TYPE_LABELS,
  type SourceClassification,
  type SourceCapability,
} from "@/lib/source-classifier";

// ─── Types ────────────────────────────────────────────────────────────────────

/** Overall generation outcome, written to the Dashboard Status sheet and response headers. */
export type GenerationStatus = "COMPLETE" | "WARNING" | "PARTIAL" | "CRITICAL";

export type WorkflowStageId =
  | "INGEST_SOURCES"
  | "HYDRATE_P4_VOYAGER"
  | "VALIDATE_P4_RECONCILIATION"
  | "FINALIZE_WORKBOOK";

export type StageStatus = "pending" | "running" | "complete" | "failed" | "awaiting_operator";

export type ExceptionSeverity = "blocking" | "operator_required" | "advisory";

export interface WorkflowException {
  code:             string;
  severity:         ExceptionSeverity;
  description:      string;
  context:          Record<string, unknown>;
  candidateFix?:    string;
  requiresOperator: boolean;
}

export interface ReconciliationCheck {
  rule:     string;
  expected: string;
  actual:   string;
  delta?:   string;
  passed:   boolean;
  severity: "INFO" | "WARNING" | "ERROR" | "CRITICAL";
}

export interface AuditEvent {
  timestamp: string;
  stageId:   WorkflowStageId;
  event:     string;
  data?:     Record<string, unknown>;
}

export interface StageResult {
  stageId:              WorkflowStageId;
  status:               StageStatus;
  durationMs:           number;
  reconciliationChecks: ReconciliationCheck[];
  exceptions:           WorkflowException[];
  auditEvents:          AuditEvent[];
}

export interface WorkflowResult {
  runId:              string;
  status:             "complete" | "failed" | "awaiting_operator";
  generationStatus:   GenerationStatus;
  stages:             StageResult[];
  outputBuffer?:      Buffer;
  snapshot?:          KpiSnapshot;
  reconciliation?:    ReconciliationReport;
  totalDurationMs:    number;
  auditLog:           AuditEvent[];
  blockedBy?:         string;
  sourceType?:        string;
  sourceCapabilities?: SourceCapability[];
  sourceLabel?:       string;
  sectionResults?:    HydrationSectionResult[];
  occRows?:           number;
  sbmRows?:           number;
}

export interface WorkflowInput {
  templateBuffer:       Buffer;
  sourceBuffers:        ArrayBuffer[];  // one or more source files
  sourceFileNames?:     string[];
  historicalSnapshots?: KpiSnapshot[];
  overrides?:           Record<string, unknown>;
}

// Internal context threaded through all stages
interface WorkflowContext {
  runId:                string;
  templateBuffer:       Buffer;
  sourceBuffers:        ArrayBuffer[];  // replaces sourceBuffer
  sourceFileNames:      string[];
  historicalSnapshots:  KpiSnapshot[];
  overrides:            Record<string, unknown>;
  sourceClassifications: SourceClassification[];
  sourceClassification?: SourceClassification;  // primary (first recognized)
  sales?:               SalesByMonthData;
  occupancy?:           OccupancyStats;
  snapshot?:            KpiSnapshot;
  budgetSource?:        BudgetSource | null;
  bookingSummaryDetail?:      BookingSummaryDetail;
  bookingSummaryPriorDetail?: BookingSummaryDetail;
  dailySpaceRawData?:      DailySpaceRawData;
  flexRawData?:            MembershipRawData;
  mailRawData?:            MembershipRawData;
  yardiIncDetailRawData?:  MembershipRawData;
  yardiAdjOccRawData?:     MembershipRawData;
  workbookBuffer?:      Buffer;
  p4ColUpdate?:         { from: string; to: string; cellsUpdated: number } | null;
  reconciliation?:      ReconciliationReport;
  sectionResults?:      HydrationSectionResult[];
}

// ─── Utilities ─────────────────────────────────────────────────────────────────

function ts(): string { return new Date().toISOString(); }

function audit(stageId: WorkflowStageId, event: string, data?: Record<string, unknown>): AuditEvent {
  return { timestamp: ts(), stageId, event, data };
}

function generateRunId(): string {
  return `wf_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

// Converts an Excel date serial to "MM/YYYY" for matching income detail Post Month
function formatSerialToPostMonth(serial: number): string {
  const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
  const m = String(date.getUTCMonth() + 1).padStart(2, "0");
  return `${m}/${date.getUTCFullYear()}`;
}

function colIdxToLetter(idx: number): string {
  let letter = "";
  let n = idx + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}

// ─── Data Bridge builder ──────────────────────────────────────────────────────
//
// Replaces the complex template-clone pipeline with a simple, clean Excel file
// containing only the parsed data.  A VBA macro in the master template reads
// this file and populates the dashboard natively — zero OOXML surgery required.
//
// Sheets produced:
//   Meta    — key/value KPI snapshot + current-month column hint for Find+Replace
//   SBM     — account labels × months (matches 12-Month Statement row structure)
//   Budget  — budget line items (if a Budget Comparison file was uploaded)
//   Status  — human-readable source-detection summary

function buildDataBridge(ctx: WorkflowContext): Buffer {
  const wb       = XLSX.utils.book_new();
  const snap     = ctx.snapshot;
  const sales    = ctx.sales;

  // ── Sheet 1: Meta ──────────────────────────────────────────────────────────
  // The VBA macro reads cell B2 onward by key name.  Keep key names stable.
  const metaRows: (string | number)[][] = [
    ["key",                    "value"],
    ["month_code",              snap?.monthCode            ?? ""],
    ["month_serial",            snap?.monthSerial          ?? 0],
    ["location",                snap?.location             ?? ""],
    ["occupancy_pct_spaces",    snap?.occupancyPctSpaces   ?? 0],
    ["occupancy_pct_sqft",      snap?.occupancyPctSqFt     ?? 0],
    ["contracts",               snap?.contracts            ?? 0],
    ["flex_members",            snap?.flexMembers          ?? 0],
    ["mail_members",            snap?.mailMembers          ?? 0],
    ["meeting_hours",           snap?.meetingHours         ?? 0],
    ["meeting_hours_prior",     snap?.meetingHoursPrior    ?? 0],
    ["meeting_revenue_dollar",  snap?.meetingRevenueDollar ?? 0],
    ["sbm_month_count",         sales?.monthSerials.length ?? 0],
    ["sbm_row_count",           sales?.lookupByLabel.size  ?? 0],
    ["contracts_auto",          (snap?.contracts ?? 0) > 0 ? 1 : 0],
    ["flex_auto",               (snap?.flexMembers ?? 0) > 0 ? 1 : 0],
    ["mail_auto",               (snap?.mailMembers ?? 0) > 0 ? 1 : 0],
    ["occ_rows",                ctx.occupancy?.rawRows.length ?? 0],
    ["daily_space_current",     snap?.dailySpaceCurrent ?? 0],
    ["daily_space_prior",       snap?.dailySpacePrior   ?? 0],
    ["generated_at",            new Date().toISOString()],
  ];
  const metaWs = XLSX.utils.aoa_to_sheet(metaRows);
  // Force string type on month_code (B2) and generated_at (B21) so Excel doesn't
  // auto-convert "May-26" to a date serial when the file is opened.
  (["B2", "B21"] as const).forEach((addr) => {
    const cell = (metaWs as Record<string, XLSX.CellObject>)[addr];
    if (cell) { cell.t = "s"; cell.v = String(cell.v ?? ""); }
  });
  XLSX.utils.book_append_sheet(wb, metaWs, "Meta");

  // ── Sheet 2: SBM ───────────────────────────────────────────────────────────
  // Row 1: "Label", "Jan-26", "Feb-26", …  (month codes from source file)
  // Row 2+: account label, value per month
  if (sales && sales.monthCodes.length > 0) {
    const headerRow: (string | number)[] = ["Label", ...sales.monthCodes];
    const sbmRows: (string | number | null)[][] = [headerRow];
    for (const [label, values] of sales.lookupByLabel) {
      sbmRows.push([label, ...values]);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sbmRows), "SBM");
  }

  // ── Sheet 3: Occupancy ─────────────────────────────────────────────────────
  // Raw rows from the Kube Occupancy Report — VBA pastes these into P.3
  if (ctx.occupancy && ctx.occupancy.rawRows.length > 0) {
    const occRows: (string | number | null)[][] = [ctx.occupancy.rawHeaders, ...ctx.occupancy.rawRows];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(occRows), "Occupancy");
  }

  // ── Sheet 4: Budget ────────────────────────────────────────────────────────
  if (ctx.budgetSource && ctx.budgetSource.lineItems.length > 0) {
    const budgetRows: (string | number | boolean)[][] = [
      ["Label", "Budget Value", "Is Section Header", "Is Total"],
    ];
    for (const item of ctx.budgetSource.lineItems) {
      if (item.isBlankRow) continue;
      budgetRows.push([item.label, item.budgetValue ?? 0, item.isSectionHeader, item.isTotal]);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(budgetRows), "Budget");
  }

  // ── Sheet 5: MtgRoomHours ─────────────────────────────────────────────────
  // One row per P2 meeting room label.  VBA reads P2Label (col A), Hours (col C),
  // and Revenue (col D) to fill P.2 HOURS USED and REVENUE COLLECTED sections.
  if (ctx.bookingSummaryDetail && ctx.bookingSummaryDetail.rooms.length > 0) {
    const mrhRows: (string | number)[][] = [
      ["P2Label", "SpaceName", "Hours", "Revenue"],
    ];
    for (const room of ctx.bookingSummaryDetail.rooms) {
      mrhRows.push([room.p2Label, room.spaceName, room.hours, room.revenue]);
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mrhRows), "MtgRoomHours");
  }

  // ── Sheet 6: BookingSummaryRaw ─────────────────────────────────────────────
  // All current-month booking rows verbatim (non-cancelled, non-25N-Staff).
  // VBA pastes these into the "Prior Month Meeting Hours" tab, replacing the
  // prior month's data so next month's run has the correct prior-month context.
  if (ctx.bookingSummaryDetail && ctx.bookingSummaryDetail.rawRows.length > 0) {
    const bsrRows: (string | number | null)[][] = [
      ctx.bookingSummaryDetail.rawHeaders,
      ...ctx.bookingSummaryDetail.rawRows,
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bsrRows), "BookingSummaryRaw");
  }

  // ── Sheet 6b: BookingSummaryPriorRaw ───────────────────────────────────────
  // Prior month's individual booking rows — VBA pastes these into the
  // "Prior Month Meeting Hours" tab, replacing the current month's rows.
  if (ctx.bookingSummaryPriorDetail && ctx.bookingSummaryPriorDetail.rawRows.length > 0) {
    const bsprRows = [
      ctx.bookingSummaryPriorDetail.rawHeaders,
      ...ctx.bookingSummaryPriorDetail.rawRows,
    ] as (string | number | null)[][];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bsprRows), "BookingSummaryPriorRaw");
  }

  // ── Sheet 7: DailySpaceRaw ─────────────────────────────────────────────────
  if (ctx.dailySpaceRawData && ctx.dailySpaceRawData.rows.length > 0) {
    const dsrRows: string[][] = [
      ctx.dailySpaceRawData.headers,
      ...ctx.dailySpaceRawData.rows,
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(dsrRows), "DailySpaceRaw");
  }

  // ── Sheet 8: FlexRaw ──────────────────────────────────────────────────────
  if (ctx.flexRawData && ctx.flexRawData.rows.length > 0) {
    const flexRows: string[][] = [ctx.flexRawData.headers, ...ctx.flexRawData.rows];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(flexRows), "FlexRaw");
  }

  // ── Sheet 9: MailRaw ──────────────────────────────────────────────────────
  if (ctx.mailRawData && ctx.mailRawData.rows.length > 0) {
    const mailRows: string[][] = [ctx.mailRawData.headers, ...ctx.mailRawData.rows];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mailRows), "MailRaw");
  }

  // ── Sheet 10: YardiIncDetailRaw ───────────────────────────────────────────
  // ALL rows from income detail files (no charge-code filter).
  // VBA pastes into "WA Yardi Inc Detail By Acct" (rows 9+) so the SUMIF
  // formulas in "WA Yardi Adj Occupancy Rep" (L108/L109/L110) recalculate.
  if (ctx.yardiIncDetailRawData && ctx.yardiIncDetailRawData.rows.length > 0) {
    const yardiRows: string[][] = [ctx.yardiIncDetailRawData.headers, ...ctx.yardiIncDetailRawData.rows];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(yardiRows), "YardiIncDetailRaw");
  }

  // ── Sheet 11: YardiAdjOccRaw ─────────────────────────────────────────────
  // Raw data rows for "WA Yardi Adj Occupancy Rep" tab.
  // VBA pastes these into the tab's data rows while preserving formula cells (rows 108+).
  if (ctx.yardiAdjOccRawData && ctx.yardiAdjOccRawData.rows.length > 0) {
    const adjOccRows: string[][] = [ctx.yardiAdjOccRawData.headers, ...ctx.yardiAdjOccRawData.rows];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(adjOccRows), "YardiAdjOccRaw");
  }

  // ── Sheet 12: Status ──────────────────────────────────────────────────────
  const sourceList = ctx.sourceClassifications
    .map((c, i) => `${ctx.sourceFileNames[i] ?? `File ${i + 1}`}: ${SOURCE_TYPE_LABELS[c.sourceType]}`)
    .join(" | ");
  const statusRows: (string | number)[][] = [
    ["Data Bridge — Source Detection Summary"],
    ["Files uploaded",   ctx.sourceBuffers.length],
    ["Sources",          sourceList || "—"],
    ["Month detected",   snap?.monthCode        ?? "Unknown"],
    ["Location",         snap?.location         ?? "Unknown"],
    ["Occupancy (spaces)", snap ? `${(snap.occupancyPctSpaces * 100).toFixed(1)}%` : "0%"],
    ["Contracts",        snap?.contracts        ?? 0],
    ["SBM rows",         sales?.lookupByLabel.size ?? 0],
    ["SBM months",       sales?.monthCodes.join(", ") ?? "—"],
    ["Generated",        new Date().toISOString()],
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(statusRows), "Status");

  return Buffer.from(XLSX.write(wb, { type: "buffer", bookType: "xlsx" }));
}

function buildSerialToColMap(ws: XLSX.WorkSheet): Map<number, string> {
  const result = new Map<number, string>();
  const range  = XLSX.utils.decode_range(ws["!ref"] ?? "A1:A1");
  // First pass: look for numeric date serials (standard Voyager SBM format)
  for (let r = 0; r <= Math.min(range.e.r, 5); r++) {
    const row: { col: number; serial: number }[] = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && typeof cell.v === "number" && cell.v > 40_000 && cell.v < 60_000) {
        row.push({ col: c, serial: cell.v });
      }
    }
    if (row.length >= 3) {
      for (const { col, serial } of row) result.set(serial, colIdxToLetter(col));
      return result;
    }
  }
  // Second pass: text month headers ("Jan 2026") — used when template has 12-Month Statement format
  const TM_RE  = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})$/i;
  const TM_IDX: Record<string, number> = { jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11 };
  for (let r = 0; r <= Math.min(range.e.r, 5); r++) {
    const hits: { col: number; serial: number }[] = [];
    for (let c = 0; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell?.v) continue;
      const m = TM_RE.exec(String(cell.v).trim());
      if (m) {
        const idx  = TM_IDX[m[1].toLowerCase()];
        const year = parseInt(m[2]);
        if (idx !== undefined && year) {
          const serial = Math.floor(Date.UTC(year, idx, 1) / 86400000) + 25569;
          hits.push({ col: c, serial });
        }
      }
    }
    if (hits.length >= 3) {
      for (const { col, serial } of hits) result.set(serial, colIdxToLetter(col));
      return result;
    }
  }
  return result;
}

function detectP4ReferencedCol(p4Ws: XLSX.WorkSheet, sbmName: string): string | null {
  const counts  = new Map<string, number>();
  const escaped = sbmName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`(?:'${escaped}'|"${escaped}"|${escaped})!([A-Z]+)(\\d+)`, "g");
  for (const addr of Object.keys(p4Ws)) {
    if (addr.startsWith("!")) continue;
    const cell = p4Ws[addr] as XLSX.CellObject;
    if (!cell.f) continue;
    let m: RegExpExecArray | null;
    pattern.lastIndex = 0;
    while ((m = pattern.exec(cell.f)) !== null) {
      counts.set(m[1], (counts.get(m[1]) ?? 0) + 1);
    }
  }
  if (counts.size === 0) return null;
  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0][0];
}

function applyP4ColUpdate(
  p4Ws: XLSX.WorkSheet, sbmName: string, fromLetter: string, toLetter: string,
): { sheet: XLSX.WorkSheet; updatedCells: number } {
  const escaped = sbmName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(
    `((?:'${escaped}'|"${escaped}"|${escaped})!)${fromLetter}(\\d+)`, "g",
  );
  let updatedCells = 0;
  const out: XLSX.WorkSheet = {};
  for (const key of Object.keys(p4Ws)) {
    if (key.startsWith("!")) { out[key] = JSON.parse(JSON.stringify(p4Ws[key])); continue; }
    const cell = JSON.parse(JSON.stringify(p4Ws[key] as XLSX.CellObject)) as XLSX.CellObject;
    if (cell.f) {
      const updated = cell.f.replace(pattern, `$1${toLetter}$2`);
      if (updated !== cell.f) { cell.f = updated; cell.v = 0; updatedCells++; }
    }
    out[key] = cell;
  }
  return { sheet: out, updatedCells };
}

// ─── Stage: INGEST_SOURCES ────────────────────────────────────────────────────

async function runIngestSources(ctx: WorkflowContext): Promise<StageResult> {
  const t0    = Date.now();
  const stage = "INGEST_SOURCES" as const;
  const exceptions: WorkflowException[]   = [];
  const events:     AuditEvent[]          = [];
  const checks:     ReconciliationCheck[] = [];

  // Parse all source buffers
  const parsedWbs: XLSX.WorkBook[] = [];
  for (let i = 0; i < ctx.sourceBuffers.length; i++) {
    try {
      const wb = XLSX.read(ctx.sourceBuffers[i], { type: "array", cellDates: false, cellStyles: true, cellNF: true });
      parsedWbs.push(wb);
    } catch (e) {
      exceptions.push({
        code: "INGEST_PARSE_FAILED", severity: "blocking",
        description: `Source file ${i + 1} cannot be read — it may be corrupted or in an unsupported format: ${e instanceof Error ? e.message : String(e)}`,
        context: {}, requiresOperator: false,
      });
      return {
        stageId: stage, status: "failed", durationMs: Date.now() - t0,
        reconciliationChecks: checks, exceptions, auditEvents: [audit(stage, "parse_failed")],
      };
    }
  }

  if (parsedWbs.length === 0) {
    exceptions.push({
      code: "INGEST_PARSE_FAILED", severity: "blocking",
      description: "No source files provided",
      context: {}, requiresOperator: false,
    });
    return {
      stageId: stage, status: "failed", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: [audit(stage, "parse_failed")],
    };
  }

  // Classify each source workbook
  const classifications: SourceClassification[] = parsedWbs.map((wb) => detectSourceType(wb));
  ctx.sourceClassifications = classifications;

  // Primary classification = first recognized source (fallback to first)
  const primaryClassification = classifications.find((c) => c.sourceType !== "UNKNOWN") ?? classifications[0];
  ctx.sourceClassification = primaryClassification;

  for (let i = 0; i < classifications.length; i++) {
    const cl = classifications[i];
    events.push(audit(stage, "source_classified", {
      fileIndex:     i,
      sourceType:    cl.sourceType,
      sourceLabel:   SOURCE_TYPE_LABELS[cl.sourceType],
      capabilities:  cl.capabilities,
      confidence:    cl.confidence,
      hints:         cl.detectedHints,
    }));
    checks.push({
      rule:     `Source file ${i + 1} type detected`,
      expected: "recognized",
      actual:   `${SOURCE_TYPE_LABELS[cl.sourceType]} (${cl.confidence} confidence)`,
      passed:   cl.sourceType !== "UNKNOWN",
      severity: cl.sourceType === "UNKNOWN" ? "WARNING" : "INFO",
    });
  }

  // Parse Sales By Month — try date-serial format first, then 12-Month Statement text format
  let sales: ReturnType<typeof parseSalesByMonth> | null = null;
  let salesFound = false;
  for (const wb of parsedWbs) {
    try {
      sales = parseSalesByMonth(wb);
      if (sales) { salesFound = true; break; }
    } catch { /* try next */ }
  }
  if (!salesFound) {
    for (const wb of parsedWbs) {
      const result = parse12MonthStatement(wb);
      if (result) { sales = result; salesFound = true; break; }
    }
  }
  if (!salesFound) {
    exceptions.push({
      code: "SBM_SHEET_MISSING", severity: "operator_required",
      description: "Sales By Month sheet not found in any uploaded file — P.2 revenue data will not be updated",
      context:     { availableSheets: parsedWbs.flatMap((wb) => wb.SheetNames) },
      candidateFix: "Upload a Voyager P&L or 12-Month Statement file — the system recognizes both date-serial and text-month column headers",
      requiresOperator: true,
    });
    checks.push({
      rule: "Sales By Month sheet present in source", expected: "present", actual: "missing",
      passed: false, severity: "ERROR",
    });
  }

  // Parse occupancy — try each workbook, stop when found
  let occupancy: ReturnType<typeof parseOccupancyStats> | null = null;
  for (const wb of parsedWbs) {
    try {
      occupancy = parseOccupancyStats(wb);
      break;
    } catch { /* try next */ }
  }
  if (!occupancy) {
    exceptions.push({
      code: "OCC_SHEET_MISSING", severity: "advisory",
      description: "Occupancy sheet not found in any uploaded file — occupancy data will not be updated",
      context:     { availableSheets: parsedWbs.flatMap((wb) => wb.SheetNames) },
      candidateFix: "Upload a file that includes the P.3 Occupancy tab, or upload the Kube Occupancy report separately",
      requiresOperator: false,
    });
  }

  const snapshot: KpiSnapshot = (sales && occupancy)
    ? extractCurrentSnapshot(sales, occupancy)
    : {
        monthCode:            "Unknown",
        monthSerial:          0,
        occupancyPctSpaces:   0,
        occupancyPctSqFt:     0,
        contracts:            0,
        flexMembers:          0,
        mailMembers:          0,
        meetingHours:         0,
        meetingRevenueDollar: 0,
        location:             "Unknown",
      };

  // Override contracts from occupancy file (col H = Contract Name, count distinct non-blank)
  if (occupancy && occupancy.contractsCount > 0) {
    snapshot.contracts = occupancy.contractsCount;
  }

  // Override flex/mail member counts and collect daily space revenue from income detail files
  if (snapshot.monthSerial > 0) {
    const targetMonth = formatSerialToPostMonth(snapshot.monthSerial);
    // Prior month: subtract ~30 days and reformat (handles month rollover correctly)
    const priorDate   = new Date(Math.round((snapshot.monthSerial - 25569) * 86400 * 1000));
    priorDate.setUTCMonth(priorDate.getUTCMonth() - 1);
    const priorMonth  = `${String(priorDate.getUTCMonth() + 1).padStart(2, "0")}/${priorDate.getUTCFullYear()}`;

    const allFlex       = new Set<string>();
    const allMail       = new Set<string>();
    const allDailySpace = new Map<string, number>();

    for (const wb of parsedWbs) {
      const detail = parseIncomeDetailMembers(wb, targetMonth);
      if (detail) {
        for (const c of detail.flexCompanies) allFlex.add(c);
        for (const c of detail.mailCompanies) allMail.add(c);
        for (const [month, rev] of detail.dailySpaceByMonth) {
          allDailySpace.set(month, (allDailySpace.get(month) ?? 0) + rev);
        }
      }
    }
    if (allFlex.size > 0) snapshot.flexMembers = allFlex.size;
    if (allMail.size > 0) snapshot.mailMembers = allMail.size;
    if (allDailySpace.has(targetMonth)) snapshot.dailySpaceCurrent = allDailySpace.get(targetMonth);
    if (allDailySpace.has(priorMonth))  snapshot.dailySpacePrior   = allDailySpace.get(priorMonth);

    // Daily Space raw rows — deduplicated by Invoice # across files (col 0)
    // so uploading multiple income detail exports doesn't create duplicates.
    const dsSeenInvoices = new Set<string>();
    const dsHeaders: string[] = [];
    const dsRows:    string[][] = [];
    let   dsYear = 0;
    for (const wb of parsedWbs) {
      const raw = parseDailySpaceRaw(wb, targetMonth);
      if (!raw) continue;
      if (dsHeaders.length === 0) { dsHeaders.push(...raw.headers); dsYear = raw.year; }
      const invoiceColIdx = raw.headers.findIndex((h) => /invoice/i.test(h));
      for (const row of raw.rows) {
        const invoiceKey = invoiceColIdx >= 0 ? row[invoiceColIdx] : row.join("|");
        if (!dsSeenInvoices.has(invoiceKey)) {
          dsSeenInvoices.add(invoiceKey);
          dsRows.push(row);
        }
      }
    }
    if (dsRows.length > 0) {
      ctx.dailySpaceRawData = { headers: dsHeaders, rows: dsRows, year: dsYear };
    }

    // Flex membership raw rows — deduplicated by Invoice # across files
    const flexSeenInvoices = new Set<string>();
    const flexHeaders: string[] = [];
    const flexRows:    string[][] = [];
    for (const wb of parsedWbs) {
      const raw = parseFlexRaw(wb, targetMonth);
      if (!raw) continue;
      if (flexHeaders.length === 0) flexHeaders.push(...raw.headers);
      const invoiceColIdx = raw.headers.findIndex((h) => /invoice/i.test(h));
      for (const row of raw.rows) {
        const key = invoiceColIdx >= 0 ? row[invoiceColIdx] : row.join("|");
        if (!flexSeenInvoices.has(key)) { flexSeenInvoices.add(key); flexRows.push(row); }
      }
    }
    if (flexRows.length > 0) ctx.flexRawData = { headers: flexHeaders, rows: flexRows };

    // Mail membership raw rows — deduplicated by Invoice # across files
    const mailSeenInvoices = new Set<string>();
    const mailHeaders: string[] = [];
    const mailRows:    string[][] = [];
    for (const wb of parsedWbs) {
      const raw = parseMailRaw(wb, targetMonth);
      if (!raw) continue;
      if (mailHeaders.length === 0) mailHeaders.push(...raw.headers);
      const invoiceColIdx = raw.headers.findIndex((h) => /invoice/i.test(h));
      for (const row of raw.rows) {
        const key = invoiceColIdx >= 0 ? row[invoiceColIdx] : row.join("|");
        if (!mailSeenInvoices.has(key)) { mailSeenInvoices.add(key); mailRows.push(row); }
      }
    }
    if (mailRows.length > 0) ctx.mailRawData = { headers: mailHeaders, rows: mailRows };

    // Yardi Income Detail — ALL rows across all income detail files, no charge-code filter.
    // Populates "YardiIncDetailRaw" bridge sheet → VBA pastes into "WA Yardi Inc Detail By Acct"
    // so SUMIF formulas in "WA Yardi Adj Occupancy Rep" (L108/L109/L110) recalculate.
    const yardiSeenInvoices = new Set<string>();
    const yardiHeaders: string[] = [];
    const yardiRows:    string[][] = [];
    for (const wb of parsedWbs) {
      const raw = parseYardiIncDetailAllRaw(wb);
      if (!raw) continue;
      if (yardiHeaders.length === 0) yardiHeaders.push(...raw.headers);
      const invCol = raw.headers.findIndex((h) => /invoice/i.test(h));
      for (const row of raw.rows) {
        const key = invCol >= 0 ? row[invCol] : row.join("|");
        if (!yardiSeenInvoices.has(key)) { yardiSeenInvoices.add(key); yardiRows.push(row); }
      }
    }
    if (yardiRows.length > 0) ctx.yardiIncDetailRawData = { headers: yardiHeaders, rows: yardiRows };

    // Yardi Adjusted Occupancy — raw data rows for "WA Yardi Adj Occupancy Rep" tab.
    // Formula cells (rows 108+) in that tab are NEVER touched; only data rows are refreshed.
    // Detection: occupancy-style headers that are NOT income detail and NOT Kube booking.
    const yardiAdjHeaders: string[] = [];
    const yardiAdjRows:    string[][] = [];
    for (const wb of parsedWbs) {
      const raw = parseYardiAdjOccAllRaw(wb);
      if (!raw) continue;
      if (yardiAdjHeaders.length === 0) yardiAdjHeaders.push(...raw.headers);
      for (const row of raw.rows) yardiAdjRows.push(row);
    }
    if (yardiAdjRows.length > 0) ctx.yardiAdjOccRawData = { headers: yardiAdjHeaders, rows: yardiAdjRows };

    // Meeting hours from Kube Booking Summary files.
    // Use max (not sum) across files — overlapping date ranges cause the same
    // bookings to appear in multiple exports, so summing would double-count.
    let totalMeetingHours = 0;
    let priorMeetingHours = 0;
    for (const wb of parsedWbs) {
      const bs = parseBookingSummary(wb, snapshot.monthSerial);
      if (bs) {
        if (bs.currentHours > totalMeetingHours) totalMeetingHours = bs.currentHours;
        if (bs.priorHours   > priorMeetingHours) priorMeetingHours = bs.priorHours;
      }
    }
    if (totalMeetingHours > 0) snapshot.meetingHours      = totalMeetingHours;
    if (priorMeetingHours > 0) snapshot.meetingHoursPrior = priorMeetingHours;

    // Booking summary by-room detail — pick the file with the most total hours
    // (same max logic as above; avoids double-counting from overlapping exports).
    let bestDetail: BookingSummaryDetail | null = null;
    let bestDetailHours = 0;
    for (const wb of parsedWbs) {
      const detail = parseBookingSummaryDetail(wb, snapshot.monthSerial);
      if (detail) {
        const total = detail.rooms.reduce((s, r) => s + r.hours, 0);
        if (total > bestDetailHours) { bestDetailHours = total; bestDetail = detail; }
      }
    }
    if (bestDetail) ctx.bookingSummaryDetail = bestDetail;

    // Prior month booking detail — for the "Prior Month Meeting Hours" tab.
    // Uses the first day of the month before the current snapshot month.
    const curDate2  = new Date(Math.round((snapshot.monthSerial - 25569) * 86400 * 1000));
    const priorYear2 = curDate2.getUTCMonth() === 0 ? curDate2.getUTCFullYear() - 1 : curDate2.getUTCFullYear();
    const priorMo2   = curDate2.getUTCMonth() === 0 ? 11 : curDate2.getUTCMonth() - 1;
    const priorMonthSerial = Math.round(Date.UTC(priorYear2, priorMo2, 1) / 86400000 + 25569);

    let bestPriorDetail: BookingSummaryDetail | null = null;
    let bestPriorHours = 0;
    for (const wb of parsedWbs) {
      const detail = parseBookingSummaryDetail(wb, priorMonthSerial);
      if (detail) {
        const total = detail.rooms.reduce((s, r) => s + r.hours, 0);
        if (total > bestPriorHours) { bestPriorHours = total; bestPriorDetail = detail; }
      }
    }
    if (bestPriorDetail) ctx.bookingSummaryPriorDetail = bestPriorDetail;
  }

  // Budget source — first workbook that has it
  let budgetSource: BudgetSource | null = null;
  for (const wb of parsedWbs) {
    const bs = parseBudgetSource(wb);
    if (bs) { budgetSource = bs; break; }
  }

  ctx.sales        = sales ?? undefined;
  ctx.occupancy    = occupancy ?? undefined;
  ctx.snapshot     = snapshot;
  ctx.budgetSource = budgetSource;

  events.push(audit(stage, "sources_parsed", {
    fileCount:    parsedWbs.length,
    monthCode:    snapshot.monthCode,
    monthSerial:  snapshot.monthSerial,
    location:     snapshot.location,
    hasBudget:    budgetSource !== null,
    sbmRowCount:  sales?.lookupByLabel.size ?? 0,
    monthCount:   sales?.monthSerials.length ?? 0,
    hasSales:     sales !== null,
    hasOccupancy: occupancy !== null,
  }));

  if (sales !== null && sales.monthSerials.length === 0) {
    exceptions.push({
      code: "SBM_NO_MONTHS", severity: "operator_required",
      description: "Sales By Month sheet contains no recognizable month columns",
      context:     { sheetName: sales.sheetName },
      candidateFix: "Verify the Sales By Month sheet has a header row with Excel date serials in the first 5 rows",
      requiresOperator: true,
    });
  } else if (sales !== null) {
    checks.push({
      rule: "Sales By Month: month columns detected", expected: ">= 1",
      actual: String(sales.monthSerials.length), passed: true, severity: "INFO",
    });
  }

  if (snapshot.monthSerial === 0) {
    exceptions.push({
      code: "SNAPSHOT_NO_DATE", severity: "operator_required",
      description: "Could not determine current reporting period from the uploaded files",
      context:     {},
      candidateFix: "Verify the source file contains Sales By Month data with valid month date headers",
      requiresOperator: true,
    });
  } else {
    checks.push({
      rule: "Current reporting period identified", expected: "valid month",
      actual: snapshot.monthCode, passed: true, severity: "INFO",
    });
  }

  // Only INGEST_PARSE_FAILED is truly blocking — everything else allows continuation
  return {
    stageId:    stage,
    status:     "complete",
    durationMs: Date.now() - t0,
    reconciliationChecks: checks, exceptions, auditEvents: events,
  };
}

// ─── Stage: HYDRATE_P4_VOYAGER ────────────────────────────────────────────────

async function runHydrateP4(ctx: WorkflowContext): Promise<StageResult> {
  const t0    = Date.now();
  const stage = "HYDRATE_P4_VOYAGER" as const;
  const exceptions: WorkflowException[]   = [];
  const events:     AuditEvent[]          = [];
  const checks:     ReconciliationCheck[] = [];

  // 1. Hydrate workbook using the existing engine (preserves formulas, styles, pivots)
  let hydrationBuffer: Buffer;
  try {
    const result = hydrateWorkbook(ctx.templateBuffer, ctx.sourceBuffers, ctx.historicalSnapshots);
    hydrationBuffer     = result.buffer;
    ctx.sectionResults  = result.sectionResults;
    if (result.warnings.length > 0) {
      events.push(audit(stage, "hydration_warnings", { warnings: result.warnings }));
    }
    events.push(audit(stage, "hydration_complete", {
      tabsRefreshed:  result.tabsRefreshed,
      tabsPreserved:  result.tabsPreserved,
      pivotReport:    result.pivotReport.summary,
      sectionCount:   result.sectionResults.length,
      hydratedCount:  result.sectionResults.filter(s => s.status === "hydrated").length,
    }));
  } catch (e) {
    // HYDRATION_FAILED is truly blocking — without this, there is no workbook
    exceptions.push({
      code: "HYDRATION_FAILED", severity: "blocking",
      description: `Workbook hydration failed — the template may be incompatible with the source file: ${e instanceof Error ? e.message : String(e)}`,
      context: {}, requiresOperator: false,
    });
    return {
      stageId: stage, status: "failed", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: events,
    };
  }

  // 2. Parse the hydrated workbook to locate P.4 and Sales By Month
  const outWb   = XLSX.read(hydrationBuffer, { type: "buffer", cellDates: false, cellStyles: true, cellNF: true });
  const sbmName = outWb.SheetNames.find((n) => /sales.?by.?month/i.test(n));
  const p4Name  = outWb.SheetNames.find((n) => n.startsWith("P.4"));

  if (!sbmName || !outWb.Sheets[sbmName]) {
    exceptions.push({
      code: "SBM_SHEET_MISSING", severity: "advisory",
      description: "Sales By Month sheet not found in hydrated output — P.4 column reference update skipped",
      context: { sheetNames: outWb.SheetNames }, requiresOperator: false,
    });
    ctx.workbookBuffer = hydrationBuffer;
    ctx.p4ColUpdate    = null;
    events.push(audit(stage, "p4_col_update_skipped", { reason: "sbm_missing" }));
    return {
      stageId: stage, status: "complete", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: events,
    };
  }

  if (!p4Name || !outWb.Sheets[p4Name]) {
    exceptions.push({
      code: "P4_SHEET_MISSING", severity: "advisory",
      description: "P.4 sheet not found — budget and variance column update skipped",
      context: { sheetNames: outWb.SheetNames }, requiresOperator: false,
    });
    ctx.workbookBuffer = hydrationBuffer;
    ctx.p4ColUpdate    = null;
    events.push(audit(stage, "p4_col_update_skipped", { reason: "p4_missing" }));
    return {
      stageId: stage, status: "complete", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: events,
    };
  }

  // 3. Build serial → column-letter map from the SBM header row
  const sbmWs        = outWb.Sheets[sbmName];
  const serialToCol  = buildSerialToColMap(sbmWs);
  const currentSerial = ctx.snapshot?.monthSerial ?? 0;

  if (serialToCol.size === 0) {
    exceptions.push({
      code: "SBM_NO_DATE_HEADER", severity: "advisory",
      description: "Could not locate date-serial header row in Sales By Month — P.4 column update skipped",
      context: { sbmName }, requiresOperator: false,
    });
    ctx.workbookBuffer = hydrationBuffer;
    ctx.p4ColUpdate    = null;
    events.push(audit(stage, "p4_col_update_skipped", { reason: "no_date_header" }));
    return {
      stageId: stage, status: "complete", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: events,
    };
  }

  // 4. Resolve the current month's column letter (exact match, then ±1 day tolerance)
  const currentColLetter =
    serialToCol.get(currentSerial) ??
    [...serialToCol.entries()].find(([s]) => Math.abs(s - currentSerial) <= 1)?.[1];

  if (!currentColLetter) {
    // Not truly blocking — hydration still produced a workbook
    exceptions.push({
      code: "CURRENT_MONTH_COL_NOT_FOUND", severity: "operator_required",
      description: `Column for current period (${ctx.snapshot?.monthCode ?? "unknown"}) not found in Sales By Month — P.4 formulas may reference the wrong month`,
      context: { currentSerial, availableSerials: [...serialToCol.keys()] },
      candidateFix: "Verify that Sales By Month includes the current reporting period in its header row",
      requiresOperator: true,
    });
    ctx.workbookBuffer = hydrationBuffer;
    ctx.p4ColUpdate    = null;
    events.push(audit(stage, "p4_col_update_skipped", { reason: "current_col_unknown" }));
    return {
      stageId: stage, status: "complete", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: events,
    };
  }

  // 5. Detect which column P.4 currently references in its SBM formulas
  const p4Ws            = outWb.Sheets[p4Name];
  const detectedFromCol = detectP4ReferencedCol(p4Ws, sbmName);

  if (!detectedFromCol) {
    exceptions.push({
      code: "P4_NO_SBM_REFS", severity: "advisory",
      description: `P.4 contains no formula references to '${sbmName}' — column update not applicable`,
      context: { p4Name, sbmName }, requiresOperator: false,
    });
    ctx.workbookBuffer = hydrationBuffer;
    ctx.p4ColUpdate    = null;
    events.push(audit(stage, "p4_col_update_skipped", { reason: "no_refs_found" }));
    return {
      stageId: stage, status: "complete", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: events,
    };
  }

  if (detectedFromCol === currentColLetter) {
    checks.push({
      rule: "P.4 column references current month",
      expected: currentColLetter, actual: detectedFromCol,
      passed: true, severity: "INFO",
    });
    ctx.workbookBuffer = hydrationBuffer;
    ctx.p4ColUpdate    = { from: detectedFromCol, to: currentColLetter, cellsUpdated: 0 };
    events.push(audit(stage, "p4_col_already_current", { colLetter: currentColLetter, month: ctx.snapshot?.monthCode }));
    return {
      stageId: stage, status: "complete", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: events,
    };
  }

  // 6. Apply column reference update (CFO step 11: !G → !H etc.)
  // Guard: both columns must be within the SBM sheet's actual column range.
  // If detectP4ReferencedCol returns a far-right baseline column (e.g. BG/BH), it has
  // picked up a non-current-month reference and shifting it would produce #REF! errors.
  const sbmRange    = XLSX.utils.decode_range(sbmWs["!ref"] ?? "A1:A1");
  const fromColIdx  = XLSX.utils.decode_col(detectedFromCol);
  const toColIdx    = XLSX.utils.decode_col(currentColLetter);
  if (fromColIdx > sbmRange.e.c || toColIdx > sbmRange.e.c) {
    exceptions.push({
      code: "P4_COL_UPDATE_SKIPPED", severity: "advisory",
      description: `P.4 column update skipped — detected column '${detectedFromCol}' (idx ${fromColIdx}) or target '${currentColLetter}' (idx ${toColIdx}) is outside the SBM sheet column range (max idx ${sbmRange.e.c}). P.4 formula references preserved unchanged.`,
      context: { fromColIdx, toColIdx, sbmMaxCol: sbmRange.e.c }, requiresOperator: false,
    });
    ctx.workbookBuffer = hydrationBuffer;
    ctx.p4ColUpdate    = null;
    return {
      stageId: stage, status: "complete", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: events,
    };
  }
  const { sheet: updatedP4, updatedCells } = applyP4ColUpdate(p4Ws, sbmName, detectedFromCol, currentColLetter);
  outWb.Sheets[p4Name] = updatedP4;

  checks.push({
    rule:     "P.4 column references updated to current month",
    expected: currentColLetter,
    actual:   `${detectedFromCol} → ${currentColLetter}`,
    passed:   true, severity: "INFO",
  });
  events.push(audit(stage, "p4_col_refs_updated", {
    from: detectedFromCol, to: currentColLetter,
    cellsUpdated: updatedCells, month: ctx.snapshot?.monthCode,
  }));
  ctx.p4ColUpdate = { from: detectedFromCol, to: currentColLetter, cellsUpdated: updatedCells };

  // 7. Re-apply recalc flags and re-write (XLSX write strips pivot XML)
  applyRecalcFlags(outWb);
  const rawUpdated = Buffer.from(XLSX.write(outWb, { type: "buffer", bookType: "xlsx" }));

  // 8. Store raw workbook — pivot re-injection + style preservation run at the very end
  //    of runWorkflow (after injectAuditSheet) so that XLSX.write inside injectAuditSheet
  //    does not destroy the carefully-restored template styles.
  ctx.workbookBuffer = rawUpdated;
  events.push(audit(stage, "workbook_staged_for_finalization", {}));

  return {
    stageId: stage, status: "complete", durationMs: Date.now() - t0,
    reconciliationChecks: checks, exceptions, auditEvents: events,
  };
}

// ─── Stage: VALIDATE_P4_RECONCILIATION ───────────────────────────────────────

async function runValidateP4(ctx: WorkflowContext): Promise<StageResult> {
  const t0    = Date.now();
  const stage = "VALIDATE_P4_RECONCILIATION" as const;
  const exceptions: WorkflowException[]   = [];
  const events:     AuditEvent[]          = [];
  const checks:     ReconciliationCheck[] = [];

  if (!ctx.workbookBuffer) {
    // Hydration failed upstream — nothing to validate, not our fault
    return { stageId: stage, status: "complete", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: events };
  }

  const wb     = XLSX.read(ctx.workbookBuffer, { type: "buffer", cellStyles: false, cellNF: false });
  const p4Name = wb.SheetNames.find((n) => n.startsWith("P.4"));
  const sbmName = wb.SheetNames.find((n) => /sales.?by.?month/i.test(n));

  if (!p4Name) {
    checks.push({
      rule: "P.4 sheet present", expected: "present", actual: "missing",
      passed: false, severity: "WARNING",
    });
    events.push(audit(stage, "p4_absent", {}));
    return { stageId: stage, status: "complete", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: events };
  }

  const p4Ws = wb.Sheets[p4Name];

  // ── Check 1: Formula errors in P.4 ─────────────────────────────────────────
  const ERROR_VALUES = new Set(["#REF!", "#VALUE!", "#NAME?", "#NULL!", "#DIV/0!", "#N/A", "#NUM!"]);
  const errorCells: string[] = [];
  for (const addr of Object.keys(p4Ws)) {
    if (addr.startsWith("!")) continue;
    const cell = p4Ws[addr] as XLSX.CellObject;
    const vStr = String(cell.v ?? "");
    const wStr = String((cell as unknown as Record<string, unknown>).w ?? "");
    if (ERROR_VALUES.has(vStr) || ERROR_VALUES.has(wStr)) errorCells.push(addr);
  }

  if (errorCells.length > 0) {
    const sample = errorCells.slice(0, 5).join(", ");
    checks.push({
      rule:     "P.4 formula errors",
      expected: "0", actual: String(errorCells.length),
      delta:    `cells: ${sample}${errorCells.length > 5 ? ` (+${errorCells.length - 5} more)` : ""}`,
      passed:   false, severity: "ERROR",
    });
    // operator_required (not blocking) — workbook still exports with errors flagged in audit tab
    exceptions.push({
      code:        "P4_FORMULA_ERRORS",
      severity:    "operator_required",
      description: `${errorCells.length} formula error(s) in P.4 — affected cells will show #REF! or #VALUE!`,
      context:     { errorCells: errorCells.slice(0, 10) },
      candidateFix: "Open P.4 and manually correct the column references in formulas showing #REF! errors",
      requiresOperator: true,
    });
  } else {
    checks.push({
      rule: "P.4 formula errors", expected: "0", actual: "0",
      passed: true, severity: "INFO",
    });
  }

  // ── Check 2: P.4 references the current month column ───────────────────────
  if (sbmName && ctx.p4ColUpdate !== undefined && ctx.p4ColUpdate !== null) {
    const expectedCol = ctx.p4ColUpdate.to;
    const actualCol   = detectP4ReferencedCol(p4Ws, sbmName);
    if (actualCol) {
      const passed = actualCol === expectedCol;
      checks.push({
        rule:     "P.4 formulas reference current month column in Sales By Month",
        expected: expectedCol, actual: actualCol ?? "unknown",
        passed, severity: passed ? "INFO" : "ERROR",
      });
      if (!passed) {
        exceptions.push({
          code: "P4_WRONG_COL_REF", severity: "operator_required",
          description: `P.4 still references column ${actualCol} — current month data is in column ${expectedCol}`,
          context: { expectedCol, actualCol, month: ctx.snapshot?.monthCode },
          candidateFix: `Manually find-and-replace !${actualCol} with !${expectedCol} in P.4 formula references`,
          requiresOperator: true,
        });
      }
    }
  }

  // ── Check 3: Gross profit source value present (structural precondition for step 12) ──
  if (ctx.sales) {
    const { lookupByLabel, currentMonthIdx, monthCodes } = ctx.sales;
    const monthCode = monthCodes[currentMonthIdx];
    const gpKeys = ["gross profit", "total gross profit", "gross margin"];
    let expectedGP: number | null = null;
    let gpKey = "";
    for (const k of gpKeys) {
      const row = lookupByLabel.get(k);
      if (row) { expectedGP = row[currentMonthIdx] ?? null; gpKey = k; break; }
    }
    if (expectedGP == null) {
      checks.push({
        rule: "Gross profit row present in Sales By Month",
        expected: "row found", actual: "not found",
        passed: false, severity: "WARNING",
      });
    } else if (expectedGP === 0) {
      checks.push({
        rule: "Gross profit non-zero for current month",
        expected: "> 0", actual: "$0",
        passed: false, severity: "WARNING",
      });
    } else {
      const fmt = `$${Math.round(expectedGP).toLocaleString("en-US")}`;
      checks.push({
        rule: `Gross profit available for ${monthCode}`,
        expected: fmt, actual: fmt,
        passed: true, severity: "INFO",
      });
      events.push(audit(stage, "gross_profit_confirmed", { key: gpKey, month: monthCode, value: expectedGP }));
    }
  }

  // ── Check 4: Named range integrity ─────────────────────────────────────────
  const namedRanges = (wb.Workbook as Record<string, unknown> | undefined)?.Names;
  if (Array.isArray(namedRanges) && namedRanges.length > 0) {
    const broken = namedRanges.filter((nm) => {
      const ref   = String((nm as Record<string, unknown>).Ref ?? "");
      const match = ref.match(/^'?([^'!]+)'?!/);
      return match ? !wb.SheetNames.includes(match[1]) : false;
    });
    checks.push({
      rule: "Named range integrity", expected: "0 broken",
      actual: `${broken.length} broken`, passed: broken.length === 0,
      severity: broken.length > 0 ? "ERROR" : "INFO",
    });
  }

  events.push(audit(stage, "p4_validation_complete", {
    checksRun: checks.length,
    passed:    checks.filter((c) => c.passed).length,
    failed:    checks.filter((c) => !c.passed).length,
  }));

  // Always complete — issues surface in audit tab, never block export
  return {
    stageId: stage, status: "complete", durationMs: Date.now() - t0,
    reconciliationChecks: checks, exceptions, auditEvents: events,
  };
}

// ─── Stage: FINALIZE_WORKBOOK ─────────────────────────────────────────────────

async function runFinalize(ctx: WorkflowContext): Promise<StageResult> {
  const t0    = Date.now();
  const stage = "FINALIZE_WORKBOOK" as const;
  const exceptions: WorkflowException[]   = [];
  const events:     AuditEvent[]          = [];
  const checks:     ReconciliationCheck[] = [];

  if (!ctx.workbookBuffer || !ctx.sales || !ctx.occupancy || !ctx.snapshot) {
    // Context gap — log it but don't crash the pipeline
    exceptions.push({
      code: "MISSING_CONTEXT", severity: "operator_required",
      description: "Reconciliation skipped — required source data was not available from prior stages",
      context: {
        hasBuffer: !!ctx.workbookBuffer, hasSales: !!ctx.sales,
        hasOccupancy: !!ctx.occupancy,  hasSnapshot: !!ctx.snapshot,
      },
      requiresOperator: false,
    });
    return { stageId: stage, status: "complete", durationMs: Date.now() - t0,
      reconciliationChecks: checks, exceptions, auditEvents: events };
  }

  const allSnapshots = [
    ...(ctx.historicalSnapshots ?? []),
    ...((ctx.historicalSnapshots ?? []).some((s) => s.monthCode === ctx.snapshot!.monthCode)
      ? [] : [ctx.snapshot!]),
  ];

  const reconciliation = reconcileOutput(
    ctx.workbookBuffer, ctx.sales, ctx.occupancy, ctx.snapshot,
    allSnapshots, ctx.budgetSource ?? null, "hydration",
  );
  ctx.reconciliation = reconciliation;

  // Map reconciliation findings into our check format
  for (const f of reconciliation.findings) {
    checks.push({
      rule:     `[${f.category}] ${f.description}`,
      expected: f.expected ?? "pass",
      actual:   f.actual   ?? (f.severity === "INFO" ? "ok" : "fail"),
      delta:    f.delta,
      passed:   f.severity === "INFO",
      severity: f.severity,
    });
  }

  // Critical reconciliation findings → operator_required exception (NOT blocking)
  // The finance team sees these in the audit tab and can investigate before distributing
  const criticals = reconciliation.findings.filter((f) => f.severity === "CRITICAL");
  if (criticals.length > 0) {
    exceptions.push({
      code:        "RECONCILIATION_CRITICAL",
      severity:    "operator_required",
      description: `${criticals.length} reconciliation item(s) require review before distributing — ${reconciliation.summary}`,
      context: {
        summary:  reconciliation.summary,
        findings: criticals.map((f) => ({
          category: f.category, sheet: f.sheet, cell: f.cell,
          description: f.description, expected: f.expected, actual: f.actual,
        })),
      },
      candidateFix: "Review flagged items in the Dashboard Status tab and correct source data before distributing the workbook",
      requiresOperator: true,
    });
  }

  events.push(audit(stage, "reconciliation_complete", {
    summary:           reconciliation.summary,
    isClean:           reconciliation.isClean,
    criticalCount:     reconciliation.criticalCount,
    errorCount:        reconciliation.errorCount,
    warningCount:      reconciliation.warningCount,
    staleFormulaCount: reconciliation.staleFormulaCount,
  }));

  events.push(audit(stage, "workbook_finalized", {
    runId:        ctx.runId,
    month:        ctx.snapshot.monthCode,
    location:     ctx.snapshot.location,
    bufferSizeKb: Math.round(ctx.workbookBuffer.length / 1024),
    p4ColUpdate:  ctx.p4ColUpdate,
  }));

  // Always complete — issues surface in audit tab
  return {
    stageId: stage, status: "complete", durationMs: Date.now() - t0,
    reconciliationChecks: checks, exceptions, auditEvents: events,
  };
}

// ─── Generation status ────────────────────────────────────────────────────────

// Codes that represent truly unrecoverable failures (no workbook can be produced)
const TRUE_BLOCKING_CODES = new Set(["INGEST_PARSE_FAILED", "HYDRATION_FAILED"]);

function computeGenerationStatus(stageResults: StageResult[]): GenerationStatus {
  const allChecks     = stageResults.flatMap((s) => s.reconciliationChecks);
  const allExceptions = stageResults.flatMap((s) => s.exceptions);

  // True blocking → CRITICAL (handled separately in orchestrator, but defensive check)
  if (allExceptions.some((e) => TRUE_BLOCKING_CODES.has(e.code))) return "CRITICAL";

  // Formula crashes, reconciliation criticals, wrong column references → PARTIAL
  if (
    allExceptions.some((e) => e.code === "P4_FORMULA_ERRORS" || e.code === "RECONCILIATION_CRITICAL") ||
    allChecks.some((c) => !c.passed && (c.severity === "CRITICAL" || c.severity === "ERROR"))
  ) return "PARTIAL";

  // Operator action required on non-critical items → PARTIAL
  if (allExceptions.some((e) => e.severity === "operator_required")) return "PARTIAL";

  // Warnings only → WARNING
  if (allChecks.some((c) => !c.passed && c.severity === "WARNING")) return "WARNING";

  return "COMPLETE";
}

// ─── Dashboard Status sheet ───────────────────────────────────────────────────

const STAGE_LABELS: Record<WorkflowStageId, string> = {
  "INGEST_SOURCES":             "Data Ingestion",
  "HYDRATE_P4_VOYAGER":         "Budget Hydration (P.4)",
  "VALIDATE_P4_RECONCILIATION": "P.4 Reconciliation",
  "FINALIZE_WORKBOOK":          "Workbook Finalization",
};

const STATUS_META: Record<GenerationStatus, { bg: string; fg: string; label: string }> = {
  COMPLETE: { bg: "16A34A", fg: "FFFFFF", label: "COMPLETE — Ready to Distribute"  },
  WARNING:  { bg: "D97706", fg: "FFFFFF", label: "REVIEW RECOMMENDED"               },
  PARTIAL:  { bg: "DC2626", fg: "FFFFFF", label: "ACTION REQUIRED — Review Issues"  },
  CRITICAL: { bg: "7F1D1D", fg: "FFFFFF", label: "GENERATION FAILED"                },
};

const SEV_COLOR: Record<string, { bg: string; fg: string }> = {
  CRITICAL: { bg: "991B1B", fg: "FFFFFF" },
  ERROR:    { bg: "DC2626", fg: "FFFFFF" },
  WARNING:  { bg: "D97706", fg: "FFFFFF" },
  INFO:     { bg: "6B7280", fg: "FFFFFF" },
};

interface IssueRow {
  severity: "CRITICAL" | "ERROR" | "WARNING" | "INFO";
  stage:    string;
  description: string;
  action:   string;
}

function cleanRule(rule: string): string {
  return rule.replace(/^\[[\w\s]+\]\s*/, "");
}

function actionFor(rule: string): string {
  const r = rule.toLowerCase();
  if (r.includes("formula error") || r.includes("#ref"))
    return "Open P.4 and correct any #REF! errors — the column reference update may need manual adjustment";
  if (r.includes("column references") && r.includes("current month"))
    return "Column references updated automatically — verify P.4 on first open in Excel";
  if (r.includes("column references"))
    return "Manually update P.4 formula references to point to the current reporting month column";
  if (r.includes("gross profit"))
    return "Confirm Sales By Month has a Gross Profit row matching the current period";
  if (r.includes("named range"))
    return "Resolve broken named range references before distributing the workbook";
  if (r.includes("budget") || r.includes("variance"))
    return "Upload Voyager budget export and verify the P.4 variance column";
  if (r.includes("occupancy"))
    return "Upload Kube Occupancy report and verify P.2/P.3 data alignment";
  if (r.includes("revenue") || r.includes("income"))
    return "Verify the Sales By Month total income row matches the Voyager P&L total";
  if (r.includes("snapshot"))
    return "Verify current month snapshot data was written correctly to the P.1 section";
  if (r.includes("membership") || r.includes("check figure"))
    return "Review income detail for cross-month charge reversals on this membership tab";
  if (r.includes("p.4 sheet") || r.includes("p4 sheet"))
    return "Upload Voyager P&L export to populate the P.4 budget section";
  if (r.includes("month column") || r.includes("reporting period"))
    return "Verify the source file contains data for the current reporting period";
  return "Review with finance team before distributing";
}

function collectIssueRows(stageResults: StageResult[]): IssueRow[] {
  const ORDER: Record<string, number> = { CRITICAL: 0, ERROR: 1, WARNING: 2, INFO: 3 };
  const rows: IssueRow[] = [];

  for (const stage of stageResults) {
    const label = STAGE_LABELS[stage.stageId];

    for (const check of stage.reconciliationChecks) {
      if (check.passed || check.severity === "INFO") continue;
      let desc = cleanRule(check.rule);
      if (
        check.actual && check.expected &&
        check.actual !== check.expected &&
        check.actual !== "fail" && check.expected !== "pass"
      ) {
        desc += ` (Expected: ${check.expected} / Found: ${check.actual})`;
      }
      rows.push({
        severity:    check.severity,
        stage:       label,
        description: desc,
        action:      actionFor(check.rule),
      });
    }

    for (const exc of stage.exceptions) {
      if (exc.severity === "advisory") continue;
      rows.push({
        severity:    exc.severity === "blocking" ? "CRITICAL" : "ERROR",
        stage:       label,
        description: exc.description,
        action:      exc.candidateFix ?? "Review with finance team",
      });
    }
  }

  return rows.sort((a, b) => (ORDER[a.severity] ?? 9) - (ORDER[b.severity] ?? 9));
}

function buildAuditWorksheet(
  stageResults:     StageResult[],
  ctx:              Readonly<WorkflowContext>,
  generationStatus: GenerationStatus,
): XLSX.WorkSheet {
  const ws:     XLSX.WorkSheet          = {};
  const merges: XLSX.Range[]            = [];
  const rowH:   Record<number, number>  = {};

  // Brand palette
  const NAVY      = "0F2044";
  const GOLD      = "C9A84C";
  const WHITE     = "FFFFFF";
  const LT_GRAY   = "F3F4F6";
  const MED_GRAY  = "E5E7EB";
  const DARK      = "111827";
  const STRIPE    = "F8FAFC";

  const sc = STATUS_META[generationStatus];

  function setCell(r: number, c: number, cell: XLSX.CellObject): void {
    ws[XLSX.utils.encode_cell({ r, c })] = cell;
  }

  function mkCell(
    v: string | number,
    opts: {
      bg?:     string;
      fg?:     string;
      bold?:   boolean;
      sz?:     number;
      wrap?:   boolean;
      center?: boolean;
      right?:  boolean;
      border?: boolean;
    } = {},
  ): XLSX.CellObject {
    return {
      v,
      t: typeof v === "number" ? "n" : "s",
      s: {
        font: {
          name:  "Calibri",
          sz:    opts.sz ?? 10,
          bold:  opts.bold ?? false,
          color: { rgb: opts.fg ?? DARK },
        },
        fill: opts.bg ? { fgColor: { rgb: opts.bg } } : undefined,
        alignment: {
          horizontal: opts.center ? "center" : opts.right ? "right" : "left",
          vertical:   "center",
          wrapText:   opts.wrap ?? false,
        },
        border: opts.border ? {
          top:    { style: "thin" as const, color: { rgb: MED_GRAY } },
          bottom: { style: "thin" as const, color: { rgb: MED_GRAY } },
          left:   { style: "thin" as const, color: { rgb: MED_GRAY } },
          right:  { style: "thin" as const, color: { rgb: MED_GRAY } },
        } : undefined,
      },
    };
  }

  let r = 0;

  const location   = ctx.snapshot?.location ?? "—";
  const monthCount = ctx.sales?.monthSerials.length ?? 0;
  const hasBudget  = ctx.budgetSource !== null && ctx.budgetSource !== undefined;
  const histCount  = ctx.historicalSnapshots.length;
  const hasOccupancy = (ctx.occupancy !== undefined && ctx.occupancy !== null);

  // Completion % — count hydrated sections out of 4 expected (SBM, Occupancy, Budget, Historical)
  const hydratedCount = [
    monthCount > 0,
    hasOccupancy,
    hasBudget,
    histCount > 0,
  ].filter(Boolean).length;
  const completionPct = Math.round((hydratedCount / 4) * 100 / 25) * 25;

  // ── Row 0: Main header ──────────────────────────────────────────────────────
  const period  = ctx.snapshot?.monthCode ?? "Unknown";
  const headerTitle = `DASHBOARD STATUS  —  25N Coworking  —  ${period}`;
  setCell(r, 0, mkCell(headerTitle, { bg: NAVY, fg: GOLD, bold: true, sz: 13 }));
  merges.push({ s: { r, c: 0 }, e: { r, c: 5 } });
  rowH[r] = 30; r++;

  // ── Row 1: Status badge + completion % + timestamp ─────────────────────────
  setCell(r, 0, mkCell(sc.label, { bg: sc.bg, fg: sc.fg, bold: true, sz: 9, center: true }));
  merges.push({ s: { r, c: 0 }, e: { r, c: 1 } });
  setCell(r, 2, mkCell(`Completion: ${completionPct}%`, { bg: LT_GRAY, center: true }));
  const genTime = new Date().toLocaleString("en-US", {
    month: "short", day: "numeric", year: "numeric", hour: "2-digit", minute: "2-digit",
  });
  setCell(r, 3, mkCell(`Generated: ${genTime}`, { bg: LT_GRAY }));
  merges.push({ s: { r, c: 3 }, e: { r, c: 4 } });
  setCell(r, 5, mkCell("Template Hydration", { bg: LT_GRAY, center: true }));
  rowH[r] = 22; r++;

  // ── Row 2: Location + source info ──────────────────────────────────────────
  setCell(r, 0, mkCell(`Location: ${location}`, { bg: LT_GRAY }));
  setCell(r, 2, mkCell(`Source Periods: ${monthCount}`, { bg: LT_GRAY }));
  setCell(r, 3, mkCell(`Files Uploaded: ${ctx.sourceBuffers.length}`, { bg: LT_GRAY }));
  setCell(r, 5, mkCell(`Historical: ${histCount} mo`, { bg: LT_GRAY, center: true }));
  rowH[r] = 20; r++;

  // ── Spacer ──────────────────────────────────────────────────────────────────
  rowH[r] = 6; r++;

  // ── SOURCE FILES section ────────────────────────────────────────────────────
  setCell(r, 0, mkCell("SOURCE FILES", { bg: NAVY, fg: WHITE, bold: true, sz: 9 }));
  merges.push({ s: { r, c: 0 }, e: { r, c: 5 } });
  rowH[r] = 22; r++;

  const srcHdrs = ["File Name", "Detected Type", "Sections Updated", "Status", "", ""];
  for (let c = 0; c < 4; c++) {
    setCell(r, c, mkCell(srcHdrs[c], { bg: GOLD, fg: NAVY, bold: true, sz: 9, center: c > 2 }));
  }
  rowH[r] = 18; r++;

  // Rows for each uploaded source
  const classifications = ctx.sourceClassifications ?? [];
  const fileNames       = ctx.sourceFileNames ?? [];
  for (let i = 0; i < classifications.length; i++) {
    const cl       = classifications[i];
    const rowBg    = i % 2 === 0 ? undefined : STRIPE;
    const caps     = cl.capabilities.length > 0
      ? cl.capabilities.map((c) => c.replace(/_/g, " ")).join(", ")
      : "—";
    const fileName = fileNames[i] ?? `File ${i + 1}`;
    const typeLabel = cl.sourceType === "UNKNOWN"
      ? "Unrecognized — no matching data extracted"
      : SOURCE_TYPE_LABELS[cl.sourceType];
    setCell(r, 0, mkCell(fileName,  { bg: rowBg, sz: 9, border: true }));
    setCell(r, 1, mkCell(typeLabel, { bg: rowBg, sz: 9, border: true }));
    setCell(r, 2, mkCell(caps,      { bg: rowBg, sz: 9, wrap: true, border: true }));
    const isUnknown = cl.sourceType === "UNKNOWN";
    setCell(r, 3, mkCell(isUnknown ? "No Data Extracted" : "Uploaded",
      { bg: isUnknown ? "D97706" : "16A34A", fg: WHITE, bold: true, sz: 9, center: true, border: true }));
    merges.push({ s: { r, c: 3 }, e: { r, c: 5 } });
    rowH[r] = 22; r++;
  }

  // Rows for MISSING sources (what wasn't uploaded)
  const missingRows: Array<{ name: string; hint: string }> = [];
  const allCaps = new Set(classifications.flatMap((c) => c.capabilities));
  if (!allCaps.has("SALES_BY_MONTH")) missingRows.push({ name: "Voyager P&L Export",    hint: "Required for P.2 revenue + Sales By Month tab" });
  if (!allCaps.has("OCCUPANCY"))      missingRows.push({ name: "Kube Occupancy Report",  hint: "Required for P.3 occupancy section" });
  if (!allCaps.has("BUDGET"))         missingRows.push({ name: "Voyager Budget Export",  hint: "Required for P.4 budget variance" });

  for (let i = 0; i < missingRows.length; i++) {
    const m = missingRows[i];
    const rowBg = (classifications.length + i) % 2 === 0 ? undefined : STRIPE;
    setCell(r, 0, mkCell("—", { bg: rowBg, sz: 9, border: true }));
    setCell(r, 1, mkCell(m.name, { bg: rowBg, sz: 9, border: true }));
    setCell(r, 2, mkCell(m.hint, { bg: rowBg, sz: 9, wrap: true, border: true }));
    setCell(r, 3, mkCell("Not Uploaded", { bg: "D97706", fg: WHITE, bold: true, sz: 9, center: true, border: true }));
    merges.push({ s: { r, c: 3 }, e: { r, c: 5 } });
    rowH[r] = 22; r++;
  }

  rowH[r] = 8; r++;

  // ── Issues section (ERROR and CRITICAL only) ────────────────────────────────
  const allIssues = collectIssueRows(stageResults);
  const issues = allIssues.filter((i) => i.severity === "CRITICAL" || i.severity === "ERROR");

  if (allIssues.length === 0) {
    setCell(r, 0, mkCell("All checks passed — workbook is ready for distribution.", {
      bg: "16A34A", fg: WHITE, bold: true,
    }));
    merges.push({ s: { r, c: 0 }, e: { r, c: 5 } });
    rowH[r] = 24; r++;
  } else if (issues.length > 0) {
    setCell(r, 0, mkCell("ISSUES REQUIRING ATTENTION", { bg: NAVY, fg: WHITE, bold: true, sz: 9 }));
    merges.push({ s: { r, c: 0 }, e: { r, c: 5 } });
    rowH[r] = 22; r++;

    const issueHdrs = ["#", "Category", "Issue", "What To Do", "", ""];
    for (let c = 0; c < 4; c++) {
      setCell(r, c, mkCell(issueHdrs[c], { bg: GOLD, fg: NAVY, bold: true, sz: 9, center: c < 2 }));
    }
    rowH[r] = 18; r++;

    // CRITICAL first, then ERROR
    const sorted = [...issues].sort((a, b) => (a.severity === "CRITICAL" ? 0 : 1) - (b.severity === "CRITICAL" ? 0 : 1));
    for (let i = 0; i < sorted.length; i++) {
      const issue  = sorted[i];
      const ss     = SEV_COLOR[issue.severity] ?? SEV_COLOR["INFO"];
      const rowBg  = i % 2 === 0 ? undefined : STRIPE;
      setCell(r, 0, mkCell(String(i + 1), { bg: ss.bg, fg: ss.fg, bold: true, sz: 9, center: true, border: true }));
      setCell(r, 1, mkCell(issue.stage,        { bg: rowBg, sz: 9, border: true }));
      setCell(r, 2, mkCell(issue.description,  { bg: rowBg, sz: 9, wrap: true, border: true }));
      setCell(r, 3, mkCell(issue.action,        { bg: rowBg, sz: 9, wrap: true, border: true }));
      merges.push({ s: { r, c: 3 }, e: { r, c: 5 } });
      rowH[r] = 52; r++;
    }
    rowH[r] = 8; r++;
  }

  // ── WORKBOOK SECTIONS ───────────────────────────────────────────────────────
  setCell(r, 0, mkCell("WORKBOOK SECTIONS", { bg: NAVY, fg: WHITE, bold: true, sz: 9 }));
  merges.push({ s: { r, c: 0 }, e: { r, c: 5 } });
  rowH[r] = 22; r++;

  const secHdrs = ["Section", "Tab Name", "Status", "Detail"];
  const secColWidths = [0, 1, 2, 3];
  for (const c of secColWidths) {
    setCell(r, c, mkCell(secHdrs[c], { bg: GOLD, fg: NAVY, bold: true, sz: 9, center: c > 1 }));
  }
  rowH[r] = 18; r++;

  const today = new Date().toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });

  const sections: Array<{ name: string; tab: string; ok: boolean; note: string }> = [
    {
      name: "P.1 Dashboard", tab: "P1. Dashboard",
      ok: true,
      note: "Snapshot rows 18+ updated — review occupancy %, contracts, flex members",
    },
    {
      name: "P.2 Monthly Income", tab: "P2. Monthly Income",
      ok: monthCount > 0,
      note: monthCount > 0 ? `Revenue pulled from ${monthCount} period(s) via Sales By Month formulas` : "Template version preserved — upload Voyager P&L to populate",
    },
    {
      name: "P.3 Occupancy", tab: "P3. Occupancy",
      ok: hasOccupancy,
      note: hasOccupancy ? "Space-level occupancy data loaded from Kube report" : "Template version preserved — upload Kube Occupancy report to update",
    },
    {
      name: "P.4 Budget-to-Actual", tab: "P4. Budget-to-Actual",
      ok: hasBudget,
      note: hasBudget ? "Budget column loaded; variance formulas reference current month column" : "Template version preserved — upload Voyager Budget export to populate",
    },
    {
      name: "Sales By Month", tab: "2026 Sales By Month",
      ok: monthCount > 0,
      note: monthCount > 0 ? `${monthCount} month column(s), ${ctx.sales?.lookupByLabel.size ?? 0} line items mapped from source` : "Template version preserved — upload Voyager P&L export to update",
    },
  ];

  for (let i = 0; i < sections.length; i++) {
    const sec   = sections[i];
    const rowBg = i % 2 === 0 ? undefined : STRIPE;
    setCell(r, 0, mkCell(sec.name, { bg: rowBg, sz: 9, bold: true, border: true }));
    setCell(r, 1, mkCell(sec.tab,  { bg: rowBg, sz: 9, border: true }));
    setCell(r, 2, mkCell(sec.ok ? "✓ Hydrated" : "Template",
      { bg: sec.ok ? "16A34A" : LT_GRAY, fg: sec.ok ? WHITE : DARK, bold: sec.ok, sz: 9, center: true, border: true }));
    setCell(r, 3, mkCell(sec.note, { bg: rowBg, sz: 9, wrap: true, border: true }));
    merges.push({ s: { r, c: 3 }, e: { r, c: 5 } });
    rowH[r] = 32; r++;
  }

  rowH[r] = 8; r++;

  // ── Footer note ─────────────────────────────────────────────────────────────
  setCell(r, 0, mkCell(
    `Generated ${today}  ·  Open each section tab to verify data before distributing.`,
    { bg: LT_GRAY, fg: "6B7280", sz: 8 },
  ));
  merges.push({ s: { r, c: 0 }, e: { r, c: 5 } });
  rowH[r] = 16; r++;

  // ── Sheet metadata ──────────────────────────────────────────────────────────
  ws["!cols"] = [
    { wch: 42 },  // A — File Name (long filenames need space)
    { wch: 32 },  // B — Detected Type
    { wch: 36 },  // C — Sections Updated / Issue description
    { wch: 48 },  // D — Status / What To Do (longest text column)
    { wch: 4  },  // E — merge overflow
    { wch: 4  },  // F — merge overflow
  ];

  const rowsArr: (XLSX.RowInfo | undefined)[] = [];
  for (let i = 0; i < r; i++) rowsArr[i] = rowH[i] ? { hpx: rowH[i] } : { hpx: 18 };
  ws["!rows"]   = rowsArr as XLSX.RowInfo[];
  ws["!merges"] = merges;
  ws["!ref"]    = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: r - 1, c: 5 } });

  return ws;
}

function injectAuditSheet(
  workbookBuffer: Buffer,
  auditWs:        XLSX.WorkSheet,
): Buffer {
  const wb    = XLSX.read(workbookBuffer, { type: "buffer", cellDates: false, cellStyles: true, cellNF: true });
  const newWb = XLSX.utils.book_new();

  // Dashboard Status goes FIRST — it's the landing tab reviewers open first
  XLSX.utils.book_append_sheet(newWb, auditWs, "Dashboard Status");
  for (const name of wb.SheetNames) {
    if (name === "Dashboard Status") continue; // drop any pre-existing copy
    XLSX.utils.book_append_sheet(newWb, wb.Sheets[name], name);
  }

  // Preserve workbook-level metadata (named ranges, calc settings)
  if (wb.Workbook) newWb.Workbook = wb.Workbook;

  applyRecalcFlags(newWb);
  // Return raw XLSX buffer only — pivot re-injection and style preservation
  // are applied by the caller (runWorkflow) after this function returns.
  return Buffer.from(XLSX.write(newWb, { type: "buffer", bookType: "xlsx" }));
}

// ─── Workflow orchestrator ────────────────────────────────────────────────────

const STAGE_DEFINITIONS: Array<{
  id: WorkflowStageId;
  fn: (ctx: WorkflowContext) => Promise<StageResult>;
}> = [
  { id: "INGEST_SOURCES",             fn: runIngestSources },
  { id: "HYDRATE_P4_VOYAGER",         fn: runHydrateP4    },
  { id: "VALIDATE_P4_RECONCILIATION", fn: runValidateP4   },
  { id: "FINALIZE_WORKBOOK",          fn: runFinalize      },
];

export async function runWorkflow(input: WorkflowInput): Promise<WorkflowResult> {
  const t0    = Date.now();
  const runId = generateRunId();

  const ctx: WorkflowContext = {
    runId,
    templateBuffer:        input.templateBuffer,
    sourceBuffers:         input.sourceBuffers,
    sourceFileNames:       input.sourceFileNames ?? [],
    historicalSnapshots:   input.historicalSnapshots ?? [],
    overrides:             input.overrides ?? {},
    sourceClassifications: [],
  };

  const stageResults: StageResult[] = [];
  const auditLog:     AuditEvent[]  = [];

  for (const { id, fn } of STAGE_DEFINITIONS) {
    let result: StageResult;
    try {
      result = await fn(ctx);
    } catch (e) {
      result = {
        stageId: id, status: "failed", durationMs: 0,
        reconciliationChecks: [],
        exceptions: [{
          code: "STAGE_UNHANDLED_ERROR", severity: "blocking",
          description: `Unexpected error in ${STAGE_LABELS[id]}: ${e instanceof Error ? e.message : String(e)}`,
          context: {}, requiresOperator: false,
        }],
        auditEvents: [audit(id, "unhandled_error", { error: String(e) })],
      };
    }

    stageResults.push(result);
    auditLog.push(...result.auditEvents);

    // Only abort for conditions where producing any workbook is impossible
    const trueBlock = result.exceptions.find((e) => TRUE_BLOCKING_CODES.has(e.code));
    if (trueBlock) {
      return {
        runId,
        status:             "complete",
        generationStatus:   "CRITICAL",
        stages:             stageResults,
        totalDurationMs:    Date.now() - t0,
        auditLog,
        blockedBy:          trueBlock.description,
        sourceType:         ctx.sourceClassification?.sourceType,
        sourceCapabilities: ctx.sourceClassification?.capabilities,
        sourceLabel:        ctx.sourceClassification
          ? SOURCE_TYPE_LABELS[ctx.sourceClassification.sourceType]
          : undefined,
      };
    }
    // All other issues: continue — they surface in the audit sheet
  }

  // ── Build Data Bridge output ──────────────────────────────────────────────
  // Instead of trying to clone and surgically modify the complex master template
  // (which causes Excel corruption across 8+ OOXML failure modes), we output a
  // simple Data Bridge Excel file.  A VBA macro in the master .xlsm reads this
  // file and populates the dashboard natively — zero format corruption possible.
  const dataBridgeBuffer = buildDataBridge(ctx);
  const generationStatus = computeGenerationStatus(stageResults);

  return {
    runId,
    status:              "complete",
    generationStatus,
    stages:              stageResults,
    outputBuffer:        dataBridgeBuffer,
    snapshot:            ctx.snapshot,
    reconciliation:      ctx.reconciliation,
    totalDurationMs:     Date.now() - t0,
    auditLog,
    sourceType:          ctx.sourceClassification?.sourceType,
    sourceCapabilities:  ctx.sourceClassification?.capabilities,
    sourceLabel:         ctx.sourceClassification
      ? SOURCE_TYPE_LABELS[ctx.sourceClassification.sourceType]
      : undefined,
    sectionResults:      ctx.sectionResults,
    occRows:             ctx.occupancy?.rawRows.length ?? 0,
    sbmRows:             ctx.sales?.lookupByLabel.size ?? 0,
  };
}
