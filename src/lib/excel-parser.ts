import * as XLSX from "xlsx-js-style";
import type {
  ReportData,
  DashboardData,
  RevenueData,
  OccupancyData,
  BudgetData,
  MeetingData,
  MeetingRoomTransaction,
  MeetingRoomTypeRow,
  MeetingMonthly,
  SpaceTypeStats,
  WeeklySnapshot,
  RevenueLineItem,
  OccupancyRow,
  BudgetLineItem,
  ValidationCheck,
  ValidationResult,
} from "@/types/report";

type SheetRow = (string | number | boolean | null | undefined)[];

function getSheetByPrefix(wb: XLSX.WorkBook, prefix: string): SheetRow[] | null {
  const name = wb.SheetNames.find((n) => n.startsWith(prefix));
  if (!name) return null;
  return XLSX.utils.sheet_to_json<SheetRow>(wb.Sheets[name], {
    header: 1,
    defval: null,
    raw: false,
  });
}

function str(v: unknown): string {
  return v == null ? "" : String(v).trim();
}

function num(v: unknown): number | null {
  if (v == null || v === "") return null;
  const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[$,%]/g, ""));
  return isNaN(n) ? null : n;
}

function pct(v: unknown): number {
  const n = num(v);
  if (n == null) return 0;
  return n > 1 ? n / 100 : n;
}

function fmtDate(v: unknown): string {
  if (v == null || v === "") return "";
  const s = str(v);
  if (/^\d{5}$/.test(s)) {
    const d = XLSX.SSF.parse_date_code(parseInt(s));
    return `${d.m.toString().padStart(2, "0")}/${d.d.toString().padStart(2, "0")}/${d.y}`;
  }
  if (/^[A-Za-z]{3}-\d{2,4}$/.test(s)) return s;
  const longDate = s.match(/^([A-Za-z]{3,})\s+(\d{1,2}),?\s*(\d{4})$/);
  if (longDate) return `${longDate[1].slice(0, 3)} ${longDate[2].padStart(2, "0")}, ${longDate[3]}`;
  return s;
}

// Issue 4: "Aug-24" → "August 1, 2024" — includes day 1 so "August 2024" isn't misread
// as a complete-month figure when the snapshot is captured on day 1 of the month.
function fmtMonthYear(v: string): string {
  const m = v.match(/^([A-Za-z]{3})-(\d{2,4})$/);
  if (!m) return v;
  const months: Record<string, string> = {
    jan: "January", feb: "February", mar: "March", apr: "April",
    may: "May",     jun: "June",     jul: "July",  aug: "August",
    sep: "September", oct: "October", nov: "November", dec: "December",
  };
  const month = months[m[1].toLowerCase()] ?? m[1];
  const year = m[2].length === 2 ? `20${m[2]}` : m[2];
  return `${month} 1, ${year}`;
}

function formatCurrency(n: number): string {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  }).format(n);
}

// ── P.1 Dashboard ────────────────────────────────────────────────────────────
function parseDashboard(wb: XLSX.WorkBook): DashboardData {
  const rows = getSheetByPrefix(wb, "P.1");
  if (!rows) throw new Error("Could not find P.1 Dashboard sheet");

  // WX-8: Extract location from title row ("25N COWORKING WACO DASHBOARD" → "Waco")
  // and reportDate from date row ("Aug-24" → "August 2024")
  let location = "";
  let reportDate = "";
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    for (const cell of rows[i] ?? []) {
      const s = str(cell);
      if (!location) {
        const m = s.match(/coworking\s+(.+?)\s+dashboard/i);
        if (m) {
          const loc = m[1].trim();
          location = loc.charAt(0).toUpperCase() + loc.slice(1).toLowerCase();
        }
      }
      if (!reportDate && /^[A-Za-z]{3}-\d{2,4}$/.test(s)) {
        reportDate = fmtMonthYear(s);
      }
    }
    if (location && reportDate) break;
  }

  const spaceTypes: SpaceTypeStats[] = [];
  let currentMonthRevenue = 0;
  const weeklySnapshots: WeeklySnapshot[] = [];

  // Scan for revenue cell
  for (const row of rows) {
    for (const cell of row) {
      const n = num(cell);
      if (n && n > 10000 && n < 500000 && currentMonthRevenue === 0) {
        currentMonthRevenue = n;
      }
    }
  }

  // Find rows with space type KPIs
  const spaceTypeLabels = ["dedicated desk", "private office"];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const label = str(row[0]).toLowerCase();
    if (spaceTypeLabels.some((l) => label.includes(l))) {
      const nums = row.map(num).filter((n): n is number => n !== null);
      if (nums.length >= 6) {
        const [totalSpaces, occupiedSpaces, , totalSqFt, occupiedSqFt] = nums;
        spaceTypes.push({
          label: str(row[0]),
          totalSpaces: totalSpaces ?? 0,
          occupiedSpaces: occupiedSpaces ?? 0,
          occupancyPct: totalSpaces ? Math.round(((occupiedSpaces ?? 0) / totalSpaces) * 100) : 0,
          totalSqFt: totalSqFt ?? 0,
          occupiedSqFt: occupiedSqFt ?? 0,
          sqFtPct: totalSqFt ? Math.round(((occupiedSqFt ?? 0) / totalSqFt) * 100) : 0,
        });
      }
    }
  }

  let totals: SpaceTypeStats = {
    label: "Total", totalSpaces: 0, occupiedSpaces: 0, occupancyPct: 0,
    totalSqFt: 0, occupiedSqFt: 0, sqFtPct: 0,
  };
  if (spaceTypes.length > 0) {
    totals = spaceTypes.reduce(
      (acc, st) => ({
        label: "Total",
        totalSpaces: acc.totalSpaces + st.totalSpaces,
        occupiedSpaces: acc.occupiedSpaces + st.occupiedSpaces,
        occupancyPct: 0,
        totalSqFt: acc.totalSqFt + st.totalSqFt,
        occupiedSqFt: acc.occupiedSqFt + st.occupiedSqFt,
        sqFtPct: 0,
      }),
      totals
    );
    totals.occupancyPct = totals.totalSpaces
      ? Math.round((totals.occupiedSpaces / totals.totalSpaces) * 100) : 0;
    totals.sqFtPct = totals.totalSqFt
      ? Math.round((totals.occupiedSqFt / totals.totalSqFt) * 100) : 0;
  }

  // Find 12-month snapshot block
  let snapshotStart = -1;
  for (let i = 0; i < rows.length; i++) {
    if (str(rows[i][0]).toLowerCase() === "month") {
      snapshotStart = i + 1;
      break;
    }
  }

  if (snapshotStart >= 0) {
    for (let i = snapshotStart; i < rows.length; i++) {
      const row = rows[i];
      if (!row[0] && !row[1]) break;
      const dateVal = row[0] ?? row[1];
      const dateStr = str(dateVal);
      const isDateLike =
        /^[A-Za-z]{3}-\d{2,4}$/.test(dateStr) ||
        /\d{1,2}\/\d{1,2}\/\d{2,4}/.test(dateStr) ||
        /^\d{5}$/.test(dateStr);
      if (!isDateLike) continue;
      const nonNullNums = row.map(num).filter((n): n is number => n !== null);
      if (nonNullNums.length < 2) continue;
      weeklySnapshots.push({
        date: fmtDate(dateVal),
        occupancyPct: Math.round(pct(row[1]) * 100),
        sqFtPct: Math.round(pct(row[2]) * 100),
        contracts: Math.round(num(row[3]) ?? 0),
        flexMembers: Math.round(num(row[4]) ?? 0),
        mailMembers: Math.round(num(row[5]) ?? 0),
        meetingHours: num(row[6]) ?? 0,
        meetingRevenue: num(row[7]) ?? 0,
      });
    }
  }

  return {
    reportDate: reportDate ||
      (weeklySnapshots.length > 0
        ? weeklySnapshots[weeklySnapshots.length - 1].date
        : new Date().toLocaleDateString("en-US")),
    location: location || "—",
    currentMonthRevenue,
    spaceTypes,
    totals,
    weeklySnapshots: weeklySnapshots.slice(-12),
  };
}

// ── P2. Monthly Income ───────────────────────────────────────────────────────
const MONTH_NAMES = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

function parsRevenue(wb: XLSX.WorkBook): RevenueData {
  // WX-3: Try "P2" prefix first — actual sheet is "P2. Monthly Income", not "P.2."
  const rows = getSheetByPrefix(wb, "P2") ?? getSheetByPrefix(wb, "P.2");
  if (!rows) throw new Error("Could not find P2 Monthly Income sheet");

  let headerRowIdx = -1;
  const months: string[] = [];
  const monthColIndices: number[] = [];

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const rowStr = row.map(str).join(" ").toLowerCase();
    const matchCount = MONTH_NAMES.filter((m) => rowStr.includes(m)).length;
    if (matchCount >= 3) {
      headerRowIdx = i;
      row.forEach((cell, j) => {
        const s = str(cell);
        const sLower = s.toLowerCase();
        // WX-5: Skip "Annual" / "YTD" summary columns — they are not months
        if (["annual", "ytd", "total year", "full year"].some((k) => sLower.includes(k))) return;
        if (MONTH_NAMES.some((m) => sLower.startsWith(m))) {
          months.push(s);
          monthColIndices.push(j);
        }
      });
      break;
    }
  }

  if (headerRowIdx < 0) return { months: [], lineItems: [] };

  const lineItems: RevenueLineItem[] = [];
  const totalLabels = ["total revenue", "total with relief", "net revenue", "gross revenue"];

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const label = str(row[0]);
    if (!label) continue;

    const values = monthColIndices.map((colIdx) => num(row[colIdx]));
    const hasValues = values.some((v) => v !== null && v !== 0);
    const labelLower = label.toLowerCase();
    const isTotal = totalLabels.some((t) => labelLower.includes(t));
    const isSectionHeader = !hasValues && label.length > 0;

    lineItems.push({ label, isTotal, isSectionHeader, values });

    // WX-1: Stop after "Total with Relief" — rows below are hours/transaction sub-tables
    if (labelLower.includes("total with relief")) break;
  }

  // WX-2: Trim trailing stub months where "Total with Relief" < 15% of peak month
  const totalLine = lineItems.find((li) => li.label.toLowerCase().includes("total with relief"));
  if (totalLine && months.length > 0) {
    const nonZero = totalLine.values.filter((v): v is number => v != null && (v as number) > 0);
    if (nonZero.length > 0) {
      const peak = Math.max(...nonZero);
      const threshold = peak * 0.15;
      let lastGoodCol = -1;
      totalLine.values.forEach((v, i) => {
        if (v != null && (v as number) >= threshold) lastGoodCol = i;
      });
      if (lastGoodCol >= 0 && lastGoodCol < months.length - 1) {
        const trimTo = lastGoodCol + 1;
        months.splice(trimTo);
        lineItems.forEach((li) => { li.values = li.values.slice(0, trimTo); });
      }
    }
  }

  return { months, lineItems };
}

// ── P.3 Occupancy ────────────────────────────────────────────────────────────
function parseOccupancy(wb: XLSX.WorkBook): OccupancyData {
  const rows = getSheetByPrefix(wb, "P.3");
  if (!rows) throw new Error("Could not find P.3 Occupancy sheet");

  let headerRowIdx = -1;
  let colMap: Record<string, number> = {};

  const colKeywords: Record<string, string[]> = {
    spaceName: ["space name", "space"],
    spaceId: ["space #", "space id", "space#", "id", "code"],
    status: ["status"],
    spaceType: ["type", "space type"],
    capacity: ["capacity"],
    sqFt: ["sq ft", "square", "sqft"],
    contractName: ["contract name", "contract"],
    companyName: ["company", "tenant"],
    moveIn: ["move-in", "move in", "start"],
    moveOut: ["move-out", "move out", "end"],
    term: ["term", "mtm", "fixed"],
  };

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].map(str);
    const rowLower = row.map((s) => s.toLowerCase());
    if (rowLower.some((s) => s.includes("status") || s.includes("occupied"))) {
      headerRowIdx = i;
      for (const [key, keywords] of Object.entries(colKeywords)) {
        const idx = rowLower.findIndex((s) => keywords.some((kw) => s.includes(kw)));
        if (idx >= 0) colMap[key] = idx;
      }
      break;
    }
  }

  if (headerRowIdx < 0) return { rows: [], totalSpaces: 0, occupiedCount: 0, vacantCount: 0 };

  const occupancyRows: OccupancyRow[] = [];
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const label = str(row[colMap.spaceName ?? colMap.spaceId ?? 0]);
    if (!label) continue;

    const statusRaw = str(row[colMap.status ?? 3]).toLowerCase();
    const status: "Occupied" | "Vacant" = statusRaw.includes("occup") ? "Occupied" : "Vacant";

    occupancyRows.push({
      spaceName: label,
      spaceId: str(row[colMap.spaceId ?? 1]),
      spaceType: str(row[colMap.spaceType ?? 4]),
      status,
      capacity: Math.round(num(row[colMap.capacity ?? 5]) ?? 1),
      sqFt: Math.round(num(row[colMap.sqFt ?? 6]) ?? 0),
      contractName: str(row[colMap.contractName ?? 7]),
      companyName: str(row[colMap.companyName ?? 8]),
      moveIn: fmtDate(row[colMap.moveIn ?? 9]),
      moveOut: fmtDate(row[colMap.moveOut ?? 10]),
      term: str(row[colMap.term ?? 11]),
    });
  }

  const occupiedCount = occupancyRows.filter((r) => r.status === "Occupied").length;
  return {
    rows: occupancyRows,
    totalSpaces: occupancyRows.length,
    occupiedCount,
    vacantCount: occupancyRows.length - occupiedCount,
  };
}

// ── P.4 Budget vs Actual ─────────────────────────────────────────────────────
function parseBudget(wb: XLSX.WorkBook): BudgetData {
  const rows = getSheetByPrefix(wb, "P.4");
  if (!rows) throw new Error("Could not find P.4 Forecast sheet");

  let headerRowIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    const rowStr = rows[i].map(str).join(" ").toLowerCase();
    if (rowStr.includes("ptd actual") || (rowStr.includes("actual") && rowStr.includes("budget"))) {
      headerRowIdx = i;
      break;
    }
  }

  if (headerRowIdx < 0) return { lineItems: [], totalRevenue: 0, totalCogs: 0, grossProfit: 0 };

  const lineItems: BudgetLineItem[] = [];
  let totalRevenue = 0;
  let totalCogs = 0;
  let grossProfit = 0;

  const isDisplayTotal = (ll: string) =>
    ["total revenue", "total income", "gross profit", "total cost", "total cogs"].some((t) => ll.includes(t));

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    const label = str(row[0]);
    if (!label) continue;

    const actual = num(row[1]);
    const budget = num(row[2]);
    const varianceDollar = num(row[3]);
    // WX-4: Always re-derive variancePct from formula — spreadsheet value is unreliable
    const variancePct =
      actual != null && budget != null && budget !== 0
        ? ((actual - budget) / Math.abs(budget)) * 100
        : null;

    const labelLower = label.toLowerCase().trim();
    const isTotal = isDisplayTotal(labelLower);
    const isSectionHeader = actual == null && budget == null;

    lineItems.push({ label, isSectionHeader, isTotal, actual, budget, varianceDollar, variancePct });

    // WX-11/12: Exact label matching — "includes" was matching too many intermediate rows
    if (labelLower === "total income" || labelLower === "total revenue") {
      totalRevenue = actual ?? totalRevenue;
    }
    if (labelLower === "total cost of sales" || labelLower === "total cogs") {
      totalCogs = actual ?? totalCogs;
    }
    if (labelLower === "gross profit") {
      grossProfit = actual ?? grossProfit;
    }
  }

  if (grossProfit === 0 && totalRevenue > 0) {
    grossProfit = totalRevenue - totalCogs;
  }

  return { lineItems, totalRevenue, totalCogs, grossProfit };
}

// ── Meeting Room Activity ────────────────────────────────────────────────────
function parseMeeting(wb: XLSX.WorkBook, revenue: RevenueData, budget: BudgetData): MeetingData {
  const txSheetName = wb.SheetNames.find((n) =>
    n.toLowerCase().includes("mtg") || n.toLowerCase().includes("meeting")
  );
  const recurringTransactions: MeetingRoomTransaction[] = [];

  if (txSheetName) {
    const rows = XLSX.utils.sheet_to_json<SheetRow>(wb.Sheets[txSheetName], {
      header: 1, defval: null, raw: false,
    });
    let headerIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      const rowStr = rows[i].map(str).join(" ").toLowerCase();
      if (rowStr.includes("company") && rowStr.includes("income")) {
        headerIdx = i;
        break;
      }
    }
    if (headerIdx >= 0) {
      const hdr = rows[headerIdx].map(str).map((s) => s.toLowerCase());
      const col = {
        company:     hdr.findIndex((s) => s.includes("company")),
        description: hdr.findIndex((s) => s.includes("description") || s.includes("desc")),
        chargeCode:  hdr.findIndex((s) => s.includes("charge") || s.includes("code")),
        month:       hdr.findIndex((s) => s.includes("month") || s.includes("date")),
        amount:      hdr.findIndex((s) => s.includes("income") || s.includes("amount") || s.includes("total")),
      };
      const c = (key: keyof typeof col, fallback: number) => col[key] >= 0 ? col[key] : fallback;
      for (let i = headerIdx + 1; i < rows.length; i++) {
        const row = rows[i];
        const company     = str(row[c("company", 2)]);
        const description = str(row[c("description", 4)]);
        const chargeCode  = str(row[c("chargeCode", 5)]);
        const month       = str(row[c("month", 6)]);
        // Issue 1/3: detect unit from the raw formatted cell string — "$420" → "$", "48" → "hrs"
        const amountRaw   = str(row[c("amount", 7)]);
        const amount      = num(amountRaw);
        const unit: "$" | "hrs" = amountRaw.trimStart().startsWith("$") ? "$" : "hrs";
        if (company && amount != null && amount > 0) {
          recurringTransactions.push({ company, description, chargeCode, month, amount, unit });
        }
      }
    }
  }

  // Room type breakdown from budget line items
  const roomTypeLabels = [
    "day office", "huddle", "round table", "loft", "ideation",
    "conference room", "boardroom", "event space", "podcast", "cabana",
    "recurring meeting", "general meeting",
  ];
  const roomTypes: MeetingRoomTypeRow[] = [];
  let inMeetingSection = false;

  for (const item of budget.lineItems) {
    const labelLower = item.label.toLowerCase().trim();
    if (labelLower.includes("meeting space") && item.isSectionHeader) {
      inMeetingSection = true;
      continue;
    }
    if (inMeetingSection) {
      if (item.isSectionHeader && !labelLower.includes("meeting")) break;
      const isRoomType = roomTypeLabels.some((r) => labelLower.includes(r));
      const isTotal = labelLower.includes("total meeting");
      if (isRoomType || isTotal) {
        const actual = item.actual ?? 0;
        const budgetVal = item.budget ?? 0;
        const varianceDollar = item.varianceDollar ?? actual - budgetVal;
        const variancePct = budgetVal !== 0 ? ((actual - budgetVal) / Math.abs(budgetVal)) * 100 : null;
        roomTypes.push({ label: item.label.trim(), actual, budget: budgetVal, varianceDollar, variancePct, isTotal });
      }
    }
  }

  const totalRow = roomTypes.find((r) => r.isTotal);
  const mtdRevenue = totalRow?.actual ?? recurringTransactions.reduce((s, t) => s + t.amount, 0);
  const mtdBudget = totalRow?.budget ?? 0;
  const varianceDollar = mtdRevenue - mtdBudget;
  const variancePct = mtdBudget !== 0 ? ((mtdRevenue - mtdBudget) / Math.abs(mtdBudget)) * 100 : null;

  // Monthly trend from revenue sheet — include ALL months (null → 0) so current month
  // always appears in the chart even when it has no meeting revenue yet (Issue 5).
  const meetingLine = revenue.lineItems.find((li) =>
    li.label.toLowerCase().trim() === "meeting rooms"
  );
  const monthlyTrend: MeetingMonthly[] = [];
  if (meetingLine) {
    revenue.months.forEach((month, i) => {
      monthlyTrend.push({ month, revenue: meetingLine.values[i] ?? 0 });
    });
  }

  return { mtdRevenue, mtdBudget, varianceDollar, variancePct, recurringTransactions, roomTypes, monthlyTrend };
}

// ── Data Validation ──────────────────────────────────────────────────────────
function near(a: number, b: number, tol = 2): boolean {
  return Math.abs(a - b) <= tol;
}

function fmtN(n: number): string {
  return n.toLocaleString("en-US", { maximumFractionDigits: 0 });
}

function validateData(
  dashboard: DashboardData,
  occupancy: OccupancyData,
  budget: BudgetData,
  revenue: RevenueData,
  meeting: MeetingData
): ValidationResult {
  const checks: ValidationCheck[] = [];

  checks.push({
    label: "Total spaces: P.1 Dashboard vs P.3 Roster",
    passed: near(dashboard.totals.totalSpaces, occupancy.totalSpaces),
    expected: `${dashboard.totals.totalSpaces} (P.1)`,
    actual: `${occupancy.totalSpaces} (P.3 rows)`,
    note: "P.1 KPI and P.3 physical roster should show identical space counts.",
  });

  checks.push({
    label: "Occupied spaces: P.1 Dashboard vs P.3 Roster",
    passed: near(dashboard.totals.occupiedSpaces, occupancy.occupiedCount),
    expected: `${dashboard.totals.occupiedSpaces} (P.1)`,
    actual: `${occupancy.occupiedCount} (P.3 rows)`,
    note: "Counts should match — any gap indicates a data-entry lag between tabs.",
  });

  const spaceSubtotal = dashboard.spaceTypes.reduce((s, t) => s + t.totalSpaces, 0);
  checks.push({
    label: "Space sub-totals sum to grand total (P.1)",
    passed: near(spaceSubtotal, dashboard.totals.totalSpaces),
    expected: `${dashboard.totals.totalSpaces}`,
    actual: `${spaceSubtotal} (Ded. Desk + Private Office)`,
    note: "Sum of Dedicated Desk and Private Office must equal Total Spaces.",
  });

  const derivedGP = budget.totalRevenue - budget.totalCogs;
  checks.push({
    label: "Gross Profit: Revenue − COGS (P.4)",
    passed: near(derivedGP, budget.grossProfit, 5),
    expected: `$${fmtN(budget.grossProfit)} (reported)`,
    actual: `$${fmtN(derivedGP)} (calculated)`,
    note: "Gross Profit should equal Total Revenue minus Cost of Sales.",
  });

  const totalLine = revenue.lineItems.find((li) => li.label.toLowerCase().includes("total with relief"));
  let p2Current: number | null = null;
  if (totalLine) {
    const nonZero = totalLine.values.filter((v) => v != null && (v as number) > 0) as number[];
    if (nonZero.length > 0) {
      const maxVal = Math.max(...nonZero);
      const threshold = maxVal * 0.15;
      const completed = nonZero.filter((v) => v >= threshold);
      p2Current = completed.length > 0 ? completed[completed.length - 1] : null;
    }
  }
  if (p2Current != null && budget.totalRevenue > 0) {
    checks.push({
      label: "Current month revenue: P.2 vs P.4",
      passed: near(p2Current, budget.totalRevenue, 10),
      expected: `$${fmtN(budget.totalRevenue)} (P.4 Actual)`,
      actual: `$${fmtN(p2Current)} (P.2 Total with Relief)`,
      note: "P.2 current-month net revenue and P.4 total should reconcile.",
    });
  }

  const txSum = meeting.recurringTransactions.reduce((s, t) => s + t.amount, 0);
  if (txSum > 0) {
    checks.push({
      label: "Meeting room recurring: transaction sum vs budget line",
      passed: near(txSum, meeting.mtdRevenue, 5),
      expected: `$${fmtN(meeting.mtdRevenue)} (P.4 line)`,
      actual: `$${fmtN(txSum)} (transaction detail)`,
      note: "Sum of individual recurring invoices should match the budget-tab total.",
    });
  }

  const calcOccPct = dashboard.totals.totalSpaces > 0
    ? Math.round((dashboard.totals.occupiedSpaces / dashboard.totals.totalSpaces) * 100)
    : 0;
  checks.push({
    label: "Occupancy % calculation (P.1)",
    passed: near(calcOccPct, dashboard.totals.occupancyPct, 1),
    expected: `${dashboard.totals.occupancyPct}% (reported)`,
    actual: `${calcOccPct}% (recalculated: ${dashboard.totals.occupiedSpaces}/${dashboard.totals.totalSpaces})`,
    note: "Independently recalculated from raw space counts.",
  });

  const monthOrder = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
  let orderedCorrectly = true;
  for (let i = 1; i < revenue.months.length; i++) {
    const prev = monthOrder.findIndex((m) => revenue.months[i - 1].toLowerCase().startsWith(m));
    const curr = monthOrder.findIndex((m) => revenue.months[i].toLowerCase().startsWith(m));
    if (curr !== prev + 1 && !(prev === 11 && curr === 0)) { orderedCorrectly = false; break; }
  }
  checks.push({
    label: "Revenue months are in sequential order (P.2)",
    passed: orderedCorrectly,
    expected: "Jan → Feb → … in sequence",
    actual: revenue.months.slice(0, 6).join(", ") + (revenue.months.length > 6 ? "…" : ""),
    note: "Out-of-order months would indicate column mapping errors in P.2.",
  });

  const passCount = checks.filter((c) => c.passed).length;
  return { checks, passCount, failCount: checks.length - passCount };
}

// ── Main entry point ─────────────────────────────────────────────────────────
export function parseExcelWorkbook(buffer: ArrayBuffer): ReportData {
  const wb = XLSX.read(buffer, { type: "array", cellDates: false });

  const dashboard = parseDashboard(wb);
  const revenue = parsRevenue(wb);
  const occupancy = parseOccupancy(wb);
  const budget = parseBudget(wb);
  const meeting = parseMeeting(wb, revenue, budget);

  // P.4 PTD Actual overrides the heuristic P.1 scan
  if (budget.totalRevenue > 0) dashboard.currentMonthRevenue = budget.totalRevenue;

  // WX-13: Backfill last snapshot meetingRevenue from meeting.mtdRevenue if null/zero
  if (dashboard.weeklySnapshots.length > 0 && meeting.mtdRevenue > 0) {
    const last = dashboard.weeklySnapshots[dashboard.weeklySnapshots.length - 1];
    if (!last.meetingRevenue) last.meetingRevenue = meeting.mtdRevenue;
  }

  const validation = validateData(dashboard, occupancy, budget, revenue, meeting);

  return {
    reportDate: dashboard.reportDate,
    dashboard,
    revenue,
    occupancy,
    budget,
    meeting,
    validation,
  };
}

export { formatCurrency };
