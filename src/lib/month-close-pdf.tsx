import React from "react";
import {
  Document, Page, Text, View, StyleSheet, Font,
} from "@react-pdf/renderer";
import type { MonthCloseData, MonthCloseItem, CheckStatus } from "@/types/month-close";

Font.register({ family: "Helvetica", fonts: [] });

// ── Palette ───────────────────────────────────────────────────────────────────
const C = {
  navy:        "#0F2044",
  navyMid:     "#1A3461",
  gold:        "#C9A84C",
  green:       "#16A34A",
  greenBg:     "#F0FDF4",
  greenDark:   "#14532D",
  greenText:   "#15803D",
  red:         "#DC2626",
  redBg:       "#FFF1F2",
  redDark:     "#7F1D1D",
  redText:     "#B91C1C",
  orange:      "#EA580C",
  orangeBg:    "#FFF7ED",
  orangeDark:  "#78350F",
  orangeText:  "#C2410C",
  blue:        "#1E40AF",
  blueBg:      "#DBEAFE",
  gray50:      "#F9FAFB",
  gray100:     "#F3F4F6",
  gray200:     "#E5E7EB",
  gray300:     "#D1D5DB",
  gray400:     "#9CA3AF",
  gray600:     "#4B5563",
  gray700:     "#374151",
  gray800:     "#1F2937",
  white:       "#FFFFFF",
};

const s = StyleSheet.create({
  page: {
    fontFamily: "Helvetica",
    fontSize: 8,
    paddingTop: 42,
    paddingBottom: 44,
    paddingHorizontal: 44,
    backgroundColor: C.white,
  },

  // Summary page header (inline flow)
  pageHeader: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "flex-end",
    marginBottom: 18,
    paddingBottom: 10,
    borderBottomWidth: 2,
    borderBottomColor: C.navy,
  },
  // Detail page header (absolute + fixed — repeats on every overflow page)
  pageHeaderFixed: {
    position: "absolute",
    top: 12,
    left: 44,
    right: 44,
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "flex-end",
    paddingBottom: 8,
    borderBottomWidth: 2,
    borderBottomColor: C.navy,
  },
  brandName:   { fontSize: 14, fontFamily: "Helvetica-Bold", color: C.navy },
  brandSub:    { fontSize: 7, color: C.gold, marginTop: 2, letterSpacing: 0.8 },
  reportTitle: { fontSize: 11, fontFamily: "Helvetica-Bold", color: C.navy, textAlign: "right" },
  reportMeta:  { fontSize: 7, color: C.gray600, textAlign: "right", marginTop: 2 },

  // Section titles
  sectionBlock: { alignItems: "center", marginTop: 14, marginBottom: 7 },
  sectionTitle: {
    fontSize: 8.5,
    fontFamily: "Helvetica-Bold",
    color: C.navy,
    textTransform: "uppercase",
    letterSpacing: 1.0,
    textAlign: "center",
  },
  sectionRule: {
    width: 28,
    height: 2,
    backgroundColor: C.gold,
    marginTop: 3,
    alignSelf: "center",
  },
  sectionNote: { fontSize: 6.5, color: C.gray600, marginBottom: 6, marginHorizontal: 2 },

  // KPI cards
  kpiRow:      { flexDirection: "row", gap: 10, marginBottom: 14 },
  kpiCard:     { flex: 1, backgroundColor: C.navy, borderRadius: 5, padding: 10, alignItems: "center" },
  kpiCardGreen: { backgroundColor: C.greenDark },
  kpiCardRed:   { backgroundColor: C.redDark },
  kpiCardGold:  { backgroundColor: C.orangeDark },
  kpiCardGray:  { backgroundColor: C.gray700 },
  kpiLabel: { fontSize: 6, color: C.gold, textTransform: "uppercase", letterSpacing: 0.8, marginBottom: 3, textAlign: "center" },
  kpiValue: { fontSize: 20, fontFamily: "Helvetica-Bold", color: C.white, textAlign: "center" },
  kpiSub:   { fontSize: 5.5, color: C.gray400, marginTop: 3, textAlign: "center", lineHeight: 1.5 },

  // Banners
  warnBanner: {
    flexDirection: "row",
    backgroundColor: "#FEF3C7",
    borderWidth: 1,
    borderColor: "#F59E0B",
    borderRadius: 4,
    padding: 7,
    marginBottom: 10,
  },
  warnText: { fontSize: 7.5, color: "#92400E", fontFamily: "Helvetica-Bold" },
  warnNote: { fontSize: 6.5, color: "#92400E", marginTop: 2 },
  clearBanner: {
    backgroundColor: C.greenBg,
    borderWidth: 1,
    borderColor: "#86EFAC",
    borderRadius: 4,
    padding: 8,
    alignItems: "center",
    marginBottom: 8,
  },
  clearText: { fontSize: 7.5, fontFamily: "Helvetica-Bold", color: C.green },

  // Tables — paddingRight on th/td prevents text bleeding into adjacent cells
  table:   { width: "100%", marginBottom: 4 },
  tHead:   { flexDirection: "row", backgroundColor: C.navy, paddingVertical: 5, paddingHorizontal: 5, borderRadius: 3 },
  tRow:    { flexDirection: "row", paddingVertical: 3.5, paddingHorizontal: 5, borderBottomWidth: 1, borderBottomColor: C.gray100, alignItems: "center" },
  tRowAlt: { backgroundColor: C.gray50 },
  th:      { fontSize: 6.5, fontFamily: "Helvetica-Bold", color: C.white, paddingRight: 6 },
  td:      { fontSize: 7, color: C.gray800, paddingRight: 6 },
  tdBold:  { fontSize: 7, fontFamily: "Helvetica-Bold", color: C.gray800, paddingRight: 6 },
  tdGray:  { fontSize: 7, color: C.gray400, paddingRight: 6 },
  tdNa:    { fontSize: 7, color: C.gray400, paddingRight: 6 },

  // Row backgrounds by status — N/A uses gray100 (more distinct than review's white)
  rowConfirmed:     { backgroundColor: "#F0FDF4" },
  rowAttention:     { backgroundColor: "#FFF1F2" },
  rowCannotCompute: { backgroundColor: "#FFF7ED" },
  rowNa:            { backgroundColor: C.gray100 },

  // External pull list
  pullGroup:     { marginBottom: 7 },
  pullGroupHead: { fontSize: 7.5, fontFamily: "Helvetica-Bold", color: C.navy, marginBottom: 3 },
  pullItem:      { fontSize: 7, color: C.gray600, marginBottom: 2, marginLeft: 8 },

  // Footer
  footer: {
    position: "absolute", bottom: 14, left: 44, right: 44,
    flexDirection: "row", justifyContent: "space-between",
    borderTopWidth: 1, borderTopColor: C.gray200, paddingTop: 4,
  },
  footerText: { fontSize: 6, color: C.gray400 },
});

// ── Helpers ───────────────────────────────────────────────────────────────────

function trunc(str: string, max = 54): string {
  return str.length > max ? str.slice(0, max - 1) + "…" : str;
}

function fmt$(n: number | null): string {
  if (n == null) return "—";
  return new Intl.NumberFormat("en-US", {
    style: "currency", currency: "USD", maximumFractionDigits: 0,
  }).format(n);
}

function fmtGap(n: number | null): string {
  if (n == null) return "—";
  const base = new Intl.NumberFormat("en-US", {
    style: "currency", currency: "USD", maximumFractionDigits: 0,
  }).format(Math.abs(n));
  return n >= 0 ? `+${base}` : `-${base}`;
}

function statusLabel(status: CheckStatus): string {
  switch (status) {
    case "confirmed":      return "CONFIRMED";
    case "variance":       return "VARIANCE";
    case "missing":        return "MISSING";
    case "over_limit":     return "OVER LIMIT";
    case "cannot_compute": return "NO DATA";
    case "review":         return "REVIEW";
    case "na":             return "N/A";
  }
}

function badgeBg(status: CheckStatus): string {
  switch (status) {
    case "confirmed":      return "#DCFCE7";
    case "variance":
    case "missing":
    case "over_limit":     return "#FEE2E2";
    case "cannot_compute": return "#FED7AA";
    case "review":         return C.blueBg;
    case "na":             return C.gray200;
  }
}

function badgeFg(status: CheckStatus): string {
  switch (status) {
    case "confirmed":      return C.greenText;
    case "variance":
    case "missing":
    case "over_limit":     return C.redText;
    case "cannot_compute": return C.orangeText;
    case "review":         return C.blue;
    case "na":             return C.gray600;
  }
}

function rowStyle(status: CheckStatus) {
  switch (status) {
    case "confirmed":      return s.rowConfirmed;
    case "variance":
    case "missing":
    case "over_limit":     return s.rowAttention;
    case "cannot_compute": return s.rowCannotCompute;
    case "na":             return s.rowNa;
    default:               return {};
  }
}

// ── Status Badge ──────────────────────────────────────────────────────────────

function StatusBadge({ status }: { status: CheckStatus }) {
  return (
    <View style={{
      backgroundColor: badgeBg(status),
      paddingHorizontal: 4,
      paddingVertical: 2,
      borderRadius: 3,
      alignSelf: "flex-start",
    }}>
      <Text style={{
        fontSize: 5.5,
        fontFamily: "Helvetica-Bold",
        color: badgeFg(status),
      }}>
        {statusLabel(status)}
      </Text>
    </View>
  );
}

// ── External report grouping ──────────────────────────────────────────────────

const ACCOUNT_REPORT_MAP: Record<string, string> = {
  "1110": "Bank Reconciliation",  "1120": "Bank Reconciliation",
  "1123": "Bank Reconciliation",  "1124": "Bank Reconciliation",
  "1200": "AR Aging Report",      "1210": "AR Aging Report",
  "1211": "AR Aging Report",      "1220": "AR Aging Report",
  "1230": "AR Aging Report",
  "1300": "Prepaid Expense Schedule", "1310": "Prepaid Expense Schedule",
  "1320": "Prepaid Expense Schedule", "1330": "Prepaid Expense Schedule",
  "1500": "Fixed Asset Schedule", "1510": "Fixed Asset Schedule",
  "1511": "Fixed Asset Schedule", "1590": "Fixed Asset Schedule",
  "1800": "Security Deposit Schedule", "1810": "Security Deposit Schedule",
  "1820": "Security Deposit Schedule",
  "2100": "AP Aging Report",      "2200": "AP Aging Report",
  "2210": "AP Aging Report",
  "2230": "Accrued Liabilities Schedule",
  "2260": "Deferred Revenue Schedule",
  "2280": "Deferred Revenue / Prepayments",
  "3100": "Equity Schedule",      "3200": "Equity Schedule",
  "3800": "Retained Earnings Schedule",
};

function getReportGroup(account_codes: string[]): string {
  for (const code of account_codes) {
    if (ACCOUNT_REPORT_MAP[code]) return ACCOUNT_REPORT_MAP[code];
  }
  const first = account_codes[0] ?? "";
  if (first.startsWith("11")) return "Bank Reconciliation";
  if (first.startsWith("12")) return "AR Aging Report";
  if (first.startsWith("13")) return "Prepaid Expense Schedule";
  if (first.startsWith("15")) return "Fixed Asset Schedule";
  if (first.startsWith("18")) return "Security Deposit Schedule";
  if (first.startsWith("22")) return "AP / Accrued Liabilities";
  if (first.startsWith("32") || first.startsWith("38")) return "Equity Schedule";
  return "Additional Reports Required";
}

// ── Shared layout ─────────────────────────────────────────────────────────────

function PageHeader({ location, month }: { location: string; month: string }) {
  return (
    <View style={s.pageHeader}>
      <View>
        <Text style={s.brandName}>25N Coworking</Text>
        <Text style={s.brandSub}>{location.toUpperCase()}</Text>
      </View>
      <View>
        <Text style={s.reportTitle}>Month-End Close Report</Text>
        <Text style={s.reportMeta}>{month} · Confidential</Text>
      </View>
    </View>
  );
}

// Repeats at fixed position on every page of the continuous detail section
function FixedPageHeader({ location, month }: { location: string; month: string }) {
  return (
    <View style={s.pageHeaderFixed} fixed>
      <View>
        <Text style={s.brandName}>25N Coworking</Text>
        <Text style={s.brandSub}>{location.toUpperCase()}</Text>
      </View>
      <View>
        <Text style={s.reportTitle}>Month-End Close Report</Text>
        <Text style={s.reportMeta}>{month} · Full Checklist</Text>
      </View>
    </View>
  );
}

function PageFooter({ location, month }: { location: string; month: string }) {
  return (
    <View style={s.footer} fixed>
      <Text style={s.footerText}>Confidential — 25N Coworking · {location} · {month} Close</Text>
      <Text style={s.footerText} render={({ pageNumber, totalPages }) => `Page ${pageNumber} of ${totalPages}`} />
    </View>
  );
}

function SectionTitle({ children }: { children: React.ReactNode }) {
  return (
    <View style={s.sectionBlock}>
      <Text style={s.sectionTitle}>{children}</Text>
      <View style={s.sectionRule} />
    </View>
  );
}

// ── Attention item row (summary page — Actual / Expected / Gap / Note) ─────────

function AttentionRow({ item }: { item: MonthCloseItem }) {
  const gap = item.actual != null && item.expected != null ? item.actual - item.expected : null;
  const gapColor = gap == null ? C.gray400 : gap > 0 ? C.redText : C.greenText;

  return (
    <View style={[s.tRow, s.rowAttention]}>
      <Text style={[s.tdBold, { flex: 0.8 }]}>{item.id}</Text>
      <Text style={[s.td, { flex: 3.5 }]}>{trunc(item.description, 50)}</Text>
      <Text style={[s.tdBold, { flex: 1.3, textAlign: "right" }]}>{fmt$(item.actual)}</Text>
      <Text style={[s.td,     { flex: 1.3, textAlign: "right" }]}>{fmt$(item.expected)}</Text>
      <Text style={[s.td,     { flex: 1.2, textAlign: "right", color: gapColor, fontFamily: "Helvetica-Bold" }]}>{fmtGap(gap)}</Text>
      <Text style={[s.td,     { flex: 2.5, color: C.gray600 }]}>{trunc(item.note || "—", 40)}</Text>
    </View>
  );
}

// ── Detail item row (full checklist — with Expected col + badge status) ─────────

function DetailRow({ item, idx }: { item: MonthCloseItem; idx: number }) {
  const isNa = item.status === "na";
  const isReviewAlt = idx % 2 === 1 && item.status === "review";
  const textStyle = isNa ? s.tdNa : s.td;
  const boldStyle = isNa ? s.tdNa : s.tdBold;

  return (
    <View style={[s.tRow, rowStyle(item.status), isReviewAlt ? s.tRowAlt : {}]}>
      <Text style={[boldStyle, { flex: 0.7 }]}>{item.id}</Text>
      <Text style={[s.tdGray,  { flex: 1.2 }]}>{item.account_codes.join(", ") || "—"}</Text>
      <Text style={[textStyle, { flex: 3.0 }]}>{trunc(item.description, 54)}</Text>
      <View style={{ flex: 1.3, justifyContent: "center" }}>
        <StatusBadge status={item.status} />
      </View>
      <Text style={[boldStyle, { flex: 1.3, textAlign: "right" }]}>
        {item.actual != null ? fmt$(item.actual) : "—"}
      </Text>
      <Text style={[textStyle, { flex: 1.3, textAlign: "right", color: isNa ? C.gray400 : C.gray600 }]}>
        {item.expected != null ? fmt$(item.expected) : "—"}
      </Text>
      <Text style={[textStyle, { flex: 2.0, color: C.gray600 }]}>{trunc(item.note || "", 40)}</Text>
    </View>
  );
}

// ── Page 1: Summary ───────────────────────────────────────────────────────────

function SummaryPage({ data }: { data: MonthCloseData }) {
  const { location, month, counts, budget_aligned, budget_mtd_month, attention_items, external_items } = data;

  const externalCount = external_items.length;
  const glReviewCount = Math.max(0, counts.review - externalCount);

  const pullGroups: Record<string, MonthCloseItem[]> = {};
  for (const item of external_items) {
    const group = getReportGroup(item.account_codes);
    pullGroups[group] = pullGroups[group] ?? [];
    pullGroups[group].push(item);
  }

  return (
    <Page size="LETTER" style={s.page}>
      <PageHeader location={location} month={month} />

      <View style={s.kpiRow}>
        <View style={[s.kpiCard, s.kpiCardGreen]}>
          <Text style={s.kpiLabel}>Confirmed</Text>
          <Text style={s.kpiValue}>{counts.confirmed}</Text>
          <Text style={s.kpiSub}>automated checks passed</Text>
        </View>
        <View style={[s.kpiCard, s.kpiCardRed]}>
          <Text style={s.kpiLabel}>Needs Attention</Text>
          <Text style={s.kpiValue}>{counts.attention}</Text>
          <Text style={s.kpiSub}>variances or missing entries</Text>
        </View>
        <View style={[s.kpiCard, s.kpiCardGold]}>
          <Text style={s.kpiLabel}>Human Review</Text>
          <Text style={s.kpiValue}>{counts.review}</Text>
          <Text style={s.kpiSub}>{glReviewCount} GL items{"\n"}{externalCount} external reports</Text>
        </View>
        <View style={[s.kpiCard, s.kpiCardGray]}>
          <Text style={s.kpiLabel}>Not Applicable</Text>
          <Text style={s.kpiValue}>{counts.na}</Text>
          <Text style={s.kpiSub}>confirmed reviewed + skipped</Text>
        </View>
      </View>

      {!budget_aligned && (
        <View style={s.warnBanner}>
          <View>
            <Text style={s.warnText}>Budget Alignment Warning</Text>
            <Text style={s.warnNote}>
              Budget file reflects {budget_mtd_month}, not {month}. Budget variance data is unreliable for this period — do not use for month-end sign-off.
            </Text>
          </View>
        </View>
      )}

      <SectionTitle>Items Requiring Attention ({counts.attention})</SectionTitle>
      {attention_items.length === 0 ? (
        <View style={s.clearBanner}>
          <Text style={s.clearText}>All automated checks passed — no variances or missing entries detected</Text>
        </View>
      ) : (
        <View style={s.table}>
          <View style={s.tHead}>
            <Text style={[s.th, { flex: 0.8 }]}>ID</Text>
            <Text style={[s.th, { flex: 3.5 }]}>Description</Text>
            <Text style={[s.th, { flex: 1.3, textAlign: "right" }]}>Actual</Text>
            <Text style={[s.th, { flex: 1.3, textAlign: "right" }]}>Expected</Text>
            <Text style={[s.th, { flex: 1.2, textAlign: "right" }]}>Gap</Text>
            <Text style={[s.th, { flex: 2.5 }]}>Note</Text>
          </View>
          {attention_items.map((item, i) => (
            <AttentionRow key={i} item={item} />
          ))}
        </View>
      )}

      {/* wrap={false} keeps section title from orphaning from its first pull item */}
      {external_items.length > 0 && (
        <View wrap={false}>
          <SectionTitle>External Reports Required to Complete Close</SectionTitle>
          {Object.entries(pullGroups).map(([group, items]) => (
            <View key={group} style={s.pullGroup}>
              <Text style={s.pullGroupHead}>{group}</Text>
              {items.map((item, i) => (
                <Text key={i} style={s.pullItem}>
                  • {item.id} — {trunc(item.description, 62)}
                </Text>
              ))}
            </View>
          ))}
        </View>
      )}

      <PageFooter location={location} month={month} />
    </Page>
  );
}

// ── Pages 2+: Checklist Detail (all sections, continuous flow) ────────────────

function ChecklistPages({ data }: { data: MonthCloseData }) {
  const { location, month, sections } = data;

  return (
    <Page size="LETTER" style={s.page}>
      {/* Fixed header repeats on every overflow page */}
      <FixedPageHeader location={location} month={month} />

      {sections.map((section) => (
        <View key={section.id}>
          {/* SectionTitle + column headers stay together — prevents orphan titles */}
          <View wrap={false}>
            <SectionTitle>Section {section.id} — {section.title}</SectionTitle>
            {section.note && (
              <Text style={s.sectionNote}>Note: {section.note}</Text>
            )}
            <View style={s.tHead}>
              <Text style={[s.th, { flex: 0.7 }]}>ID</Text>
              <Text style={[s.th, { flex: 1.2 }]}>Account(s)</Text>
              <Text style={[s.th, { flex: 3.0 }]}>Description</Text>
              <Text style={[s.th, { flex: 1.3 }]}>Status</Text>
              <Text style={[s.th, { flex: 1.3, textAlign: "right" }]}>Actual</Text>
              <Text style={[s.th, { flex: 1.3, textAlign: "right" }]}>Expected</Text>
              <Text style={[s.th, { flex: 2.0 }]}>Note</Text>
            </View>
          </View>
          {section.items.map((item, i) => (
            <DetailRow key={i} item={item} idx={i} />
          ))}
        </View>
      ))}

      <PageFooter location={location} month={month} />
    </Page>
  );
}

// ── Main Document ─────────────────────────────────────────────────────────────

export function MonthCloseDocument({ data }: { data: MonthCloseData }) {
  return (
    <Document
      title={`25N ${data.location} — ${data.month} Month-End Close`}
      author="25N Coworking"
    >
      <SummaryPage data={data} />
      <ChecklistPages data={data} />
    </Document>
  );
}
