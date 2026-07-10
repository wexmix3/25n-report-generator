import React from "react";
import { Document, Page, Text, View, StyleSheet } from "@react-pdf/renderer";

// ── Palette — matches financial-packet-pdf.tsx for visual consistency ────────
const C = {
  navy:    "#0F2044",
  gold:    "#C9A84C",
  green:   "#16A34A",
  red:     "#DC2626",
  amber:   "#D97706",
  gray50:  "#F9FAFB",
  gray100: "#F3F4F6",
  gray200: "#E5E7EB",
  gray300: "#D1D5DB",
  gray400: "#9CA3AF",
  gray600: "#4B5563",
  gray700: "#374151",
  gray800: "#1F2937",
  white:   "#FFFFFF",
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
  pageHeader: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "flex-end",
    marginBottom: 18,
    paddingBottom: 10,
    borderBottomWidth: 2,
    borderBottomColor: C.navy,
  },
  brandName:   { fontSize: 14, fontFamily: "Helvetica-Bold", color: C.navy },
  brandSub:    { fontSize: 7, color: C.gold, marginTop: 2, letterSpacing: 0.8 },
  reportTitle: { fontSize: 11, fontFamily: "Helvetica-Bold", color: C.navy, textAlign: "right" },
  reportMeta:  { fontSize: 7, color: C.gray600, textAlign: "right", marginTop: 2 },
  footer: {
    position: "absolute", bottom: 14, left: 44, right: 44,
    flexDirection: "row", justifyContent: "space-between",
    borderTopWidth: 1, borderTopColor: C.gray200, paddingTop: 4,
  },
  footerText: { fontSize: 6, color: C.gray400 },
  sectionBlock: { alignItems: "center", marginTop: 16, marginBottom: 8 },
  sectionTitle: {
    fontSize: 8.5, fontFamily: "Helvetica-Bold", color: C.navy,
    textTransform: "uppercase", letterSpacing: 1.0, textAlign: "center",
  },
  sectionRule: { width: 28, height: 2, backgroundColor: C.gold, marginTop: 3, alignSelf: "center" },
  sectionNote: { fontSize: 6.5, color: C.gray600, textAlign: "center", marginTop: 4, marginHorizontal: 40 },
  kpiRow:  { flexDirection: "row", gap: 10, marginBottom: 14 },
  kpiCard: { flex: 1, backgroundColor: C.navy, borderRadius: 5, padding: 10, alignItems: "center" },
  kpiLabel: { fontSize: 6, color: C.gold, textTransform: "uppercase", letterSpacing: 0.8, marginBottom: 3, textAlign: "center" },
  kpiValue: { fontSize: 18, fontFamily: "Helvetica-Bold", color: C.white, textAlign: "center" },
  kpiSub:   { fontSize: 5.5, color: C.gray400, marginTop: 3, textAlign: "center" },
  table:   { width: "100%", marginBottom: 6 },
  tHead:   { flexDirection: "row", backgroundColor: C.navy, paddingVertical: 5, paddingHorizontal: 5, borderRadius: 3 },
  tRow:    { flexDirection: "row", paddingVertical: 3, paddingHorizontal: 5, borderBottomWidth: 1, borderBottomColor: C.gray100 },
  tRowAlt: { backgroundColor: C.gray50 },
  th:      { fontSize: 6.5, fontFamily: "Helvetica-Bold", color: C.white, paddingRight: 4 },
  td:      { fontSize: 7, color: C.gray800, paddingRight: 4 },
  tdMono:  { fontSize: 7, fontFamily: "Courier", color: C.gray800, paddingRight: 4 },
  tdGray:  { fontSize: 7, color: C.gray400, paddingRight: 4 },
  tdRed:   { fontSize: 7, fontFamily: "Helvetica-Bold", color: C.red, paddingRight: 4 },
  tdAmber: { fontSize: 7, fontFamily: "Helvetica-Bold", color: C.amber, paddingRight: 4 },
  tdGreen: { fontSize: 7, fontFamily: "Helvetica-Bold", color: C.green, paddingRight: 4 },
  emptyState: { fontSize: 8, color: C.green, textAlign: "center", marginVertical: 10, fontFamily: "Helvetica-Bold" },
});

// ── Types ──────────────────────────────────────────────────────────────────

interface VarianceFlag {
  account: string;
  name: string;
  prior: number;
  current: number;
  variance: number;
  variance_pct: number | null;
  rule_triggered: string;
}

interface Violation {
  account: string;
  account_name: string;
  control: string;
  violation_type: string;
  message: string;
  amount: number;
}

interface JournalEntryAccount {
  account: string;
  account_name: string;
  transaction_count: number;
  total_amount: number;
}

export interface GLVarianceData {
  location: string;
  month: string;
  prior_month: string;
  generated_at: string;
  flag_count: number;
  clean_count: number;
  flags: VarianceFlag[];
  violation_count: number;
  violations: Violation[];
  journal_entry_account_count: number;
  journal_entry_accounts: JournalEntryAccount[];
}

// ── Helpers ────────────────────────────────────────────────────────────────

function fmt$(n: number | null | undefined): string {
  if (n == null) return "—";
  const abs = Math.abs(n);
  const str = abs >= 1000
    ? `$${abs.toLocaleString("en-US", { maximumFractionDigits: 0 })}`
    : `$${Math.round(abs)}`;
  return n < 0 ? `(${str})` : str;
}

function fmtPct(n: number | null | undefined): string {
  if (n == null) return "new";
  return `${n > 0 ? "+" : ""}${n.toFixed(1)}%`;
}

// ── Shared layout ─────────────────────────────────────────────────────────

function Header({ location, month }: { location: string; month: string }) {
  return (
    <View style={s.pageHeader} fixed>
      <View>
        <Text style={s.brandName}>25N Coworking</Text>
        <Text style={s.brandSub}>{location.toUpperCase()}</Text>
      </View>
      <View>
        <Text style={s.reportTitle}>GL Variance Report</Text>
        <Text style={s.reportMeta}>{month}</Text>
      </View>
    </View>
  );
}

function Footer({ location, month }: { location: string; month: string }) {
  return (
    <View style={s.footer} fixed>
      <Text style={s.footerText}>Confidential — 25N Coworking · {location} · {month}</Text>
      <Text style={s.footerText} render={({ pageNumber, totalPages }) => `Page ${pageNumber} of ${totalPages}`} />
    </View>
  );
}

function SectionTitle({ children, note }: { children: React.ReactNode; note?: string }) {
  return (
    <View style={s.sectionBlock}>
      <Text style={s.sectionTitle}>{children}</Text>
      <View style={s.sectionRule} />
      {note && <Text style={s.sectionNote}>{note}</Text>}
    </View>
  );
}

// ── Document ──────────────────────────────────────────────────────────────

export function GLVarianceDocument({ data }: { data: GLVarianceData }) {
  const sortedFlags = [...data.flags].sort((a, b) => a.account.localeCompare(b.account));
  const sortedViolations = [...data.violations].sort((a, b) => a.account.localeCompare(b.account));
  const sortedJE = [...data.journal_entry_accounts].sort((a, b) => a.account.localeCompare(b.account));

  return (
    <Document title={`25N ${data.location} — ${data.month} GL Variance Report`} author="25N Coworking / Wex Advisory">
      <Page size="LETTER" style={s.page}>
        <Header location={data.location} month={data.month} />

        <View style={s.kpiRow}>
          <View style={[s.kpiCard, { backgroundColor: data.flag_count === 0 ? "#14532D" : C.navy }]}>
            <Text style={s.kpiLabel}>Variance Flags</Text>
            <Text style={s.kpiValue}>{data.flag_count}</Text>
            <Text style={s.kpiSub}>vs {data.prior_month}</Text>
          </View>
          <View style={[s.kpiCard, { backgroundColor: data.violation_count === 0 ? "#14532D" : "#7F1D1D" }]}>
            <Text style={s.kpiLabel}>Control # Issues</Text>
            <Text style={s.kpiValue}>{data.violation_count}</Text>
            <Text style={s.kpiSub}>wrong prefix / vendor mismatch</Text>
          </View>
          <View style={[s.kpiCard, { backgroundColor: data.journal_entry_account_count === 0 ? "#14532D" : "#78350F" }]}>
            <Text style={s.kpiLabel}>Needs JE Review</Text>
            <Text style={s.kpiValue}>{data.journal_entry_account_count}</Text>
            <Text style={s.kpiSub}>accounts with journal entries</Text>
          </View>
          <View style={s.kpiCard}>
            <Text style={s.kpiLabel}>Clean Accounts</Text>
            <Text style={s.kpiValue}>{data.clean_count}</Text>
            <Text style={s.kpiSub}>no flags this period</Text>
          </View>
        </View>

        {/* Variance flags */}
        <SectionTitle note={`Prior $0 → flag if current exceeds $500. Prior over $1,000 → flag if change exceeds the larger of $500 or 20% of prior. Prior $1–$1,000 → flag if change exceeds $500.`}>
          MoM Variance Flags ({data.flag_count})
        </SectionTitle>
        {sortedFlags.length === 0 ? (
          <Text style={s.emptyState}>No variances exceed thresholds vs {data.prior_month}.</Text>
        ) : (
          <View style={s.table}>
            <View style={s.tHead}>
              <Text style={[s.th, { flex: 0.9 }]}>Acct</Text>
              <Text style={[s.th, { flex: 2.6 }]}>Description</Text>
              <Text style={[s.th, { flex: 1.2, textAlign: "right" }]}>Prior</Text>
              <Text style={[s.th, { flex: 1.2, textAlign: "right" }]}>Current</Text>
              <Text style={[s.th, { flex: 1.2, textAlign: "right" }]}>Change</Text>
              <Text style={[s.th, { flex: 0.9, textAlign: "right" }]}>%</Text>
            </View>
            {sortedFlags.map((f, i) => {
              const isRev = f.account.startsWith("4") || f.account.startsWith("9");
              const unfav = isRev ? f.variance < 0 : f.variance > 0;
              return (
                <View key={f.account} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]} wrap={false}>
                  <Text style={[s.tdMono, { flex: 0.9 }]}>{f.account}</Text>
                  <Text style={[s.td, { flex: 2.6 }]}>{f.name}</Text>
                  <Text style={[s.tdGray, { flex: 1.2, textAlign: "right" }]}>{fmt$(f.prior)}</Text>
                  <Text style={[s.td, { flex: 1.2, textAlign: "right" }]}>{fmt$(f.current)}</Text>
                  <Text style={[unfav ? s.tdRed : s.tdGreen, { flex: 1.2, textAlign: "right" }]}>{fmt$(f.variance)}</Text>
                  <Text style={[unfav ? s.tdRed : s.tdGreen, { flex: 0.9, textAlign: "right" }]}>{fmtPct(f.variance_pct)}</Text>
                </View>
              );
            })}
          </View>
        )}

        {/* Journal entry review */}
        <SectionTitle note="A journal entry means a human made a manual adjustment — flagged for independent review regardless of whether the account breached the variance threshold above.">
          Journal Entry Review Required ({data.journal_entry_account_count})
        </SectionTitle>
        {sortedJE.length === 0 ? (
          <Text style={s.emptyState}>No journal entries posted this period.</Text>
        ) : (
          <View style={s.table}>
            <View style={s.tHead}>
              <Text style={[s.th, { flex: 0.9 }]}>Acct</Text>
              <Text style={[s.th, { flex: 3 }]}>Description</Text>
              <Text style={[s.th, { flex: 1.3, textAlign: "right" }]}># Entries</Text>
              <Text style={[s.th, { flex: 1.3, textAlign: "right" }]}>Net Amount</Text>
            </View>
            {sortedJE.map((e, i) => (
              <View key={e.account} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]} wrap={false}>
                <Text style={[s.tdMono, { flex: 0.9 }]}>{e.account}</Text>
                <Text style={[s.td, { flex: 3 }]}>{e.account_name}</Text>
                <Text style={[s.tdAmber, { flex: 1.3, textAlign: "right" }]}>{e.transaction_count}</Text>
                <Text style={[s.td, { flex: 1.3, textAlign: "right" }]}>{fmt$(e.total_amount)}</Text>
              </View>
            ))}
          </View>
        )}

        <Footer location={data.location} month={data.month} />
      </Page>

      {/* Page 2: Control # / vendor violations — "are the account numbers correct" */}
      <Page size="LETTER" style={s.page}>
        <Header location={data.location} month={data.month} />
        <SectionTitle note="Every posted transaction should carry a Control # prefix appropriate for its account section (R = recurring/rent, K = cash disbursement, C = charge/revenue, J = journal entry, P = purchase/vendor payment). These didn't match what's expected for the account — or, for vendor-tagged accounts, the vendor on the transaction doesn't match the vendor of record.">
          Control # &amp; Vendor Issues ({data.violation_count})
        </SectionTitle>
        {sortedViolations.length === 0 ? (
          <Text style={s.emptyState}>No control # or vendor mismatches found — every transaction ties out correctly.</Text>
        ) : (
          <View style={s.table}>
            <View style={s.tHead}>
              <Text style={[s.th, { flex: 0.8 }]}>Acct</Text>
              <Text style={[s.th, { flex: 1.6 }]}>Description</Text>
              <Text style={[s.th, { flex: 1 }]}>Control #</Text>
              <Text style={[s.th, { flex: 1, textAlign: "right" }]}>Amount</Text>
              <Text style={[s.th, { flex: 3.6 }]}>Issue</Text>
            </View>
            {sortedViolations.map((v, i) => (
              <View key={`${v.account}-${v.control}-${i}`} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]} wrap={false}>
                <Text style={[s.tdMono, { flex: 0.8 }]}>{v.account}</Text>
                <Text style={[s.td, { flex: 1.6 }]}>{v.account_name}</Text>
                <Text style={[s.tdMono, { flex: 1 }]}>{v.control}</Text>
                <Text style={[s.td, { flex: 1, textAlign: "right" }]}>{fmt$(v.amount)}</Text>
                <Text style={[s.tdRed, { flex: 3.6, fontFamily: "Helvetica" }]}>
                  {v.violation_type === "wrong_control_prefix" ? "Unexpected control # prefix for this account" : v.violation_type === "wrong_vendor" ? "Vendor on transaction doesn't match vendor of record" : v.violation_type}
                </Text>
              </View>
            ))}
          </View>
        )}
        <Footer location={data.location} month={data.month} />
      </Page>
    </Document>
  );
}
