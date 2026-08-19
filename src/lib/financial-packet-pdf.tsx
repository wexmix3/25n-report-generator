import React from "react";
import { Document, Page, Text, View, StyleSheet, Font } from "@react-pdf/renderer";

Font.register({ family: "Helvetica", fonts: [] });

// ── Palette ───────────────────────────────────────────────────────────────────
const C = {
  navy:    "#0F2044",
  gold:    "#C9A84C",
  green:   "#16A34A",
  red:     "#DC2626",
  orange:  "#E07A3E",
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
  fixedHeader: {
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
  footer: {
    position: "absolute", bottom: 14, left: 44, right: 44,
    flexDirection: "row", justifyContent: "space-between",
    borderTopWidth: 1, borderTopColor: C.gray200, paddingTop: 4,
  },
  footerText: { fontSize: 6, color: C.gray400 },
  sectionBlock: { alignItems: "center", marginTop: 16, marginBottom: 8 },
  sectionTitle: {
    fontSize: 8.5,
    fontFamily: "Helvetica-Bold",
    color: C.navy,
    textTransform: "uppercase",
    letterSpacing: 1.0,
    textAlign: "center",
  },
  sectionRule: {
    width: 28, height: 2,
    backgroundColor: C.gold,
    marginTop: 3, alignSelf: "center",
  },
  kpiRow:  { flexDirection: "row", gap: 10, marginBottom: 14 },
  kpiCard: { flex: 1, backgroundColor: C.navy, borderRadius: 5, padding: 10, alignItems: "center" },
  kpiLabel: { fontSize: 6, color: C.gold, textTransform: "uppercase", letterSpacing: 0.8, marginBottom: 3, textAlign: "center" },
  kpiValue: { fontSize: 18, fontFamily: "Helvetica-Bold", color: C.white, textAlign: "center" },
  kpiSub:   { fontSize: 5.5, color: C.gray400, marginTop: 3, textAlign: "center" },
  table:   { width: "100%", marginBottom: 6 },
  tHead:   { flexDirection: "row", backgroundColor: C.navy, paddingVertical: 5, paddingHorizontal: 5, borderRadius: 3 },
  tRow:    { flexDirection: "row", paddingVertical: 3, paddingHorizontal: 5, borderBottomWidth: 1, borderBottomColor: C.gray100 },
  tRowAlt: { backgroundColor: C.gray50 },
  tRowSub: { backgroundColor: C.gray100, borderTopWidth: 1, borderTopColor: C.gray300 },
  th:      { fontSize: 6.5, fontFamily: "Helvetica-Bold", color: C.white, paddingRight: 4 },
  td:      { fontSize: 7, color: C.gray800, paddingRight: 4 },
  tdBold:  { fontSize: 7, fontFamily: "Helvetica-Bold", color: C.gray800, paddingRight: 4 },
  tdGray:  { fontSize: 7, color: C.gray400, paddingRight: 4 },
  tdRed:   { fontSize: 7, fontFamily: "Helvetica-Bold", color: C.red, paddingRight: 4 },
  tdGreen: { fontSize: 7, fontFamily: "Helvetica-Bold", color: C.green, paddingRight: 4 },
});

// ── Types (matching build_packet.py output) ───────────────────────────────────

interface LineItem { actual: number; budget: number; variance: number; }
interface IS {
  revenue: {
    membership: LineItem; registration_access: LineItem; workspace_rental: LineItem;
    meeting_space: LineItem; package_revenue: LineItem; member_amenities: LineItem;
    miscellaneous: LineItem; _total: LineItem;
  };
  cos: { direct_cos: LineItem; community: LineItem; _total: LineItem; };
  gross_profit: LineItem & { margin_pct: number };
  opex: {
    payroll: LineItem; facilities: LineItem; insurance: LineItem; admin_legal: LineItem;
    marketing: LineItem; meals_entertainment: LineItem; office_supplies: LineItem;
    technology: LineItem; travel: LineItem; utilities: LineItem; other: LineItem; _total: LineItem;
  };
  net_operating_income: LineItem & { margin_pct: number };
  net_income: { actual: number; margin_pct: number };
}

interface CashFlow {
  beginning_cash: number;
  net_cash_flow: number;
  ending_cash: number;
  net_cash_flow_ytd?: number;
  ytd_complete?: boolean;
  error?: string;
}

interface BSSummary {
  cash_and_bank: number; receivables: number; prepaid_and_other: number;
  fixed_assets: number; other_assets: number; total_assets: number;
  current_liabilities: number; long_term_liabilities: number;
  equity: number; total_liabilities_and_equity: number; balance_check: number;
}

interface AgingTotals {
  current: number; d0_30: number; d31_60: number; d61_90: number;
  over_90: number; total_owed: number; future?: number;
}

interface APVendor { payee_name: string; current: number; d0_30: number; d31_60: number; d61_90: number; over_90: number; }

export interface FinancialPacketData {
  location: string;
  month: string;
  generated_at: string;
  income_statement: IS;
  balance_sheet?: { summary: BSSummary; error?: string };
  ar_aging?: { customer_count: number; totals: AgingTotals; delinquent_customers: Record<string, { d61_90: number; over_90: number }>; error?: string };
  ap_aging?: { vendor_count: number; totals: AgingTotals; overdue_vendors: Record<string, APVendor>; overdue_count: number; error?: string };
  cash_flow?: CashFlow;
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function fmt$(n: number | null | undefined): string {
  if (n == null) return "—";
  const abs = Math.abs(n);
  const str = abs >= 1000
    ? `$${abs.toLocaleString("en-US", { maximumFractionDigits: 0 })}`
    : `$${Math.round(abs)}`;
  return n < 0 ? `(${str})` : str;
}

function fmtPct(n: number | null | undefined): string {
  if (n == null) return "—";
  return `${n.toFixed(1)}%`;
}

function variance$(actual: number, budget: number): string {
  return fmt$(actual - budget);
}

// ── Shared layout ─────────────────────────────────────────────────────────────

function Header({ location, month, subtitle }: { location: string; month: string; subtitle: string }) {
  return (
    <View style={s.pageHeader}>
      <View>
        <Text style={s.brandName}>25N Coworking</Text>
        <Text style={s.brandSub}>{location.toUpperCase()}</Text>
      </View>
      <View>
        <Text style={s.reportTitle}>Financial Packet</Text>
        <Text style={s.reportMeta}>{month} · {subtitle}</Text>
      </View>
    </View>
  );
}

function FixedHeader({ location, month, subtitle }: { location: string; month: string; subtitle: string }) {
  return (
    <View style={s.fixedHeader} fixed>
      <View>
        <Text style={s.brandName}>25N Coworking</Text>
        <Text style={s.brandSub}>{location.toUpperCase()}</Text>
      </View>
      <View>
        <Text style={s.reportTitle}>Financial Packet</Text>
        <Text style={s.reportMeta}>{month} · {subtitle}</Text>
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

function SectionTitle({ children }: { children: React.ReactNode }) {
  return (
    <View style={s.sectionBlock}>
      <Text style={s.sectionTitle}>{children}</Text>
      <View style={s.sectionRule} />
    </View>
  );
}

// ── Page 1: Cover / KPIs ──────────────────────────────────────────────────────

function CoverPage({ data }: { data: FinancialPacketData }) {
  const is = data.income_statement;
  const rev = is.revenue._total.actual;
  const noi = is.net_operating_income.actual;
  const ni = is.net_income.actual;

  return (
    <Page size="LETTER" style={s.page}>
      <Header location={data.location} month={data.month} subtitle="Confidential" />

      <View style={s.kpiRow}>
        <View style={s.kpiCard}>
          <Text style={s.kpiLabel}>Revenue</Text>
          <Text style={s.kpiValue}>{fmt$(rev)}</Text>
          <Text style={s.kpiSub}>vs {fmt$(is.revenue._total.budget)} budget</Text>
        </View>
        <View style={[s.kpiCard, { backgroundColor: ni >= 0 ? "#14532D" : "#7F1D1D" }]}>
          <Text style={s.kpiLabel}>Net Margin</Text>
          <Text style={s.kpiValue}>{fmtPct(is.net_income.margin_pct)}</Text>
          <Text style={s.kpiSub}>{fmt$(ni)} net income</Text>
        </View>
        <View style={[s.kpiCard, { backgroundColor: noi >= 0 ? "#14532D" : "#7F1D1D" }]}>
          <Text style={s.kpiLabel}>Net Op. Income</Text>
          <Text style={s.kpiValue}>{fmt$(noi)}</Text>
          <Text style={s.kpiSub}>{fmtPct(is.net_operating_income.margin_pct)} margin</Text>
        </View>
        <View style={[s.kpiCard, { backgroundColor: ni >= 0 ? "#14532D" : "#7F1D1D" }]}>
          <Text style={s.kpiLabel}>Net Income</Text>
          <Text style={s.kpiValue}>{fmt$(ni)}</Text>
          <Text style={s.kpiSub}>{fmtPct(is.net_income.margin_pct)} margin</Text>
        </View>
      </View>

      {/* Quick IS summary */}
      <SectionTitle>Income Statement Summary</SectionTitle>
      <View style={s.table}>
        <View style={s.tHead}>
          <Text style={[s.th, { flex: 3 }]}>Line Item</Text>
          <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Actual</Text>
          <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Budget</Text>
          <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Variance</Text>
        </View>
        {[
          { label: "Revenue",             li: is.revenue._total, bold: true, isExpense: false },
          { label: "  Workspace Rental",  li: is.revenue.workspace_rental, isExpense: false },
          { label: "  Meeting Space",     li: is.revenue.meeting_space, isExpense: false },
          { label: "  Package Revenue",   li: is.revenue.package_revenue, isExpense: false },
          { label: "Cost of Sales",       li: is.cos._total, bold: true, isExpense: true },
          { label: "Gross Profit",        li: is.gross_profit, bold: true, isExpense: false },
          { label: "Total OPEX",          li: is.opex._total, bold: true, isExpense: true },
          { label: "  Payroll",           li: is.opex.payroll, isExpense: true },
          { label: "  Facilities",        li: is.opex.facilities, isExpense: true },
          { label: "  Insurance",         li: is.opex.insurance, isExpense: true },
          { label: "NOI",                 li: is.net_operating_income, bold: true, isExpense: false },
        ].map(({ label, li, bold, isExpense }, i) => {
          const favorable = isExpense ? li.actual <= li.budget : li.actual >= li.budget;
          return (
          <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}, bold ? s.tRowSub : {}]}>
            <Text style={[bold ? s.tdBold : s.td, { flex: 3 }]}>{label}</Text>
            <Text style={[bold ? s.tdBold : s.td, { flex: 1.5, textAlign: "right" }]}>{fmt$(li.actual)}</Text>
            <Text style={[s.tdGray, { flex: 1.5, textAlign: "right" }]}>{fmt$(li.budget)}</Text>
            <Text style={[favorable ? s.tdGreen : s.tdRed, { flex: 1.5, textAlign: "right" }]}>
              {variance$(li.actual, li.budget)}
            </Text>
          </View>
          );
        })}
      </View>

      <Footer location={data.location} month={data.month} />
    </Page>
  );
}

// ── Page 2: Full Income Statement ─────────────────────────────────────────────

function ISPage({ data }: { data: FinancialPacketData }) {
  const is = data.income_statement;

  const rows: { label: string; li: { actual: number; budget: number }; indent?: boolean; bold?: boolean; isExpense?: boolean }[] = [
    { label: "REVENUE", li: is.revenue._total, bold: true },
    { label: "Workspace Rental", li: is.revenue.workspace_rental, indent: true },
    { label: "Meeting Space", li: is.revenue.meeting_space, indent: true },
    { label: "Package Revenue", li: is.revenue.package_revenue, indent: true },
    { label: "Member Amenities", li: is.revenue.member_amenities, indent: true },
    { label: "Membership", li: is.revenue.membership, indent: true },
    { label: "Registration & Access", li: is.revenue.registration_access, indent: true },
    { label: "Miscellaneous", li: is.revenue.miscellaneous, indent: true },
    { label: "COST OF SALES", li: is.cos._total, bold: true, isExpense: true },
    { label: "Direct COS", li: is.cos.direct_cos, indent: true, isExpense: true },
    { label: "Community", li: is.cos.community, indent: true, isExpense: true },
    { label: `GROSS PROFIT (${fmtPct(is.gross_profit.margin_pct)})`, li: is.gross_profit, bold: true },
    { label: "OPERATING EXPENSES", li: is.opex._total, bold: true, isExpense: true },
    { label: "Payroll", li: is.opex.payroll, indent: true, isExpense: true },
    { label: "Facilities", li: is.opex.facilities, indent: true, isExpense: true },
    { label: "Insurance", li: is.opex.insurance, indent: true, isExpense: true },
    { label: "Utilities", li: is.opex.utilities, indent: true, isExpense: true },
    { label: "Marketing", li: is.opex.marketing, indent: true, isExpense: true },
    { label: "Meals & Entertainment", li: is.opex.meals_entertainment, indent: true, isExpense: true },
    { label: "Office Expense", li: is.opex.office_supplies, indent: true, isExpense: true },
    { label: "Technology", li: is.opex.technology, indent: true, isExpense: true },
    { label: "Travel", li: is.opex.travel, indent: true, isExpense: true },
    { label: "Admin / Legal", li: is.opex.admin_legal, indent: true, isExpense: true },
    { label: "Other OPEX", li: is.opex.other, indent: true, isExpense: true },
    { label: `NET OPERATING INCOME (${fmtPct(is.net_operating_income.margin_pct)})`, li: is.net_operating_income, bold: true },
  ];

  return (
    <Page size="LETTER" style={s.page}>
      <FixedHeader location={data.location} month={data.month} subtitle="Income Statement" />
      <View style={{ marginTop: 40 }}>
        <SectionTitle>Income Statement — MTD vs Budget</SectionTitle>
        <View style={s.table}>
          <View style={s.tHead}>
            <Text style={[s.th, { flex: 3.5 }]}>Account</Text>
            <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Actual MTD</Text>
            <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Budget</Text>
            <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Variance</Text>
          </View>
          {rows.map(({ label, li, indent, bold, isExpense }, i) => {
            const favorable = isExpense ? li.actual <= li.budget : li.actual >= li.budget;
            return (
            <View key={i} style={[s.tRow, i % 2 === 1 && !bold ? s.tRowAlt : {}, bold ? s.tRowSub : {}]}>
              <Text style={[bold ? s.tdBold : s.td, { flex: 3.5, paddingLeft: indent ? 12 : 0 }]}>{label}</Text>
              <Text style={[bold ? s.tdBold : s.td, { flex: 1.5, textAlign: "right" }]}>{fmt$(li.actual)}</Text>
              <Text style={[s.tdGray, { flex: 1.5, textAlign: "right" }]}>{fmt$(li.budget)}</Text>
              <Text style={[favorable ? s.tdGreen : s.tdRed, { flex: 1.5, textAlign: "right" }]}>
                {variance$(li.actual, li.budget)}
              </Text>
            </View>
            );
          })}
        </View>
      </View>
      <Footer location={data.location} month={data.month} />
    </Page>
  );
}

// ── Page 3: Balance Sheet ─────────────────────────────────────────────────────

function BSPage({ data }: { data: FinancialPacketData }) {
  const bs = data.balance_sheet;
  if (!bs || bs.error) return null;
  const sm = bs.summary;

  const assetRows = [
    { label: "Cash & Bank", val: sm.cash_and_bank },
    { label: "Receivables", val: sm.receivables },
    { label: "Prepaid & Other Current Assets", val: sm.prepaid_and_other },
    { label: "Fixed Assets", val: sm.fixed_assets },
    { label: "Other Assets", val: sm.other_assets },
  ];
  const liabRows = [
    { label: "Current Liabilities", val: sm.current_liabilities },
    { label: "Long-term Liabilities", val: sm.long_term_liabilities },
    { label: "Equity", val: sm.equity },
  ];

  return (
    <Page size="LETTER" style={s.page}>
      <FixedHeader location={data.location} month={data.month} subtitle="Balance Sheet" />
      <View style={{ marginTop: 40 }}>
        <SectionTitle>Balance Sheet — End of {data.month}</SectionTitle>

        {/* Assets */}
        <View style={s.table}>
          <View style={s.tHead}>
            <Text style={[s.th, { flex: 3 }]}>ASSETS</Text>
            <Text style={[s.th, { flex: 2, textAlign: "right" }]}>Balance</Text>
          </View>
          {assetRows.map(({ label, val }, i) => (
            <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
              <Text style={[s.td, { flex: 3, paddingLeft: 10 }]}>{label}</Text>
              <Text style={[s.td, { flex: 2, textAlign: "right" }]}>{fmt$(val)}</Text>
            </View>
          ))}
          <View style={[s.tRow, s.tRowSub]}>
            <Text style={[s.tdBold, { flex: 3 }]}>Total Assets</Text>
            <Text style={[s.tdBold, { flex: 2, textAlign: "right" }]}>{fmt$(sm.total_assets)}</Text>
          </View>
        </View>

        {/* Liabilities + Equity */}
        <View style={[s.table, { marginTop: 10 }]}>
          <View style={s.tHead}>
            <Text style={[s.th, { flex: 3 }]}>LIABILITIES &amp; EQUITY</Text>
            <Text style={[s.th, { flex: 2, textAlign: "right" }]}>Balance</Text>
          </View>
          {liabRows.map(({ label, val }, i) => (
            <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
              <Text style={[s.td, { flex: 3, paddingLeft: 10 }]}>{label}</Text>
              <Text style={[s.td, { flex: 2, textAlign: "right" }]}>{fmt$(val)}</Text>
            </View>
          ))}
          <View style={[s.tRow, s.tRowSub]}>
            <Text style={[s.tdBold, { flex: 3 }]}>Total Liabilities + Equity</Text>
            <Text style={[s.tdBold, { flex: 2, textAlign: "right" }]}>{fmt$(sm.total_liabilities_and_equity)}</Text>
          </View>
        </View>

        {Math.abs(sm.balance_check) > 1 && (
          <View style={{ marginTop: 8, backgroundColor: "#FEF3C7", borderRadius: 3, padding: 6 }}>
            <Text style={{ fontSize: 6.5, color: "#92400E" }}>
              Balance check: {fmt$(sm.balance_check)} difference — review retained earnings / 9xxx accounts
            </Text>
          </View>
        )}
      </View>
      <Footer location={data.location} month={data.month} />
    </Page>
  );
}

// ── Page 4: Cash Flow ─────────────────────────────────────────────────────────

function CashFlowPage({ data }: { data: FinancialPacketData }) {
  const cf = data.cash_flow;
  if (!cf || cf.error) return null;

  const rows: { label: string; val: number; bold?: boolean }[] = [
    { label: `Beginning Cash — ${data.month}`, val: cf.beginning_cash },
    { label: `Net Cash Flow — ${data.month}`, val: cf.net_cash_flow },
    { label: `Ending Cash — ${data.month}`, val: cf.ending_cash, bold: true },
  ];
  if (cf.net_cash_flow_ytd !== undefined) {
    rows.push({
      label: `Net Cash Flow — YTD${cf.ytd_complete === false ? " (partial)" : ""}`,
      val: cf.net_cash_flow_ytd,
    });
  }

  return (
    <Page size="LETTER" style={s.page}>
      <FixedHeader location={data.location} month={data.month} subtitle="Cash Flow" />
      <View style={{ marginTop: 40 }}>
        <SectionTitle>Cash Flow</SectionTitle>
        <View style={s.table}>
          {rows.map(({ label, val, bold }, i) => (
            <View key={i} style={[s.tRow, bold ? s.tRowSub : (i % 2 === 1 ? s.tRowAlt : {})]}>
              <Text style={[bold ? s.tdBold : s.td, { flex: 3 }]}>{label}</Text>
              <Text style={[bold ? s.tdBold : s.td, { flex: 2, textAlign: "right" }]}>{fmt$(val)}</Text>
            </View>
          ))}
        </View>
      </View>
      <Footer location={data.location} month={data.month} />
    </Page>
  );
}

// ── Page 5: AR Aging ──────────────────────────────────────────────────────────

function ARPage({ data }: { data: FinancialPacketData }) {
  const ar = data.ar_aging;
  if (!ar || ar.error) return null;
  const t = ar.totals;
  const delinquent = Object.entries(ar.delinquent_customers ?? {});

  return (
    <Page size="LETTER" style={s.page}>
      <FixedHeader location={data.location} month={data.month} subtitle="AR Aging" />
      <View style={{ marginTop: 40 }}>
        <SectionTitle>Accounts Receivable Aging</SectionTitle>
        <View style={s.table}>
          <View style={s.tHead}>
            <Text style={[s.th, { flex: 3 }]}>Aging Bucket</Text>
            <Text style={[s.th, { flex: 2, textAlign: "right" }]}>Amount</Text>
          </View>
          {[
            { label: "Current", val: t.current },
            { label: "0–30 Days", val: t.d0_30 },
            { label: "31–60 Days", val: t.d31_60 },
            { label: "61–90 Days", val: t.d61_90 },
            { label: "Over 90 Days", val: t.over_90 },
          ].map(({ label, val }, i) => {
            const overdue = i >= 2 && val !== 0;
            return (
              <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
                <Text style={[overdue ? s.tdRed : s.td, { flex: 3, paddingLeft: 10 }]}>{label}</Text>
                <Text style={[overdue ? s.tdRed : s.td, { flex: 2, textAlign: "right" }]}>{fmt$(val)}</Text>
              </View>
            );
          })}
          <View style={[s.tRow, s.tRowSub]}>
            <Text style={[s.tdBold, { flex: 3 }]}>Total Receivable</Text>
            <Text style={[s.tdBold, { flex: 2, textAlign: "right" }]}>{fmt$(t.total_owed)}</Text>
          </View>
        </View>

        {delinquent.length > 0 && (
          <>
            <SectionTitle>Delinquent Accounts (61+ Days)</SectionTitle>
            <View style={s.table}>
              <View style={s.tHead}>
                <Text style={[s.th, { flex: 4 }]}>Customer</Text>
                <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>61–90</Text>
                <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Over 90</Text>
              </View>
              {delinquent.map(([customer, c], i) => (
                <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
                  <Text style={[s.td, { flex: 4 }]}>{customer}</Text>
                  <Text style={[c.d61_90 !== 0 ? s.tdRed : s.tdGray, { flex: 1.5, textAlign: "right" }]}>{fmt$(c.d61_90)}</Text>
                  <Text style={[c.over_90 !== 0 ? s.tdRed : s.tdGray, { flex: 1.5, textAlign: "right" }]}>{fmt$(c.over_90)}</Text>
                </View>
              ))}
            </View>
          </>
        )}
      </View>
      <Footer location={data.location} month={data.month} />
    </Page>
  );
}

// ── Page 5: AP Aging ──────────────────────────────────────────────────────────

function APPage({ data }: { data: FinancialPacketData }) {
  const ap = data.ap_aging;
  if (!ap || ap.error) return null;
  const t = ap.totals;
  const overdue = Object.entries(ap.overdue_vendors);

  return (
    <Page size="LETTER" style={s.page}>
      <FixedHeader location={data.location} month={data.month} subtitle="AP Aging" />
      <View style={{ marginTop: 40 }}>
        <SectionTitle>Accounts Payable Aging</SectionTitle>
        <View style={s.table}>
          <View style={s.tHead}>
            <Text style={[s.th, { flex: 3 }]}>Aging Bucket</Text>
            <Text style={[s.th, { flex: 2, textAlign: "right" }]}>Amount</Text>
          </View>
          {/* "Current" is the invoice face value on every AP line, not a distinct
              aging bucket — it's already reflected inside 0-30/31-60/etc, so it's
              excluded here to avoid double-counting the summary. */}
          {[
            { label: "0–30 Days", val: t.d0_30 },
            { label: "31–60 Days", val: t.d31_60 },
            { label: "61–90 Days", val: t.d61_90 },
            { label: "Over 90 Days", val: t.over_90 },
          ].map(({ label, val }, i) => {
            const overdue = i >= 1 && val !== 0;
            return (
              <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
                <Text style={[overdue ? s.tdRed : s.td, { flex: 3, paddingLeft: 10 }]}>{label}</Text>
                <Text style={[overdue ? s.tdRed : s.td, { flex: 2, textAlign: "right" }]}>{fmt$(val)}</Text>
              </View>
            );
          })}
          <View style={[s.tRow, s.tRowSub]}>
            <Text style={[s.tdBold, { flex: 3 }]}>Total Payable</Text>
            <Text style={[s.tdBold, { flex: 2, textAlign: "right" }]}>{fmt$(t.total_owed)}</Text>
          </View>
        </View>

        {overdue.length > 0 && (
          <>
            <SectionTitle>Overdue Vendors (31+ Days)</SectionTitle>
            <View style={s.table}>
              <View style={s.tHead}>
                <Text style={[s.th, { flex: 3 }]}>Vendor</Text>
                <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>31–60</Text>
                <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>61–90</Text>
                <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Over 90</Text>
              </View>
              {overdue.sort((a, b) => (b[1].d31_60 + b[1].d61_90 + b[1].over_90) - (a[1].d31_60 + a[1].d61_90 + a[1].over_90)).map(([, v], i) => (
                <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
                  <Text style={[s.td, { flex: 3 }]}>{v.payee_name}</Text>
                  <Text style={[v.d31_60 !== 0 ? s.tdRed : s.tdGray, { flex: 1.5, textAlign: "right" }]}>{fmt$(v.d31_60)}</Text>
                  <Text style={[v.d61_90 !== 0 ? s.tdRed : s.tdGray, { flex: 1.5, textAlign: "right" }]}>{fmt$(v.d61_90)}</Text>
                  <Text style={[v.over_90 !== 0 ? s.tdRed : s.tdGray, { flex: 1.5, textAlign: "right" }]}>{fmt$(v.over_90)}</Text>
                </View>
              ))}
            </View>
          </>
        )}
      </View>
      <Footer location={data.location} month={data.month} />
    </Page>
  );
}

// ── Main Document ─────────────────────────────────────────────────────────────

export function FinancialPacketDocument({ data }: { data: FinancialPacketData }) {
  return (
    <Document
      title={`25N ${data.location} — ${data.month} Financial Packet`}
      author="25N Coworking / Wex Advisory"
    >
      <CoverPage data={data} />
      <ISPage data={data} />
      {data.balance_sheet && !data.balance_sheet.error && <BSPage data={data} />}
      {data.cash_flow && !data.cash_flow.error && <CashFlowPage data={data} />}
      {data.ar_aging && !data.ar_aging.error && <ARPage data={data} />}
      {data.ap_aging && !data.ap_aging.error && <APPage data={data} />}
    </Document>
  );
}
