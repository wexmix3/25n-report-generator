import React from "react";
import {
  Document, Page, Text, View, StyleSheet, Font,
  Svg, Rect, Line, Path, G,
} from "@react-pdf/renderer";
import type {
  ReportData, BudgetLineItem, OccupancyRow, RevenueLineItem,
  MeetingRoomTypeRow, MeetingRoomTransaction, ValidationCheck,
} from "@/types/report";

Font.register({ family: "Helvetica", fonts: [] });

// ── Palette ──────────────────────────────────────────────────────────────────
const C = {
  navy:     "#0F2044",
  navyMid:  "#1A3461",
  gold:     "#C9A84C",
  goldLight:"#F0DFA0",
  green:    "#16A34A",
  red:      "#DC2626",
  gray50:   "#F9FAFB",
  gray100:  "#F3F4F6",
  gray200:  "#E5E7EB",
  gray300:  "#D1D5DB",
  gray400:  "#9CA3AF",
  gray600:  "#4B5563",
  gray800:  "#1F2937",
  white:    "#FFFFFF",
  occGreen: "#DCFCE7",
  occRed:   "#FEE2E2",
  barBlue:  "#3B6ECC",
  barGold:  "#C9A84C",
};

// ── Styles ───────────────────────────────────────────────────────────────────
const s = StyleSheet.create({
  page: {
    fontFamily: "Helvetica",
    fontSize: 8,
    paddingTop: 40,
    paddingBottom: 44,
    paddingHorizontal: 44,
    backgroundColor: C.white,
  },

  // ── Page header
  pageHeader: {
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "flex-end",
    marginBottom: 20,
    paddingBottom: 12,
    borderBottomWidth: 2,
    borderBottomColor: C.navy,
  },
  brandName:  { fontSize: 15, fontFamily: "Helvetica-Bold", color: C.navy },
  brandSub:   { fontSize: 8,  color: C.gold, marginTop: 2, letterSpacing: 1 },
  reportTitle:{ fontSize: 12, fontFamily: "Helvetica-Bold", color: C.navy, textAlign: "right" },
  reportDate: { fontSize: 8,  color: C.gray600, textAlign: "right", marginTop: 2 },

  // ── Section titles
  sectionBlock: { alignItems: "center", marginTop: 12, marginBottom: 8 },
  sectionTitle: {
    fontSize: 9,
    fontFamily: "Helvetica-Bold",
    color: C.navy,
    textTransform: "uppercase",
    letterSpacing: 1.2,
    textAlign: "center",
  },
  sectionRule: {
    width: 32,
    height: 2,
    backgroundColor: C.gold,
    marginTop: 4,
    alignSelf: "center",
  },

  // ── KPI cards
  kpiRow:  { flexDirection: "row", gap: 10, marginBottom: 16 },
  kpiCard: {
    flex: 1,
    backgroundColor: C.navy,
    borderRadius: 6,
    padding: 12,
    alignItems: "center",
  },
  kpiLabel: { fontSize: 6.5, color: C.gold, textTransform: "uppercase", letterSpacing: 0.8, marginBottom: 4, textAlign: "center" },
  kpiValue: { fontSize: 17, fontFamily: "Helvetica-Bold", color: C.white, textAlign: "center" },
  kpiSub:   { fontSize: 6.5, color: C.gray400, marginTop: 3, textAlign: "center" },

  // ── Space type cards
  spaceRow:   { flexDirection: "row", gap: 10, marginBottom: 14 },
  spaceCard:  { flex: 1, borderWidth: 1, borderColor: C.gray200, borderRadius: 6, padding: 10 },
  spaceCardHeader: {
    fontFamily: "Helvetica-Bold", fontSize: 8, color: C.navy,
    marginBottom: 8, paddingBottom: 5,
    borderBottomWidth: 1, borderBottomColor: C.gray200, textAlign: "center",
  },
  spaceMetaRow: { flexDirection: "row", justifyContent: "space-between", marginBottom: 3 },
  spaceLabel:   { fontSize: 7, color: C.gray600 },
  spaceValue:   { fontSize: 7, fontFamily: "Helvetica-Bold", color: C.gray800 },
  progressTrack: { height: 5, backgroundColor: C.gray200, borderRadius: 3, marginTop: 7, marginBottom: 2 },
  progressFill:  { height: 5, backgroundColor: C.navy, borderRadius: 3 },
  progressPct:   { fontSize: 7, color: C.navy, fontFamily: "Helvetica-Bold", textAlign: "right", marginTop: 1 },

  // ── Tables
  table: { width: "100%", marginBottom: 4 },
  tHead: { flexDirection: "row", backgroundColor: C.navy, paddingVertical: 5, paddingHorizontal: 6, borderRadius: 3 },
  tRow:  { flexDirection: "row", paddingVertical: 4, paddingHorizontal: 6, borderBottomWidth: 1, borderBottomColor: C.gray100 },
  tRowAlt:     { backgroundColor: C.gray50 },
  tRowTotal:   { backgroundColor: C.navy, paddingVertical: 7, borderTopWidth: 2, borderTopColor: C.gold },
  tRowSection: { backgroundColor: C.gray100, paddingVertical: 5, borderLeftWidth: 3, borderLeftColor: C.gold },
  th:          { fontSize: 7, fontFamily: "Helvetica-Bold", color: C.white },
  td:          { fontSize: 7, color: C.gray800 },
  tdBold:      { fontSize: 7, fontFamily: "Helvetica-Bold", color: C.gray800 },
  tdSection:   { fontSize: 7.5, fontFamily: "Helvetica-Bold", color: C.navy, letterSpacing: 0.4 },
  tdTotal:     { fontSize: 7.5, fontFamily: "Helvetica-Bold", color: C.goldLight },

  // ── Status pills
  pillOcc: { fontSize: 6, color: C.green, backgroundColor: C.occGreen, paddingHorizontal: 5, paddingVertical: 2, borderRadius: 4, textAlign: "center" },
  pillVac: { fontSize: 6, color: C.red,   backgroundColor: C.occRed,   paddingHorizontal: 5, paddingVertical: 2, borderRadius: 4, textAlign: "center" },

  // ── Variance colors
  varGreen:   { fontSize: 7, color: C.green },
  varRed:     { fontSize: 7, color: C.red },
  varNeutral: { fontSize: 7, color: C.gray600 },

  // ── Footer
  footer: {
    position: "absolute", bottom: 16, left: 44, right: 44,
    flexDirection: "row", justifyContent: "space-between",
    borderTopWidth: 1, borderTopColor: C.gray200, paddingTop: 5,
  },
  footerText: { fontSize: 6.5, color: C.gray400 },

  // ── Supplemental / notes
  noteText:   { fontSize: 8, color: C.gray800, marginBottom: 5, lineHeight: 1.6 },
  noteHead:   { fontSize: 9, fontFamily: "Helvetica-Bold", color: C.navy, marginTop: 12, marginBottom: 4, letterSpacing: 0.5 },
  noteBullet: { fontSize: 7.5, color: C.gray600, marginBottom: 3, marginLeft: 10, lineHeight: 1.5 },

  // ── Validation
  passRow: { backgroundColor: "#F0FDF4" },
  failRow: { backgroundColor: "#FFF1F2" },
  passText:{ fontSize: 7, color: C.green, fontFamily: "Helvetica-Bold" },
  failText:{ fontSize: 7, color: C.red,   fontFamily: "Helvetica-Bold" },
});

// ── Formatters ───────────────────────────────────────────────────────────────
function trunc(s: string, max = 36): string {
  return s.length > max ? s.slice(0, max - 1) + "…" : s;
}

function fmt$(n: number | null): string {
  if (n == null) return "—";
  return new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 }).format(n);
}
function fmtPct(n: number | null): string {
  if (n == null) return "—";
  return `${n > 0 ? "+" : ""}${n.toFixed(1)}%`;
}

// ── Shared layout pieces ─────────────────────────────────────────────────────
// WX-8: location prop replaces hardcoded "WACO, TEXAS"
function PageHeader({ title, date, location }: { title: string; date: string; location: string }) {
  return (
    <View style={s.pageHeader}>
      <View>
        <Text style={s.brandName}>25N Coworking</Text>
        <Text style={s.brandSub}>{location.toUpperCase()}</Text>
      </View>
      <View>
        <Text style={s.reportTitle}>{title}</Text>
        <Text style={s.reportDate}>As of {date}</Text>
      </View>
    </View>
  );
}

// WX-8: location prop replaces hardcoded "Waco"
function PageFooter({ location }: { location: string }) {
  return (
    <View style={s.footer} fixed>
      <Text style={s.footerText}>Confidential — 25N Coworking · {location}</Text>
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

// ── SVG Charts ───────────────────────────────────────────────────────────────
const CHART_W     = 490;
const CHART_PAD_L = 44;
const CHART_TOP_PAD = 10;
const BAR_INNER_H = 72;   // reduced from 88 — gives charts room to avoid page cut-off
const BAR_CHART_H = 108;  // reduced from 128
const LINE_INNER_H = 68;  // reduced from 82
const LINE_CHART_H = 120; // reduced from 146

function fmtK(v: number): string {
  if (v === 0) return "$0";
  return v >= 1000 ? `$${Math.round(v / 1000)}k` : `$${Math.round(v)}`;
}

function BarChart({ values, labels, color = C.barBlue, highlightLast = false, yFmt }:
  { values: number[]; labels: string[]; color?: string; highlightLast?: boolean; yFmt?: (v: number) => string }) {
  const INNER_W = CHART_W - CHART_PAD_L - 6;
  const max = Math.max(...values, 1);
  const gap  = INNER_W / values.length;
  const barW = gap * 0.55;
  const fmt  = yFmt ?? fmtK;

  const TOP = CHART_TOP_PAD;
  return (
    <Svg width={CHART_W} height={BAR_CHART_H}>
      <Line x1={CHART_PAD_L} y1={TOP} x2={CHART_PAD_L} y2={TOP + BAR_INNER_H} stroke={C.gray300} strokeWidth={1} />
      <Line x1={CHART_PAD_L} y1={TOP + BAR_INNER_H} x2={CHART_W - 4} y2={TOP + BAR_INNER_H} stroke={C.gray200} strokeWidth={1} />
      {[0.25, 0.5, 0.75, 1].map((pct, i) => {
        const y = TOP + BAR_INNER_H - pct * BAR_INNER_H;
        return (
          <G key={i}>
            <Line x1={CHART_PAD_L} y1={y} x2={CHART_W - 4} y2={y} stroke={C.gray100} strokeWidth={1} />
            <Text x={CHART_PAD_L - 3} y={y + 2.5} style={{ fontSize: 5.5, fill: C.gray400 } as any} textAnchor="end">
              {fmt(pct * max)}
            </Text>
          </G>
        );
      })}
      {values.map((v, i) => {
        const barH = Math.max((v / max) * BAR_INNER_H, v > 0 ? 1 : 0);
        const x = CHART_PAD_L + i * gap + (gap - barW) / 2;
        const y = TOP + BAR_INNER_H - barH;
        const fill = highlightLast && i === values.length - 1 ? C.barGold : color;
        return (
          <G key={i}>
            <Rect x={x} y={y} width={barW} height={barH} fill={fill} rx={2} />
          </G>
        );
      })}
      {labels.map((lbl, i) => {
        const anchor = i === 0 ? "start" : i === labels.length - 1 ? "end" : "middle";
        return (
          <G key={i}>
            <Text
              x={CHART_PAD_L + i * gap + gap / 2}
              y={TOP + BAR_INNER_H + 14}
              style={{ fontSize: 6, fill: C.gray600 } as any}
              textAnchor={anchor}
            >
              {lbl}
            </Text>
          </G>
        );
      })}
    </Svg>
  );
}

function LineChart({ series, labels, yFmt }: {
  series: { values: number[]; color: string; label: string }[];
  labels: string[];
  yFmt?: (v: number) => string;
}) {
  const INNER_W = CHART_W - CHART_PAD_L - 6;
  const allVals = series.flatMap((s) => s.values);
  const max  = Math.max(...allVals, 1);
  const step = INNER_W / Math.max(labels.length - 1, 1);
  const fmt  = yFmt ?? ((v: number) => `${Math.round(v)}`);

  const TOP = CHART_TOP_PAD;
  return (
    <Svg width={CHART_W} height={LINE_CHART_H}>
      <Line x1={CHART_PAD_L} y1={TOP} x2={CHART_PAD_L} y2={TOP + LINE_INNER_H} stroke={C.gray200} strokeWidth={1} />
      <Line x1={CHART_PAD_L} y1={TOP + LINE_INNER_H} x2={CHART_W - 4} y2={TOP + LINE_INNER_H} stroke={C.gray200} strokeWidth={1} />
      {[0.25, 0.5, 0.75, 1].map((pct, i) => {
        const y = TOP + LINE_INNER_H - pct * LINE_INNER_H;
        return (
          <G key={i}>
            <Line x1={CHART_PAD_L} y1={y} x2={CHART_W - 4} y2={y} stroke={C.gray100} strokeWidth={1} />
            <Text x={CHART_PAD_L - 3} y={y + 2.5} style={{ fontSize: 5.5, fill: C.gray400 } as any} textAnchor="end">
              {fmt(pct * max)}
            </Text>
          </G>
        );
      })}
      {series.map((serie, si) => {
        const pts = serie.values.map((v, i) => ({
          x: CHART_PAD_L + i * step,
          y: TOP + LINE_INNER_H - (v / max) * LINE_INNER_H,
        }));
        const d = pts.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x.toFixed(1)} ${p.y.toFixed(1)}`).join(" ");
        return (
          <G key={si}>
            <Path d={d} stroke={serie.color} strokeWidth={1.5} fill="none" />
            {pts.map((p, i) => (
              <Rect key={i} x={p.x - 2} y={p.y - 2} width={4} height={4} rx={2} fill={serie.color} />
            ))}
          </G>
        );
      })}
      {labels.map((lbl, i) => {
        const anchor = i === 0 ? "start" : i === labels.length - 1 ? "end" : "middle";
        return (
          <Text key={i} x={CHART_PAD_L + i * step} y={TOP + LINE_INNER_H + 13} style={{ fontSize: 5.5, fill: C.gray600 } as any} textAnchor={anchor}>
            {lbl}
          </Text>
        );
      })}
      {series.map((sr, i) => (
        <G key={i}>
          <Rect x={CHART_PAD_L + i * 100} y={TOP + LINE_INNER_H + 24} width={10} height={3} fill={sr.color} rx={1} />
          <Text x={CHART_PAD_L + i * 100 + 14} y={TOP + LINE_INNER_H + 28} style={{ fontSize: 6, fill: C.gray600 } as any}>{sr.label}</Text>
        </G>
      ))}
    </Svg>
  );
}

// ── Page 1: Executive Dashboard ──────────────────────────────────────────────
function DashboardPage({ data }: { data: ReportData }) {
  const { dashboard, occupancy } = data;
  const { totals, spaceTypes, weeklySnapshots, currentMonthRevenue, location } = dashboard;

  return (
    <Page size="LETTER" style={s.page}>
      <PageHeader title="Executive Dashboard & Key Metrics" date={dashboard.reportDate} location={location} />

      <View style={s.kpiRow}>
        {[
          { label: "Spaces Occupied",   value: `${totals.occupiedSpaces}/${totals.totalSpaces}`, sub: `${totals.occupancyPct}% occupancy rate` },
          { label: "Sq Ft Occupied",    value: totals.occupiedSqFt.toLocaleString(), sub: `${totals.sqFtPct}% of ${totals.totalSqFt.toLocaleString()} sq ft` },
          { label: "Active Contracts",  value: String(occupancy.occupiedCount), sub: "Current members" },
          { label: `${dashboard.reportDate.split(" ")[0]} Revenue`, value: fmt$(currentMonthRevenue), sub: "Actual — P4 Confirmed" },
        ].map((k) => (
          <View key={k.label} style={s.kpiCard}>
            <Text style={s.kpiLabel}>{k.label}</Text>
            <Text style={s.kpiValue}>{k.value}</Text>
            <Text style={s.kpiSub}>{k.sub}</Text>
          </View>
        ))}
      </View>

      <SectionTitle>Space Type Breakdown</SectionTitle>
      <View style={s.spaceRow}>
        {spaceTypes.map((st) => (
          <View key={st.label} style={s.spaceCard}>
            <Text style={s.spaceCardHeader}>{st.label}</Text>
            <View style={s.spaceMetaRow}>
              <Text style={s.spaceLabel}>Spaces Occupied</Text>
              <Text style={s.spaceValue}>{st.occupiedSpaces} / {st.totalSpaces}</Text>
            </View>
            <View style={s.progressTrack}>
              <View style={[s.progressFill, { width: `${st.occupancyPct}%` }]} />
            </View>
            <Text style={s.progressPct}>{st.occupancyPct}%</Text>
            <View style={[s.spaceMetaRow, { marginTop: 5 }]}>
              <Text style={s.spaceLabel}>Sq Ft</Text>
              <Text style={s.spaceValue}>{st.occupiedSqFt.toLocaleString()} / {st.totalSqFt.toLocaleString()}</Text>
            </View>
            <View style={s.spaceMetaRow}>
              <Text style={s.spaceLabel}>Sq Ft %</Text>
              <Text style={s.spaceValue}>{st.sqFtPct}%</Text>
            </View>
          </View>
        ))}
      </View>

      {weeklySnapshots.length > 0 && (
        <>
          {/* WX-7: "Monthly Snapshots" clarifies these are month-end data points */}
          <SectionTitle>Historical Trend — Last {weeklySnapshots.length} Monthly Snapshots</SectionTitle>
          <View style={s.table}>
            <View style={s.tHead}>
              {/* Mtg Hrs = contracted package hours; Mtg Billed Hrs = total hours used/billed — both in hours, not $ */}
              {["Month", "Occ %", "Sq Ft %", "Contracts", "Flex", "Mail", "Pkg Hrs", "Billed Hrs"].map((h, i) => (
                <Text key={i} style={[s.th, { flex: i === 0 ? 2 : 1, textAlign: i > 0 ? "right" : "left" }]}>{h}</Text>
              ))}
            </View>
            {weeklySnapshots.map((snap, i) => (
              <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
                <Text style={[s.td, { flex: 2 }]}>{snap.date}</Text>
                <Text style={[s.td, { flex: 1, textAlign: "right" }]}>{snap.occupancyPct}%</Text>
                <Text style={[s.td, { flex: 1, textAlign: "right" }]}>{snap.sqFtPct}%</Text>
                <Text style={[s.td, { flex: 1, textAlign: "right" }]}>{snap.contracts}</Text>
                <Text style={[s.td, { flex: 1, textAlign: "right" }]}>{snap.flexMembers}</Text>
                <Text style={[s.td, { flex: 1, textAlign: "right" }]}>{snap.mailMembers}</Text>
                <Text style={[s.td, { flex: 1, textAlign: "right" }]}>{snap.meetingHours.toFixed(0)}</Text>
                <Text style={[s.td, { flex: 1, textAlign: "right" }]}>{snap.meetingRevenue.toFixed(0)}</Text>
              </View>
            ))}
          </View>
        </>
      )}

      <PageFooter location={location} />
    </Page>
  );
}

// ── Page 2: Monthly Revenue ──────────────────────────────────────────────────
function RevenuePage({ data }: { data: ReportData }) {
  const { revenue } = data;
  const location = data.dashboard.location;
  const displayMonths = revenue.months;
  const lastIdx = displayMonths.length - 1;
  // Hide section headers (they have no figures) — data rows are self-descriptive
  const visibleItems = revenue.lineItems.filter((item) => item.label.trim() !== "" && !item.isSectionHeader);
  // YTD totals: sum each revenue line across all displayed months
  const totalLine = revenue.lineItems.find((li) => li.label.toLowerCase().includes("total with relief"));
  const ytdTotal = totalLine ? totalLine.values.reduce<number>((sum, v) => sum + (v ?? 0), 0) : 0;
  const currentMonthTotal = totalLine && lastIdx >= 0 ? (totalLine.values[lastIdx] ?? 0) : 0;

  return (
    <Page size="LETTER" style={s.page}>
      <PageHeader title="Monthly Revenue" date={data.dashboard.reportDate} location={location} />
      <View style={s.table}>
        <View style={s.tHead}>
          <Text style={[s.th, { flex: 3 }]}>Revenue Category</Text>
          {displayMonths.map((m, i) => (
            <Text key={i} style={[s.th, { flex: 1, textAlign: "right" }, i === lastIdx ? { color: C.gold } : {}]}>
              {/* WX-16: Month is complete — no "(PTD)" suffix */}
              {m.slice(0, 3)}
            </Text>
          ))}
        </View>
        {visibleItems.map((item: RevenueLineItem, i) => {
          if (item.isSectionHeader) return (
            <View key={i} style={[s.tRow, { backgroundColor: C.gray100, borderTopWidth: 1, borderTopColor: C.gray200 }]}>
              <Text style={[s.td, { flex: 3 + displayMonths.length }]}>{item.label}</Text>
            </View>
          );
          if (item.isTotal) return (
            <View key={i} style={[s.tRow, s.tRowTotal]}>
              <Text style={[s.tdTotal, { flex: 3 }]}>{item.label}</Text>
              {displayMonths.map((_, j) => (
                <Text key={j} style={[s.tdTotal, { flex: 1, textAlign: "right" }]}>
                  {item.values[j] != null ? fmt$(item.values[j] as number) : "—"}
                </Text>
              ))}
            </View>
          );
          return (
            <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
              <Text style={[s.td, { flex: 3 }]}>{item.label}</Text>
              {displayMonths.map((_, j) => (
                <Text key={j} style={[s.td, { flex: 1, textAlign: "right" }]}>
                  {item.values[j] != null ? fmt$(item.values[j] as number) : "—"}
                </Text>
              ))}
            </View>
          );
        })}
      </View>
      {/* YTD Summary strip — visible only when there are at least 2 months of data */}
      {displayMonths.length >= 2 && ytdTotal > 0 && (
        <View style={{ flexDirection: "row", marginTop: 12, gap: 10 }}>
          <View style={{ flex: 1, backgroundColor: C.navy, borderRadius: 6, padding: 10, alignItems: "center" }}>
            <Text style={{ fontSize: 6.5, color: C.gold, textTransform: "uppercase", letterSpacing: 0.8, marginBottom: 3 }}>
              {data.dashboard.reportDate.split(" ")[0]} Revenue (Current Month)
            </Text>
            <Text style={{ fontSize: 14, fontFamily: "Helvetica-Bold", color: C.white }}>{fmt$(currentMonthTotal)}</Text>
            <Text style={{ fontSize: 6, color: C.gray400, marginTop: 2 }}>Net after discounts — P2 final column</Text>
          </View>
          <View style={{ flex: 1, backgroundColor: C.navy, borderRadius: 6, padding: 10, alignItems: "center" }}>
            <Text style={{ fontSize: 6.5, color: C.gold, textTransform: "uppercase", letterSpacing: 0.8, marginBottom: 3 }}>
              Year-to-Date Revenue ({displayMonths[0]?.slice(0, 3)}–{displayMonths[lastIdx]?.slice(0, 3)})
            </Text>
            <Text style={{ fontSize: 14, fontFamily: "Helvetica-Bold", color: C.white }}>{fmt$(ytdTotal)}</Text>
            <Text style={{ fontSize: 6, color: C.gray400, marginTop: 2 }}>Cumulative net revenue — {displayMonths.length} months</Text>
          </View>
          <View style={{ flex: 1, backgroundColor: C.navy, borderRadius: 6, padding: 10, alignItems: "center" }}>
            <Text style={{ fontSize: 6.5, color: C.gold, textTransform: "uppercase", letterSpacing: 0.8, marginBottom: 3 }}>
              Monthly Average
            </Text>
            <Text style={{ fontSize: 14, fontFamily: "Helvetica-Bold", color: C.white }}>{fmt$(Math.round(ytdTotal / displayMonths.length))}</Text>
            <Text style={{ fontSize: 6, color: C.gray400, marginTop: 2 }}>Avg monthly net revenue YTD</Text>
          </View>
        </View>
      )}
      <Text style={{ fontSize: 6, color: C.gray400, marginTop: 6, marginBottom: 2 }}>
        All figures in USD. Source: P2 Monthly Income tab — includes recurring membership revenue (Dedicated Desks, Private Offices, Flex, Mail) and non-recurring revenue (Meeting Rooms, Day Passes, and other income). "Total with Relief" is net of member discounts across all categories. Current month ({displayMonths[lastIdx] ?? "—"}) represents a complete period.
      </Text>
      <PageFooter location={location} />
    </Page>
  );
}

// ── Page 3: Space Inventory ──────────────────────────────────────────────────
function OccupancyPage({ data }: { data: ReportData }) {
  const { occupancy } = data;
  const location = data.dashboard.location;
  const sorted = [...occupancy.rows].sort((a, b) => {
    if (a.status !== b.status) return a.status === "Occupied" ? -1 : 1;
    return a.spaceType.localeCompare(b.spaceType);
  });

  return (
    <Page size="LETTER" style={s.page}>
      <PageHeader title="Space Inventory" date={data.dashboard.reportDate} location={location} />
      <View style={[s.kpiRow, { marginBottom: 14 }]}>
        {[
          { label: "Total Spaces",   value: String(occupancy.totalSpaces) },
          { label: "Occupied",       value: String(occupancy.occupiedCount) },
          { label: "Vacant",         value: String(occupancy.vacantCount) },
          // WX-12: Guard against division by zero
          { label: "Occupancy Rate", value: occupancy.totalSpaces > 0 ? `${Math.round((occupancy.occupiedCount / occupancy.totalSpaces) * 100)}%` : "—" },
        ].map((k) => (
          <View key={k.label} style={s.kpiCard}>
            <Text style={s.kpiLabel}>{k.label}</Text>
            <Text style={s.kpiValue}>{k.value}</Text>
          </View>
        ))}
      </View>
      <Text style={{ fontSize: 6.5, color: C.gray600, marginBottom: 6 }}>
        Occupied spaces listed first. Status determined from P3 roster tab. Sq Ft reflects contracted or assigned square footage.
      </Text>
      <View style={s.table}>
        <View style={s.tHead}>
          {[
            { h: "Space ID",        f: 1.1 },
            { h: "Type",            f: 1.4 },
            { h: "Status",          f: 1.0 },
            { h: "Sq Ft",           f: 0.9 },
            { h: "Company / Tenant",f: 2.8 },
            { h: "Move-In",         f: 1.5 },
            { h: "Move-Out",        f: 1.5 },
            { h: "Term",            f: 0.8 },
          ].map(({ h, f }) => (
            <Text key={h} style={[s.th, { flex: f }]}>{h}</Text>
          ))}
        </View>
        {sorted.map((row: OccupancyRow, i) => (
          <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
            <Text style={[s.td, { flex: 1.1 }]}>{row.spaceId || row.spaceName || "—"}</Text>
            <Text style={[s.td, { flex: 1.4 }]}>{row.spaceType || "—"}</Text>
            <Text style={[s.td, { flex: 1.0, color: row.status === "Occupied" ? C.green : C.red, fontFamily: "Helvetica-Bold" }]}>{row.status}</Text>
            {/* "sf" suffix makes sq ft self-documenting on overflow pages where the column header isn't visible */}
            <Text style={[s.td, { flex: 0.9, textAlign: "right" }]}>{row.sqFt > 0 ? `${row.sqFt} sf` : "—"}</Text>
            {/* paddingLeft separates the company dash from the right-aligned sq ft number */}
            <Text style={[s.td, { flex: 2.8, paddingLeft: 6 }]}>{trunc(row.companyName || row.contractName || "—", 32)}</Text>
            <Text style={[s.td, { flex: 1.5 }]}>{row.moveIn || "—"}</Text>
            <Text style={[s.td, { flex: 1.5 }]}>{row.moveOut || "—"}</Text>
            <Text style={[s.td, { flex: 0.8 }]}>{row.term || "—"}</Text>
          </View>
        ))}
        {/* Totals row — flex groups match column header groups to avoid number collision */}
        <View style={[s.tRow, s.tRowTotal]}>
          <Text style={[s.tdTotal, { flex: 5 }]}>
            Total — {sorted.length} Spaces · {sorted.reduce((sum, r) => sum + (r.sqFt || 0), 0).toLocaleString()} sf
          </Text>
          <Text style={[s.tdTotal, { flex: 3, textAlign: "right", color: C.green }]}>
            {occupancy.occupiedCount} Occupied
          </Text>
          <Text style={[s.tdTotal, { flex: 3, textAlign: "right", color: "#F87171" }]}>
            {occupancy.vacantCount} Vacant
          </Text>
        </View>
      </View>
      <PageFooter location={location} />
    </Page>
  );
}

// ── Page 4: Budget vs Actual ─────────────────────────────────────────────────
function BudgetPage({ data }: { data: ReportData }) {
  const { budget } = data;
  const location = data.dashboard.location;
  let isExpenseSection = false;
  const visibleItems = budget.lineItems
    .filter((item) => {
      if (!item.label) return false;
      if (item.isSectionHeader || item.isTotal) return true;
      const hasActual = item.actual != null && item.actual !== 0;
      const hasBudget = item.budget != null && item.budget !== 0;
      return hasActual || hasBudget;
    })
    .map((item): BudgetLineItem & { isExpense: boolean } => {
      if (item.isSectionHeader) {
        const ll = item.label.toLowerCase();
        if (ll.includes("cost") || ll.includes("cogs") || ll.includes("expense")) isExpenseSection = true;
        else if (ll.includes("income") || ll.includes("workspace") || ll.includes("meeting") ||
                 ll.includes("package") || ll.includes("member") || ll.includes("event") || ll.includes("registr"))
          isExpenseSection = false;
      }
      return { ...item, isExpense: isExpenseSection };
    });

  return (
    <Page size="LETTER" style={s.page}>
      <PageHeader title="Forecast & Budget vs. Actual" date={data.dashboard.reportDate} location={location} />
      {budget.grossProfit > 0 && (
        <View style={[s.kpiRow, { marginBottom: 14 }]}>
          {[
            { label: "Total Revenue",  value: fmt$(budget.totalRevenue) },
            { label: "Cost of Sales",  value: fmt$(budget.totalCogs) },
            { label: "Gross Profit",   value: fmt$(budget.grossProfit) },
            { label: "Gross Margin",   value: budget.totalRevenue > 0 ? `${((budget.grossProfit / budget.totalRevenue) * 100).toFixed(1)}%` : "—" },
          ].map((k) => (
            <View key={k.label} style={s.kpiCard}>
              <Text style={s.kpiLabel}>{k.label}</Text>
              <Text style={[s.kpiValue, { fontSize: 14 }]}>{k.value}</Text>
            </View>
          ))}
        </View>
      )}
      <Text style={{ fontSize: 6.5, color: C.gray600, marginBottom: 8 }}>
        All figures in USD. Variance: Green = favorable (revenue above plan or costs below plan). Red = unfavorable.
      </Text>
      <View style={s.table}>
        <View style={s.tHead}>
          <Text style={[s.th, { flex: 4 }]}>Line Item</Text>
          <Text style={[s.th, { flex: 2, textAlign: "right" }]}>Actual ($)</Text>
          <Text style={[s.th, { flex: 2, textAlign: "right" }]}>Budget ($)</Text>
          <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Var ($)</Text>
          <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Var (%)</Text>
        </View>
        {visibleItems.map((item, i) => {
          if (item.isSectionHeader) return (
            <View key={i} style={[s.tRow, { backgroundColor: C.gray100, borderTopWidth: 1, borderTopColor: C.gray200 }]}>
              <Text style={[s.td, { flex: 11 }]}>{item.label}</Text>
            </View>
          );
          const vp = item.variancePct;
          const pos = (vp ?? 0) >= 0;
          const varSt = vp == null ? s.varNeutral : item.isExpense
            ? (pos ? s.varRed : s.varGreen)
            : (pos ? s.varGreen : s.varRed);
          if (item.isTotal) return (
            <View key={i} style={[s.tRow, s.tRowTotal]}>
              <Text style={[s.tdTotal, { flex: 4 }]}>{item.label}</Text>
              <Text style={[s.tdTotal, { flex: 2, textAlign: "right" }]}>{fmt$(item.actual)}</Text>
              <Text style={[s.tdTotal, { flex: 2, textAlign: "right" }]}>{fmt$(item.budget)}</Text>
              <Text style={[s.tdTotal, { flex: 1.5, textAlign: "right" }]}>{item.varianceDollar != null ? fmt$(item.varianceDollar) : "—"}</Text>
              <Text style={[s.tdTotal, { flex: 1.5, textAlign: "right" }]}>{vp != null ? fmtPct(vp) : "—"}</Text>
            </View>
          );
          return (
            <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
              <Text style={[s.td, { flex: 4 }]}>{item.label}</Text>
              <Text style={[s.td, { flex: 2, textAlign: "right" }]}>{fmt$(item.actual)}</Text>
              <Text style={[s.td, { flex: 2, textAlign: "right" }]}>{fmt$(item.budget)}</Text>
              <Text style={[s.td, { flex: 1.5, textAlign: "right" }]}>{item.varianceDollar != null ? fmt$(item.varianceDollar) : "—"}</Text>
              <Text style={[varSt, { flex: 1.5, textAlign: "right" }]}>{vp != null ? fmtPct(vp) : "—"}</Text>
            </View>
          );
        })}
      </View>
      <PageFooter location={location} />
    </Page>
  );
}

// ── Page 5: Meeting Room Activity ────────────────────────────────────────────
function MeetingPage({ data }: { data: ReportData }) {
  const { meeting } = data;
  const location = data.dashboard.location;
  const monthName = data.dashboard.reportDate.split(" ")[0]; // "August"
  const varPos = (meeting.variancePct ?? 0) >= 0;

  // Issue 5: always use every month in the trend — current month must always appear,
  // even if revenue is $0. Parser already excludes future stub months via WX-2.
  const trendMonths = meeting.monthlyTrend;

  return (
    <Page size="LETTER" style={s.page}>
      <PageHeader title="Meeting Room Activity" date={data.dashboard.reportDate} location={location} />
      <View style={s.kpiRow}>
        {[
          { label: `${monthName} Revenue`,  value: fmt$(meeting.mtdRevenue),  sub: "Actual collected — complete month" },
          { label: `${monthName} Budget`,   value: fmt$(meeting.mtdBudget),   sub: "Annual plan target" },
          { label: "Variance ($)",   value: (meeting.varianceDollar >= 0 ? "+" : "") + fmt$(meeting.varianceDollar), sub: "vs. budget", color: varPos ? C.green : "#F87171" },
          { label: "Variance (%)",   value: meeting.variancePct != null ? fmtPct(meeting.variancePct) : "N/A", sub: "budget attainment", color: meeting.variancePct == null ? C.gold : varPos ? C.green : "#F87171" },
        ].map((k) => (
          <View key={k.label} style={s.kpiCard}>
            <Text style={s.kpiLabel}>{k.label}</Text>
            <Text style={[s.kpiValue, k.color ? { color: k.color } : {}]}>{k.value}</Text>
            <Text style={s.kpiSub}>{k.sub}</Text>
          </View>
        ))}
      </View>

      {meeting.recurringTransactions.length > 0 && (() => {
        // Detect dominant unit across all transactions — all-hours vs. all-dollars (rarely mixed)
        const allHrs = meeting.recurringTransactions.every((t) => t.unit === "hrs");
        const fmtTx  = (t: MeetingRoomTransaction) =>
          t.unit === "hrs" ? `${t.amount.toFixed(1)} hrs` : fmt$(t.amount);
        const amtHeader  = allHrs ? "Hours (hrs)" : "Amount ($)";
        const totalValue = meeting.recurringTransactions.reduce((acc, t) => acc + t.amount, 0);
        const totalFmt   = allHrs ? `${totalValue.toFixed(1)} hrs` : fmt$(totalValue);
        return (
          <>
            <SectionTitle>Active Recurring Agreements — Current Period</SectionTitle>
            <View style={s.table}>
              <View style={s.tHead}>
                <Text style={[s.th, { flex: 3 }]}>Company</Text>
                <Text style={[s.th, { flex: 4 }]}>Package / Description</Text>
                <Text style={[s.th, { flex: 2 }]}>Charge Code</Text>
                <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>{amtHeader}</Text>
              </View>
              {meeting.recurringTransactions.map((tx: MeetingRoomTransaction, i) => (
                <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
                  <Text style={[s.tdBold, { flex: 3 }]}>{tx.company}</Text>
                  <Text style={[s.td, { flex: 4 }]}>{tx.description}</Text>
                  <Text style={[s.td, { flex: 2 }]}>{tx.chargeCode}</Text>
                  <Text style={[s.tdBold, { flex: 1.5, textAlign: "right" }]}>{fmtTx(tx)}</Text>
                </View>
              ))}
              <View style={[s.tRow, s.tRowTotal]}>
                <Text style={[s.tdTotal, { flex: 9 }]}>Total Recurring</Text>
                <Text style={[s.tdTotal, { flex: 1.5, textAlign: "right" }]}>{totalFmt}</Text>
              </View>
            </View>
          </>
        );
      })()}

      {meeting.roomTypes.length > 0 && (
        <>
          <SectionTitle>Room Type Breakdown — Budget vs. Actual</SectionTitle>
          <View style={s.table}>
            <View style={s.tHead}>
              <Text style={[s.th, { flex: 4 }]}>Room Type</Text>
              <Text style={[s.th, { flex: 2, textAlign: "right" }]}>Actual</Text>
              <Text style={[s.th, { flex: 2, textAlign: "right" }]}>Budget</Text>
              <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Var $</Text>
              <Text style={[s.th, { flex: 1.5, textAlign: "right" }]}>Var %</Text>
            </View>
            {meeting.roomTypes.filter((room) => room.isTotal || room.actual > 0 || room.budget > 0).map((room: MeetingRoomTypeRow, i) => {
              const vp = room.variancePct;
              const pos = (vp ?? 0) >= 0;
              if (room.isTotal) return (
                <View key={i} style={[s.tRow, s.tRowTotal]}>
                  <Text style={[s.tdTotal, { flex: 4 }]}>{room.label}</Text>
                  <Text style={[s.tdTotal, { flex: 2, textAlign: "right" }]}>{fmt$(room.actual)}</Text>
                  <Text style={[s.tdTotal, { flex: 2, textAlign: "right" }]}>{fmt$(room.budget)}</Text>
                  <Text style={[s.tdTotal, { flex: 1.5, textAlign: "right" }]}>{fmt$(room.varianceDollar)}</Text>
                  <Text style={[s.tdTotal, { flex: 1.5, textAlign: "right" }]}>{vp != null ? fmtPct(vp) : "—"}</Text>
                </View>
              );
              return (
                <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
                  <Text style={[s.td, { flex: 4 }]}>{room.label}</Text>
                  <Text style={[s.td, { flex: 2, textAlign: "right" }]}>{room.actual > 0 ? fmt$(room.actual) : "—"}</Text>
                  <Text style={[s.td, { flex: 2, textAlign: "right" }]}>{room.budget > 0 ? fmt$(room.budget) : "—"}</Text>
                  <Text style={[s.td, { flex: 1.5, textAlign: "right" }]}>{room.actual > 0 || room.budget > 0 ? fmt$(room.varianceDollar) : "—"}</Text>
                  <Text style={[vp != null ? (pos ? s.varGreen : s.varRed) : s.varNeutral, { flex: 1.5, textAlign: "right" }]}>
                    {vp != null && (room.actual > 0 || room.budget > 0) ? fmtPct(vp) : "—"}
                  </Text>
                </View>
              );
            })}
          </View>
        </>
      )}

      {/* SectionTitle is INSIDE wrap={false} so it never orphans on a prior page */}
      {trendMonths.length > 0 && (
        <>
          <View wrap={false}>
          <SectionTitle>Monthly Meeting Room Revenue Trend</SectionTitle>
          <View style={{ alignItems: "center", marginBottom: 4 }}>
            <BarChart
              values={trendMonths.map((m) => m.revenue)}
              labels={trendMonths.map((m) => m.month.slice(0, 3))}
              color={C.barBlue}
              highlightLast
              yFmt={fmtK}
            />
          </View>
          <Text style={{ fontSize: 6.5, color: C.gray400, textAlign: "center", marginBottom: 8 }}>
            Gold bar = {monthName} (current period — {trendMonths[trendMonths.length - 1]?.revenue === 0 ? "no revenue recorded yet" : "complete month"}); blue bars = prior months. Includes both recurring (contracted packages) and non-recurring (ad-hoc) meeting room revenue. Source: P2 Monthly Income — "Meeting Rooms" total line.
          </Text>
          {/* YTD summary strip + month-by-month breakdown — fills the page */}
          {(() => {
            const dataMonths = trendMonths.filter((m) => m.revenue > 0);
            const ytdMtg  = trendMonths.reduce((s, m) => s + m.revenue, 0);
            const avgMtg  = dataMonths.length > 0 ? ytdMtg / dataMonths.length : 0;
            const peakMtg = Math.max(...trendMonths.map((m) => m.revenue), 0);
            return (
              <>
                <View style={{ flexDirection: "row", gap: 10, marginTop: 4, marginBottom: 10 }}>
                  <View style={{ flex: 1, backgroundColor: C.gray50, borderWidth: 1, borderColor: C.gray200, borderRadius: 5, padding: 8, alignItems: "center" }}>
                    <Text style={{ fontSize: 6, color: C.gray600, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 2 }}>YTD Meeting Revenue</Text>
                    <Text style={{ fontSize: 11, fontFamily: "Helvetica-Bold", color: C.navy }}>{fmt$(ytdMtg)}</Text>
                  </View>
                  <View style={{ flex: 1, backgroundColor: C.gray50, borderWidth: 1, borderColor: C.gray200, borderRadius: 5, padding: 8, alignItems: "center" }}>
                    <Text style={{ fontSize: 6, color: C.gray600, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 2 }}>Monthly Average</Text>
                    <Text style={{ fontSize: 11, fontFamily: "Helvetica-Bold", color: C.navy }}>{fmt$(Math.round(avgMtg))}</Text>
                    <Text style={{ fontSize: 5.5, color: C.gray400, marginTop: 1 }}>excl. months w/ no data</Text>
                  </View>
                  <View style={{ flex: 1, backgroundColor: C.gray50, borderWidth: 1, borderColor: C.gray200, borderRadius: 5, padding: 8, alignItems: "center" }}>
                    <Text style={{ fontSize: 6, color: C.gray600, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 2 }}>Peak Month</Text>
                    <Text style={{ fontSize: 11, fontFamily: "Helvetica-Bold", color: C.navy }}>{fmt$(peakMtg)}</Text>
                    <Text style={{ fontSize: 5.5, color: C.gray400, marginTop: 1 }}>
                      {trendMonths.find((m) => m.revenue === peakMtg)?.month ?? "—"}
                    </Text>
                  </View>
                </View>

                {/* Month-by-month breakdown table */}
                <View style={s.table}>
                  <View style={s.tHead}>
                    <Text style={[s.th, { flex: 2 }]}>Month</Text>
                    <Text style={[s.th, { flex: 2, textAlign: "right" }]}>Revenue ($)</Text>
                    <Text style={[s.th, { flex: 2, textAlign: "right" }]}>vs. Avg ($)</Text>
                    <Text style={[s.th, { flex: 2, textAlign: "right" }]}>vs. Avg (%)</Text>
                    <Text style={[s.th, { flex: 2, textAlign: "right" }]}>% of YTD</Text>
                  </View>
                  {trendMonths.map((m, i) => {
                    const isLast   = i === trendMonths.length - 1;
                    const noData   = m.revenue === 0;
                    const vsAvg    = m.revenue - avgMtg;
                    const vsAvgPct = avgMtg > 0 ? (vsAvg / avgMtg) * 100 : null;
                    const pctYtd   = ytdMtg > 0 ? (m.revenue / ytdMtg) * 100 : null;
                    const varColor = noData ? s.varNeutral : vsAvg >= 0 ? s.varGreen : s.varRed;
                    if (isLast && noData) return (
                      <View key={i} style={[s.tRow, s.tRowAlt]}>
                        <Text style={[s.tdBold, { flex: 2 }]}>{m.month.slice(0, 3)} (current)</Text>
                        <Text style={[s.varNeutral, { flex: 2, textAlign: "right" }]}>No data yet</Text>
                        <Text style={[s.varNeutral, { flex: 2, textAlign: "right" }]}>—</Text>
                        <Text style={[s.varNeutral, { flex: 2, textAlign: "right" }]}>—</Text>
                        <Text style={[s.varNeutral, { flex: 2, textAlign: "right" }]}>—</Text>
                      </View>
                    );
                    return (
                      <View key={i} style={[s.tRow, i % 2 === 1 ? s.tRowAlt : {}]}>
                        <Text style={[s.tdBold, { flex: 2 }]}>{m.month.slice(0, 3)}{isLast ? " ★" : ""}</Text>
                        <Text style={[s.td, { flex: 2, textAlign: "right" }]}>{fmt$(m.revenue)}</Text>
                        <Text style={[varColor, { flex: 2, textAlign: "right" }]}>
                          {vsAvg >= 0 ? "+" : ""}{fmt$(Math.round(vsAvg))}
                        </Text>
                        <Text style={[varColor, { flex: 2, textAlign: "right" }]}>
                          {vsAvgPct != null ? fmtPct(vsAvgPct) : "—"}
                        </Text>
                        <Text style={[s.td, { flex: 2, textAlign: "right" }]}>
                          {pctYtd != null ? `${pctYtd.toFixed(1)}%` : "—"}
                        </Text>
                      </View>
                    );
                  })}
                  <View style={[s.tRow, s.tRowTotal]}>
                    <Text style={[s.tdTotal, { flex: 2 }]}>YTD Total</Text>
                    <Text style={[s.tdTotal, { flex: 2, textAlign: "right" }]}>{fmt$(ytdMtg)}</Text>
                    <Text style={[s.tdTotal, { flex: 2, textAlign: "right" }]}>—</Text>
                    <Text style={[s.tdTotal, { flex: 2, textAlign: "right" }]}>—</Text>
                    <Text style={[s.tdTotal, { flex: 2, textAlign: "right" }]}>100%</Text>
                  </View>
                </View>
                <Text style={{ fontSize: 6, color: C.gray400, marginTop: 4 }}>
                  ★ = current period ({monthName}). "vs. Avg" compares each month to the YTD monthly average excluding months with no data.
                </Text>
              </>
            );
          })()}
          </View>{/* closes wrap={false} */}
        </>
      )}

      <PageFooter location={location} />
    </Page>
  );
}

// ── Page 6: Visual Appendix ──────────────────────────────────────────────────
function AppendixPage({ data }: { data: ReportData }) {
  const { dashboard, revenue, meeting } = data;
  const location = dashboard.location;
  const snaps = dashboard.weeklySnapshots.slice(-12);

  const occValues    = snaps.map((s) => s.occupancyPct);
  const sqftValues   = snaps.map((s) => s.sqFtPct);
  const snapLabels   = snaps.map((s) => s.date.slice(0, 6));

  const revLine = revenue.lineItems.find((li) => li.label.toLowerCase().includes("total with relief"));
  const allRevValues = revLine ? revenue.months.map((_, i) => (revLine.values[i] ?? 0) as number) : [];
  const allRevLabels = revenue.months.map((m) => m.slice(0, 3));
  const revMax = Math.max(...allRevValues, 1);
  const lastSignificant = allRevValues.reduce((last, v, i) => v >= revMax * 0.01 ? i : last, -1);
  const revValues = lastSignificant >= 0 ? allRevValues.slice(0, lastSignificant + 1) : allRevValues;
  const revLabels = lastSignificant >= 0 ? allRevLabels.slice(0, lastSignificant + 1) : allRevLabels;

  // Issue 5: use full trend — parser already excludes future stubs; current month always shown.
  const mtgValues = meeting.monthlyTrend.map((m) => m.revenue);
  const mtgLabels = meeting.monthlyTrend.map((m) => m.month.slice(0, 3));

  return (
    <Page size="LETTER" style={s.page}>
      <PageHeader title="Visual Appendix" date={dashboard.reportDate} location={location} />

      {snaps.length > 1 && (
        <>
          <View wrap={false}>
          <SectionTitle>Occupancy % Trend (Last {snaps.length} Snapshots)</SectionTitle>
          <View style={{ alignItems: "center", marginBottom: 4 }}>
            <LineChart
              series={[
                { values: occValues,  color: C.navy,    label: "By Spaces (%)" },
                { values: sqftValues, color: C.barGold, label: "By Sq Ft (%)" },
              ]}
              labels={snapLabels}
              yFmt={(v) => `${Math.round(v)}%`}
            />
          </View>
          <Text style={{ fontSize: 6.5, color: C.gray400, textAlign: "center", marginBottom: 6 }}>
            Source: P1 Dashboard. Navy line = space-count occupancy %; Gold line = sq ft occupancy %. All months shown are complete periods.
          </Text>
          {/* Mini data table for occupancy trend */}
          <View style={[s.table, { marginBottom: 10 }]}>
            <View style={s.tHead}>
              <Text style={[s.th, { flex: 2 }]}>Month</Text>
              {snaps.map((sn, i) => (
                <Text key={i} style={[s.th, { flex: 1, textAlign: "right" }]}>{sn.date.slice(0, 6)}</Text>
              ))}
            </View>
            <View style={s.tRow}>
              <Text style={[s.tdBold, { flex: 2 }]}>Occ % (Spaces)</Text>
              {snaps.map((sn, i) => (
                <Text key={i} style={[s.td, { flex: 1, textAlign: "right" }]}>{sn.occupancyPct}%</Text>
              ))}
            </View>
            <View style={[s.tRow, s.tRowAlt]}>
              <Text style={[s.tdBold, { flex: 2 }]}>Occ % (Sq Ft)</Text>
              {snaps.map((sn, i) => (
                <Text key={i} style={[s.td, { flex: 1, textAlign: "right" }]}>{sn.sqFtPct}%</Text>
              ))}
            </View>
          </View>
          </View>{/* closes wrap={false} for occupancy chart block */}
          {/* Occupancy summary tiles — peak month and average */}
          {(() => {
            const avgOccSpaces = snaps.length > 0 ? Math.round(snaps.reduce((sum, sn) => sum + sn.occupancyPct, 0) / snaps.length) : 0;
            const avgOccSqFt  = snaps.length > 0 ? Math.round(snaps.reduce((sum, sn) => sum + sn.sqFtPct, 0) / snaps.length) : 0;
            const peakSnap    = snaps.reduce((best, sn) => sn.occupancyPct > best.occupancyPct ? sn : best, snaps[0]);
            return (
              <View style={{ flexDirection: "row", gap: 10, marginTop: 6, marginBottom: 10 }}>
                <View style={{ flex: 1, backgroundColor: C.gray50, borderWidth: 1, borderColor: C.gray200, borderRadius: 5, padding: 8, alignItems: "center" }}>
                  <Text style={{ fontSize: 6, color: C.gray600, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 2 }}>Avg Monthly Occupancy (Spaces)</Text>
                  <Text style={{ fontSize: 11, fontFamily: "Helvetica-Bold", color: C.navy }}>{avgOccSpaces}%</Text>
                  <Text style={{ fontSize: 5.5, color: C.gray400, marginTop: 1 }}>{snaps.length}-month average</Text>
                </View>
                <View style={{ flex: 1, backgroundColor: C.gray50, borderWidth: 1, borderColor: C.gray200, borderRadius: 5, padding: 8, alignItems: "center" }}>
                  <Text style={{ fontSize: 6, color: C.gray600, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 2 }}>Avg Monthly Occupancy (Sq Ft)</Text>
                  <Text style={{ fontSize: 11, fontFamily: "Helvetica-Bold", color: C.navy }}>{avgOccSqFt}%</Text>
                  <Text style={{ fontSize: 5.5, color: C.gray400, marginTop: 1 }}>{snaps.length}-month average</Text>
                </View>
                <View style={{ flex: 1, backgroundColor: C.gray50, borderWidth: 1, borderColor: C.gray200, borderRadius: 5, padding: 8, alignItems: "center" }}>
                  <Text style={{ fontSize: 6, color: C.gray600, textTransform: "uppercase", letterSpacing: 0.6, marginBottom: 2 }}>Peak Month Occupancy</Text>
                  <Text style={{ fontSize: 11, fontFamily: "Helvetica-Bold", color: C.navy }}>{peakSnap?.occupancyPct ?? 0}%</Text>
                  <Text style={{ fontSize: 5.5, color: C.gray400, marginTop: 1 }}>{peakSnap?.date ?? "—"} — by spaces</Text>
                </View>
              </View>
            );
          })()}
        </>
      )}

      {revValues.length > 0 && (
        <>
          <View wrap={false}>
          <SectionTitle>Total Monthly Revenue — Recurring + Non-Recurring, Net After Discounts</SectionTitle>
          <View style={{ alignItems: "center", marginBottom: 4 }}>
            <BarChart values={revValues} labels={revLabels} color={C.barBlue} highlightLast yFmt={fmtK} />
          </View>
          <Text style={{ fontSize: 6.5, color: C.gray400, textAlign: "center", marginBottom: 6 }}>
            Gold bar = {dashboard.reportDate.split(" ")[0]} (current period); blue bars = prior months. Includes all recurring revenue (memberships, contracts) and non-recurring revenue (meeting rooms, day passes). Source: P2 — "Total with Relief" (net after member discounts across all categories).
          </Text>
          {/* Mini data table for revenue trend */}
          <View style={[s.table, { marginBottom: 10 }]}>
            <View style={s.tHead}>
              <Text style={[s.th, { flex: 2 }]}>Month</Text>
              {revLabels.map((lbl, i) => (
                <Text key={i} style={[s.th, { flex: 1, textAlign: "right" }]}>{lbl}</Text>
              ))}
            </View>
            <View style={[s.tRow, { backgroundColor: C.navy, paddingVertical: 7 }]}>
              <Text style={[s.tdTotal, { flex: 2 }]}>Net Revenue ($)</Text>
              {revValues.map((v, i) => (
                <Text key={i} style={[s.tdTotal, { flex: 1, textAlign: "right" }]}>{fmtK(v)}</Text>
              ))}
            </View>
          </View>
          </View>{/* closes wrap={false} for revenue chart block */}
        </>
      )}

      {mtgValues.length > 0 && (
        <>
          <View wrap={false}>
          <SectionTitle>Total Meeting Room Revenue by Month (Recurring + Ad-Hoc)</SectionTitle>
          <View style={{ alignItems: "center", marginBottom: 4 }}>
            <BarChart values={mtgValues} labels={mtgLabels} color={C.barBlue} highlightLast yFmt={fmtK} />
          </View>
          <Text style={{ fontSize: 6.5, color: C.gray400, textAlign: "center", marginBottom: 8 }}>
            Gold bar = current month ({mtgLabels[mtgLabels.length - 1] ?? "—"}); blue bars = prior months. Total includes both recurring (contracted packages) and non-recurring (ad-hoc) meeting room revenue. Source: P2 Monthly Income — "Meeting Rooms" total line.
          </Text>
          </View>{/* closes wrap={false} for meeting chart block */}
        </>
      )}

      <PageFooter location={location} />
    </Page>
  );
}

// ── Page 7: Supplemental Notes + Data Validation ─────────────────────────────
function SupplementalPage({ data }: { data: ReportData }) {
  const { validation, budget } = data;
  const location = data.dashboard.location;
  const marginPct = budget.totalRevenue > 0
    ? `${((budget.grossProfit / budget.totalRevenue) * 100).toFixed(1)}%`
    : "N/A";

  return (
    <Page size="LETTER" style={s.page}>
      <PageHeader title="Supplemental Notes & Methodology" date={data.dashboard.reportDate} location={location} />

      <SectionTitle>Data Accuracy Audit — Automated Cross-Tab Reconciliation</SectionTitle>
      <Text style={[s.noteText, { textAlign: "center", marginBottom: 8 }]}>
        {validation.passCount} of {validation.checks.length} cross-tab checks passed
        {validation.failCount > 0
          ? ` — ${validation.failCount} discrepanc${validation.failCount === 1 ? "y requires" : "ies require"} review before distribution`
          : " — All figures independently verified"}
      </Text>
      <View style={s.table}>
        <View style={s.tHead}>
          <Text style={[s.th, { flex: 1, textAlign: "center" }]}>✓</Text>
          <Text style={[s.th, { flex: 4 }]}>Check</Text>
          <Text style={[s.th, { flex: 3 }]}>Expected</Text>
          <Text style={[s.th, { flex: 3 }]}>Actual (recalculated)</Text>
        </View>
        {validation.checks.map((chk: ValidationCheck, i) => (
          <View key={i} style={[s.tRow, chk.passed ? s.passRow : s.failRow]}>
            <Text style={[chk.passed ? s.passText : s.failText, { flex: 1, textAlign: "center" }]}>
              {chk.passed ? "PASS" : "FAIL"}
            </Text>
            <Text style={[s.td, { flex: 4 }]}>{chk.label}</Text>
            <Text style={[s.td, { flex: 3 }]}>{chk.expected}</Text>
            <Text style={[s.td, { flex: 3 }]}>{chk.actual}</Text>
          </View>
        ))}
      </View>

      <SectionTitle>Methodology Notes</SectionTitle>

      <Text style={s.noteHead}>REPORT OVERVIEW</Text>
      <Text style={s.noteBullet}>• This report is generated monthly from the 25N Coworking Waco operational workbook. It is prepared for CFO review and covers occupancy, revenue, budget performance, and meeting room activity.</Text>
      <Text style={s.noteBullet}>• All data reflects the state of the workbook as of the snapshot date shown in each page header. Figures are not restated after publication.</Text>
      <Text style={s.noteBullet}>• The report is organized into seven sections: Monthly Revenue, Space Inventory, Meeting Room Activity, Budget vs. Actual, Executive Dashboard &amp; Key Metrics, Visual Appendix, and this Supplemental Notes page.</Text>

      <Text style={s.noteHead}>DATA SOURCES</Text>
      <Text style={s.noteBullet}>• P1 Dashboard: Executive summary tab. Contains KPI snapshot, space type breakdown, and month-over-month historical trend table. Used for occupancy KPIs and the dashboard page historical chart.</Text>
      <Text style={s.noteBullet}>• P2 Monthly Income: Revenue ledger by category across all months YTD. Rows cover Flex memberships, Mail memberships, Daily Space, Meeting Rooms, and other income. "Total with Relief" is the net revenue figure after member discounts and promotional credits.</Text>
      <Text style={s.noteBullet}>• P3 Occupancy: Live space roster listing every leasable unit. Columns include space ID, type (Dedicated Desk / Private Office), status, sq ft, assigned tenant, contract name, move-in date, move-out date, and term type (MTM = month-to-month; Fixed = fixed-term lease).</Text>
      <Text style={s.noteBullet}>• P4 Forecast & Budget: Point-in-time comparison of actual results vs. the annual operating budget. Organized by revenue category and cost category. Variance columns are recalculated from the Actual and Budget values — the spreadsheet's own variance figures are not used.</Text>
      <Text style={s.noteBullet}>• Recurring & Promo Mtg Rm Hrs: Transaction-level detail sheet for contracted meeting room agreements. Each row represents one active agreement for the period, including company, package description, charge code, and hours (or dollar amount where applicable).</Text>

      <Text style={s.noteHead}>SPACE CLASSIFICATIONS</Text>
      <Text style={s.noteBullet}>• Dedicated Desk (Ded. Desk): An assigned, reserved workstation in an open-floor area. Member has exclusive use of a specific seat and associated storage. Typically smaller sq ft than a private office.</Text>
      <Text style={s.noteBullet}>• Private Office (Priv. Office): An enclosed, lockable office suite leased exclusively to one tenant or company. Ranges from single-person offices (~46 sq ft) to large suites (400+ sq ft).</Text>
      <Text style={s.noteBullet}>• Occupancy is tracked separately for each space type because they carry different sq ft footprints and price points. A 70% space-count occupancy rate may mask a higher sq ft occupancy if large private offices are occupied and small dedicated desks are vacant — or vice versa.</Text>

      <Text style={s.noteHead}>MEMBERSHIP & REVENUE CATEGORIES</Text>
      <Text style={s.noteBullet}>• Flex Memberships: Month-to-month coworking access passes (hot desk, drop-in, or flex-desk tiers). Revenue is billed at the start of each period. Member count fluctuates more than contract-based spaces.</Text>
      <Text style={s.noteBullet}>• Mail Memberships: Virtual office / business address service. Lowest-revenue category per member but highly scalable with minimal space consumption. Members receive mail handling without occupying a physical desk.</Text>
      <Text style={s.noteBullet}>• Daily Space: Walk-in day passes and short-term bookings not tied to a recurring membership. Represents the most volatile revenue line — weather, events, and local business cycles drive this category.</Text>
      <Text style={s.noteBullet}>• Meeting Rooms: Revenue from reserving conference rooms, boardrooms, huddle rooms, and event spaces. Split into Recurring (contracted packages) and General (ad-hoc bookings). See the Meeting Room Activity page for detail.</Text>

      <Text style={s.noteHead}>OCCUPANCY CALCULATIONS</Text>
      <Text style={s.noteBullet}>• Occupancy % by Spaces: Occupied spaces ÷ Total leasable spaces. A space is "Occupied" when it has an active tenant assignment in P3; "Vacant" when it has no current assignment regardless of lease expiration date.</Text>
      <Text style={s.noteBullet}>• Occupancy % by Sq Ft: Occupied sq ft ÷ Total leasable sq ft. This weighted metric accounts for the fact that a single large private office (e.g., 400 sq ft) has the same unit count as a 33 sq ft dedicated desk, but contributes ~12× more revenue potential per space.</Text>
      <Text style={s.noteBullet}>• The two occupancy metrics will diverge when large-format spaces are disproportionately occupied or vacant. Management should monitor both to identify whether demand is concentrated in one space type.</Text>

      <Text style={s.noteHead}>REVENUE RECOGNITION & "TOTAL WITH RELIEF"</Text>
      <Text style={s.noteBullet}>• Gross Revenue: Sum of all invoiced amounts before any member discounts or promotional credits. Shown as individual line items in P2.</Text>
      <Text style={s.noteBullet}>• "Total with Relief": Net revenue after deducting member relief credits (promotional rates, complimentary memberships, or goodwill adjustments). This is the operative top-line revenue figure used throughout this report, and is the basis for gross profit calculation.</Text>
      <Text style={s.noteBullet}>• Revenue is recognized on a cash/billing basis consistent with prior periods. Month-over-month changes reflect actual collected amounts, not accruals.</Text>

      <Text style={s.noteHead}>COGS & GROSS PROFIT</Text>
      <Text style={s.noteBullet}>• Cost of Sales (COGS): Direct costs specifically attributable to revenue-generating activities. At Waco, this is primarily the printer/copier lease (~$347/month). It does not include payroll, rent, utilities, or other operating overhead — those appear below the gross profit line.</Text>
      <Text style={s.noteBullet}>• Gross Profit = "Total with Relief" − Total COGS. This metric isolates operational efficiency before fixed-overhead costs. Current gross margin: {marginPct}.</Text>
      <Text style={s.noteBullet}>• The high gross margin is characteristic of the flex-space model: once a location is leased and staffed, incremental revenue from new members carries very low marginal cost. Margin compression would signal an increase in direct variable costs or a decline in occupancy at the same fixed-cost base.</Text>

      <Text style={s.noteHead}>MEETING ROOM ACCOUNTING</Text>
      <Text style={s.noteBullet}>• Recurring Agreements: Pre-purchased hour bundles or day-office packages contracted by corporate clients. These are billed on a fixed monthly schedule, making them the most predictable component of meeting room revenue. Tracked individually in the "Recurring & Promo Mtg Rm Hrs" tab.</Text>
      <Text style={s.noteBullet}>• General / Ad-Hoc Bookings: Walk-in and on-demand room reservations. Revenue is event-driven and therefore more volatile. Management has set a $7,000/month budget target for this channel, reflecting an internal sales objective for growing non-contracted bookings.</Text>
      <Text style={s.noteBullet}>• Meeting hours columns in the Executive Dashboard (Pkg Hrs / Billed Hrs) reflect hours logged in the P1 snapshot table — both columns track hours, not dollar amounts. Revenue figures for meeting rooms are sourced from P2 and P4.</Text>
      <Text style={s.noteBullet}>• A large negative variance in General Meeting Space revenue vs. budget is an opportunity gap — it reflects unbooked capacity against a stretch target, not a pricing or billing error.</Text>

      <Text style={s.noteHead}>BUDGET VS. ACTUAL — VARIANCE CONVENTIONS</Text>
      <Text style={s.noteBullet}>• Variance ($): Actual minus Budget. Positive = above plan; Negative = below plan.</Text>
      <Text style={s.noteBullet}>• Variance (%): (Actual − Budget) ÷ |Budget|. Expressed as a signed percentage. Recalculated independently from the formula — spreadsheet variance fields are not used directly.</Text>
      <Text style={s.noteBullet}>• Color Coding: Green = favorable outcome (revenue above plan OR costs below plan). Red = unfavorable (revenue below plan OR costs above plan). The direction of "favorable" flips between revenue and expense line items automatically.</Text>
      <Text style={s.noteBullet}>• "N/A" or "—": Budget was $0 for that line item; percentage variance is mathematically undefined (division by zero). Does not indicate missing data.</Text>
      <Text style={s.noteBullet}>• Variance columns in the Budget vs. Actual page are recalculated by the report generator from the raw Actual and Budget figures to ensure consistency. Any rounding differences from the source spreadsheet are expected and immaterial.</Text>

      <PageFooter location={location} />
    </Page>
  );
}

// ── Main Document ────────────────────────────────────────────────────────────
export function ReportDocument({ data }: { data: ReportData }) {
  return (
    <Document
      title={`25N Coworking Occupancy Report — ${data.dashboard.reportDate}`}
      author="25N Coworking"
    >
      <RevenuePage data={data} />
      <OccupancyPage data={data} />
      <MeetingPage data={data} />
      <BudgetPage data={data} />
      <DashboardPage data={data} />
      <AppendixPage data={data} />
      <SupplementalPage data={data} />
    </Document>
  );
}
