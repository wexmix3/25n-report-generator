import React from "react";
import {
  Document, Page, Text, View, StyleSheet, renderToBuffer,
} from "@react-pdf/renderer";

const C = {
  navy:    "#0F2044",
  gold:    "#C9A84C",
  gray50:  "#F9FAFB",
  gray100: "#F3F4F6",
  gray200: "#E5E7EB",
  gray400: "#9CA3AF",
  gray600: "#4B5563",
  gray800: "#1F2937",
  white:   "#FFFFFF",
  green:   "#D1FAE5",
  greenTx: "#065F46",
};

const s = StyleSheet.create({
  page: {
    fontFamily: "Helvetica",
    fontSize: 9,
    paddingTop: 44,
    paddingBottom: 44,
    paddingHorizontal: 48,
    backgroundColor: C.white,
  },

  // Header band
  headerBand: {
    backgroundColor: C.navy,
    marginHorizontal: -48,
    marginTop: -44,
    paddingHorizontal: 48,
    paddingTop: 28,
    paddingBottom: 22,
    marginBottom: 28,
  },
  headerTitle: {
    fontFamily: "Helvetica-Bold",
    fontSize: 18,
    color: C.white,
    letterSpacing: 0.3,
  },
  headerSub: {
    fontSize: 9,
    color: C.gold,
    marginTop: 4,
    letterSpacing: 1,
  },

  // URL badge
  urlBadge: {
    backgroundColor: C.green,
    borderRadius: 4,
    paddingHorizontal: 10,
    paddingVertical: 6,
    marginBottom: 24,
    flexDirection: "row",
    alignItems: "center",
    gap: 6,
  },
  urlLabel: { fontSize: 8, color: C.greenTx, fontFamily: "Helvetica-Bold" },
  urlValue: { fontSize: 9, color: C.greenTx, fontFamily: "Helvetica-Bold" },

  // Section headers
  sectionHeader: {
    backgroundColor: C.navy,
    paddingHorizontal: 10,
    paddingVertical: 6,
    marginBottom: 8,
    borderRadius: 3,
  },
  sectionTitle: {
    fontFamily: "Helvetica-Bold",
    fontSize: 10,
    color: C.white,
    letterSpacing: 0.5,
  },
  sectionNote: {
    fontSize: 7.5,
    color: C.gold,
    marginTop: 1.5,
  },

  // Steps
  stepRow: {
    flexDirection: "row",
    marginBottom: 7,
    alignItems: "flex-start",
  },
  stepNum: {
    width: 18,
    height: 18,
    borderRadius: 9,
    backgroundColor: C.navy,
    alignItems: "center",
    justifyContent: "center",
    marginRight: 8,
    flexShrink: 0,
    marginTop: 0.5,
  },
  stepNumText: {
    fontFamily: "Helvetica-Bold",
    fontSize: 8,
    color: C.white,
  },
  stepText: { fontSize: 9, color: C.gray800, lineHeight: 1.5, flex: 1 },
  stepBold: { fontFamily: "Helvetica-Bold" },

  // Sub-steps (indented)
  subRow: {
    flexDirection: "row",
    marginBottom: 4,
    paddingLeft: 26,
    alignItems: "flex-start",
  },
  bullet: { fontSize: 9, color: C.gold, marginRight: 6, marginTop: 1 },
  subText: { fontSize: 9, color: C.gray600, lineHeight: 1.5, flex: 1 },

  // Divider
  divider: {
    borderBottomWidth: 1,
    borderBottomColor: C.gray200,
    marginVertical: 16,
  },

  // Table
  table: { marginBottom: 8 },
  tableHeader: {
    flexDirection: "row",
    backgroundColor: C.gray100,
    paddingHorizontal: 8,
    paddingVertical: 5,
    borderTopLeftRadius: 3,
    borderTopRightRadius: 3,
  },
  tableHeaderCell: {
    fontFamily: "Helvetica-Bold",
    fontSize: 8,
    color: C.gray600,
  },
  tableRow: {
    flexDirection: "row",
    paddingHorizontal: 8,
    paddingVertical: 5,
    borderBottomWidth: 1,
    borderBottomColor: C.gray100,
  },
  tableRowAlt: {
    backgroundColor: C.gray50,
  },
  tableCellIssue: { fontSize: 8.5, color: C.gray800, flex: 1.2 },
  tableCellFix:   { fontSize: 8.5, color: C.gray600, flex: 2 },

  // Footer
  footer: {
    position: "absolute",
    bottom: 22,
    left: 48,
    right: 48,
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
    borderTopWidth: 1,
    borderTopColor: C.gray200,
    paddingTop: 8,
  },
  footerLeft:  { fontSize: 7, color: C.gray400 },
  footerRight: { fontSize: 7, color: C.gold },

  mb4:  { marginBottom: 4 },
  mb8:  { marginBottom: 8 },
  mb16: { marginBottom: 16 },
  mt4:  { marginTop: 4 },
  mt8:  { marginTop: 8 },
});

function Step({ num, children }: { num: number; children: React.ReactNode }) {
  return (
    <View style={s.stepRow}>
      <View style={s.stepNum}>
        <Text style={s.stepNumText}>{num}</Text>
      </View>
      <Text style={s.stepText}>{children}</Text>
    </View>
  );
}

function Sub({ children }: { children: React.ReactNode }) {
  return (
    <View style={s.subRow}>
      <Text style={s.bullet}>›</Text>
      <Text style={s.subText}>{children}</Text>
    </View>
  );
}

function SectionHeader({ title, note }: { title: string; note?: string }) {
  return (
    <View style={[s.sectionHeader, s.mb8]}>
      <Text style={s.sectionTitle}>{title}</Text>
      {note ? <Text style={s.sectionNote}>{note}</Text> : null}
    </View>
  );
}

const TOOL_URL = "https://report-generator-phi-virid.vercel.app";

const troubleshootRows = [
  { issue: 'Macro says "file not found"', fix: "Move Data Bridge.xlsx to your Desktop before running the macro" },
  { issue: "Login says wrong password",   fix: "Contact Max" },
  { issue: "Numbers look off",            fix: "Re-upload the correct source files in Step 1, re-download Data Bridge, re-run macro" },
  { issue: "Macro doesn't appear in list",fix: "Re-import MonthlyUpdate.bas via Alt+F11 > File > Import File" },
];

function OnboardingDoc() {
  return (
    <Document title="25N Monthly Close Tool — Onboarding Guide">
      <Page size="LETTER" style={s.page}>

        {/* Header band */}
        <View style={s.headerBand}>
          <Text style={s.headerTitle}>25N Monthly Close Tool</Text>
          <Text style={s.headerSub}>ONBOARDING GUIDE</Text>
        </View>

        {/* URL badge */}
        <View style={s.urlBadge}>
          <Text style={s.urlLabel}>TOOL URL  </Text>
          <Text style={s.urlValue}>{TOOL_URL}</Text>
        </View>

        {/* One-time setup */}
        <SectionHeader title="ONE-TIME SETUP" note="First month only — do this once, never again" />
        <Step num={1}>Open your <Text style={s.stepBold}>MASTER.xlsm</Text> in Excel</Step>
        <Step num={2}>Press <Text style={s.stepBold}>Alt + F11</Text> to open the VBA editor</Step>
        <Step num={3}>In the menu bar: <Text style={s.stepBold}>File → Import File</Text></Step>
        <Step num={4}>Select <Text style={s.stepBold}>MonthlyUpdate.bas</Text> (provided by Max)</Step>
        <Step num={5}>Close the VBA editor, then save the workbook as <Text style={s.stepBold}>.xlsm</Text> (macro-enabled format)</Step>

        <View style={s.divider} />

        {/* Monthly workflow */}
        <SectionHeader title="EVERY MONTH" note="Run these two steps on close day" />

        {/* Step 1 */}
        <View style={s.mb8}>
          <View style={[s.stepRow, s.mb4]}>
            <View style={[s.stepNum, { backgroundColor: C.gold }]}>
              <Text style={s.stepNumText}>1</Text>
            </View>
            <Text style={[s.stepText, { fontFamily: "Helvetica-Bold", fontSize: 10 }]}>
              Generate Data Bridge
            </Text>
          </View>
          <Sub>Log in at the tool URL above</Sub>
          <Sub>Under <Text style={{ fontFamily: "Helvetica-Bold" }}>Step 1 – Generate Data Bridge</Text>, upload your source exports (Voyager 12-Month Statement, Kube occupancy report)</Sub>
          <Sub>Click <Text style={{ fontFamily: "Helvetica-Bold" }}>Generate</Text> → download <Text style={{ fontFamily: "Helvetica-Bold" }}>Data Bridge.xlsx</Text></Sub>
          <Sub>Open <Text style={{ fontFamily: "Helvetica-Bold" }}>MASTER.xlsm</Text> in Excel</Sub>
          <Sub>Press <Text style={{ fontFamily: "Helvetica-Bold" }}>Alt + F8</Text> → select <Text style={{ fontFamily: "Helvetica-Bold" }}>MonthlyUpdate</Text> → click <Text style={{ fontFamily: "Helvetica-Bold" }}>Run</Text></Sub>
          <Sub>When prompted, navigate to and select the <Text style={{ fontFamily: "Helvetica-Bold" }}>Data Bridge.xlsx</Text> you downloaded</Sub>
          <Sub>The macro runs automatically (~30 seconds) — dashboard is updated</Sub>
        </View>

        {/* Step 2 */}
        <View style={s.mb16}>
          <View style={[s.stepRow, s.mb4]}>
            <View style={[s.stepNum, { backgroundColor: C.gold }]}>
              <Text style={s.stepNumText}>2</Text>
            </View>
            <Text style={[s.stepText, { fontFamily: "Helvetica-Bold", fontSize: 10 }]}>
              Generate PDF Report  <Text style={{ fontFamily: "Helvetica", fontSize: 8, color: C.gray400 }}>(optional)</Text>
            </Text>
          </View>
          <Sub>Back in the web tool, under <Text style={{ fontFamily: "Helvetica-Bold" }}>Step 2 – Generate PDF Report</Text></Sub>
          <Sub>Upload the completed Excel dashboard</Sub>
          <Sub>Click <Text style={{ fontFamily: "Helvetica-Bold" }}>Generate</Text> → download the PDF for CFO review</Sub>
        </View>

        {/* Troubleshooting */}
        <SectionHeader title="TROUBLESHOOTING" />
        <View style={s.table}>
          <View style={s.tableHeader}>
            <Text style={[s.tableHeaderCell, { flex: 1.2 }]}>Issue</Text>
            <Text style={[s.tableHeaderCell, { flex: 2 }]}>Fix</Text>
          </View>
          {troubleshootRows.map((row, i) => (
            <View key={i} style={[s.tableRow, i % 2 === 1 ? s.tableRowAlt : {}]}>
              <Text style={s.tableCellIssue}>{row.issue}</Text>
              <Text style={s.tableCellFix}>{row.fix}</Text>
            </View>
          ))}
        </View>

        {/* Footer */}
        <View style={s.footer} fixed>
          <Text style={s.footerLeft}>25N Coworking — Monthly Close Tool — Confidential</Text>
          <Text style={s.footerRight}>25N COWORKING</Text>
        </View>

      </Page>
    </Document>
  );
}

export async function GET() {
  const buf = await renderToBuffer(<OnboardingDoc />);
  return new Response(new Uint8Array(buf), {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": 'attachment; filename="25N-Monthly-Close-Onboarding.pdf"',
    },
  });
}
