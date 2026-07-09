/**
 * Validation script: compare historical vs generated Excel dashboards.
 * Run: node validate.js
 */
const XLSX = require("xlsx-js-style");
const path = require("path");
const fs = require("fs");

const HIST_PATH = path.join(__dirname, "Waco Sales Dashboard 8.1.2024.xlsx");
const GEN_PATH  = path.join(__dirname, "Waco Sales Dashboard 8.1.2024 - Generated Dashboard.xlsx");

if (!fs.existsSync(GEN_PATH)) {
  console.error("Generated file not found. Run the generator first.");
  process.exit(1);
}

const hist = XLSX.readFile(HIST_PATH, { cellStyles: true, cellNF: true, cellFormula: true });
const gen  = XLSX.readFile(GEN_PATH,  { cellStyles: true, cellNF: true, cellFormula: true });

const DASH_SHEETS = {
  p1: ["P.1 Dashboard NEW", "P.1 Dashboard"],
  p2: ["P2. Monthly Income"],
  p4: ["P.4 Forecast & Budget to Actual"],
};

function findSheet(wb, candidates) {
  for (const n of candidates) if (wb.Sheets[n]) return { name: n, sheet: wb.Sheets[n] };
  return null;
}

function cellAddr(r, c) {
  return XLSX.utils.encode_cell({ r, c });
}

function rangeOf(ws) {
  const ref = ws["!ref"];
  if (!ref) return null;
  return XLSX.utils.decode_range(ref);
}

function fmtStr(cell) {
  if (!cell) return "";
  return cell.z || "";
}

function valStr(cell) {
  if (!cell) return "(empty)";
  if (cell.f) return `formula:${cell.f}`;
  if (cell.v !== undefined) return String(cell.v);
  return "(empty)";
}

function compareSheet(label, histSheet, genSheet) {
  console.log(`\n${"=".repeat(60)}`);
  console.log(`Sheet: ${label}`);
  console.log("=".repeat(60));

  // --- Merges ---
  const hm = (histSheet["!merges"] || []).map(m => XLSX.utils.encode_range(m)).sort();
  const gm = (genSheet["!merges"]  || []).map(m => XLSX.utils.encode_range(m)).sort();
  const onlyHist = hm.filter(x => !gm.includes(x));
  const onlyGen  = gm.filter(x => !hm.includes(x));
  if (onlyHist.length === 0 && onlyGen.length === 0) {
    console.log("Merges: MATCH ✓");
  } else {
    console.log(`Merges: DIFF — hist-only: ${onlyHist.length}, gen-only: ${onlyGen.length}`);
    if (onlyHist.length) console.log("  hist-only:", onlyHist.slice(0, 10).join(", "));
    if (onlyGen.length)  console.log("  gen-only:",  onlyGen.slice(0, 10).join(", "));
  }

  // --- Row Heights ---
  const hr = histSheet["!rows"] || [];
  const gr = genSheet["!rows"]  || [];
  const maxRows = Math.max(hr.length, gr.length);
  let rowDiffs = 0;
  const rowDiffDetails = [];
  for (let i = 0; i < maxRows; i++) {
    const h = hr[i] || {};
    const g = gr[i] || {};
    const hHpt = h.hpt ?? "(default)";
    const gHpt = g.hpt ?? "(default)";
    const hHid = h.hidden ?? false;
    const gHid = g.hidden ?? false;
    if (hHpt !== gHpt || hHid !== gHid) {
      rowDiffs++;
      rowDiffDetails.push(`  row ${i+1}: hist={hpt:${hHpt},hidden:${hHid}} gen={hpt:${gHpt},hidden:${gHid}}`);
    }
  }
  if (rowDiffs === 0) {
    console.log("Row heights: MATCH ✓");
  } else {
    console.log(`Row heights: DIFF — ${rowDiffs} rows differ`);
    rowDiffDetails.slice(0, 15).forEach(d => console.log(d));
    if (rowDiffDetails.length > 15) console.log(`  ... and ${rowDiffDetails.length - 15} more`);
  }

  // --- Col Widths ---
  const hc = histSheet["!cols"] || [];
  const gc = genSheet["!cols"]  || [];
  console.log('DBG gc[0]:', JSON.stringify(gc[0])); const maxCols = Math.max(hc.length, gc.length);
  let colDiffs = 0;
  for (let i = 0; i < maxCols; i++) {
    const h = hc[i] || {};
    const g = gc[i] || {};
    // Compare wpx if available, else wch (rounded to 2dp)
    const hW = h.wpx !== undefined ? `wpx:${h.wpx}` : (h.wch !== undefined ? `wch:${h.wch.toFixed(2)}` : "(default)");
    const gW = g.wpx !== undefined ? `wpx:${g.wpx}` : (g.wch !== undefined ? `wch:${g.wch.toFixed(2)}` : "(default)");
    if (hW !== gW) {
      colDiffs++;
      if (colDiffs <= 8) console.log(`  col ${XLSX.utils.encode_col(i)}: hist=${hW} gen=${gW}`);
    }
  }
  if (colDiffs === 0) {
    console.log("Col widths: MATCH ✓");
  } else {
    console.log(`Col widths: DIFF — ${colDiffs} cols differ`);
  }

  // --- Cell values and formats ---
  const hRange = rangeOf(histSheet);
  const gRange = rangeOf(genSheet);
  if (!hRange || !gRange) { console.log("(cannot compare cells — no range)"); return; }

  const maxR = Math.max(hRange.e.r, gRange.e.r);
  const maxC = Math.max(hRange.e.c, gRange.e.c);

  let fmtDiffs = 0;
  let valDiffs = 0;
  const fmtDiffRows = [];
  const valDiffRows = [];

  for (let r = 0; r <= maxR; r++) {
    for (let c = 0; c <= maxC; c++) {
      const addr = cellAddr(r, c);
      const hCell = histSheet[addr];
      const gCell = genSheet[addr];

      // Skip empty-vs-empty
      if (!hCell && !gCell) continue;

      const hFmt = fmtStr(hCell);
      const gFmt = fmtStr(gCell);
      const hVal = valStr(hCell);
      const gVal = valStr(gCell);

      if (hFmt !== gFmt) {
        fmtDiffs++;
        if (fmtDiffRows.length < 20) {
          fmtDiffRows.push(`  ${addr}: hist_fmt="${hFmt}" gen_fmt="${gFmt}" val="${hVal}"`);
        }
      }
      if (hVal !== gVal) {
        valDiffs++;
        if (valDiffRows.length < 20) {
          valDiffRows.push(`  ${addr}: hist="${hVal}" gen="${gVal}"`);
        }
      }
    }
  }

  console.log(`Format diffs: ${fmtDiffs}${fmtDiffs === 0 ? " ✓" : ""}`);
  fmtDiffRows.forEach(d => console.log(d));
  if (fmtDiffs > 20) console.log(`  ... (showing first 20 of ${fmtDiffs})`);

  console.log(`Value diffs:  ${valDiffs}${valDiffs === 0 ? " ✓" : ""}`);
  valDiffRows.forEach(d => console.log(d));
  if (valDiffs > 20) console.log(`  ... (showing first 20 of ${valDiffs})`);
}

// Run comparisons
for (const [key, candidates] of Object.entries(DASH_SHEETS)) {
  const h = findSheet(hist, candidates);
  const g = findSheet(gen,  candidates);
  if (!h) { console.log(`\n[SKIP] ${key}: not found in historical workbook`); continue; }
  if (!g) { console.log(`\n[SKIP] ${key}: not found in generated workbook`);  continue; }
  compareSheet(`${key.toUpperCase()} (hist="${h.name}", gen="${g.name}")`, h.sheet, g.sheet);
}

console.log("\n\nDone.");
