// ─── Pivot Engine ─────────────────────────────────────────────────────────────
// Detects, preserves, and refresh-marks pivot tables in XLSX workbooks.
//
// OOXML pivot tables live in two ZIP parts outside the worksheet data:
//   xl/pivotTables/pivotTable*.xml     — pivot table definition
//   xl/pivotCache/pivotCacheDefinition*.xml — cached source data
//
// xlsx-js-style does NOT parse or preserve these parts when doing read→write.
// The pivot engine uses ZIP-level manipulation (via fflate) to:
//   1. Detect pivots in the TEMPLATE buffer
//   2. Set refreshOnLoad="1" on all pivot cache definitions
//   3. Inject the pivot parts into the OUTPUT buffer
//   4. Patch [Content_Types].xml and workbook.xml.rels to register them
//
// When the finance team opens the workbook, Excel refreshes all pivots from
// the updated source data automatically — no manual Refresh All needed.
//
// Compatibility:
//   - Excel for Windows/Mac: refreshOnLoad is fully honoured
//   - Excel 365 (web):       pivots refresh on open
//   - Google Sheets import:  pivot tables are not supported (documented below)

import * as fflate from "fflate";
import * as zlib from "node:zlib";

export interface PivotPartInfo {
  path:       string;   // e.g. "xl/pivotTables/pivotTable1.xml"
  cacheRef?:  string;   // e.g. "xl/pivotCache/pivotCacheDefinition1.xml"
  hasRefreshOnLoad: boolean;
}

export interface PivotReport {
  templateHasPivots: boolean;
  outputHasPivots:   boolean;
  pivotsInjected:    boolean;
  pivotParts:        PivotPartInfo[];
  warnings:          string[];
  summary:           string;
}

// ─── ZIP utilities (synchronous) ─────────────────────────────────────────────

function unzipSync(buffer: Buffer): Record<string, Uint8Array> {
  return fflate.unzipSync(new Uint8Array(buffer));
}

function zipSync(files: Record<string, Uint8Array>): Buffer {
  const zipped = fflate.zipSync(files, { level: 6 });
  return Buffer.from(zipped);
}

// ─── XML helpers ─────────────────────────────────────────────────────────────

function setRefreshOnLoad(xml: string): string {
  // Add or update refreshOnLoad="1" attribute on <pivotCacheDefinition>
  return xml.replace(
    /<pivotCacheDefinition(\s)/,
    '<pivotCacheDefinition refreshOnLoad="1"$1',
  ).replace(
    /refreshOnLoad="0"/g,
    'refreshOnLoad="1"',
  );
}

function hasAttr(xml: string, attr: string): boolean {
  return new RegExp(`\\b${attr}\\s*=`).test(xml);
}

function decodeXml(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString("utf8");
}

function encodeXml(xml: string): Uint8Array {
  return new Uint8Array(Buffer.from(xml, "utf8"));
}

// ─── Pivot detection ─────────────────────────────────────────────────────────

/**
 * Detect pivot table XML parts in an XLSX buffer.
 * Returns info about each pivot part found.
 */
export function detectPivots(buffer: Buffer): PivotPartInfo[] {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(buffer);
  } catch {
    return [];
  }

  const parts: PivotPartInfo[] = [];

  for (const path of Object.keys(files)) {
    if (path.startsWith("xl/pivotTables/pivotTable")) {
      const xml = decodeXml(files[path]);
      // Extract the cache reference from the pivot table's .rels file
      const relsPath = path.replace(/xl\/pivotTables\/([^/]+)$/, "xl/pivotTables/_rels/$1.rels");
      let cacheRef: string | undefined;
      if (files[relsPath]) {
        const relsXml = decodeXml(files[relsPath]);
        const m = relsXml.match(/Target="[^"]*pivotCacheDefinition([^"]+)"/);
        if (m) cacheRef = `xl/pivotCache/pivotCacheDefinition${m[1]}`;
      }
      parts.push({ path, cacheRef, hasRefreshOnLoad: xml.includes('refreshOnLoad="1"') });
    }
  }

  return parts;
}

// ─── Pivot injection ─────────────────────────────────────────────────────────

/**
 * Extract pivot XML parts from the template, set refreshOnLoad on all cache
 * definitions, and inject them into the output buffer.
 *
 * This re-attaches the pivot tables to the hydrated workbook so that when
 * Excel opens it, the pivots auto-refresh from the updated source data.
 *
 * If the template or output can't be parsed, returns the output buffer unchanged
 * and records the failure in the report.
 */
export function injectPivotParts(
  templateBuffer: Buffer,
  outputBuffer: Buffer,
): { buffer: Buffer; report: PivotReport } {
  const warnings: string[] = [];

  let templateFiles: Record<string, Uint8Array>;
  let outputFiles:   Record<string, Uint8Array>;

  try {
    templateFiles = unzipSync(templateBuffer);
  } catch (e) {
    return {
      buffer: outputBuffer,
      report: {
        templateHasPivots: false,
        outputHasPivots:   false,
        pivotsInjected:    false,
        pivotParts:        [],
        warnings:          [`Could not parse template ZIP: ${e}`],
        summary:           "Template ZIP parse failed — pivots not injected",
      },
    };
  }

  try {
    outputFiles = unzipSync(outputBuffer);
  } catch (e) {
    return {
      buffer: outputBuffer,
      report: {
        templateHasPivots: false,
        outputHasPivots:   false,
        pivotsInjected:    false,
        pivotParts:        [],
        warnings:          [`Could not parse output ZIP: ${e}`],
        summary:           "Output ZIP parse failed — pivots not injected",
      },
    };
  }

  // Gather all pivot-related paths from the template
  const pivotPaths = Object.keys(templateFiles).filter(
    (p) => p.startsWith("xl/pivotTables/") || p.startsWith("xl/pivotCache/"),
  );

  const templateHasPivots = pivotPaths.length > 0;
  const outputHasPivots   = Object.keys(outputFiles).some(
    (p) => p.startsWith("xl/pivotTables/") || p.startsWith("xl/pivotCache/"),
  );

  if (!templateHasPivots) {
    return {
      buffer: outputBuffer,
      report: {
        templateHasPivots: false,
        outputHasPivots:   false,
        pivotsInjected:    false,
        pivotParts:        [],
        warnings:          [],
        summary:           "No pivot tables in template — nothing to inject",
      },
    };
  }

  // Copy all pivot paths from template to output, patching cache defs
  const injectedParts: PivotPartInfo[] = [];

  for (const path of pivotPaths) {
    const bytes = templateFiles[path];
    if (!bytes) continue;

    let xml = decodeXml(bytes);

    // Set refreshOnLoad on cache definitions
    if (path.startsWith("xl/pivotCache/pivotCacheDefinition")) {
      const alreadyHas = xml.includes('refreshOnLoad="1"');
      xml = setRefreshOnLoad(xml);
      if (!alreadyHas) {
        injectedParts.push({ path, hasRefreshOnLoad: true });
      }
    }

    outputFiles[path] = encodeXml(xml);
  }

  // Also copy the _rels files for pivot tables
  const pivotRelsPrefix = "xl/pivotTables/_rels/";
  for (const path of Object.keys(templateFiles)) {
    if (path.startsWith(pivotRelsPrefix)) {
      outputFiles[path] = templateFiles[path];
    }
  }

  // Patch [Content_Types].xml to register the pivot parts
  const contentTypesKey = "[Content_Types].xml";
  if (outputFiles[contentTypesKey]) {
    let ctXml = decodeXml(outputFiles[contentTypesKey]);

    // Ensure pivot content types are registered
    const pivotTableCT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml';
    const pivotCacheCT  = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml';
    const pivotRecordCT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml';

    const addDefaultIfMissing = (ct: string) => {
      if (!ctXml.includes(ct)) {
        // Add before </Types>
        ctXml = ctXml.replace(
          "</Types>",
          `<Override ContentType="${ct}" PartName="PLACEHOLDER"/></Types>`,
        );
      }
    };

    // Instead of adding <Default>, add <Override> entries for each pivot part
    for (const path of pivotPaths) {
      const partName = `/${path}`;
      if (!ctXml.includes(partName)) {
        let ct = "";
        if (path.startsWith("xl/pivotTables/pivotTable")) ct = pivotTableCT;
        else if (path.startsWith("xl/pivotCache/pivotCacheDefinition")) ct = pivotCacheCT;
        else if (path.startsWith("xl/pivotCache/pivotCacheRecords")) ct = pivotRecordCT;

        if (ct) {
          ctXml = ctXml.replace(
            "</Types>",
            `<Override PartName="${partName}" ContentType="${ct}"/></Types>`,
          );
        }
      }
    }

    // Clean up any PLACEHOLDER entries we may have created
    ctXml = ctXml.replace(/<Override ContentType="[^"]*" PartName="PLACEHOLDER"\/>/g, "");

    outputFiles[contentTypesKey] = encodeXml(ctXml);
  }

  // Patch xl/_rels/workbook.xml.rels to register pivot cache relationships
  const wbRelsKey = "xl/_rels/workbook.xml.rels";
  if (outputFiles[wbRelsKey] && templateFiles[wbRelsKey]) {
    let outRels = decodeXml(outputFiles[wbRelsKey]);
    const tplRels = decodeXml(templateFiles[wbRelsKey]);

    // Extract pivot cache relationship entries from template
    const pivotCacheRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition';
    const relRegex = new RegExp(`<Relationship[^/]*/pivotCache[^>]+>`, 'g');
    const tplPivotRels = tplRels.match(relRegex) ?? [];

    for (const rel of tplPivotRels) {
      if (!outRels.includes(rel)) {
        outRels = outRels.replace("</Relationships>", `${rel}</Relationships>`);
      }
    }

    outputFiles[wbRelsKey] = encodeXml(outRels);
  }

  let finalBuffer: Buffer;
  try {
    finalBuffer = zipSync(outputFiles);
  } catch (e) {
    warnings.push(`Failed to re-zip output with pivot parts: ${e}`);
    return {
      buffer: outputBuffer,
      report: {
        templateHasPivots,
        outputHasPivots,
        pivotsInjected: false,
        pivotParts: injectedParts,
        warnings,
        summary: "Re-zip failed — original output returned",
      },
    };
  }

  const pivotTableCount = pivotPaths.filter((p) => p.startsWith("xl/pivotTables/pivotTable")).length;
  const summary = `Injected ${pivotTableCount} pivot table(s) from template with refreshOnLoad=1 — pivots will auto-refresh on Excel open`;

  return {
    buffer: finalBuffer,
    report: {
      templateHasPivots,
      outputHasPivots: true,
      pivotsInjected: true,
      pivotParts: injectedParts,
      warnings,
      summary,
    },
  };
}

// ─── Style + theme deep preservation ─────────────────────────────────────────
//
// xlsx-js-style rebuilds xl/styles.xml with freshly-assigned style indices when
// writing.  Even if we copy the original styles.xml back afterwards, the cell
// <c s="N"> attributes in the worksheet XMLs still reference the NEW (wrong)
// indices.  This causes every cell to render with the wrong format/fill/border.
//
// Deep preservation does THREE things in one ZIP-level pass:
//
//  1. Copy xl/styles.xml and xl/theme/theme1.xml from template verbatim.
//
//  2. For each worksheet that exists in both template and output, restore every
//     cell's s="N" attribute and every row's s="N" / customFormat attributes
//     to the original template values.  New cells (added by hydration, e.g. the
//     P1 live-row snapshot) keep their output-assigned style.
//
//  3. Copy conditional-formatting (<conditionalFormatting>) blocks, data
//     validation (<dataValidations>), print settings (<pageMargins>,
//     <pageSetup>, <printOptions>), and auto-filter (<autoFilter>) from the
//     template worksheet into the output worksheet, replacing any stale copies.
//
// This makes the output indistinguishable from the master template visually.

// ─── Sheet-name → worksheet-file map ─────────────────────────────────────────

function buildSheetFileMap(files: Record<string, Uint8Array>): Map<string, string> {
  const result = new Map<string, string>();
  try {
    const wbRaw  = files["xl/workbook.xml"];
    const relRaw = files["xl/_rels/workbook.xml.rels"];
    if (!wbRaw || !relRaw) return result;

    const wbXml  = decodeXml(wbRaw);
    const relXml = decodeXml(relRaw);

    // rId → file path from .rels
    const ridToFile = new Map<string, string>();
    const relRe = /<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*/g;
    let m: RegExpExecArray | null;
    while ((m = relRe.exec(relXml)) !== null) {
      // Target may be "worksheets/sheet1.xml" or "../worksheets/sheet1.xml"
      const target = m[2].replace(/^\.\.\//, "");
      ridToFile.set(m[1], target.startsWith("xl/") ? target : `xl/${target}`);
    }

    // Sheet name → rId from workbook.xml (handles r:id and rId namespace variants)
    const sheetRe = /<sheet\b[^>]*\bname="([^"]+)"[^>]*(?:r:id|rId)="([^"]+)"[^>]*/gi;
    while ((m = sheetRe.exec(wbXml)) !== null) {
      const file = ridToFile.get(m[2]);
      if (file) result.set(m[1], file);
    }
  } catch { /* ignore — caller handles empty map */ }
  return result;
}

// ─── Cell and row s-attribute restoration ────────────────────────────────────

function restoreCellStyleIndices(templateWsXml: string, outputWsXml: string): string {
  // Build: cell address → s index (or null = no s attr) from template
  const styleMap  = new Map<string, string | null>();
  // Intentionally broad: match <c ... r="ADDR" ... > and <c ... r="ADDR" ... />
  const cellTagRe = /<c\b([^>]*?)(\s*\/)?>/g;
  let m: RegExpExecArray | null;

  while ((m = cellTagRe.exec(templateWsXml)) !== null) {
    const attrs = m[1];
    const addrM = /\br="([A-Z]+\d+)"/.exec(attrs);
    if (!addrM) continue;
    const sM = /\bs="(\d+)"/.exec(attrs);
    styleMap.set(addrM[1], sM ? sM[1] : null);
  }
  if (styleMap.size === 0) return outputWsXml;

  return outputWsXml.replace(/<c\b([^>]*?)(\s*\/)?>/g, (full, attrs, sc) => {
    const addrM = /\br="([A-Z]+\d+)"/.exec(attrs);
    if (!addrM) return full;
    const addr = addrM[1];
    if (!styleMap.has(addr)) return full; // new cell — preserve output style

    const tmplS = styleMap.get(addr);
    const slash  = sc ?? "";

    if (tmplS === null) {
      // Template cell has no s attr — strip from output
      const stripped = attrs.replace(/\s*\bs="\d+"/, "");
      return `<c${stripped}${slash}>`;
    }

    const newAttrs = /\bs="\d+"/.test(attrs)
      ? attrs.replace(/\bs="\d+"/, `s="${tmplS}"`)
      : attrs.replace(/(\br="[A-Z]+\d+")/, `$1 s="${tmplS}"`);
    return `<c${newAttrs}${slash}>`;
  });
}

function restoreRowStyleIndices(templateWsXml: string, outputWsXml: string): string {
  // Build: row number → {s, customFormat} from template
  const rowMap = new Map<string, { s: string | null; cf: string | null }>();
  const rowTagRe = /<row\b([^>]*?)(\s*\/)?>/g;
  let m: RegExpExecArray | null;

  while ((m = rowTagRe.exec(templateWsXml)) !== null) {
    const attrs = m[1];
    const rM  = /\br="(\d+)"/.exec(attrs);
    if (!rM) continue;
    const sM  = /\bs="(\d+)"/.exec(attrs);
    const cfM = /\bcustomFormat="(\d+)"/.exec(attrs);
    rowMap.set(rM[1], { s: sM ? sM[1] : null, cf: cfM ? cfM[1] : null });
  }
  if (rowMap.size === 0) return outputWsXml;

  return outputWsXml.replace(/<row\b([^>]*?)(\s*\/)?>/g, (full, attrs, sc) => {
    const rM = /\br="(\d+)"/.exec(attrs);
    if (!rM) return full;
    const rowNum = rM[1];
    if (!rowMap.has(rowNum)) return full; // new row

    const info  = rowMap.get(rowNum)!;
    const slash = sc ?? "";

    let newAttrs = attrs;
    if (info.s === null) {
      newAttrs = newAttrs.replace(/\s*\bs="\d+"/, "").replace(/\s*\bcustomFormat="\d+"/, "");
    } else {
      newAttrs = /\bs="\d+"/.test(newAttrs)
        ? newAttrs.replace(/\bs="\d+"/, `s="${info.s}"`)
        : newAttrs.replace(/(\br="\d+")/, `$1 s="${info.s}"`);
      if (info.cf) {
        newAttrs = /\bcustomFormat="\d+"/.test(newAttrs)
          ? newAttrs.replace(/\bcustomFormat="\d+"/, `customFormat="${info.cf}"`)
          : newAttrs + ` customFormat="${info.cf}"`;
      }
    }
    return `<row${newAttrs}${slash}>`;
  });
}

// ─── Worksheet-feature restoration ───────────────────────────────────────────

function restoreWorksheetFeature(tag: string, templateXml: string, outputXml: string): string {
  // Matches both <tag .../> and <tag ...>...</tag>
  const blockRe = new RegExp(`<${tag}\\b[\\s\\S]*?(?:<\\/${tag}>|\\/>)`, "g");
  const tmplBlocks = templateXml.match(blockRe);

  // Remove stale copies from output regardless
  let result = outputXml.replace(blockRe, "");

  if (!tmplBlocks || tmplBlocks.length === 0) return result;

  // Re-inject before </worksheet>
  return result.replace("</worksheet>", `${tmplBlocks.join("")}</worksheet>`);
}

// ─── P1 snapshot structural row insertion ────────────────────────────────────
//
// The P1 "Dashboard" sheet has a rolling 12-month snapshot block:
//
//   row N-12 … row N-1 : data rows  (col B = date serial, C-I = KPI values)
//   row N               : =AVERAGE(C(N-12):C(N-1)) formulas
//
// This function inserts a new data row immediately above the AVERAGE row,
// shifts the AVERAGE row down, extends the AVERAGE range by 1, and hides the
// oldest row when the window exceeds 12 months.  It operates entirely at the
// ZIP/XML level so no style information is disturbed.

export interface P1SnapshotSpec {
  sheetName:            string;
  monthSerial:          number;   // Excel date serial for 1st of month
  occupancyPctSpaces:   number;   // decimal (0–1)
  occupancyPctSqFt:     number;
  contracts:            number;
  flexMembers:          number;
  mailMembers:          number;
  meetingHours:         number;
  meetingRevenueDollar: number;
}

export function insertP1SnapshotRow(
  outputBuffer: Buffer,
  spec: P1SnapshotSpec,
): { buffer: Buffer; action: "inserted" | "updated" | "skipped"; warnings: string[] } {
  const warnings: string[] = [];

  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(outputBuffer);
  } catch (e) {
    return { buffer: outputBuffer, action: "skipped", warnings: [`ZIP parse failed: ${e}`] };
  }

  const sheetMap = buildSheetFileMap(files);
  const wsPath   = sheetMap.get(spec.sheetName);
  if (!wsPath || !files[wsPath]) {
    return { buffer: outputBuffer, action: "skipped", warnings: [`Sheet "${spec.sheetName}" not found`] };
  }

  let wsXml = decodeXml(files[wsPath]);

  // ── Parse all <row> elements ──────────────────────────────────────────────
  // We need row number, visibility, and cell content.

  interface RowInfo {
    num:    number;       // 1-based Excel row number
    xml:    string;       // full <row ...>...</row> or <row .../> text
    hidden: boolean;
    cells:  Map<string, string>;  // col letter → raw value (v element text or formula)
    hasDateSerial: boolean;       // col B is an Excel date serial 40000–70000
    hasAverage: boolean;          // contains AVERAGE( formula
  }

  function parseRows(xml: string): RowInfo[] {
    const rowRe = /<row\b[^>]*>[\s\S]*?<\/row>|<row\b[^>]*\/>/g;
    const rows: RowInfo[] = [];
    let m: RegExpExecArray | null;
    while ((m = rowRe.exec(xml)) !== null) {
      const rowXml = m[0];
      const rM = /\br="(\d+)"/.exec(rowXml);
      if (!rM) continue;
      const num = parseInt(rM[1], 10);

      const hidden = /\bhidden="1"/.test(rowXml);
      const cells  = new Map<string, string>();

      // Extract cells: <c r="B5" ...><v>...</v></c> or <c r="B5" ...><f>...</f></c>
      const cellRe = /<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b[^>]*\/>/g;
      let cm: RegExpExecArray | null;
      while ((cm = cellRe.exec(rowXml)) !== null) {
        const cattrs = cm[1] ?? "";
        const cbody  = cm[2] ?? "";
        const addrM  = /\br="([A-Z]+\d+)"/.exec(cattrs);
        if (!addrM) continue;
        const col = addrM[1].replace(/\d+/, "");
        // Prefer <v> value, then <f> formula
        const vM = /<v>([\s\S]*?)<\/v>/.exec(cbody);
        const fM = /<f>([\s\S]*?)<\/f>/.exec(cbody);
        cells.set(col, vM ? vM[1] : fM ? fM[1] : "");
      }

      const bVal = cells.get("B") ?? "";
      const bNum = parseFloat(bVal);
      const hasDateSerial = bNum >= 40000 && bNum <= 70000;
      const hasAverage    = rowXml.includes("AVERAGE(");

      rows.push({ num, xml: rowXml, hidden, cells, hasDateSerial, hasAverage });
    }
    return rows;
  }

  const rows = parseRows(wsXml);

  // ── Identify snapshot rows and AVERAGE row ────────────────────────────────
  const snapRows = rows.filter((r) => r.hasDateSerial && !r.hasAverage);
  const avgRow   = rows.find((r) => r.hasAverage);

  if (snapRows.length === 0) {
    warnings.push("No snapshot rows found (no col-B date serials 40000–70000)");
    return { buffer: outputBuffer, action: "skipped", warnings };
  }
  if (!avgRow) {
    warnings.push("AVERAGE row not found — skipping structural insertion");
    return { buffer: outputBuffer, action: "skipped", warnings };
  }

  // ── Check if this month already exists ───────────────────────────────────
  const existingRow = snapRows.find((r) => {
    const bVal = r.cells.get("B") ?? "";
    return Math.abs(parseFloat(bVal) - spec.monthSerial) < 1;
  });

  // Column mapping: B=date, C=occSpaces, D=occSqFt, E=contracts,
  //                 F=flex, G=mail, H=meetingHrs, I=meetingRev
  const colValues: [string, number][] = [
    ["B", spec.monthSerial],
    ["C", spec.occupancyPctSpaces],
    ["D", spec.occupancyPctSqFt],
    ["E", spec.contracts],
    ["F", spec.flexMembers],
    ["G", spec.mailMembers],
    ["H", spec.meetingHours],
    ["I", spec.meetingRevenueDollar],
  ];

  // ── Helper: set a cell's <v> value in a row XML snippet ──────────────────
  function setCellValue(rowXml: string, col: string, rowNum: number, value: number): string {
    const addr = `${col}${rowNum}`;
    // Match an existing cell for this column
    const cellRe = new RegExp(`(<c\\b[^>]*\\br="${addr}"[^>]*>)[\\s\\S]*?(<\\/c>)`);
    const selfClose = new RegExp(`<c\\b[^>]*\\br="${addr}"[^>]*/>`);

    if (cellRe.test(rowXml)) {
      return rowXml.replace(cellRe, `$1<v>${value}</v>$2`);
    }
    if (selfClose.test(rowXml)) {
      return rowXml.replace(selfClose, (m) => m.replace("/>", `><v>${value}</v></c>`));
    }
    // Cell doesn't exist — insert before </row>
    return rowXml.replace("</row>", `<c r="${addr}"><v>${value}</v></c></row>`);
  }

  // ── Helper: renumber all r= attributes in a row XML snippet ──────────────
  function renumberRow(rowXml: string, newNum: number): string {
    // Update <row r="N"> attribute
    let result = rowXml.replace(/(<row\b[^>]*\br=")(\d+)(")/, `$1${newNum}$3`);
    // Update each cell address: letter(s) + old row number → new row number
    const oldNum = parseInt((/\br="(\d+)"/.exec(rowXml) ?? ["", "0"])[1], 10);
    result = result.replace(/\br="([A-Z]+)(\d+)"/g, (_, col, n) => {
      return parseInt(n, 10) === oldNum ? `r="${col}${newNum}"` : `r="${col}${n}"`;
    });
    return result;
  }

  // ── Helper: shift all rows with num >= fromNum down by 1 ─────────────────
  function shiftRowsDown(xml: string, fromNum: number): string {
    return xml.replace(/<row\b[^>]*>[\s\S]*?<\/row>|<row\b[^>]*\/>/g, (rowXml) => {
      const rM = /\br="(\d+)"/.exec(rowXml);
      if (!rM) return rowXml;
      const n = parseInt(rM[1], 10);
      if (n < fromNum) return rowXml;
      return renumberRow(rowXml, n + 1);
    });
  }

  // ── Helper: extend AVERAGE range end by 1 row ───────────────────────────
  // targetRow = the last row of the old range (= insertAtRow - 1).
  // Handles $-prefixed absolute references, e.g. AVERAGE(C$58:C$69) with targetRow=69
  // → AVERAGE(C$58:C$70).
  function extendAverageRanges(xml: string, targetRow: number): string {
    return xml.replace(/<f>([^<]*AVERAGE\([^)]+\))<\/f>/g, (_full, formula) => {
      const extended = formula.replace(/([A-Z]+)(\$?)(\d+)\)/g, (_: string, col: string, dollar: string, n: string) => {
        const num = parseInt(n, 10);
        return num === targetRow ? `${col}${dollar}${num + 1})` : `${col}${dollar}${n})`;
      });
      return `<f>${extended}</f>`;
    });
  }

  // ── Helper: add hidden="1" to a specific row ──────────────────────────────
  function hideRow(xml: string, rowNum: number): string {
    return xml.replace(/<row\b([^>]*)>/g, (full, attrs) => {
      const rM = /\br="(\d+)"/.exec(attrs);
      if (!rM || parseInt(rM[1], 10) !== rowNum) return full;
      if (/\bhidden="1"/.test(attrs)) return full;
      return `<row${attrs} hidden="1">`;
    });
  }

  // ── UPDATE existing row ───────────────────────────────────────────────────
  if (existingRow) {
    let updated = wsXml;
    for (const [col, val] of colValues) {
      // Need to replace the value within the existing row only
      const rowEsc = existingRow.xml.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const newRowXml = setCellValue(existingRow.xml, col, existingRow.num, val);
      updated = updated.replace(existingRow.xml, newRowXml);
    }
    // Un-hide if it was hidden
    if (existingRow.hidden) {
      updated = updated.replace(/<row\b([^>]*)\bhidden="1"([^>]*)>/g, (full, pre, post) => {
        const rM = /\br="(\d+)"/.exec(pre + post);
        if (!rM || parseInt(rM[1], 10) !== existingRow.num) return full;
        return `<row${pre}${post}>`;
      });
    }
    files[wsPath] = encodeXml(updated);
    try {
      return { buffer: zipSync(files), action: "updated", warnings };
    } catch (e) {
      return { buffer: outputBuffer, action: "skipped", warnings: [`Re-zip failed: ${e}`] };
    }
  }

  // ── INSERT new row ────────────────────────────────────────────────────────
  const lastSnap    = snapRows[snapRows.length - 1];  // most recent existing snapshot
  const insertAtRow = avgRow.num;                      // insert immediately above AVERAGE

  // 1. Shift AVERAGE row and everything below it down by 1
  let modified = shiftRowsDown(wsXml, insertAtRow);

  // 2. Extend AVERAGE formula ranges — targetRow = old last snap row (insertAtRow - 1)
  modified = extendAverageRanges(modified, insertAtRow - 1);

  // 3. Clone the last snapshot row at the new insertAtRow position,
  //    stripping formulas and hidden attr, then set fresh values.
  let newRowXml = renumberRow(lastSnap.xml, insertAtRow);
  // Strip any hidden attr from the cloned row
  newRowXml = newRowXml.replace(/\s*\bhidden="1"/, "");
  // Strip formula cells (<f> elements) — we'll write plain values
  newRowXml = newRowXml.replace(/<f>[^<]*<\/f>/g, "");
  // Strip existing <v> values; we'll inject fresh ones
  newRowXml = newRowXml.replace(/<v>[^<]*<\/v>/g, "");
  // Write column values
  for (const [col, val] of colValues) {
    newRowXml = setCellValue(newRowXml, col, insertAtRow, val);
  }

  // 4. Inject the new row before the (now-shifted) AVERAGE row
  const newAvgRowNum = insertAtRow + 1;
  const avgRowRe = new RegExp(`(<row\\b[^>]*\\br="${newAvgRowNum}"[^>]*>[\\s\\S]*?<\\/row>|<row\\b[^>]*\\br="${newAvgRowNum}"[^>]*/?>)`);
  if (avgRowRe.test(modified)) {
    modified = modified.replace(avgRowRe, `${newRowXml}$1`);
  } else {
    // Fallback: append before </sheetData>
    warnings.push(`Could not locate shifted AVERAGE row at ${newAvgRowNum}; appending before </sheetData>`);
    modified = modified.replace("</sheetData>", `${newRowXml}</sheetData>`);
  }

  // 5. Hide oldest visible snapshot row if window exceeds 12
  const WINDOW = 12;
  // Re-parse to get updated snapshot rows
  const updatedRows  = parseRows(modified);
  const visibleSnaps = updatedRows.filter((r) => r.hasDateSerial && !r.hasAverage && !r.hidden);
  if (visibleSnaps.length > WINDOW) {
    // Sort by col-B value ascending; oldest is smallest serial
    visibleSnaps.sort((a, b) => {
      const aV = parseFloat(a.cells.get("B") ?? "0");
      const bV = parseFloat(b.cells.get("B") ?? "0");
      return aV - bV;
    });
    const rowToHide = visibleSnaps[0].num;
    modified = hideRow(modified, rowToHide);
    warnings.push(`Rolling window: hid row ${rowToHide} (oldest month)`);
  }

  files[wsPath] = encodeXml(modified);
  try {
    return { buffer: zipSync(files), action: "inserted", warnings };
  } catch (e) {
    return { buffer: outputBuffer, action: "skipped", warnings: [`Re-zip failed: ${e}`] };
  }
}

// ─── External link infrastructure copy ───────────────────────────────────────
//
// The template workbook may contain <externalReferences> in workbook.xml and
// worksheet formulas that reference external workbooks via [N] notation, e.g.
// COUNTIF([4]Occupancy!E:E,B10).  xlsx-js-style's XLSX.write() does not emit
// <externalReferences> in workbook.xml, so Excel cannot resolve [N] references
// and reports "found a problem with some content".
//
// This function copies all externalLink* files from the template and registers
// them in [Content_Types].xml, workbook.xml.rels, and workbook.xml — using
// freshly-assigned rIds that do not conflict with the output's existing rIds.

function copyExternalLinks(
  templateFiles: Record<string, Uint8Array>,
  outputFiles:   Record<string, Uint8Array>,
): void {
  // Collect external link file paths from template
  const extPaths = Object.keys(templateFiles).filter(
    (p) => p.startsWith("xl/externalLinks/") || p.startsWith("xl/_rels/externalLinks/"),
  );
  if (extPaths.length === 0) return;

  // Copy files verbatim
  for (const p of extPaths) outputFiles[p] = templateFiles[p];

  // ── Build rId remapping: template rIds → fresh rIds that avoid output conflicts
  const outRelPath = "xl/_rels/workbook.xml.rels";
  if (!templateFiles[outRelPath] || !outputFiles[outRelPath]) return;

  const tmplRelXml = decodeXml(templateFiles[outRelPath]);
  let   outRelXml  = decodeXml(outputFiles[outRelPath]);

  // Find max existing rId number in output rels
  let maxId = 0;
  const idRe = /\bId="rId(\d+)"/g;
  let m: RegExpExecArray | null;
  while ((m = idRe.exec(outRelXml)) !== null) {
    const n = parseInt(m[1], 10);
    if (n > maxId) maxId = n;
  }

  // Extract external link Relationship entries from template and remap their rIds
  const extRelRe = /<Relationship\b[^>]*\/externalLink[^>]*\/?>/g;
  const oldToNew = new Map<string, string>();
  const newRels:  string[] = [];

  while ((m = extRelRe.exec(tmplRelXml)) !== null) {
    const entry = m[0];
    const oldId = /\bId="([^"]+)"/.exec(entry)?.[1];
    if (!oldId) continue;
    const newId  = `rId${++maxId}`;
    oldToNew.set(oldId, newId);
    newRels.push(entry.replace(`Id="${oldId}"`, `Id="${newId}"`));
  }
  if (newRels.length === 0) return;

  // Append new rels to output workbook.xml.rels
  outRelXml = outRelXml.replace("</Relationships>", `${newRels.join("")}</Relationships>`);
  outputFiles[outRelPath] = encodeXml(outRelXml);

  // ── Patch xl/workbook.xml: add <externalReferences> with remapped rIds ────
  if (outputFiles["xl/workbook.xml"] && templateFiles["xl/workbook.xml"]) {
    let outWbXml = decodeXml(outputFiles["xl/workbook.xml"]);
    // Only add if not already present (idempotent)
    if (!outWbXml.includes("<externalReferences>")) {
      const tmplWbXml = decodeXml(templateFiles["xl/workbook.xml"]);
      const extRefsM  = /<externalReferences>[\s\S]*?<\/externalReferences>/.exec(tmplWbXml);
      if (extRefsM) {
        let extRefsXml = extRefsM[0];
        for (const [oldId, newId] of oldToNew) {
          extRefsXml = extRefsXml.replace(
            new RegExp(`r:id="${oldId}"`, "g"),
            `r:id="${newId}"`,
          );
        }
        // Insert after </sheets> or before </workbook>
        outWbXml = outWbXml.includes("</sheets>")
          ? outWbXml.replace("</sheets>", `</sheets>${extRefsXml}`)
          : outWbXml.replace("</workbook>", `${extRefsXml}</workbook>`);
        outputFiles["xl/workbook.xml"] = encodeXml(outWbXml);
      }
    }
  }

  // ── Patch [Content_Types].xml: register external link overrides ───────────
  if (outputFiles["[Content_Types].xml"] && templateFiles["[Content_Types].xml"]) {
    let ctXml = decodeXml(outputFiles["[Content_Types].xml"]);
    const tmplCtXml = decodeXml(templateFiles["[Content_Types].xml"]);
    const overRe = /<Override[^>]*PartName="\/xl\/externalLinks\/externalLink[^"]*"[^>]*\/?>/g;
    while ((m = overRe.exec(tmplCtXml)) !== null) {
      const partName = /PartName="([^"]+)"/.exec(m[0])?.[1];
      if (partName && !ctXml.includes(partName)) {
        ctXml = ctXml.replace("</Types>", `${m[0]}</Types>`);
      }
    }
    outputFiles["[Content_Types].xml"] = encodeXml(ctXml);
  }
}

export function preserveTemplateStyles(
  templateBuffer: Buffer,
  outputBuffer:   Buffer,
): Buffer {
  let templateFiles: Record<string, Uint8Array>;
  let outputFiles:   Record<string, Uint8Array>;

  try {
    templateFiles = unzipSync(templateBuffer);
    outputFiles   = unzipSync(outputBuffer);
  } catch {
    return outputBuffer;
  }

  // ── 1. Copy styles.xml and theme verbatim ─────────────────────────────────
  for (const path of ["xl/styles.xml", "xl/theme/theme1.xml"]) {
    if (templateFiles[path]) outputFiles[path] = templateFiles[path];
  }

  // ── 2. Match sheets by name and restore per-worksheet formatting ──────────
  const tmplSheetMap = buildSheetFileMap(templateFiles);
  const outSheetMap  = buildSheetFileMap(outputFiles);

  for (const [sheetName, tmplPath] of tmplSheetMap) {
    const outPath = outSheetMap.get(sheetName);
    if (!outPath || !templateFiles[tmplPath] || !outputFiles[outPath]) continue;

    const tmplXml = decodeXml(templateFiles[tmplPath]);
    let   outXml  = decodeXml(outputFiles[outPath]);

    // 2a. Restore cell s-indices
    outXml = restoreCellStyleIndices(tmplXml, outXml);

    // 2b. Restore row s-indices and customFormat
    outXml = restoreRowStyleIndices(tmplXml, outXml);

    // 2c. Restore conditional formatting (critical for finance dashboard colours)
    outXml = restoreWorksheetFeature("conditionalFormatting", tmplXml, outXml);

    // 2d. Restore print settings and data validation
    for (const tag of ["dataValidations", "printOptions", "pageMargins", "pageSetup"]) {
      outXml = restoreWorksheetFeature(tag, tmplXml, outXml);
    }

    // 2e. Strip r:id from <pageSetup> — the output has no printerSettings binaries
    //     or worksheet _rels files; keeping the r:id reference causes Excel to
    //     report "found a problem with some content" on every generated file.
    outXml = outXml.replace(/<pageSetup\b([^>]*?)\/>/g, (_full, attrs) =>
      `<pageSetup${attrs.replace(/\s+r:id="[^"]*"/, "")}/>`,
    );

    outputFiles[outPath] = encodeXml(outXml);
  }

  // ── 3. Copy external link infrastructure from template ────────────────────
  // Template worksheets contain [N] external-ref formulas (e.g. [4]Occupancy!E:E).
  // Without the externalLinks files and workbook.xml <externalReferences> entry,
  // Excel cannot resolve [N] and reports "found a problem with some content".
  copyExternalLinks(templateFiles, outputFiles);

  // ── 4. Re-zip ─────────────────────────────────────────────────────────────
  try {
    return zipSync(outputFiles);
  } catch {
    return outputBuffer;
  }
}
