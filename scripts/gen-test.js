/**
 * Test harness: read source workbook and regenerate dashboard.
 */
const XLSX = require("xlsx-js-style");
const path = require("path");
const fs = require("fs");
const { generateDashboardExcel } = require("./dist-test/excel-dashboard-generator");

const SRC = path.join(__dirname, "Waco Sales Dashboard 8.1.2024.xlsx");
const OUT = path.join(__dirname, "Waco Sales Dashboard 8.1.2024 - Generated Dashboard.xlsx");

const srcBuf = fs.readFileSync(SRC);
const outBuf = generateDashboardExcel(srcBuf);
fs.writeFileSync(OUT, outBuf);
console.log("Generated:", OUT);
