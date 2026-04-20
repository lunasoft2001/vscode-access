const fs = require("fs");
const bas = fs.readFileSync("src/vba/SecondBrainBulkExport.bas", "utf8");
// For TypeScript template literal: escape \, ` and ${
const escaped = bas
  .replace(/\\/g, "\\\\")   // \ → \\
  .replace(/`/g, "\\`")     // ` → \`
  .replace(/\${/g, "\\${"); // ${ → \${
const ts = "export const BULK_EXPORT_VBA = `" + escaped + "`;";
fs.writeFileSync("src/vba/bulkExportVba.ts", ts, "utf8");
const lines = ts.split("\n");
console.log("Lines:", lines.length);
const idx = lines.findIndex(l => l.includes("01_Tablas"));
if (idx >= 0) console.log("Sample:", lines[idx]);
