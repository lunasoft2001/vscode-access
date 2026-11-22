const path = require("node:path");
const yauzl = require("yauzl");

const archivePath = process.argv[2] || path.join(".artifacts", "access-explorer.vsix");
const required = new Set([
  "extension/node_modules/@modelcontextprotocol/sdk/package.json",
  "extension/node_modules/@modelcontextprotocol/sdk/dist/cjs/client/index.js",
  "extension/node_modules/@modelcontextprotocol/sdk/dist/cjs/client/stdio.js"
]);

yauzl.open(archivePath, { lazyEntries: true }, (error, archive) => {
  if (error) {
    console.error(`ERROR: cannot open VSIX: ${error.message}`);
    process.exitCode = 1;
    return;
  }

  archive.on("error", (archiveError) => {
    console.error(`ERROR: cannot read VSIX: ${archiveError.message}`);
    process.exitCode = 1;
  });
  archive.on("entry", (entry) => {
    required.delete(entry.fileName);
    archive.readEntry();
  });
  archive.on("end", () => {
    if (required.size > 0) {
      console.error(`ERROR: VSIX is missing runtime files:\n${[...required].join("\n")}`);
      process.exitCode = 1;
    } else {
      console.log("OK: VSIX contains the MCP client runtime");
    }
  });
  archive.readEntry();
});
