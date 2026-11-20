const fs = require("node:fs");
const path = require("node:path");
const { spawnSync } = require("node:child_process");

const testDir = path.join(process.cwd(), ".test-build", "test");
const files = fs.existsSync(testDir)
    ? fs.readdirSync(testDir).filter((name) => name.endsWith(".test.js")).map((name) => path.join(testDir, name))
    : [];

if (files.length === 0) {
    console.error("ERROR: no compiled tests found in .test-build/test. Run `npm test` (it compiles them first).");
    process.exit(1);
}

// Explicit file list works on every Node version (glob arguments need Node 21+).
const result = spawnSync(process.execPath, ["--test", ...files], { stdio: "inherit" });
process.exit(result.status ?? 1);
