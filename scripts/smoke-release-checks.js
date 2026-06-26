const fs = require("node:fs");
const path = require("node:path");

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function fail(message) {
  console.error(`ERROR: ${message}`);
  process.exitCode = 1;
}

function ok(message) {
  console.log(`OK: ${message}`);
}

const root = process.cwd();
const packageJsonPath = path.join(root, "package.json");
const lockPath = path.join(root, "package-lock.json");
const changelogPath = path.join(root, "CHANGELOG.md");
const readmePaths = ["README.md", "README.es.md", "README.de.md"].map((f) => path.join(root, f));
const nlsPaths = ["package.nls.json", "package.nls.es.json", "package.nls.de.json"].map((f) => path.join(root, f));

const pkg = readJson(packageJsonPath);
const lock = readJson(lockPath);
const nlsFiles = nlsPaths.map((p) => ({ path: p, content: readJson(p) }));

if (!/^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(pkg.version)) {
  fail(`package.json version is not valid semver: ${pkg.version}`);
} else {
  ok(`Valid package version: ${pkg.version}`);
}

if (lock.version !== pkg.version) {
  fail(`package-lock.json version (${lock.version}) does not match package.json (${pkg.version})`);
} else {
  ok("Lockfile top-level version matches package.json");
}

if (lock.packages && lock.packages[""] && lock.packages[""].version !== pkg.version) {
  fail(`package-lock packages[""].version (${lock.packages[""].version}) does not match package.json (${pkg.version})`);
} else {
  ok("Lockfile root package version matches package.json");
}

const changelog = fs.readFileSync(changelogPath, "utf8");
if (!changelog.includes(`## ${pkg.version} -`)) {
  fail(`CHANGELOG.md does not contain header for version ${pkg.version}`);
} else {
  ok(`CHANGELOG contains entry for ${pkg.version}`);
}

for (const readmePath of readmePaths) {
  if (!fs.existsSync(readmePath)) {
    fail(`Missing README file: ${path.basename(readmePath)}`);
    continue;
  }
  const content = fs.readFileSync(readmePath, "utf8");
  if (content.trim().length < 500) {
    fail(`README seems too short: ${path.basename(readmePath)}`);
  } else {
    ok(`README present: ${path.basename(readmePath)}`);
  }
}

const commands = Array.isArray(pkg.contributes?.commands) ? pkg.contributes.commands : [];
for (const command of commands) {
  const title = String(command.title || "");
  const match = title.match(/^%(.+)%$/);
  if (!match) {
    continue;
  }

  const key = match[1];
  for (const nls of nlsFiles) {
    if (!(key in nls.content)) {
      fail(`Missing localization key '${key}' in ${path.basename(nls.path)}`);
    }
  }
}
ok(`Validated ${commands.length} command localization keys across package.nls files`);

if (process.exitCode && process.exitCode !== 0) {
  console.error("Release smoke checks failed.");
  process.exit(process.exitCode);
}

console.log("Release smoke checks passed.");
