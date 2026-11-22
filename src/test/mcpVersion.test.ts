import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";

const fakeVscode = {
    env: { language: "en" },
    EventEmitter: class { event = () => {}; dispose() {} },
    window: { createOutputChannel: () => ({ appendLine() {}, dispose() {} }) }
};
const nodeModule = require("node:module") as { _load: (request: string, ...rest: unknown[]) => unknown };
const originalLoad = nodeModule._load;
nodeModule._load = function (this: unknown, request: string, ...rest: unknown[]) {
    return request === "vscode" ? fakeVscode : originalLoad.call(this, request, ...rest);
};
const { McpAccessClient } = require("../mcp/mcpAccessClient") as typeof import("../mcp/mcpAccessClient");
nodeModule._load = originalLoad;

async function fixture() {
    const root = await fs.mkdtemp(path.join(os.tmpdir(), "access-mcp-version-"));
    const repo = path.join(root, "runtime", "mcp-access", "MCP-Access");
    await fs.mkdir(repo, { recursive: true });
    const config = { get: (_key: string, fallback: unknown) => fallback } as unknown as import("vscode").WorkspaceConfiguration;
    const context = { globalStorageUri: { fsPath: root } } as import("vscode").ExtensionContext;
    const client = new McpAccessClient(() => config, context);
    client["getOrCreatePrerequisiteCheck"] = async () => ({ pythonCommand: "python", serverScriptPath: path.join(repo, "access_mcp_server.py") });
    return { root, repo, client };
}

describe("managed MCP version", () => {
    it("reports the installed ZIP package version after updating, not a historical README reference", async () => {
        const { root, repo, client } = await fixture();
        try {
            await fs.writeFile(path.join(repo, "README.md"), "An old feature in v0.7.51.");
            await fs.writeFile(path.join(repo, "pyproject.toml"), '[project]\nname = "mcp"\nversion = "0.7.64"\n');
            assert.equal((await client.getMcpRuntimeInfo()).mcpVersion, "v0.7.64");

            client["commandAvailable"] = async () => false;
            client["downloadMcpAccessZipWithBackup"] = async () => {
                await fs.writeFile(path.join(repo, "pyproject.toml"), '[build-system]\nversion = "1.0.0"\n[project]\nname = "mcp"\nversion = "0.7.65"\n');
            };
            const updated = await client.updateManagedMcpRuntime();
            assert.equal(updated.mcpVersion, "v0.7.65");

            client["fetchRemoteMainCommitInfo"] = async () => ({ sha: "abc", version: "v0.7.65" });
            client["fetchLatestRemoteTagInfo"] = async () => ({ sha: "abc", version: "v0.7.65" });
            assert.equal((await client.checkManagedMcpUpdate())?.updateAvailable, false);
        } finally {
            await fs.rm(root, { recursive: true, force: true });
        }
    });

    it("does not mistake unrelated documentation or another TOML section for the installed version", async () => {
        const { root, repo, client } = await fixture();
        try {
            await fs.writeFile(path.join(repo, "README.md"), "### v0.7.51\n");
            await fs.writeFile(path.join(repo, "pyproject.toml"), '[tool.example]\nversion = "1.0.0"\n[project]\nname = "mcp"\n');
            assert.equal((await client.getMcpRuntimeInfo()).mcpVersion, "desconocida (instalado por ZIP)");
        } finally {
            await fs.rm(root, { recursive: true, force: true });
        }
    });

    it("prefers package metadata over an older Git tag and uses the commit when no exact tag exists", async () => {
        const { root, repo, client } = await fixture();
        try {
            await fs.mkdir(path.join(repo, ".git"));
            await fs.writeFile(path.join(repo, "pyproject.toml"), "[project]\nversion = '0.7.65' # installed\n");
            const calls: string[][] = [];
            client["runCommand"] = async (_command, args) => {
                calls.push(args);
                if (args.includes("describe")) { return { exitCode: 1, stdout: "", stderr: "" }; }
                if (args.includes("rev-parse")) { return { exitCode: 0, stdout: "abc1234", stderr: "" }; }
                if (args.includes("log")) { return { exitCode: 0, stdout: "Maintenance", stderr: "" }; }
                return { exitCode: 0, stdout: "https://example.test/repo.git", stderr: "" };
            };
            assert.equal((await client.getMcpRuntimeInfo()).mcpVersion, "v0.7.65");
            assert.ok(calls.some((args) => args.includes("--exact-match")));
            await fs.writeFile(path.join(repo, "pyproject.toml"), "[project]\ndynamic = ['version']\n");
            assert.equal((await client.getMcpRuntimeInfo()).mcpVersion, "commit abc1234");
        } finally {
            await fs.rm(root, { recursive: true, force: true });
        }
    });
});
