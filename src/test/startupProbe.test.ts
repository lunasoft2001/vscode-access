import { strict as assert } from "node:assert";
import { describe, it } from "node:test";

const output: string[] = [];
const fakeVscode = {
    env: { language: "en" },
    EventEmitter: class { event = () => {}; dispose() {} },
    window: { createOutputChannel: () => ({ appendLine: (line: string) => output.push(line), dispose() {} }) }
};
const nodeModule = require("node:module") as { _load: (request: string, ...rest: unknown[]) => unknown };
const originalLoad = nodeModule._load;
nodeModule._load = function (this: unknown, request: string, ...rest: unknown[]) {
    return request === "vscode" ? fakeVscode : originalLoad.call(this, request, ...rest);
};
const { McpAccessClient } = require("../mcp/mcpAccessClient") as typeof import("../mcp/mcpAccessClient");
nodeModule._load = originalLoad;

describe("startup probes", () => {
    it("reports a timed out probe even when failures are allowed and can run a later probe", async () => {
        const client = new McpAccessClient(
            () => ({} as import("vscode").WorkspaceConfiguration),
            {} as import("vscode").ExtensionContext
        );
        output.length = 0;
        await assert.rejects(
            client["runCommand"](process.execPath, ["-e", "setTimeout(() => {}, 10000)"], undefined, true, 100),
            /timed out/
        );
        assert.ok(output.some((line) => line.includes("timed out")));
        const next = await client["runCommand"](process.execPath, ["-e", "process.stdout.write('ok')"], undefined, true, 5000);
        assert.equal(next.exitCode, 0);
        assert.equal(next.stdout, "ok");
    });

    it("allows prerequisite checks to be retried after a failure", async () => {
        const client = new McpAccessClient(
            () => ({} as import("vscode").WorkspaceConfiguration),
            {} as import("vscode").ExtensionContext
        );
        let attempts = 0;
        client["runPrerequisiteCheck"] = async () => {
            if (++attempts === 1) {
                throw new Error("probe timed out");
            }
            return { pythonCommand: "python", serverScriptPath: "server.py" };
        };
        await assert.rejects(client.ensurePrerequisites(), /probe timed out/);
        await client.ensurePrerequisites();
        assert.equal(attempts, 2);
    });
});
