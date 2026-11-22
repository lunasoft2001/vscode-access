import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import type { AccessConnection } from "../models/types";
import { BackgroundMetadataBlockedError } from "../services/accessMetadataService";

const nodeModule = require("node:module") as { _load: (request: string, ...rest: unknown[]) => unknown };
const originalLoad = nodeModule._load;
nodeModule._load = function (this: unknown, request: string, ...rest: unknown[]) {
    return request === "vscode" ? {} : originalLoad.call(this, request, ...rest);
};
const { McpAccessClient } = require("../mcp/mcpAccessClient") as typeof import("../mcp/mcpAccessClient");
nodeModule._load = originalLoad;

const a: AccessConnection = { id: "a", name: "A", dbPath: "C:\\db\\a.accdb" };

describe("background MCP metadata dispatch", () => {
    it("blocks each metadata request if the database changes while connecting", async () => {
        for (const request of [
            (client: InstanceType<typeof McpAccessClient>) => client.listObjects(a, "table", undefined, { onlyIfCurrentDatabase: true }),
            (client: InstanceType<typeof McpAccessClient>) => client.getControls(a, "form", "Form", undefined, { onlyIfCurrentDatabase: true }),
            (client: InstanceType<typeof McpAccessClient>) => client.getTableFields(a, "Orders", { onlyIfCurrentDatabase: true })
        ]) {
            let release!: () => void;
            const gate = new Promise<void>((resolve) => { release = resolve; });
            let dispatched = false;
            const client: InstanceType<typeof McpAccessClient> = Object.create(McpAccessClient.prototype);
            Reflect.set(client, "connect", () => gate);
            Reflect.set(client, "lastUsedDbPath", a.dbPath);
            Reflect.set(client, "client", { callTool: () => { dispatched = true; return Promise.resolve({ structuredContent: {} }); } });
            Reflect.set(client, "getConfig", () => ({ get: (_key: string, fallback: unknown) => fallback }));
            Reflect.set(client, "output", { appendLine() {} });

            const pending = request(client);
            Reflect.set(client, "lastUsedDbPath", "C:\\db\\b.accdb");
            release();

            await assert.rejects(pending, BackgroundMetadataBlockedError);
            assert.equal(dispatched, false);
            assert.equal(client.getLastUsedDbPath(), "C:\\db\\b.accdb");
        }
    });

    it("allows a guarded request when the normalized database path still matches", async () => {
        const client: InstanceType<typeof McpAccessClient> = Object.create(McpAccessClient.prototype);
        Reflect.set(client, "connect", () => Promise.resolve());
        Reflect.set(client, "lastUsedDbPath", "c:/DB/a.accdb");
        Reflect.set(client, "client", { callTool: () => Promise.resolve({ structuredContent: { fields: [{ name: "ID" }] } }) });
        Reflect.set(client, "getConfig", () => ({ get: (_key: string, fallback: unknown) => fallback }));
        Reflect.set(client, "output", { appendLine() {} });

        assert.deepEqual(await client.getTableFields(a, "Orders", { onlyIfCurrentDatabase: true }), [
            { name: "ID", type: undefined, size: undefined, required: undefined }
        ]);
    });
});
