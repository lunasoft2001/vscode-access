import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";

const prompts: string[] = [];
const fakeVscode = {
    env: { language: "en" },
    EventEmitter: class { event = () => {}; dispose() {} },
    window: {
        createOutputChannel: () => ({ dispose() {} }),
        showWarningMessage: async (message: string) => { prompts.push(message); return undefined; }
    }
};
const nodeModule = require("node:module") as { _load: (request: string, ...rest: unknown[]) => unknown };
const originalLoad = nodeModule._load;
nodeModule._load = function (this: unknown, request: string, ...rest: unknown[]) {
    return request === "vscode" ? fakeVscode : originalLoad.call(this, request, ...rest);
};
const { rt } = require("../utils/runtimeL10n") as typeof import("../utils/runtimeL10n");
const { isAccessDatabaseOpenError } = require("../utils/accessRecovery") as typeof import("../utils/accessRecovery");
const { McpAccessClient } = require("../mcp/mcpAccessClient") as typeof import("../mcp/mcpAccessClient");
const { SecondBrainService } = require("../services/secondBrainService") as typeof import("../services/secondBrainService");
nodeModule._load = originalLoad;

describe("runtime localization", () => {
    it("selects existing translations for regional locales and falls back to English", () => {
        const cases = [
            ["en-US", 'Delete query "Orders"?'],
            ["es-ES", '¿Eliminar la consulta "Orders"?'],
            ["de-DE", 'Abfrage "Orders" löschen?'],
            ["fr", 'Delete query "Orders"?']
        ];
        for (const [language, expected] of cases) {
            fakeVscode.env.language = language;
            assert.equal(rt("query.delete.title", "Orders"), expected);
        }
    });

    it("formats new translations in each supported language", () => {
        for (const [language, expected] of [
            ["en", "Restart Access"],
            ["es", "Reiniciar Access"],
            ["de", "Access neu starten"]
        ]) {
            fakeVscode.env.language = language;
            assert.equal(rt("access.restartButton"), expected);
            assert.ok(rt("mcp.tableTimeout", "Orders").includes("Orders"));
        }
    });

    it("recognizes the localized database-open error without losing external variants", () => {
        for (const language of ["en", "es", "de"]) {
            fakeVscode.env.language = language;
            assert.equal(isAccessDatabaseOpenError(rt("mcp.alreadyOpen")), true);
        }
        assert.equal(isAccessDatabaseOpenError("La base de datos ya esta abierta"), true);
        assert.equal(isAccessDatabaseOpenError("database is already open"), true);
        assert.equal(isAccessDatabaseOpenError("permission denied"), false);
    });

    it("still offers the installer when a configured server script is missing in every language", async () => {
        const config = { get: () => "Z:\\nonexistent-access-explorer-test\\access_mcp_server.py" } as unknown as import("vscode").WorkspaceConfiguration;
        const context = {} as import("vscode").ExtensionContext;
        for (const language of ["en", "es", "de"]) {
            fakeVscode.env.language = language;
            prompts.length = 0;
            const client = new McpAccessClient(() => config, context);
            await assert.rejects(client.ensurePrerequisites(), { message: rt("mcp.install.notConfigured") });
            assert.deepEqual(prompts, [rt("mcp.install.prompt")]);
        }
    });

    it("keeps detailed SecondBrain log messages while localizing progress", async () => {
        fakeVscode.env.language = "de";
        const directory = await fs.mkdtemp(path.join(os.tmpdir(), "access-i18n-test-"));
        const source = { listRelationships: async () => { throw new Error("test stop"); } } as unknown as InstanceType<typeof McpAccessClient>;
        const service = new SecondBrainService(source, directory);
        const events: import("../services/secondBrainService").SecondBrainProgressEvent[] = [];
        try {
            await assert.rejects(service.exportSecondBrain(
                { id: "test", name: "Test", dbPath: "test.accdb" }, directory,
                { mode: "category", categoryKey: "relationships" },
                { onProgress: (event) => { events.push(event); } }
            ), /test stop/);
            assert.match(events[0].message, /^Preparando exportacion en /);
            assert.match(events[0].displayMessage ?? "", /^Export in .* wird vorbereitet$/);
            assert.equal(events[1].message, "Cargando relaciones");
            assert.equal(events[1].displayMessage, "Beziehungen werden geladen");
        } finally {
            await fs.rm(directory, { recursive: true, force: true });
        }
    });
});
