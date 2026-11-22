import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import type { AccessConnection, AccessControlInfo, AccessObjectInfo, AccessTableFieldInfo } from "../models/types";
import { AccessMetadataService } from "../services/accessMetadataService";

// Minimal stand-in for the parts of the "vscode" API the adapter touches.
class Position {
    constructor(readonly line: number, readonly character: number) {}
    translate(lineDelta = 0, characterDelta = 0): Position {
        return new Position(this.line + lineDelta, this.character + characterDelta);
    }
}
class Range {
    constructor(readonly start: Position, readonly end: Position) {}
}
class CompletionItem {
    insertText?: string;
    filterText?: string;
    detail?: string;
    sortText?: string;
    range?: Range;
    constructor(readonly label: string, readonly kind?: number) {}
}
class CompletionList {
    constructor(readonly items: CompletionItem[], readonly isIncomplete = false) {}
}
interface Registration {
    selector: { language: string };
    provider: any;
    triggers: string[];
}
const registrations: Registration[] = [];
const fakeVscode = {
    Position,
    Range,
    CompletionItem,
    CompletionList,
    CompletionItemKind: { Class: 6, Property: 9, Field: 4 },
    Disposable: { from: (...items: Array<{ dispose(): void }>) => ({ dispose: () => items.forEach((item) => item.dispose()) }) },
    languages: {
        registerCompletionItemProvider(selector: { language: string }, provider: unknown, ...triggers: string[]) {
            registrations.push({ selector, provider, triggers });
            return { dispose() {} };
        }
    }
};

const nodeModule = require("node:module") as { _load: (request: string, ...rest: unknown[]) => unknown };
const originalLoad = nodeModule._load;
nodeModule._load = function (this: unknown, request: string, ...rest: unknown[]) {
    return request === "vscode" ? fakeVscode : originalLoad.call(this, request, ...rest);
};
// eslint-disable-next-line @typescript-eslint/no-require-imports
const { registerAccessCompletionProviders, resolveSqlCompletionConnection } = require("../providers/accessCompletionProvider") as typeof import("../providers/accessCompletionProvider");
nodeModule._load = originalLoad;

const connection: AccessConnection = { id: "a", name: "Sales", dbPath: "C:\\db\\sales.accdb" };

function createDocument(text: string) {
    const lines = text.split("\n");
    return {
        lineAt: (line: number) => ({ text: lines[line] }),
        getText: () => text,
        offsetAt: (position: Position) => lines.slice(0, position.line).reduce((sum, line) => sum + line.length + 1, 0) + position.character
    } as any;
}

function setup(bound: boolean, options?: { resolveSqlConnection?: () => AccessConnection | undefined; gate?: Promise<void> }) {
    const calls: string[] = [];
    const service = new AccessMetadataService({
        async listObjects(_c, type) {
            calls.push(`list:${type}`);
            await options?.gate;
            return (type === "form" ? ["frmCustomers"] : ["Orders"]).map((name) => ({ name } as AccessObjectInfo));
        },
        async getControls() {
            calls.push("controls");
            return [{ name: "CustomerName", type_name: "TextBox" } as AccessControlInfo];
        },
        async getTableFields() {
            calls.push("fields");
            return [{ name: "OrderID", type: "LONG" } as AccessTableFieldInfo];
        }
    });
    registrations.length = 0;
    registerAccessCompletionProviders({
        metadata: service,
        resolveVbaConnection: () => (bound ? connection : undefined),
        resolveSqlConnection: options?.resolveSqlConnection ?? (() => (bound ? connection : undefined))
    });
    const find = (language: string) => registrations.find((r) => r.selector.language === language)!;
    return { calls, vba: find("vb"), sql: find("sql") };
}

const liveToken = { isCancellationRequested: false, onCancellationRequested: () => ({ dispose() {} }) } as any;

describe("access completion providers", () => {
    it("registers a VBA provider triggered by '!' and a SQL provider triggered by '.'", () => {
        const { vba, sql } = setup(true);
        assert.deepEqual(vba.triggers, ["!"]);
        assert.deepEqual(sql.triggers, ["."]);
    });

    it("offers nothing and never calls Access for documents not bound to an Access connection", async () => {
        const { calls, vba, sql } = setup(false);

        assert.equal(await vba.provider.provideCompletionItems(createDocument("Forms!"), new Position(0, 6), liveToken), undefined);
        assert.equal(await sql.provider.provideCompletionItems(createDocument("SELECT o.\nFROM Orders o"), new Position(0, 9), liveToken), undefined);
        assert.deepEqual(calls, []);
    });

    it("turns VBA form and control entries into completion items", async () => {
        const { vba } = setup(true);

        const forms = await vba.provider.provideCompletionItems(createDocument("Forms!frm"), new Position(0, 9), liveToken);
        assert.deepEqual(forms.items.map((item: CompletionItem) => item.label), ["frmCustomers"]);
        assert.equal(forms.isIncomplete, false);
        assert.deepEqual(forms.items[0].range, new Range(new Position(0, 6), new Position(0, 9)));

        const controls = await vba.provider.provideCompletionItems(createDocument("Forms!frmCustomers!"), new Position(0, 19), liveToken);
        assert.deepEqual(controls.items.map((item: CompletionItem) => item.label), ["CustomerName"]);
        assert.equal(controls.items[0].detail, "TextBox");
    });

    it("turns SQL field entries into completion items for the alias at the cursor", async () => {
        const { sql } = setup(true);

        const result = await sql.provider.provideCompletionItems(createDocument("SELECT o.\nFROM Orders o"), new Position(0, 9), liveToken);

        assert.deepEqual(result.items.map((item: CompletionItem) => item.label), ["OrderID"]);
        assert.equal(result.items[0].detail, "LONG");
    });

    it("only resolves SQL completion when the document and execution connections match", () => {
        const other = { ...connection, id: "b", dbPath: "C:\\db\\other.accdb" };
        assert.equal(resolveSqlCompletionConnection(connection, connection), connection);
        assert.equal(resolveSqlCompletionConnection(connection, { ...connection, dbPath: "c:/DB/sales.accdb" }), connection);
        assert.equal(resolveSqlCompletionConnection(connection, other), undefined);
        assert.equal(resolveSqlCompletionConnection(connection, { ...connection, dbPath: other.dbPath }), undefined);
        assert.equal(resolveSqlCompletionConnection(connection, undefined), undefined);
    });

    it("discards SQL completions if the selected connection changes during loading", async () => {
        let release!: () => void;
        const gate = new Promise<void>((resolve) => { release = resolve; });
        let selected: AccessConnection | undefined = connection;
        const { sql, calls } = setup(true, {
            gate,
            resolveSqlConnection: () => resolveSqlCompletionConnection(connection, selected)
        });
        const pending = sql.provider.provideCompletionItems(
            createDocument("SELECT o.\nFROM Orders o"), new Position(0, 9), liveToken
        );
        selected = { ...connection, id: "b", dbPath: "C:\\db\\other.accdb" };
        release();
        assert.equal(await pending, undefined);
        assert.deepEqual(calls, ["list:table"], "must not request fields for the previous connection");
    });

    it("returns nothing when the request was cancelled", async () => {
        const { vba } = setup(true);
        const cancelled = { isCancellationRequested: true, onCancellationRequested: () => ({ dispose() {} }) } as any;

        assert.equal(await vba.provider.provideCompletionItems(createDocument("Forms!"), new Position(0, 6), cancelled), undefined);
    });
});
