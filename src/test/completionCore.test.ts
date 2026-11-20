import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import type { AccessConnection, AccessControlInfo, AccessObjectInfo, AccessTableFieldInfo } from "../models/types";
import { getSqlCompletions, getVbaCompletions } from "../intellisense/completionCore";
import { AccessMetadataService } from "../services/accessMetadataService";
import type { AccessMetadataSource } from "../services/accessMetadataService";

const connection: AccessConnection = { id: "a", name: "Sales", dbPath: "C:\\db\\sales.accdb" };

function createService() {
    const calls: string[] = [];
    const source: AccessMetadataSource = {
        async listObjects(_connection, objectType) {
            calls.push(`list:${objectType}`);
            const names = objectType === "form" ? ["frmCustomers", "frmOrders", "Order Entry"] : ["Customers", "Orders"];
            return names.map((name) => ({ name } as AccessObjectInfo));
        },
        async getControls(_connection, _type, name) {
            calls.push(`controls:${name}`);
            return ["CustomerID", "CustomerName", "Save Button"].map((n) => ({ name: n, type_name: "TextBox" } as AccessControlInfo));
        },
        async getTableFields(_connection, name) {
            calls.push(`fields:${name}`);
            return [
                { name: "OrderID", type: "LONG", required: true },
                { name: "Order", type: "TEXT", size: 20 },
                { name: "Amount", type: "CURRENCY" }
            ] as AccessTableFieldInfo[];
        }
    };
    return { service: new AccessMetadataService(source), calls };
}

function sql(sample: string): [string, number] {
    const offset = sample.indexOf("|");
    return [sample.slice(0, offset) + sample.slice(offset + 1), offset];
}

describe("getVbaCompletions", () => {
    it("suggests the database forms after 'Forms!'", async () => {
        const { service } = createService();
        const entries = await getVbaCompletions(service, connection, "Forms!", "");

        assert.deepEqual(entries.map((e) => e.label), ["frmCustomers", "frmOrders", "Order Entry"]);
        assert.ok(entries.every((e) => e.kind === "form"));
        assert.equal(entries[0].insertText, "frmCustomers");
        assert.equal(entries[2].insertText, "[Order Entry]", "names that need brackets are bracketed");
    });

    it("suggests the controls of the form after 'Forms!frmCustomers!'", async () => {
        const { service, calls } = createService();
        const entries = await getVbaCompletions(service, connection, "x = forms!FRMCUSTOMERS!", "");

        assert.deepEqual(entries.map((e) => e.label), ["CustomerID", "CustomerName", "Save Button"]);
        assert.equal(entries[1].detail, "TextBox");
        assert.equal(entries[2].insertText, "[Save Button]");
        assert.deepEqual(calls, ["list:form", "controls:frmCustomers"], "uses the canonical form name");
    });

    it("replaces the typed partial, including an auto-closed bracket", async () => {
        const { service } = createService();
        const [plain] = await getVbaCompletions(service, connection, "Forms!frmC", "");
        assert.equal(plain.replaceBefore, 4);
        assert.equal(plain.replaceAfter, 0);

        const bracketed = await getVbaCompletions(service, connection, "Forms![Ord", "]");
        const entry = bracketed.find((e) => e.label === "Order Entry");
        assert.equal(entry?.insertText, "[Order Entry]");
        assert.equal(entry?.filterText, "[Order Entry]");
        assert.equal(entry?.replaceBefore, 4);
        assert.equal(entry?.replaceAfter, 1);
    });

    it("returns nothing for unknown forms and unrelated code without calling Access", async () => {
        const { service, calls } = createService();

        assert.deepEqual(await getVbaCompletions(service, connection, "Forms!frmMissing!", ""), []);
        assert.deepEqual(calls, ["list:form"]);

        const other = createService();
        assert.deepEqual(await getVbaCompletions(other.service, connection, "Dim x As Long", ""), []);
        assert.deepEqual(await getVbaCompletions(other.service, connection, "' Forms!", ""), []);
        assert.deepEqual(other.calls, []);
    });

    it("does not call Access again on the warm path", async () => {
        const { service, calls } = createService();

        await getVbaCompletions(service, connection, "Forms!", "");
        await getVbaCompletions(service, connection, "Forms!frm", "");
        await getVbaCompletions(service, connection, "Forms!frmCustomers!", "");
        await getVbaCompletions(service, connection, "Forms!frmCustomers!Cu", "");

        assert.deepEqual(calls, ["list:form", "controls:frmCustomers"]);
    });
});

describe("getSqlCompletions", () => {
    it("suggests the fields of the table behind the alias after 'o.'", async () => {
        const { service, calls } = createService();
        const [text, offset] = sql("SELECT o.|\nFROM Orders o");
        const entries = await getSqlCompletions(service, connection, text, offset);

        assert.deepEqual(entries.map((e) => e.label), ["OrderID", "Order", "Amount"]);
        assert.ok(entries.every((e) => e.kind === "field"));
        assert.equal(entries[0].detail, "LONG required");
        assert.equal(entries[1].detail, "TEXT (20)");
        assert.equal(entries[1].insertText, "[Order]", "reserved words are bracketed");
        assert.deepEqual(calls, ["list:table", "fields:Orders"]);
    });

    it("works for JOIN aliases and does not call Access again on the warm path", async () => {
        const { service, calls } = createService();
        const [text, offset] = sql("SELECT o.|\nFROM Customers c\nINNER JOIN Orders o ON c.CustomerID = o.CustomerID");

        await getSqlCompletions(service, connection, text, offset);
        const [typed, typedOffset] = sql("SELECT o.Am|\nFROM Customers c\nINNER JOIN Orders o ON c.CustomerID = o.CustomerID");
        const entries = await getSqlCompletions(service, connection, typed, typedOffset);

        assert.equal(entries[0].replaceBefore, 2);
        assert.deepEqual(calls, ["list:table", "fields:Orders"]);
    });

    it("returns nothing for unknown aliases, unknown tables and unsupported syntax", async () => {
        const { service, calls } = createService();

        assert.deepEqual(await getSqlCompletions(service, connection, ...sql("SELECT x.|\nFROM Orders o")), []);
        assert.deepEqual(await getSqlCompletions(service, connection, ...sql("SELECT q.|\nFROM SavedQuery q")), []);
        assert.deepEqual(await getSqlCompletions(service, connection, ...sql("SELECT o.|\nFROM (SELECT * FROM Orders) o")), []);
        assert.deepEqual(calls, ["list:table"], "fields are only requested for a known table");
    });
});
