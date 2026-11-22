import { strict as assert } from "node:assert";
import { describe, it } from "node:test";
import type { AccessConnection } from "../models/types";
import { checkSqlDocumentConnection } from "../services/sqlDocumentConnection";

const a: AccessConnection = { id: "a", name: "A", dbPath: "C:\\db\\a.accdb" };
const b: AccessConnection = { id: "b", name: "B", dbPath: "C:\\db\\b.accdb" };

describe("SQL document connection safety", () => {
    it("allows the bound and selected connection when it is still registered", () => {
        assert.deepEqual(checkSqlDocumentConnection(a, { ...a, dbPath: "c:/DB/a.accdb" }, [a, b]), {
            kind: "ready", connection: a
        });
    });

    it("blocks a different or absent selected connection", () => {
        assert.deepEqual(checkSqlDocumentConnection(a, b, [a, b]), { kind: "mismatch" });
        assert.deepEqual(checkSqlDocumentConnection(a, undefined, [a, b]), { kind: "mismatch" });
        assert.deepEqual(checkSqlDocumentConnection(a, { ...a, dbPath: b.dbPath }, [a, b]), { kind: "mismatch" });
    });

    it("blocks removed or repointed document connections", () => {
        assert.deepEqual(checkSqlDocumentConnection(a, a, [b]), { kind: "missing" });
        assert.deepEqual(checkSqlDocumentConnection(a, a, [{ ...a, dbPath: b.dbPath }]), { kind: "missing" });
    });

    it("blocks a document rebound or closed before dispatch", () => {
        assert.deepEqual(checkSqlDocumentConnection(a, a, [a, b], b), { kind: "mismatch" });
        assert.deepEqual(checkSqlDocumentConnection(a, a, [a, b], null), { kind: "mismatch" });
        assert.deepEqual(checkSqlDocumentConnection(a, a, [a, b], { ...a, dbPath: "c:/DB/a.accdb" }), {
            kind: "ready", connection: a
        });
    });
});
