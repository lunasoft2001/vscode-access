import type { AccessConnection } from "../models/types";
import { normalizeDbPath } from "./connectionStore";

type SqlDocumentConnectionCheck =
    | { kind: "ready"; connection: AccessConnection }
    | { kind: "missing" }
    | { kind: "mismatch" };

export function checkSqlDocumentConnection(
    bound: AccessConnection,
    selected: AccessConnection | undefined,
    connections: readonly AccessConnection[]
): SqlDocumentConnectionCheck {
    const stored = connections.find((item) => item.id === bound.id
        && normalizeDbPath(item.dbPath) === normalizeDbPath(bound.dbPath));
    if (!stored) {
        return { kind: "missing" };
    }
    if (!selected || selected.id !== stored.id
        || normalizeDbPath(selected.dbPath) !== normalizeDbPath(stored.dbPath)) {
        return { kind: "mismatch" };
    }
    return { kind: "ready", connection: stored };
}
