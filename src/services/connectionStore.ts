import * as crypto from "crypto";
import * as vscode from "vscode";
import { AccessConnection } from "../models/types";

const CONNECTIONS_KEY = "accessExplorer.connections";

export class ConnectionStore {
    constructor(private readonly context: vscode.ExtensionContext) { }

    getAll(): AccessConnection[] {
        return this.context.globalState.get<AccessConnection[]>(CONNECTIONS_KEY, []);
    }

    findByDbPath(dbPath: string): AccessConnection | undefined {
        const normalized = normalizeDbPath(dbPath);
        return this.getAll().find((connection) => normalizeDbPath(connection.dbPath) === normalized);
    }

    async add(name: string, dbPath: string): Promise<AccessConnection> {
        const connection: AccessConnection = {
            id: crypto.randomUUID(),
            name,
            dbPath
        };

        const next = [...this.getAll(), connection];
        await this.context.globalState.update(CONNECTIONS_KEY, next);
        return connection;
    }

    async upsert(name: string, dbPath: string): Promise<AccessConnection> {
        const normalized = normalizeDbPath(dbPath);
        const current = this.getAll();
        const existing = current.find((connection) => normalizeDbPath(connection.dbPath) === normalized);

        if (!existing) {
            return await this.add(name, dbPath);
        }

        const updated: AccessConnection = {
            ...existing,
            name,
            dbPath
        };
        const next = current.map((connection) => connection.id === existing.id ? updated : connection);
        await this.context.globalState.update(CONNECTIONS_KEY, next);
        return updated;
    }

    async remove(id: string): Promise<void> {
        const next = this.getAll().filter((conn) => conn.id !== id);
        await this.context.globalState.update(CONNECTIONS_KEY, next);
    }
}

function normalizeDbPath(dbPath: string): string {
    return dbPath.trim().replace(/\//g, "\\").toLowerCase();
}
