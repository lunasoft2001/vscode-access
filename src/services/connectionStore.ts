import * as crypto from "crypto";
import * as vscode from "vscode";
import { AccessConnection } from "../models/types";

const CONNECTIONS_KEY = "accessExplorer.connections";

export class ConnectionStore {
    constructor(private readonly context: vscode.ExtensionContext) { }

    getAll(): AccessConnection[] {
        return this.context.globalState.get<AccessConnection[]>(CONNECTIONS_KEY, []);
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

    async remove(id: string): Promise<void> {
        const next = this.getAll().filter((conn) => conn.id !== id);
        await this.context.globalState.update(CONNECTIONS_KEY, next);
    }
}
