import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { AccessConnection } from "../models/types";
import { McpAccessClient } from "../mcp/mcpAccessClient";
import { BULK_EXPORT_VBA } from "../vba/bulkExportVba";

const MODULE_NAME = "SecondBrainBulkExport";

/** ms - creacion + inyeccion del runner */
const INJECT_TIMEOUT_MS = 90_000;

/** ms - ejecucion del VBA completo (puede recorrer 200+ formularios) */
const RUN_TIMEOUT_MS = 900_000;   // 15 min

function sanitizeForLoadFromText(code: string): string {
    return code.replace(/[^\x09\x0A\x0D\x20-\x7E]/g, " ");
}

export class BulkExportService {
    constructor(
        private readonly mcpClient: McpAccessClient,
        _globalStoragePath: string
    ) {}

    async runJsonExport(
        connection: AccessConnection,
        mode: string,
        timeoutMs = RUN_TIMEOUT_MS
    ): Promise<Record<string, unknown>> {
        const tempPath = path.join(os.tmpdir(), `sb_bulk_${Date.now()}.json`);
        const runnerDbPath = this.resolveRunnerPath(connection.dbPath);

        await this.createAndInjectRunner(runnerDbPath);
        const runnerConn: AccessConnection = { id: "runner", name: "SecondBrainRunner", dbPath: runnerDbPath };
        try {
            await this.runVba(runnerConn, "ExportToJsonFile", connection.dbPath, tempPath, mode, timeoutMs);
            const raw = await fs.readFile(tempPath, "utf-8");
            return JSON.parse(raw) as Record<string, unknown>;
        } finally {
            try { await fs.unlink(tempPath); } catch { /* ok */ }
            await this.deleteRunner(runnerDbPath);
        }
    }

    async runFileExport(
        connection: AccessConnection,
        outputDir: string,
        mode: string,
        timeoutMs = RUN_TIMEOUT_MS
    ): Promise<void> {
        const runnerDbPath = this.resolveRunnerPath(connection.dbPath);

        await this.createAndInjectRunner(runnerDbPath);
        const runnerConn: AccessConnection = { id: "runner", name: "SecondBrainRunner", dbPath: runnerDbPath };
        try {
            await this.runVba(runnerConn, "ExportToFiles", connection.dbPath, outputDir, mode, timeoutMs);
        } finally {
            await this.deleteRunner(runnerDbPath);
        }
    }

    private resolveRunnerPath(targetDbPath: string): string {
        const dir = path.dirname(targetDbPath);
        const base = path.basename(targetDbPath, path.extname(targetDbPath));
        return path.join(dir, `${base}_Runner.accdb`);
    }

    private async createAndInjectRunner(runnerDbPath: string): Promise<void> {
        await fs.rm(runnerDbPath, { force: true });
        await this.mcpClient.createDatabase(runnerDbPath);
        const runnerConn: AccessConnection = { id: "runner", name: "SecondBrainRunner", dbPath: runnerDbPath };
        await this.mcpClient.setCode(runnerConn, "module", MODULE_NAME, sanitizeForLoadFromText(BULK_EXPORT_VBA), INJECT_TIMEOUT_MS);
        await this.mcpClient.compileModule(runnerConn, MODULE_NAME, INJECT_TIMEOUT_MS);
    }

    private async deleteRunner(runnerDbPath: string): Promise<void> {
        try { await this.mcpClient.closeAccess(); } catch { /* ok */ }
        try { await fs.rm(runnerDbPath, { force: true }); } catch { /* ok */ }
    }

    private async runVba(
        runnerConn: AccessConnection,
        subName: string,
        targetDbPath: string,
        outputPath: string,
        mode: string,
        timeoutMs: number,
    ): Promise<void> {
        const qualifiedProcedure = `${MODULE_NAME}.${subName}`;
        const args = [targetDbPath, outputPath, mode];
        try {
            await this.mcpClient.runVba(runnerConn, qualifiedProcedure, args, timeoutMs);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            throw new Error(`Exportacion VBA fallo (${message}). Usando metodo secuencial...`);
        }
    }
}
