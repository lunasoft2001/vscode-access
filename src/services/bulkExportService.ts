import * as cp from "node:child_process";
import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { promisify } from "node:util";
import { AccessConnection } from "../models/types";
import { McpAccessClient } from "../mcp/mcpAccessClient";
import { BULK_EXPORT_VBA } from "../vba/bulkExportVba";

const exec = promisify(cp.exec);

const MODULE_NAME = "SecondBrainBulkExport";

/** ms - creacion + inyeccion del runner */
const INJECT_TIMEOUT_MS = 90_000;

/** ms - ejecucion del VBA completo (puede recorrer 200+ formularios) */
const RUN_TIMEOUT_MS = 900_000;   // 15 min

function sanitizeForLoadFromText(code: string): string {
    return code.replace(/[^\x09\x0A\x0D\x20-\x7E]/g, " ");
}

export class BulkExportService {
    private readonly runnerDbPath: string;

    constructor(
        private readonly mcpClient: McpAccessClient,
        globalStoragePath: string
    ) {
        this.runnerDbPath = path.join(globalStoragePath, "SecondBrainRunner.accdb");
    }

    async runJsonExport(
        connection: AccessConnection,
        mode: string,
        timeoutMs = RUN_TIMEOUT_MS
    ): Promise<Record<string, unknown>> {
        const tempPath = path.join(os.tmpdir(), `sb_bulk_${Date.now()}.json`);
        const runnerConn = await this.ensureRunner();

        try {
            await this.runOnRunner(runnerConn, "ExportToJsonFile", connection.dbPath, tempPath, mode, timeoutMs);
            const raw = await fs.readFile(tempPath, "utf-8");
            return JSON.parse(raw) as Record<string, unknown>;
        } finally {
            try { await fs.unlink(tempPath); } catch { /* ok */ }
        }
    }

    async runFileExport(
        connection: AccessConnection,
        outputDir: string,
        mode: string,
        timeoutMs = RUN_TIMEOUT_MS
    ): Promise<void> {
        const runnerConn = await this.ensureRunner();
        await this.runOnRunner(runnerConn, "ExportToFiles", connection.dbPath, outputDir, mode, timeoutMs);
    }

    private async ensureRunner(): Promise<AccessConnection> {
        const exists = await fs.access(this.runnerDbPath).then(() => true).catch(() => false);
        if (!exists) {
            await this.createRunnerDatabase();
            await this.injectModuleIntoRunner();
        }
        return { id: "runner", name: "SecondBrainRunner", dbPath: this.runnerDbPath };
    }

    private async createRunnerDatabase(): Promise<void> {
        await fs.mkdir(path.dirname(this.runnerDbPath), { recursive: true });
        const escapedPath = this.runnerDbPath.replace(/'/g, "''");
        const script = [
            `$app = New-Object -ComObject Access.Application`,
            `$app.Visible = $false`,
            `try { $app.NewCurrentDatabase('${escapedPath}') } finally { $app.Quit() }`,
        ].join("; ");
        await exec(`powershell.exe -NoProfile -NonInteractive -Command "${script}"`, { timeout: 30_000 });
    }

    private async injectModuleIntoRunner(): Promise<void> {
        const runnerConn: AccessConnection = { id: "runner", name: "SecondBrainRunner", dbPath: this.runnerDbPath };
        await this.mcpClient.setCode(runnerConn, "module", MODULE_NAME, sanitizeForLoadFromText(BULK_EXPORT_VBA), INJECT_TIMEOUT_MS);
        await this.mcpClient.compileModule(runnerConn, MODULE_NAME, INJECT_TIMEOUT_MS);
    }

    private async runOnRunner(
        runnerConn: AccessConnection,
        subName: string,
        targetDbPath: string,
        outputPath: string,
        mode: string,
        timeoutMs: number,
        retry = true
    ): Promise<void> {
        const qualifiedProcedure = `${MODULE_NAME}.${subName}`;
        const args = [targetDbPath, outputPath, mode];

        try {
            await this.mcpClient.runVba(runnerConn, qualifiedProcedure, args, timeoutMs);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            const missingProcedure =
                /no\s+encuentra\s+el\s+procedimiento/i.test(message) ||
                /can('|')t\s+find\s+procedure/i.test(message) ||
                /could\s+not\s+find\s+the\s+procedure/i.test(message);

            if (missingProcedure && retry) {
                try { await fs.unlink(this.runnerDbPath); } catch { /* ok */ }
                await this.createRunnerDatabase();
                await this.injectModuleIntoRunner();
                await this.runOnRunner(runnerConn, subName, targetDbPath, outputPath, mode, timeoutMs, false);
                return;
            }

            throw new Error(`Exportacion VBA fallo (${message}). Usando metodo secuencial...`);
        }
    }
}
