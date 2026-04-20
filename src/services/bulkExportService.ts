import * as fs from "node:fs/promises";
import * as os from "node:os";
import * as path from "node:path";
import { AccessConnection } from "../models/types";
import { McpAccessClient } from "../mcp/mcpAccessClient";
import { BULK_EXPORT_VBA } from "../vba/bulkExportVba";

const MODULE_NAME = "SecondBrainBulkExport";

/** ms — inyección de módulo (set_code puede ser lento en bases grandes) */
const INJECT_TIMEOUT_MS = 90_000;

/** ms — ejecución del VBA completo (puede recorrer 200+ formularios) */
const RUN_TIMEOUT_MS = 900_000;   // 15 min

/** ms — borrado del módulo temporal */
const CLEANUP_TIMEOUT_MS = 20_000;

export class BulkExportService {
    constructor(private readonly mcpClient: McpAccessClient) { }

    /**
     * Inyecta el módulo VBA, ejecuta ExportToJsonFile, lee el resultado y limpia.
     * @returns metadata como objeto JS ya parseado
     */
    async runJsonExport(
        connection: AccessConnection,
        mode: string,
        timeoutMs = RUN_TIMEOUT_MS
    ): Promise<Record<string, unknown>> {
        const tempPath = path.join(os.tmpdir(), `sb_bulk_${Date.now()}.json`);

        await this.inject(connection);

        try {
            await this.callSubroutine(connection, "ExportToJsonFile", tempPath, mode, timeoutMs);
            const raw = await fs.readFile(tempPath, "utf-8");
            return JSON.parse(raw) as Record<string, unknown>;
        } finally {
            await this.cleanup(connection);
            try { await fs.unlink(tempPath); } catch { /* el archivo puede no existir si el VBA falló */ }
        }
    }

    /**
     * Inyecta el módulo VBA, ejecuta ExportToFiles a una carpeta destino y limpia.
     */
    async runFileExport(
        connection: AccessConnection,
        outputDir: string,
        mode: string,
        timeoutMs = RUN_TIMEOUT_MS
    ): Promise<void> {
        await this.inject(connection);
        try {
            await this.callSubroutine(connection, "ExportToFiles", outputDir, mode, timeoutMs);
        } finally {
            await this.cleanup(connection);
        }
    }

    // ── privados ──────────────────────────────────────────────────────────

    private async inject(connection: AccessConnection): Promise<void> {
        await this.mcpClient.setCode(
            connection,
            "module",
            MODULE_NAME,
            BULK_EXPORT_VBA,
            INJECT_TIMEOUT_MS
        );
    }

    /**
     * Llama a `MODULE_NAME.subName(arg1, mode)` via eval_vba.
     * Las rutas se escapan para que sean válidas en una cadena VBA.
     */
    private async callSubroutine(
        connection: AccessConnection,
        subName: string,
        arg1: string,
        mode: string,
        timeoutMs: number
    ): Promise<void> {
        const escapedArg1 = arg1.replace(/\\/g, "\\\\").replace(/"/g, "\\\"");
        const vbaCode = `${MODULE_NAME}.${subName} "${escapedArg1}", "${mode}"`;
        await this.mcpClient.evalVba(connection, vbaCode, timeoutMs);
    }

    private async cleanup(connection: AccessConnection): Promise<void> {
        try {
            await this.mcpClient.deleteVbaModule(connection, MODULE_NAME);
        } catch {
            // Non-fatal: el módulo puede no existir si la inyección falló
        }
    }
}
