import * as vscode from "vscode";
import * as path from "node:path";
import { AccessCategoryKey, AccessConnection } from "../models/types";
import { BulkExportService } from "./bulkExportService";

export type ExportObjectsScope =
    | { mode: "full" }
    | { mode: "category"; categoryKey: AccessCategoryKey };

export interface ExportObjectsOptions {
    onProgress?: (message: string) => void;
}

export interface ExportObjectsResult {
    outputDir: string;
    stats: Record<string, unknown>;
}

export class ExportObjectsService {
    private readonly bulkExportService: BulkExportService;

    constructor(bulkExportService: BulkExportService) {
        this.bulkExportService = bulkExportService;
    }

    async exportObjects(
        connection: AccessConnection,
        baseOutputDir: string,
        scope: ExportObjectsScope,
        options?: ExportObjectsOptions
    ): Promise<ExportObjectsResult> {
        const timestamp = new Date().toISOString().replace(/[:]/g, "-").replace(/\..+$/, "");
        const label = scope.mode === "full" ? "full" : scope.categoryKey;
        const outputDir = path.join(
            baseOutputDir,
            `export-${sanitize(connection.name)}-${label}-${timestamp}`
        );

        options?.onProgress?.(`Exportando objetos de ${connection.name} → ${outputDir}`);

        const mode = scope.mode === "full" ? "full" : scope.categoryKey;
        await this.bulkExportService.runFileExport(connection, outputDir, mode);

        options?.onProgress?.("Exportación completada");

        return {
            outputDir,
            stats: { mode, connection: connection.name, outputDir }
        };
    }
}

function sanitize(name: string): string {
    return name.replace(/[<>:"/\\|?*\x00-\x1F]/g, "_");
}
