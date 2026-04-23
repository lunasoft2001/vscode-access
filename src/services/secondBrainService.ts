import * as fs from "node:fs/promises";
import * as path from "node:path";
import { AccessCategoryKey, AccessConnection, AccessControlInfo, AccessObjectInfo, AccessPropertyInfo } from "../models/types";
import { McpAccessClient } from "../mcp/mcpAccessClient";

export type SecondBrainScope =
    | { mode: "full" }
    | { mode: "category"; categoryKey: AccessCategoryKey }
    | { mode: "object"; categoryKey: AccessCategoryKey; objectInfo: AccessObjectInfo };

interface SecondBrainMetadata {
    database: string;
    generated: string;
    source_path: string;
    columns: Array<Record<string, unknown>>;
    tables: Array<Record<string, unknown>>;
    primary_keys: Array<Record<string, unknown>>;
    foreign_keys: Array<Record<string, unknown>>;
    indexes: Array<Record<string, unknown>>;
    queries: Array<Record<string, unknown>>;
    relationships: Array<Record<string, unknown>>;
    forms: Array<Record<string, unknown>>;
    reports: Array<Record<string, unknown>>;
    macros: Array<Record<string, unknown>>;
    modules: Array<Record<string, unknown>>;
    linked_tables: Array<Record<string, unknown>>;
    startup_options: Array<Record<string, unknown>>;
    references: Array<Record<string, unknown>>;
    warnings: string[];
}

export interface SecondBrainExportResult {
    outputDir: string;
    metadataPath: string;
    vaultDir: string;
    stats: Record<string, number>;
}

export interface SecondBrainProgressEvent {
    phase: "inventory" | "object" | "write" | "done";
    message: string;
    completed?: number;
    total?: number;
}

export interface SecondBrainExportOptions {
    onProgress?: (event: SecondBrainProgressEvent) => void | Promise<void>;
    linkDensity?: "standard" | "high";
}

interface NoteDraft {
    notePath: string;
    outgoing: Set<string>;
}

export class SecondBrainService {
    private static readonly INVENTORY_TIMEOUT_MS = 300000;
    private static readonly OBJECT_TIMEOUT_MS = 180000;
    private static readonly UI_TIMEOUT_MS = 240000;
    private static readonly UI_CONTROLS_TIMEOUT_MS = 20000;
    private static readonly QUERY_TIMEOUT_MS = 15000;
    private static readonly DEGRADED_TIMEOUT_MS = 15000;
    private static readonly MCP_TIMEOUT_RETRIES = 1;
    private static readonly SLOW_OBJECT_WARNING_MS = 10000;
    private static readonly SLOW_OBJECT_REPEAT_MS = 15000;

    private timeoutDegradedMode = false;
    private _globalProgressTotal = 0;
    private _globalProgressOffset = 0;

    constructor(private readonly mcpClient: McpAccessClient, _globalStoragePath: string) {}

    async exportSecondBrain(
        connection: AccessConnection,
        baseOutputDir: string,
        scope: SecondBrainScope,
        options?: SecondBrainExportOptions
    ): Promise<SecondBrainExportResult> {
        const exportStartedAt = Date.now();
        const timestamp = new Date().toISOString().replace(/[:]/g, "-").replace(/\..+$/, "");
        const scopeLabel = this.scopeLabel(scope);
        const rootDir = path.join(baseOutputDir, `secondbrain-${sanitize(connection.name)}-${scopeLabel}-${timestamp}`);
        const vaultDir = path.join(rootDir, "db-second-brain");

        await fs.mkdir(vaultDir, { recursive: true });

        await this.reportProgress(options, {
            phase: "inventory",
            message: `Preparando exportacion en ${rootDir}`
        });

        const metadataStartedAt = Date.now();
        const metadata = await this.buildMetadata(connection, scope, options);
        await this.reportProgress(options, {
            phase: "inventory",
            message: `Metadata lista en ${formatDuration(Date.now() - metadataStartedAt)}`
        });

        const metadataPath = path.join(rootDir, "metadata.json");
        const metadataWriteStartedAt = Date.now();
        await fs.writeFile(metadataPath, JSON.stringify(metadata, null, 2), "utf-8");

        await this.reportProgress(options, {
            phase: "write",
            message: `metadata.json escrito en ${formatDuration(Date.now() - metadataWriteStartedAt)}`
        });

        const vaultStartedAt = Date.now();
        await this.generateVault(vaultDir, metadata, scope, options);

        await this.reportProgress(options, {
            phase: "write",
            message: `Vault generado en ${formatDuration(Date.now() - vaultStartedAt)}`
        });

        await this.reportProgress(options, {
            phase: "done",
            message: `SecondBrain generado correctamente en ${formatDuration(Date.now() - exportStartedAt)}`
        });

        return {
            outputDir: rootDir,
            metadataPath,
            vaultDir,
            stats: this.computeStats(metadata)
        };
    }

    private async buildMetadata(
        connection: AccessConnection,
        scope: SecondBrainScope,
        options?: SecondBrainExportOptions
    ): Promise<SecondBrainMetadata> {
        return this.buildMetadataSequential(connection, scope, options);
    }

    /** Convierte el JSON plano del VBA al tipo SecondBrainMetadata */
    private rawToMetadata(
        raw: Record<string, unknown>,
        connection: AccessConnection
    ): SecondBrainMetadata {
        const toArr = (v: unknown): Array<Record<string, unknown>> =>
            Array.isArray(v) ? (v as Array<Record<string, unknown>>) : [];

        const metadata: SecondBrainMetadata = {
            database: String(raw.database ?? path.basename(connection.dbPath, path.extname(connection.dbPath))),
            generated: String(raw.generated ?? new Date().toISOString()),
            source_path: String(raw.source_path ?? connection.dbPath),
            columns: toArr(raw.columns),
            tables: toArr(raw.tables),
            primary_keys: toArr(raw.primary_keys),
            foreign_keys: toArr(raw.foreign_keys),
            indexes: toArr(raw.indexes),
            queries: toArr(raw.queries),
            relationships: toArr(raw.relationships),
            forms: toArr(raw.forms),
            reports: toArr(raw.reports),
            macros: toArr(raw.macros),
            modules: toArr(raw.modules),
            linked_tables: toArr(raw.linked_tables),
            startup_options: toArr(raw.startup_options),
            references: toArr(raw.references),
            warnings: Array.isArray(raw.warnings) ? raw.warnings.map(String) : []
        };

        // Derivar foreign_keys si el VBA no las calcula (compat con versiones anteriores)
        if (metadata.foreign_keys.length === 0 && metadata.relationships.length > 0) {
            this.deriveForeignKeysFromRelationships(metadata);
        }

        return metadata;
    }

    /** Implementación original secuencial — usada como fallback y para scope "object" */
    private async buildMetadataSequential(
        connection: AccessConnection,
        scope: SecondBrainScope,
        options?: SecondBrainExportOptions
    ): Promise<SecondBrainMetadata> {
        const metadata: SecondBrainMetadata = {
            database: path.basename(connection.dbPath, path.extname(connection.dbPath)),
            generated: new Date().toISOString(),
            source_path: connection.dbPath,
            columns: [],
            tables: [],
            primary_keys: [],
            foreign_keys: [],
            indexes: [],
            queries: [],
            relationships: [],
            forms: [],
            reports: [],
            macros: [],
            modules: [],
            linked_tables: [],
            startup_options: [],
            references: [],
            warnings: []
        };

        const includeAll = scope.mode === "full";

        if (includeAll) {
            // Pre-fetch object lists to compute global total for progress bar
            const toolCategoriesToPreload: AccessCategoryKey[] = ["tables", "queries", "forms", "reports", "macros", "modules"];
            const preloadedLists = new Map<AccessCategoryKey, import("../models/types").AccessObjectInfo[]>();
            await this.reportProgress(options, { phase: "inventory", message: "Preparando inventario de objetos..." });
            for (const cat of toolCategoriesToPreload) {
                const toolType = toolTypeForCategory(cat);
                if (toolType) {
                    const items = await this.mcpClient.listObjects(connection, toolType, SecondBrainService.INVENTORY_TIMEOUT_MS);
                    preloadedLists.set(cat, items);
                }
            }
            this._globalProgressTotal = [...preloadedLists.values()].reduce((sum, list) => sum + list.length, 0);
            this._globalProgressOffset = 0;

            await this.collectCategory(connection, metadata, "tables", options, preloadedLists.get("tables"));
            await this.collectCategory(connection, metadata, "queries", options, preloadedLists.get("queries"));
            await this.collectCategory(connection, metadata, "forms", options, preloadedLists.get("forms"));
            await this.collectCategory(connection, metadata, "reports", options, preloadedLists.get("reports"));
            await this.collectCategory(connection, metadata, "macros", options, preloadedLists.get("macros"));
            await this.collectCategory(connection, metadata, "modules", options, preloadedLists.get("modules"));
            await this.collectCategory(connection, metadata, "relationships", options);
            await this.collectCategory(connection, metadata, "references", options);
            await this.collectLinkedAndStartup(connection, metadata, options);
            this.deriveForeignKeysFromRelationships(metadata);
            return metadata;
        }

        if (scope.mode === "category") {
            await this.collectCategory(connection, metadata, scope.categoryKey, options);
            if (scope.categoryKey === "tables" || scope.categoryKey === "relationships") {
                await this.collectLinkedAndStartup(connection, metadata, options);
                this.deriveForeignKeysFromRelationships(metadata);
            }
            return metadata;
        }

        await this.collectSingleObject(connection, metadata, scope.categoryKey, scope.objectInfo, options);
        if (scope.categoryKey === "tables" || scope.categoryKey === "relationships") {
            await this.collectLinkedAndStartup(connection, metadata, options);
            this.deriveForeignKeysFromRelationships(metadata);
        }

        return metadata;
    }

    private async collectCategory(
        connection: AccessConnection,
        metadata: SecondBrainMetadata,
        categoryKey: AccessCategoryKey,
        options?: SecondBrainExportOptions,
        preloadedObjects?: import("../models/types").AccessObjectInfo[]
    ): Promise<void> {
        if (categoryKey === "relationships") {
            await this.reportProgress(options, {
                phase: "inventory",
                message: "Cargando relaciones"
            });
            const relationships = await this.mcpClient.listRelationships(connection, SecondBrainService.INVENTORY_TIMEOUT_MS);
            for (const rel of relationships) {
                metadata.relationships.push(rel.metadata ?? { name: rel.name });
            }
            return;
        }

        if (categoryKey === "references") {
            await this.reportProgress(options, {
                phase: "inventory",
                message: "Cargando referencias VBA"
            });
            const references = await this.mcpClient.listReferences(connection, SecondBrainService.INVENTORY_TIMEOUT_MS);
            for (const ref of references) {
                metadata.references.push(ref.metadata ?? { name: ref.name });
            }
            return;
        }

        const toolType = toolTypeForCategory(categoryKey);
        if (!toolType) {
            return;
        }

        await this.reportProgress(options, {
            phase: "inventory",
            message: `Listando ${categoryKey}`
        });
        const objects = preloadedObjects ?? await this.mcpClient.listObjects(connection, toolType, SecondBrainService.INVENTORY_TIMEOUT_MS);
        const total = objects.length;
        const globalOffset = this._globalProgressOffset;
        const useGlobal = this._globalProgressTotal > 0;
        let completed = 0;
        for (const objectInfo of objects) {
            completed += 1;
            const objectStartedAt = Date.now();
            await this.reportProgress(options, {
                phase: "object",
                message: `${categoryKey}: ${objectInfo.name}`,
                completed: useGlobal ? globalOffset + completed : completed,
                total: useGlobal ? this._globalProgressTotal : total
            });

            const stopSlowObjectWarnings = this.startSlowObjectWarnings(
                options,
                categoryKey,
                objectInfo.name,
                completed,
                total,
                objectStartedAt
            );

            try {
                await this.collectSingleObject(connection, metadata, categoryKey, objectInfo, options);
                const objectDurationMs = Date.now() - objectStartedAt;
                if (objectDurationMs >= SecondBrainService.SLOW_OBJECT_WARNING_MS) {
                    await this.reportProgress(options, {
                        phase: "inventory",
                        message: `${categoryKey}: ${objectInfo.name} completado en ${formatDuration(objectDurationMs)} (${completed}/${total})`
                    });
                }
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                metadata.warnings.push(`No se pudo exportar ${categoryKey}:${objectInfo.name} -> ${message}`);
                await this.reportProgress(options, {
                    phase: "inventory",
                    message: `${categoryKey}: ${objectInfo.name} fallo tras ${formatDuration(Date.now() - objectStartedAt)} -> ${message}`
                });
            } finally {
                stopSlowObjectWarnings();
            }
        }

        if (useGlobal) {
            this._globalProgressOffset += total;
        }
    }

    private startSlowObjectWarnings(
        options: SecondBrainExportOptions | undefined,
        categoryKey: AccessCategoryKey,
        objectName: string,
        completed: number,
        total: number,
        startedAt: number
    ): () => void {
        let timer: ReturnType<typeof setTimeout> | undefined;
        let warningCount = 0;
        let stopped = false;

        const schedule = (delayMs: number): void => {
            timer = setTimeout(() => {
                if (stopped) {
                    return;
                }

                warningCount += 1;
                void this.reportProgress(options, {
                    phase: "inventory",
                    message: `${categoryKey}: ${objectName} sigue en curso tras ${formatDuration(Date.now() - startedAt)} (${completed}/${total}, aviso ${warningCount})`
                });

                schedule(SecondBrainService.SLOW_OBJECT_REPEAT_MS);
            }, delayMs);
        };

        schedule(SecondBrainService.SLOW_OBJECT_WARNING_MS);

        return () => {
            stopped = true;
            if (timer) {
                clearTimeout(timer);
            }
        };
    }

    private async recoverAfterTimeout(
        options: SecondBrainExportOptions | undefined,
        contextLabel: string,
        metadata: SecondBrainMetadata
    ): Promise<void> {
        await this.reportProgress(options, {
            phase: "inventory",
            message: `Timeout en ${contextLabel}. Reconectando MCP antes de continuar.`
        });

        try {
            await this.mcpClient.reconnect();
            await this.reportProgress(options, {
                phase: "inventory",
                message: `MCP reconectado tras timeout en ${contextLabel}`
            });
        } catch (reconnectError) {
            const reconnectMessage = reconnectError instanceof Error ? reconnectError.message : String(reconnectError);
            metadata.warnings.push(`No se pudo reconectar MCP tras timeout en ${contextLabel} -> ${reconnectMessage}`);
        }
    }

    private async runWithTimeoutRetry<T>(
        options: SecondBrainExportOptions | undefined,
        metadata: SecondBrainMetadata,
        contextLabel: string,
        operation: () => Promise<T>
    ): Promise<T> {
        let attempt = 0;
        while (true) {
            try {
                return await operation();
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                if (!isLikelyTimeoutError(message) || attempt >= SecondBrainService.MCP_TIMEOUT_RETRIES) {
                    throw error;
                }

                if (!this.timeoutDegradedMode) {
                    this.timeoutDegradedMode = true;
                    await this.reportProgress(options, {
                        phase: "inventory",
                        message: `Modo timeout agresivo activado (${SecondBrainService.DEGRADED_TIMEOUT_MS} ms) tras timeout en ${contextLabel}.`
                    });
                }

                attempt += 1;
                await this.reportProgress(options, {
                    phase: "inventory",
                    message: `Timeout en ${contextLabel}. Reintentando (${attempt}/${SecondBrainService.MCP_TIMEOUT_RETRIES}) tras reconexion MCP.`
                });
                await this.recoverAfterTimeout(options, contextLabel, metadata);
            }
        }
    }

    private resolveTimeout(baseTimeoutMs: number): number {
        if (!this.timeoutDegradedMode) {
            return baseTimeoutMs;
        }
        return Math.min(baseTimeoutMs, SecondBrainService.DEGRADED_TIMEOUT_MS);
    }

    private async collectSingleObject(
        connection: AccessConnection,
        metadata: SecondBrainMetadata,
        categoryKey: AccessCategoryKey,
        objectInfo: AccessObjectInfo,
        options?: SecondBrainExportOptions
    ): Promise<void> {
        const objectName = objectInfo.name;

        if (categoryKey === "tables") {
            const tableDoc = await this.mcpClient.getObjectDocument(
                connection,
                "tables",
                objectName,
                objectInfo.metadata,
                SecondBrainService.OBJECT_TIMEOUT_MS
            );
            const parsed = safeJson(tableDoc.content);
            const fields = Array.isArray(parsed?.fields) ? parsed.fields : [];
            const indexes = await this.mcpClient.getTableIndexes(connection, objectName, SecondBrainService.OBJECT_TIMEOUT_MS);

            metadata.tables.push({
                table_name: objectName,
                table_type: parsed?.is_linked ? "LINKED" : "LOCAL",
                record_count: parsed?.record_count,
                source_table: parsed?.source_table
            });

            fields.forEach((field: any, idx: number) => {
                metadata.columns.push({
                    table_schema: "dbo",
                    table_name: objectName,
                    column_name: fixEncoding(field?.name),
                    data_type: field?.type,
                    is_nullable: field?.required ? "NO" : "YES",
                    character_maximum_length: field?.size,
                    ordinal_position: idx + 1
                });
            });

            for (const index of indexes) {
                const indexFields = Array.isArray(index.fields) ? index.fields : [];
                for (const field of indexFields) {
                    metadata.indexes.push({
                        table_name: objectName,
                        index_name: index.name,
                        field_name: field.name,
                        is_unique: Boolean(index.unique),
                        is_primary: Boolean(index.primary),
                        sort_order: field.order ?? "asc"
                    });

                    if (index.primary && field.name) {
                        metadata.primary_keys.push({
                            table_schema: "dbo",
                            table_name: objectName,
                            column_name: field.name
                        });
                    }
                }
            }
            return;
        }

        if (categoryKey === "queries") {
            let sql = "";
            try {
                sql = fixEncoding(await this.runWithTimeoutRetry(
                    options,
                    metadata,
                    `${categoryKey}:${objectName}:sql`,
                    async () => this.mcpClient.getQuerySql(connection, objectName, this.resolveTimeout(SecondBrainService.QUERY_TIMEOUT_MS))
                )) ?? "";
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                metadata.warnings.push(`SQL no disponible para query:${objectName} -> ${message}`);
            }

            metadata.queries.push({
                name: objectName,
                type: inferQueryType(sql),
                sql
            });
            return;
        }

        if (categoryKey === "forms" || categoryKey === "reports") {
            const objectType = categoryKey === "forms" ? "form" : "report";
            let controls: AccessControlInfo[] = [];
            let properties: AccessPropertyInfo[] = [];
            let docContent = "";

            await this.reportProgress(options, {
                phase: "inventory",
                message: `${categoryKey}: ${objectName} -> controles`
            });
            try {
                controls = await this.runWithTimeoutRetry(
                    options,
                    metadata,
                    `${categoryKey}:${objectName}:controles`,
                    async () => this.mcpClient.getControls(
                        connection,
                        objectType,
                        objectName,
                        this.resolveTimeout(SecondBrainService.UI_CONTROLS_TIMEOUT_MS)
                    )
                );
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                metadata.warnings.push(`Controles no disponibles para ${categoryKey}:${objectName} -> ${message}`);
            }

            await this.reportProgress(options, {
                phase: "inventory",
                message: `${categoryKey}: ${objectName} -> propiedades`
            });
            try {
                properties = await this.runWithTimeoutRetry(
                    options,
                    metadata,
                    `${categoryKey}:${objectName}:propiedades`,
                    async () => this.mcpClient.getFormReportProperties(
                        connection,
                        objectType,
                        objectName,
                        this.resolveTimeout(SecondBrainService.UI_TIMEOUT_MS)
                    )
                );
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                metadata.warnings.push(`Propiedades no disponibles para ${categoryKey}:${objectName} -> ${message}`);
            }

            const recordSource = fixEncoding(properties.find((item) => item.name.toLowerCase() === "recordsource")?.value ?? "") ?? "";

            await this.reportProgress(options, {
                phase: "inventory",
                message: `${categoryKey}: ${objectName} -> codigo`
            });
            try {
                const doc = await this.runWithTimeoutRetry(
                    options,
                    metadata,
                    `${categoryKey}:${objectName}:codigo`,
                    async () => this.mcpClient.getObjectDocument(
                        connection,
                        categoryKey,
                        objectName,
                        objectInfo.metadata,
                        this.resolveTimeout(SecondBrainService.UI_TIMEOUT_MS)
                    )
                );
                docContent = doc.content;
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                metadata.warnings.push(`Codigo no disponible para ${categoryKey}:${objectName} -> ${message}`);
            }

            const propVal = (key: string): string | undefined =>
                properties.find((p) => p.name.toLowerCase() === key)?.value;
            const uiObj = {
                name: objectName,
                record_source: recordSource,
                allow_edits: propVal("allowedits"),
                allow_additions: propVal("allowadditions"),
                allow_deletions: propVal("allowdeletions"),
                default_view: propVal("defaultview"),
                modal: propVal("modal"),
                pop_up: propVal("popup"),
                controls: controls.map((ctrl) => ({
                    name: ctrl.name,
                    control_type: ctrl.type_name,
                    control_source: ctrl.control_source,
                    caption: fixEncoding(ctrl.caption),
                    source_object: ctrl.source_object
                })),
                code: docContent,
                procedures: extractVbaProcedures(docContent)
            };

            if (categoryKey === "forms") {
                metadata.forms.push(uiObj);
            } else {
                metadata.reports.push(uiObj);
            }
            return;
        }

        if (categoryKey === "macros") {
            const doc = await this.mcpClient.getObjectDocument(
                connection,
                "macros",
                objectName,
                objectInfo.metadata,
                SecondBrainService.OBJECT_TIMEOUT_MS
            );
            metadata.macros.push({
                name: objectName,
                actions: doc.content
            });
            return;
        }

        if (categoryKey === "modules") {
            const doc = await this.mcpClient.getObjectDocument(
                connection,
                "modules",
                objectName,
                objectInfo.metadata,
                SecondBrainService.OBJECT_TIMEOUT_MS
            );
            metadata.modules.push({
                name: objectName,
                kind: "standard",
                code: doc.content,
                procedures: extractVbaProcedures(doc.content)
            });
        }
    }

    private async collectLinkedAndStartup(
        connection: AccessConnection,
        metadata: SecondBrainMetadata,
        options?: SecondBrainExportOptions
    ): Promise<void> {
        await this.reportProgress(options, {
            phase: "inventory",
            message: "Cargando tablas vinculadas"
        });
        const linkedTables = await this.mcpClient.listLinkedTables(connection, SecondBrainService.INVENTORY_TIMEOUT_MS);
        for (const linked of linkedTables) {
            const linkedMeta = linked.metadata ?? {};
            metadata.linked_tables.push({
                table_name: linked.name,
                source_table: linkedMeta.source_table,
                connect: linkedMeta.connect_string,
                is_odbc: linkedMeta.is_odbc
            });
        }

        await this.reportProgress(options, {
            phase: "inventory",
            message: "Cargando opciones de inicio"
        });
        const startupOptions = await this.mcpClient.listStartupOptions(connection, SecondBrainService.INVENTORY_TIMEOUT_MS);
        for (const option of startupOptions) {
            const optionMeta = option.metadata ?? {};
            metadata.startup_options.push({
                option: option.name,
                value: optionMeta.value,
                source: optionMeta.source
            });
        }
    }

    private deriveForeignKeysFromRelationships(metadata: SecondBrainMetadata): void {
        for (const rel of metadata.relationships) {
            const fields = Array.isArray(rel.fields) ? rel.fields : [];
            const fkTable = String(rel.foreign_table ?? "");
            const refTable = String(rel.table ?? "");

            for (const field of fields as Array<Record<string, unknown>>) {
                metadata.foreign_keys.push({
                    fk_name: rel.name,
                    fk_schema: "dbo",
                    fk_table: fkTable,
                    fk_column: field.foreign,
                    ref_schema: "dbo",
                    ref_table: refTable,
                    ref_column: field.local,
                    update_rule: null,
                    delete_rule: null
                });
            }
        }
    }

    private async generateVault(
        vaultDir: string,
        metadata: SecondBrainMetadata,
        scope: SecondBrainScope,
        options?: SecondBrainExportOptions
    ): Promise<void> {
        const vaultStartedAt = Date.now();
        const useHighDensity = options?.linkDensity === "high";
        const folders = {
            tables: path.join(vaultDir, "tables"),
            queries: path.join(vaultDir, "queries"),
            relationships: path.join(vaultDir, "relationships"),
            forms: path.join(vaultDir, "forms"),
            reports: path.join(vaultDir, "reports"),
            macros: path.join(vaultDir, "macros"),
            modules: path.join(vaultDir, "modules"),
            mocs: path.join(vaultDir, "mocs"),
            linkedTables: path.join(vaultDir, "linked-tables"),
            startup: path.join(vaultDir, "startup")
        };

        await Promise.all(Object.values(folders).map((folder) => fs.mkdir(folder, { recursive: true })));

        await this.reportProgress(options, {
            phase: "write",
            message: "Escribiendo notas del vault"
        });

        const indexingStartedAt = Date.now();

        const columnsByTable = groupBy(metadata.columns, (item) => String(item.table_name ?? ""));
        const pksByTable = groupBy(metadata.primary_keys, (item) => String(item.table_name ?? ""));
        const idxByTable = groupBy(metadata.indexes, (item) => String(item.table_name ?? ""));
        const relationshipsByTable = new Map<string, Set<string>>();
        for (const rel of metadata.relationships) {
            const relPath = `relationships/${sanitize(String(rel.name ?? "(sin nombre)"))}`;
            const oneSide = String(rel.table ?? "");
            const manySide = String(rel.foreign_table ?? "");

            addToSetMap(relationshipsByTable, oneSide, relPath);
            addToSetMap(relationshipsByTable, manySide, relPath);
        }

        const foreignKeysFromTable = groupBy(metadata.foreign_keys, (item) => String(item.fk_table ?? ""));
        const foreignKeysToTable = groupBy(metadata.foreign_keys, (item) => String(item.ref_table ?? ""));

        const tableTargets = new Map(metadata.tables.map((table) => [String(table.table_name ?? ""), `tables/${sanitize(String(table.table_name ?? ""))}`] as const));
        const queryTargets = new Map(metadata.queries.map((query) => [String(query.name ?? ""), `queries/${sanitize(String(query.name ?? ""))}`] as const));
        const formTargets = new Map(metadata.forms.map((form) => [String(form.name ?? ""), `forms/${sanitize(String(form.name ?? ""))}`] as const));
        const reportTargets = new Map(metadata.reports.map((report) => [String(report.name ?? ""), `reports/${sanitize(String(report.name ?? ""))}`] as const));
        const macroTargets = new Map(metadata.macros.map((macro) => [String(macro.name ?? ""), `macros/${sanitize(String(macro.name ?? ""))}`] as const));
        const moduleTargets = new Map(metadata.modules.map((module) => [String(module.name ?? ""), `modules/${sanitize(String(module.name ?? ""))}`] as const));

        const domainEntries: Array<{ name: string; notePath: string; kind: string }> = [
            ...metadata.tables.map((item) => ({ name: String(item.table_name ?? ""), notePath: `tables/${sanitize(String(item.table_name ?? ""))}`, kind: "table" })),
            ...metadata.queries.map((item) => ({ name: String(item.name ?? ""), notePath: `queries/${sanitize(String(item.name ?? ""))}`, kind: "query" })),
            ...metadata.forms.map((item) => ({ name: String(item.name ?? ""), notePath: `forms/${sanitize(String(item.name ?? ""))}`, kind: "form" })),
            ...metadata.reports.map((item) => ({ name: String(item.name ?? ""), notePath: `reports/${sanitize(String(item.name ?? ""))}`, kind: "report" })),
            ...metadata.macros.map((item) => ({ name: String(item.name ?? ""), notePath: `macros/${sanitize(String(item.name ?? ""))}`, kind: "macro" })),
            ...metadata.modules.map((item) => ({ name: String(item.name ?? ""), notePath: `modules/${sanitize(String(item.name ?? ""))}`, kind: "module" }))
        ];

        const mocGroups = new Map<string, Array<{ name: string; notePath: string; kind: string }>>();
        const mocByNotePath = new Map<string, string>();

        if (useHighDensity) {
            const groupedDomains = groupBy(domainEntries, (entry) => inferDomainKey(entry.name));
            for (const [domainKey, entries] of groupedDomains.entries()) {
                if (!domainKey) {
                    continue;
                }
                const kinds = new Set(entries.map((item) => item.kind));
                if (entries.length < 4 || kinds.size < 2) {
                    continue;
                }

                const mocPath = `mocs/${sanitize(domainKey)}`;
                mocGroups.set(mocPath, entries);
                for (const entry of entries) {
                    mocByNotePath.set(entry.notePath, mocPath);
                }
            }
        }

        const allNoteTargets = new Set<string>([
            ...tableTargets.values(),
            ...queryTargets.values(),
            ...formTargets.values(),
            ...reportTargets.values(),
            ...macroTargets.values(),
            ...moduleTargets.values(),
            ...mocGroups.keys(),
            ...metadata.relationships.map((rel) => `relationships/${sanitize(String(rel.name ?? "(sin nombre)"))}`),
            ...metadata.linked_tables.map((linked) => `linked-tables/${sanitize(String(linked.table_name ?? ""))}`),
            "startup/startup-options",
            "_index",
            "_overview",
            "_health",
            "_dependencies",
            "_guide",
            "_entrypoints",
            "_critical-objects",
            "_known-issues",
            "_reading-order",
            "_ai-prompt"
        ]);

        const noteDrafts = new Map<string, NoteDraft>();
        const bodyWrites: Array<Promise<void>> = [];

        const registerNote = (notePath: string, lines: string[], outgoing: Iterable<string> = []): void => {
            const validOutgoing = new Set<string>();
            for (const target of outgoing) {
                if (target && target !== notePath && allNoteTargets.has(target)) {
                    validOutgoing.add(target);
                }
            }
            noteDrafts.set(notePath, { notePath, outgoing: validOutgoing });
            bodyWrites.push(fs.writeFile(path.join(vaultDir, `${notePath}.md`), lines.join("\n"), "utf-8"));
        };

        const queryDependenciesByQuery = new Map<string, Set<string>>();
        for (const query of metadata.queries) {
            const queryName = String(query.name ?? "");
            const sql = String(query.sql ?? "");
            const deps = new Set<string>();
            for (const candidate of extractSqlObjectCandidates(sql)) {
                const tableTarget = tableTargets.get(candidate);
                if (tableTarget) {
                    deps.add(tableTarget);
                    continue;
                }
                const queryTarget = queryTargets.get(candidate);
                if (queryTarget && queryTarget !== queryTargets.get(queryName)) {
                    deps.add(queryTarget);
                }
            }
            queryDependenciesByQuery.set(queryName, deps);
        }

        const queriesByTable = new Map<string, string[]>();
        for (const query of metadata.queries) {
            const queryName = String(query.name ?? "");
            const queryPath = queryTargets.get(queryName);
            if (!queryPath) {
                continue;
            }

            for (const dependency of queryDependenciesByQuery.get(queryName) ?? []) {
                if (!dependency.startsWith("tables/")) {
                    continue;
                }

                const queries = queriesByTable.get(dependency) ?? [];
                queries.push(queryPath);
                queriesByTable.set(dependency, queries);
            }
        }

        const formsByRecordSource = new Map<string, string[]>();
        for (const form of metadata.forms) {
            const formName = String(form.name ?? "");
            const formPath = formTargets.get(formName);
            if (!formPath) {
                continue;
            }

            const recordSource = normalizeIdentifier(String(form.record_source ?? ""));
            if (!recordSource) {
                continue;
            }

            const forms = formsByRecordSource.get(recordSource) ?? [];
            forms.push(formPath);
            formsByRecordSource.set(recordSource, forms);
        }

        await this.reportProgress(options, {
            phase: "write",
            message: `Indices del vault listos en ${formatDuration(Date.now() - indexingStartedAt)}`
        });

        for (const table of metadata.tables) {
            const name = String(table.table_name ?? "");
            const cols = columnsByTable.get(name) ?? [];
            const pks = (pksByTable.get(name) ?? []).map((item) => String(item.column_name ?? ""));
            const idx = idxByTable.get(name) ?? [];
            const tableNotePath = `tables/${sanitize(name)}`;

            const outgoing = new Set(relationshipsByTable.get(name) ?? []);
            const mocPath = mocByNotePath.get(tableNotePath);
            if (mocPath) {
                outgoing.add(mocPath);
            }

            const lines = [
                `# Table: ${name}`,
                "",
                `- Type: ${table.table_type ?? ""}`,
                table.record_count !== undefined ? `- Records: ${table.record_count}` : "",
                "",
                "## Columns",
                "",
                "| # | Name | Type | Nullable | Size |",
                "|---:|---|---|---|---:|"
            ].filter(Boolean);

            cols.forEach((col, i) => {
                lines.push(
                    `| ${i + 1} | ${col.column_name ?? ""} | ${col.data_type ?? ""} | ${col.is_nullable ?? ""} | ${col.character_maximum_length ?? ""} |`
                );
            });

            lines.push("", `## Primary Key`, "", pks.length > 0 ? `- ${pks.join(", ")}` : "- None", "", "## Indexes", "");
            if (idx.length === 0) {
                lines.push("- None");
            } else {
                lines.push("| Index | Field | Unique | Primary | Order |", "|---|---|---|---|---|");
                idx.forEach((item) => {
                    lines.push(
                        `| ${item.index_name ?? ""} | ${item.field_name ?? ""} | ${item.is_unique ?? false} | ${item.is_primary ?? false} | ${item.sort_order ?? ""} |`
                    );
                });
            }

            const fksFromTable = foreignKeysFromTable.get(name) ?? [];
            const fksToTable = foreignKeysToTable.get(name) ?? [];
            if (fksFromTable.length > 0) {
                lines.push("", "## Foreign Keys", "", "| FK Name | Column | Ref Table | Ref Column |", "|---|---|---|---|");
                for (const fk of fksFromTable) {
                    const refTarget = tableTargets.get(String(fk.ref_table ?? ""));
                    if (refTarget) { outgoing.add(refTarget); }
                    lines.push(`| ${fk.fk_name ?? ""} | ${fk.fk_column ?? ""} | ${refTarget ? wiki(refTarget) : String(fk.ref_table ?? "")} | ${fk.ref_column ?? ""} |`);
                }
            }
            if (fksToTable.length > 0) {
                lines.push("", "## Referenciado por FK", "", "| FK Name | From Table | Via Column |", "|---|---|---|");
                for (const fk of fksToTable) {
                    const fkTarget = tableTargets.get(String(fk.fk_table ?? ""));
                    lines.push(`| ${fk.fk_name ?? ""} | ${fkTarget ? wiki(fkTarget) : String(fk.fk_table ?? "")} | ${fk.fk_column ?? ""} |`);
                }
            }

            registerNote(tableNotePath, lines, outgoing);
        }

        for (const query of metadata.queries) {
            const name = String(query.name ?? "");
            const sql = String(query.sql ?? "");
            const queryNotePath = `queries/${sanitize(name)}`;
            const outgoing = queryDependenciesByQuery.get(name) ?? new Set<string>();
            const mocPath = mocByNotePath.get(queryNotePath);
            if (mocPath) {
                outgoing.add(mocPath);
            }

            const dependsOn = Array.from(outgoing).sort();
            const content = [
                `# Query: ${name}`,
                "",
                `- Type: ${query.type ?? "unknown"}`,
                "",
                "## Depends On",
                "",
                ...(dependsOn.length > 0 ? dependsOn.map((target) => `- ${wiki(target)}`) : ["- None"]),
                "",
                "```sql",
                sql || "-- SQL not available",
                "```",
                ""
            ].join("\n");
            registerNote(queryNotePath, content.split("\n"), outgoing);
        }

        for (const rel of metadata.relationships) {
            const name = String(rel.name ?? "(sin nombre)");
            const oneSide = String(rel.table ?? "");
            const manySide = String(rel.foreign_table ?? "");
            const outgoing = new Set<string>();
            const oneTarget = tableTargets.get(oneSide);
            const manyTarget = tableTargets.get(manySide);
            if (oneTarget) {
                outgoing.add(oneTarget);
            }
            if (manyTarget) {
                outgoing.add(manyTarget);
            }
            const content = [
                `# Relationship: ${name}`,
                "",
                `- One side: ${oneTarget ? wiki(oneTarget) : oneSide}`,
                `- Many side: ${manyTarget ? wiki(manyTarget) : manySide}`,
                ""
            ].join("\n");
            registerNote(`relationships/${sanitize(name)}`, content.split("\n"), outgoing);
        }

        for (const form of metadata.forms) {
            const name = String(form.name ?? "");
            const formNotePath = `forms/${sanitize(name)}`;
            const outgoing = this.collectUiDependencies(form, tableTargets, queryTargets);
            for (const link of extractVbaObjectLinks(
                String(form.code ?? ""), formTargets, reportTargets, queryTargets, macroTargets, tableTargets, moduleTargets
            )) { outgoing.add(link); }
            for (const ctrl of Array.isArray(form.controls) ? (form.controls as Array<Record<string, unknown>>) : []) {
                const typeName = String(ctrl.control_type ?? "").toLowerCase();
                if (typeName.includes("subform") || typeName.includes("subreport")) {
                    const sourceObj = normalizeIdentifier(String(ctrl.source_object ?? ""));
                    if (sourceObj) {
                        const subTarget = formTargets.get(sourceObj) ?? reportTargets.get(sourceObj);
                        if (subTarget) { outgoing.add(subTarget); }
                    }
                }
            }
            const mocPath = mocByNotePath.get(formNotePath);
            if (mocPath) {
                outgoing.add(mocPath);
            }
            registerNote(formNotePath, this.buildUiNote("Form", form, outgoing).split("\n"), outgoing);
        }

        for (const report of metadata.reports) {
            const name = String(report.name ?? "");
            const reportNotePath = `reports/${sanitize(name)}`;
            const outgoing = this.collectUiDependencies(report, tableTargets, queryTargets);
            for (const link of extractVbaObjectLinks(
                String(report.code ?? ""), formTargets, reportTargets, queryTargets, macroTargets, tableTargets, moduleTargets
            )) { outgoing.add(link); }
            for (const ctrl of Array.isArray(report.controls) ? (report.controls as Array<Record<string, unknown>>) : []) {
                const typeName = String(ctrl.control_type ?? "").toLowerCase();
                if (typeName.includes("subform") || typeName.includes("subreport")) {
                    const sourceObj = normalizeIdentifier(String(ctrl.source_object ?? ""));
                    if (sourceObj) {
                        const subTarget = reportTargets.get(sourceObj) ?? formTargets.get(sourceObj);
                        if (subTarget) { outgoing.add(subTarget); }
                    }
                }
            }
            const mocPath = mocByNotePath.get(reportNotePath);
            if (mocPath) {
                outgoing.add(mocPath);
            }
            registerNote(reportNotePath, this.buildUiNote("Report", report, outgoing).split("\n"), outgoing);
        }

        for (const macro of metadata.macros) {
            const name = String(macro.name ?? "");
            const outgoing = extractVbaObjectLinks(
                String(macro.actions ?? ""),
                formTargets,
                reportTargets,
                queryTargets,
                macroTargets,
                tableTargets,
                moduleTargets
            );
            const macroNotePath = `macros/${sanitize(name)}`;
            const mocPath = mocByNotePath.get(macroNotePath);
            if (mocPath) {
                outgoing.add(mocPath);
            }
            const content = [
                `# Macro: ${name}`,
                "",
                "```text",
                String(macro.actions ?? ""),
                "```",
                ""
            ].join("\n");
            registerNote(macroNotePath, content.split("\n"), outgoing);
        }

        for (const module of metadata.modules) {
            const name = String(module.name ?? "");
            const outgoing = extractVbaObjectLinks(
                String(module.code ?? ""),
                formTargets,
                reportTargets,
                queryTargets,
                macroTargets,
                tableTargets,
                moduleTargets
            );
            const moduleNotePath = `modules/${sanitize(name)}`;
            const mocPath = mocByNotePath.get(moduleNotePath);
            if (mocPath) {
                outgoing.add(mocPath);
            }
            const procs = Array.isArray(module.procedures)
                ? (module.procedures as Array<{ kind: string; name: string; visibility: string }>)
                : [];
            const codeLines = countVbaCodeLines(String(module.code ?? ""));
            const procSection = procs.length > 0
                ? [
                    `## Procedures (${procs.length})`,
                    "",
                    "| Kind | Name | Visibility |",
                    "|---|---|---|",
                    ...procs.map((p) => `| ${p.kind} | ${p.name} | ${p.visibility} |`),
                    ""
                ]
                : [];
            const content = [
                `# Module: ${name}`,
                "",
                `- Procedures: ${procs.length}`,
                `- Code lines: ${codeLines}`,
                "",
                ...procSection,
                "## Code",
                "",
                "```vb",
                String(module.code ?? ""),
                "```",
                ""
            ].join("\n");
            registerNote(moduleNotePath, content.split("\n"), outgoing);
        }

        for (const [mocPath, entries] of mocGroups.entries()) {
            const groupedByKind = groupBy(entries, (entry) => entry.kind);
            const orderedKinds = ["table", "query", "form", "report", "macro", "module"];
            const lines: string[] = [
                `# MOC: ${path.basename(mocPath)}`,
                "",
                "Mapa de conocimiento generado automaticamente para objetos relacionados por dominio de nombres.",
                ""
            ];

            for (const kind of orderedKinds) {
                const kindItems = (groupedByKind.get(kind) ?? []).sort((a, b) => a.name.localeCompare(b.name));
                if (kindItems.length === 0) {
                    continue;
                }
                lines.push(`## ${kind[0].toUpperCase()}${kind.slice(1)}s`, "");
                for (const item of kindItems) {
                    lines.push(`- ${wiki(item.notePath)}`);
                }
                lines.push("");
            }

            registerNote(mocPath, lines, entries.map((entry) => entry.notePath));
        }

        for (const linked of metadata.linked_tables) {
            const name = String(linked.table_name ?? "");
            const sourceTable = String(linked.source_table ?? "");
            const outgoing = new Set<string>();
            const sourceTarget = tableTargets.get(sourceTable);
            if (sourceTarget) {
                outgoing.add(sourceTarget);
            }
            const content = [
                `# Linked Table: ${name}`,
                "",
                `- Source table: ${sourceTarget ? wiki(sourceTarget) : sourceTable}`,
                `- Is ODBC: ${linked.is_odbc ?? false}`,
                `- Connect: ${linked.connect ?? ""}`,
                ""
            ].join("\n");
            registerNote(`linked-tables/${sanitize(name)}`, content.split("\n"), outgoing);
        }

        const startupContent = [
            "# Startup Options",
            "",
            "| Option | Value | Source |",
            "|---|---|---|",
            ...metadata.startup_options.map((item) => `| ${item.option ?? ""} | ${String(item.value ?? "")} | ${item.source ?? ""} |`),
            ""
        ].join("\n");
        registerNote("startup/startup-options", startupContent.split("\n"));

        const stats = this.computeStats(metadata);
        const autoExecMacros = metadata.macros
            .map((macro) => String(macro.name ?? ""))
            .filter((name) => name && name.toLowerCase().includes("autoexec"));
        const likelyEntryForms = metadata.forms
            .map((form) => ({
                name: String(form.name ?? ""),
                recordSource: String(form.record_source ?? ""),
                controls: Array.isArray(form.controls) ? form.controls.length : 0,
                procedures: Array.isArray(form.procedures) ? form.procedures.length : 0
            }))
            .filter((form) => form.name && !isLikelySubObjectName(form.name))
            .sort((left, right) => (right.controls + right.procedures * 2) - (left.controls + left.procedures * 2) || left.name.localeCompare(right.name))
            .slice(0, 12);
        const likelyEntryReports = metadata.reports
            .map((report) => ({
                name: String(report.name ?? ""),
                controls: Array.isArray(report.controls) ? report.controls.length : 0,
                procedures: Array.isArray(report.procedures) ? report.procedures.length : 0
            }))
            .filter((report) => report.name && !isLikelySubObjectName(report.name))
            .sort((left, right) => (right.controls + right.procedures * 2) - (left.controls + left.procedures * 2) || left.name.localeCompare(right.name))
            .slice(0, 10);
        const indexContent = [
            `# Access Second Brain: ${metadata.database}`,
            "",
            `Generated / Generado: ${metadata.generated}`,
            "",
            `Scope / Alcance: ${this.scopeLabel(scope)}`,
            "",
            "## Stats / Estadísticas",
            "",
            ...Object.entries(stats).map(([key, value]) => `- ${key}: ${value}`),
            "",
            "## Main Notes / Notas principales",
            "",
            "- [[_overview]]",
            "- [[_entrypoints]]",
            "- [[_critical-objects]]",
            "- [[_health]]",
            "- [[_known-issues]]",
            "- [[_dependencies]]",
            "- [[_reading-order]]",
            "- [[_ai-prompt]]",
            "- [[_guide]]",
            ""
        ].join("\n");

        // Health checks
        const formsNoSource = metadata.forms
            .filter((f) => !String(f.record_source ?? "").trim())
            .map((f) => String(f.name ?? ""));
        const tablesNoRel = metadata.tables
            .filter((t) => {
                const tname = String(t.table_name ?? "");
                return !metadata.relationships.some((r) => String(r.table ?? "") === tname || String(r.foreign_table ?? "") === tname);
            })
            .map((t) => String(t.table_name ?? ""));
        const modulesNoProcs = metadata.modules
            .filter((m) => !(Array.isArray(m.procedures) ? m.procedures : []).length)
            .map((m) => String(m.name ?? ""));
        const queriesNoSql = metadata.queries
            .filter((q) => !String(q.sql ?? "").trim())
            .map((q) => String(q.name ?? ""));
        const warningsByObject = new Map<string, string[]>();
        for (const warning of metadata.warnings) {
            collectWarningMatches(warningsByObject, warning, "queries", metadata.queries, (item) => String(item.name ?? ""));
            collectWarningMatches(warningsByObject, warning, "forms", metadata.forms, (item) => String(item.name ?? ""));
            collectWarningMatches(warningsByObject, warning, "reports", metadata.reports, (item) => String(item.name ?? ""));
            collectWarningMatches(warningsByObject, warning, "modules", metadata.modules, (item) => String(item.name ?? ""));
            collectWarningMatches(warningsByObject, warning, "macros", metadata.macros, (item) => String(item.name ?? ""));
            collectWarningMatches(warningsByObject, warning, "tables", metadata.tables, (item) => String(item.table_name ?? ""));
        }

        const temporaryIncomingByTarget = new Map<string, Set<string>>();
        for (const [sourcePath, draft] of noteDrafts.entries()) {
            for (const target of draft.outgoing) {
                addToSetMap(temporaryIncomingByTarget, target, sourcePath);
            }
        }

        const criticalObjects = [
            ...metadata.tables.map((table) => {
                const name = String(table.table_name ?? "");
                const notePath = tableTargets.get(name) ?? `tables/${sanitize(name)}`;
                const score = (temporaryIncomingByTarget.get(notePath)?.size ?? 0) * 3
                    + (relationshipsByTable.get(name)?.size ?? 0) * 2
                    + (queriesByTable.get(notePath)?.length ?? 0) * 2
                    + (formsByRecordSource.get(name)?.length ?? 0);
                return { kind: "table", name, notePath, score, detail: `${queriesByTable.get(notePath)?.length ?? 0} queries, ${formsByRecordSource.get(name)?.length ?? 0} forms` };
            }),
            ...metadata.queries.map((query) => {
                const name = String(query.name ?? "");
                const notePath = queryTargets.get(name) ?? `queries/${sanitize(name)}`;
                const deps = queryDependenciesByQuery.get(name)?.size ?? 0;
                const score = (temporaryIncomingByTarget.get(notePath)?.size ?? 0) * 3
                    + deps * 2
                    + Math.min(10, Math.floor(String(query.sql ?? "").length / 250));
                return { kind: "query", name, notePath, score, detail: `${deps} dependencies, ${String(query.type ?? "unknown")}` };
            }),
            ...metadata.forms.map((form) => {
                const name = String(form.name ?? "");
                const notePath = formTargets.get(name) ?? `forms/${sanitize(name)}`;
                const controls = Array.isArray(form.controls) ? form.controls.length : 0;
                const procedures = Array.isArray(form.procedures) ? form.procedures.length : 0;
                const score = (temporaryIncomingByTarget.get(notePath)?.size ?? 0) * 3
                    + controls + procedures * 2 + (String(form.record_source ?? "").trim() ? 2 : 0);
                return { kind: "form", name, notePath, score, detail: `${controls} controls, ${procedures} procedures` };
            }),
            ...metadata.reports.map((report) => {
                const name = String(report.name ?? "");
                const notePath = reportTargets.get(name) ?? `reports/${sanitize(name)}`;
                const controls = Array.isArray(report.controls) ? report.controls.length : 0;
                const procedures = Array.isArray(report.procedures) ? report.procedures.length : 0;
                const score = (temporaryIncomingByTarget.get(notePath)?.size ?? 0) * 3 + controls + procedures * 2;
                return { kind: "report", name, notePath, score, detail: `${controls} controls, ${procedures} procedures` };
            }),
            ...metadata.modules.map((module) => {
                const name = String(module.name ?? "");
                const notePath = moduleTargets.get(name) ?? `modules/${sanitize(name)}`;
                const procedures = Array.isArray(module.procedures) ? module.procedures.length : 0;
                const score = (temporaryIncomingByTarget.get(notePath)?.size ?? 0) * 3
                    + (noteDrafts.get(notePath)?.outgoing.size ?? 0) * 2
                    + procedures * 2
                    + Math.min(10, Math.floor(countVbaCodeLines(String(module.code ?? "")) / 40));
                return { kind: "module", name, notePath, score, detail: `${procedures} procedures` };
            })
        ]
            .filter((item) => item.name)
            .sort((left, right) => right.score - left.score || left.name.localeCompare(right.name))
            .slice(0, 25);

        const refsSection: string[] = metadata.references.length > 0
            ? [
                "",
                "## VBA References / Referencias VBA",
                "",
                "| Name / Nombre | GUID | Path |",
                "|---|---|---|",
                ...metadata.references.map((r) => `| ${r.name ?? ""} | ${r.guid ?? ""} | ${r.path ?? ""} |`),
                ""
            ]
            : ["", "## VBA References / Referencias VBA", "", "- None", ""];

        const overviewContent = [
            `# Overview / Resumen: ${metadata.database}`,
            "",
            `This SecondBrain package was generated with scope: ${this.scopeLabel(scope)}.`,
            `Este paquete SecondBrain fue generado con alcance: ${this.scopeLabel(scope)}.`,
            "",
            "## Summary / Resumen",
            "",
            ...Object.entries(stats).map(([key, value]) => `- ${key}: ${value}`),
            ...refsSection,
            "## MOCs",
            "",
            ...(mocGroups.size > 0 ? Array.from(mocGroups.keys()).sort().map((mocPath) => `- ${wiki(mocPath)}`) : ["- None"]),
            ""
        ].join("\n");

        const healthContent = [
            "# Health / Salud",
            "",
            "## Warnings / Advertencias",
            "",
            ...(metadata.warnings.length > 0 ? metadata.warnings.map((w) => `- ${w}`) : ["- None"]),
            "",
            `## Forms without RecordSource / Formularios sin RecordSource (${formsNoSource.length})`,
            "",
            ...(formsNoSource.length > 0
                ? formsNoSource.map((n) => `- ${wiki(`forms/${sanitize(n)}`)}`) : ["- None"]),
            "",
            `## Tables without relationships / Tablas sin relaciones (${tablesNoRel.length})`,
            "",
            ...(tablesNoRel.length > 0
                ? tablesNoRel.map((n) => `- ${wiki(`tables/${sanitize(n)}`)}`) : ["- None"]),
            "",
            `## Modules without procedures / Módulos sin procedimientos (${modulesNoProcs.length})`,
            "",
            ...(modulesNoProcs.length > 0
                ? modulesNoProcs.map((n) => `- ${wiki(`modules/${sanitize(n)}`)}`) : ["- None"]),
            "",
            `## Queries without SQL / Consultas sin SQL (${queriesNoSql.length})`,
            "",
            ...(queriesNoSql.length > 0
                ? queriesNoSql.map((n) => `- ${wiki(`queries/${sanitize(n)}`)}`) : ["- None"]),
            ""
        ].join("\n");

        const entrypointsContent = [
            "# Entrypoints / Puntos de entrada",
            "",
            "## Startup Options / Opciones de inicio",
            "",
            ...(metadata.startup_options.length > 0
                ? metadata.startup_options.map((item) => `- ${item.option ?? ""}: ${String(item.value ?? "") || "(empty)"}`)
                : ["- None"]),
            "",
            "## AutoExec Macros / Macros de inicio",
            "",
            ...(autoExecMacros.length > 0 ? autoExecMacros.map((name) => `- ${wiki(`macros/${sanitize(name)}`)}`) : ["- None"]),
            "",
            "## Likely Main Forms / Formularios principales probables",
            "",
            ...(likelyEntryForms.length > 0
                ? likelyEntryForms.map((form) => `- ${wiki(`forms/${sanitize(form.name)}`)} | RecordSource: ${form.recordSource || "-"} | Controls: ${form.controls} | Procedures: ${form.procedures}`)
                : ["- None"]),
            "",
            "## Likely Main Reports / Informes principales probables",
            "",
            ...(likelyEntryReports.length > 0
                ? likelyEntryReports.map((report) => `- ${wiki(`reports/${sanitize(report.name)}`)} | Controls: ${report.controls} | Procedures: ${report.procedures}`)
                : ["- None"]),
            "",
            "## Suggested First Reads / Primeras lecturas sugeridas",
            "",
            "- [[_overview]]",
            "- [[_dependencies]]",
            "- [[_critical-objects]]",
            "- [[_known-issues]]",
            "- [[_reading-order]]",
            ""
        ].join("\n");

        const criticalObjectsContent = [
            "# Critical Objects / Objetos críticos",
            "",
            "Objetos con mayor centralidad heurística para empezar análisis, depuración o refactor.",
            "",
            "| Rank | Kind | Object | Score | Signals |",
            "|---:|---|---|---:|---|",
            ...criticalObjects.map((item, index) => `| ${index + 1} | ${item.kind} | ${wiki(item.notePath)} | ${item.score} | ${item.detail} |`),
            ""
        ].join("\n");

        const timeoutWarnings = metadata.warnings.filter((warning) => isLikelyTimeoutError(warning));
        const partialObjects = Array.from(warningsByObject.entries())
            .sort((left, right) => right[1].length - left[1].length || left[0].localeCompare(right[0]));

        const knownIssuesContent = [
            "# Known Issues / Problemas conocidos",
            "",
            "## Timeout Warnings / Warnings por timeout",
            "",
            ...(timeoutWarnings.length > 0 ? timeoutWarnings.map((warning) => `- ${warning}`) : ["- None"]),
            "",
            "## Partial Objects / Objetos parciales o con incidencias",
            "",
            ...(partialObjects.length > 0
                ? partialObjects.flatMap(([notePath, warnings]) => [`- ${wiki(notePath)}`, ...warnings.map((warning) => `  - ${warning}`)])
                : ["- None"]),
            "",
            "## Structural Smells / Señales estructurales",
            "",
            `- Forms without RecordSource: ${formsNoSource.length}`,
            `- Tables without relationships: ${tablesNoRel.length}`,
            `- Modules without procedures: ${modulesNoProcs.length}`,
            `- Queries without SQL: ${queriesNoSql.length}`,
            "",
            "## Review First / Revisar primero",
            "",
            ...(criticalObjects.slice(0, 10).map((item) => `- ${wiki(item.notePath)}`)),
            ""
        ].join("\n");

        const readingOrderContent = [
            "# Reading Order / Orden de lectura",
            "",
            "## UI bug / Problema de interfaz",
            "",
            "1. [[_entrypoints]]",
            "2. Nota del formulario o informe afectado",
            "3. Su RecordSource en [[_dependencies]] o en la nota de query/tabla",
            "4. El módulo o macro relacionado",
            "",
            "## Data bug / Problema de datos",
            "",
            "1. [[_dependencies]]",
            "2. Nota de la tabla afectada",
            "3. Queries que leen o escriben esa tabla",
            "4. Forms/reports que usan esas queries",
            "",
            "## VBA bug / Problema en VBA",
            "",
            "1. [[_critical-objects]]",
            "2. Nota del módulo o code-behind",
            "3. Objetos enlazados desde la sección Links",
            "4. [[_known-issues]] para ver si hay extracción parcial",
            "",
            "## Performance issue / Problema de rendimiento",
            "",
            "1. [[_known-issues]]",
            "2. [[_critical-objects]]",
            "3. Queries con SQL ausente o timeouts",
            "4. Forms/reports con muchos controles o incidencias",
            ""
        ].join("\n");

        const aiPromptContent = [
            "# AI Prompt / Prompt para IA",
            "",
            "Copia y pega este prompt en tu herramienta de IA junto con las notas relevantes de este Second Brain.",
            "",
            "## Prompt base",
            "",
            "```text",
            `Analiza este Second Brain de la base Access ${metadata.database}.`,
            "Usa primero _overview.md, _dependencies.md, _critical-objects.md, _known-issues.md y _reading-order.md para orientarte.",
            "Después lee solo las notas concretas relacionadas con el problema.",
            "",
            "Objetivo:",
            "- adquirir contexto rápido sin inventar información",
            "- localizar el origen probable del problema",
            "- identificar tablas, queries, forms, reports, macros y módulos implicados",
            "- señalar riesgos, dependencias y posibles efectos secundarios",
            "- proponer una estrategia mínima y segura para corregir o refactorizar",
            "",
            "Reglas:",
            "- trata como parciales los objetos listados en _known-issues.md",
            "- no asumas comportamientos que no estén respaldados por notas o metadata",
            "- si falta contexto, indica exactamente qué nota adicional necesitas leer",
            "- prioriza soluciones pequeñas, locales y verificables",
            "```",
            "",
            "## Qué compartir con la IA según el problema",
            "",
            "- Query rota: _dependencies.md + nota de la query + tablas relacionadas",
            "- Form/report roto: nota del form/report + su RecordSource + módulo/macro relacionado",
            "- Error VBA: nota del módulo + objetos enlazados desde Links",
            "- Problema general: _overview.md + _critical-objects.md + _known-issues.md + _dependencies.md",
            "",
            "## Suggested context pack / Pack sugerido",
            "",
            "- [[_overview]]",
            "- [[_dependencies]]",
            "- [[_critical-objects]]",
            "- [[_known-issues]]",
            "- [[_reading-order]]",
            ""
        ].join("\n");

        // _dependencies note: cross-reference map
        const depsOutgoing = new Set<string>();
        const depsLines: string[] = [
            "# Dependencies / Dependencias",
            "",
            "## Forms → Data Source / Formularios → Fuente de datos",
            "",
            "| Form / Formulario | RecordSource |",
            "|---|---|"
        ];
        for (const form of metadata.forms) {
            const fName = String(form.name ?? "");
            const fPath = formTargets.get(fName);
            const rs = String(form.record_source ?? "");
            const rsNorm = normalizeIdentifier(rs);
            const rsTarget = tableTargets.get(rsNorm) ?? queryTargets.get(rsNorm);
            if (fPath) {
                depsOutgoing.add(fPath);
                if (rsTarget) { depsOutgoing.add(rsTarget); }
                depsLines.push(`| ${wiki(fPath)} | ${rsTarget ? wiki(rsTarget) : (rs || "-")} |`);
            }
        }
        depsLines.push("", "## Queries → Used Objects / Consultas → Objetos usados", "", "| Query / Consulta | Type / Tipo | Depends on / Depende de |", "|---|---|---|");
        for (const query of metadata.queries) {
            const qName = String(query.name ?? "");
            const qPath = queryTargets.get(qName);
            const qDeps = queryDependenciesByQuery.get(qName) ?? new Set<string>();
            if (qPath) {
                depsOutgoing.add(qPath);
                for (const d of qDeps) { depsOutgoing.add(d); }
                const depsStr = Array.from(qDeps).sort().map((t) => wiki(t)).join(", ") || "-";
                depsLines.push(`| ${wiki(qPath)} | ${query.type ?? ""} | ${depsStr} |`);
            }
        }
        depsLines.push("", "## Tables → Who uses them / Tablas → Quién las usa", "", "| Table / Tabla | Queries | Forms / Formularios |", "|---|---|---|");
        for (const table of metadata.tables) {
            const tName = String(table.table_name ?? "");
            const tPath = tableTargets.get(tName);
            if (!tPath) { continue; }
            depsOutgoing.add(tPath);
            const usedByQueries = (queriesByTable.get(tPath) ?? []).map((queryPath) => wiki(queryPath)).join(", ") || "-";
            const usedByForms = (formsByRecordSource.get(tName) ?? []).map((formPath) => wiki(formPath)).join(", ") || "-";
            depsLines.push(`| ${wiki(tPath)} | ${usedByQueries} | ${usedByForms} |`);
        }
        depsLines.push("");

        const guideContent = [
            "# Second Brain — Guide / Guía de uso",
            "",
            `> Auto-generated from / Generado desde: **${metadata.database}**  `,
            `> Date / Fecha: ${metadata.generated}`,
            "",
            "---",
            "",
            "## 🇬🇧 English",
            "",
            "### What is this vault?",
            "",
            "This Obsidian vault contains a complete map of an Access database:",
            "tables, queries, forms, reports, VBA modules, relationships and dependencies.",
            "Each object is a note with its metadata, outgoing links and automatic backlinks.",
            "",
            "### How to open in Obsidian",
            "",
            "1. Open Obsidian → **Open folder as vault**",
            "2. Select the root folder of this export",
            "3. Trust the vault when prompted",
            "4. Start from [[_index]] or use **Graph View** to navigate dependencies visually",
            "",
            "#### Key notes",
            "",
            "| Note | Description |",
            "|---|---|",
            `| ${wiki("_index")} | Entry point with general statistics |`,
            `| ${wiki("_overview")} | Summary, MOCs and VBA references |`,
            `| ${wiki("_entrypoints")} | Likely entry forms, reports, startup options and AutoExec macros |`,
            `| ${wiki("_critical-objects")} | Objects with highest diagnostic and refactor priority |`,
            `| ${wiki("_health")} | Quality checks: forms without source, orphan tables… |`,
            `| ${wiki("_known-issues")} | Timeouts, partial objects and warnings grouped for review |`,
            `| ${wiki("_dependencies")} | Cross-reference map: forms→tables, queries→tables |`,
            `| ${wiki("_reading-order")} | Recommended reading sequence by problem type |`,
            `| ${wiki("_ai-prompt")} | Ready-to-copy prompt and context pack for AI analysis |`,
            "",
            "#### Folder structure",
            "",
            "| Folder | Contents |",
            "|---|---|",
            "| `tables/` | One note per table with columns, PKs, FKs and indexes |",
            "| `queries/` | One note per query with SQL and dependencies |",
            "| `forms/` | One note per form with controls, properties and VBA code |",
            "| `reports/` | One note per report with controls and VBA code |",
            "| `modules/` | One note per module with procedure index and code |",
            "| `macros/` | One note per macro with its actions |",
            "| `relationships/` | One note per table relationship |",
            "| `linked-tables/` | Linked tables with their ODBC or MDB origin |",
            "| `mocs/` | Knowledge maps grouped by domain (high density mode) |",
            "| `startup/` | Application startup options |",
            "",
            "### How to use with AI (Claude, ChatGPT, Copilot…)",
            "",
            "**Option A — Paste content directly**",
            "",
            "1. Open the relevant note (e.g. `tables/MyTable.md`)",
            "2. Copy its full content",
            "3. Paste it into your AI conversation with your question",
            "",
            "**Option B — Upload the full vault as context**",
            "",
            "If your AI tool supports file/folder uploads:",
            "",
            "1. Compress the `db-second-brain/` folder into a ZIP",
            "2. Upload it as context to the session",
            "3. Ask for full analysis: *\"Which forms use table X?\"*,",
            "   *\"Explain the data flow of this module\"*, *\"Are there orphan tables?\"*",
            "",
            "**Example prompts**",
            "",
            "```",
            "Analyze this Access database Second Brain.",
            "Identify critical dependencies and potential design issues.",
            "```",
            "",
            "```",
            "Based on _dependencies.md, which tables are most critical?",
            "What would happen if I delete table X?",
            "```",
            "",
            "```",
            "Read module Y and explain what each procedure does.",
            "Generate English documentation for each function.",
            "```",
            "",
            "```",
            "Review _health.md and suggest fixes for each issue found.",
            "```",
            "",
            "### Obsidian tips",
            "",
            "- Use **Ctrl+G** to open Graph View and navigate visually",
            "- Use **Ctrl+P** → *Switcher* to search any note by name",
            "- Filter the graph by folder (e.g. only `tables/`) to see a sub-graph",
            "- Install the **Dataview** plugin to run SQL-like queries over notes",
            "- With Dataview you can list all forms that use a specific table",
            "",
            "---",
            "",
            "## 🇪🇸 Español",
            "",
            "### ¿Qué es este vault?",
            "",
            "Este vault Obsidian contiene un mapa completo de la base de datos Access:",
            "tablas, consultas, formularios, informes, módulos VBA, relaciones y dependencias.",
            "Cada objeto es una nota con sus metadatos, vínculos y backlinks automáticos.",
            "",
            "### Cómo abrirlo en Obsidian",
            "",
            "1. Abre Obsidian → **Abrir carpeta como vault**",
            "2. Selecciona la carpeta raíz de este export",
            "3. Confía en el vault cuando lo pida",
            "4. Empieza desde [[_index]] o usa **Graph View** para ver el grafo de dependencias",
            "",
            "#### Notas clave",
            "",
            "| Nota | Descripción |",
            "|---|---|",
            `| ${wiki("_index")} | Punto de entrada con estadísticas generales |`,
            `| ${wiki("_overview")} | Resumen, MOCs y referencias VBA |`,
            `| ${wiki("_entrypoints")} | Formularios, informes y macros de inicio más probables |`,
            `| ${wiki("_critical-objects")} | Objetos con mayor prioridad diagnóstica y de refactor |`,
            `| ${wiki("_health")} | Checks de calidad: formularios sin fuente, tablas huérfanas… |`,
            `| ${wiki("_known-issues")} | Timeouts, objetos parciales y advertencias agrupadas |`,
            `| ${wiki("_dependencies")} | Mapa cruzado: formularios→tablas, consultas→tablas |`,
            `| ${wiki("_reading-order")} | Orden recomendado de lectura según el problema |`,
            `| ${wiki("_ai-prompt")} | Prompt listo para usar con IA y pack de contexto |`,
            "",
            "#### Estructura de carpetas",
            "",
            "| Carpeta | Contenido |",
            "|---|---|",
            "| `tables/` | Una nota por tabla con columnas, PKs, FKs e índices |",
            "| `queries/` | Una nota por consulta con SQL y dependencias |",
            "| `forms/` | Una nota por formulario con controles, props y código VBA |",
            "| `reports/` | Una nota por informe con controles y código VBA |",
            "| `modules/` | Una nota por módulo con índice de procedimientos y código |",
            "| `macros/` | Una nota por macro con sus acciones |",
            "| `relationships/` | Una nota por relación entre tablas |",
            "| `linked-tables/` | Tablas vinculadas con su origen ODBC o MDB |",
            "| `mocs/` | Mapas de conocimiento agrupados por dominio (high density) |",
            "| `startup/` | Opciones de inicio de la aplicación |",
            "",
            "### Cómo usarlo con IA (Claude, ChatGPT, Copilot…)",
            "",
            "**Opción A — Pegar el contenido directamente**",
            "",
            "1. Abre la nota relevante (por ejemplo `tables/MiTabla.md`)",
            "2. Copia su contenido completo",
            "3. Pégalo en tu conversación con la IA junto con tu pregunta",
            "",
            "**Opción B — Exportar el vault completo como contexto**",
            "",
            "Si tu herramienta de IA soporta subida de archivos o carpetas:",
            "",
            "1. Comprime la carpeta `db-second-brain/` en un ZIP",
            "2. Súbelo como contexto a la sesión",
            "3. Pide análisis completos: *\"¿Qué formularios usan la tabla X?\"*,",
            "   *\"Explica el flujo de datos de este módulo\"*, *\"¿Hay tablas sin relaciones?\"*",
            "",
            "**Prompts de ejemplo**",
            "",
            "```",
            "Analiza el siguiente Second Brain de una base de datos Access.",
            "Identifica las dependencias críticas y posibles problemas de diseño.",
            "```",
            "",
            "```",
            "Basándote en la nota _dependencies.md, ¿qué tablas son más críticas?",
            "¿Qué pasaría si elimino la tabla X?",
            "```",
            "",
            "```",
            "Lee el módulo Y y explica qué hace cada procedimiento.",
            "Genera documentación en español para cada función.",
            "```",
            "",
            "```",
            "Revisa _health.md y propón soluciones para cada issue encontrado.",
            "```",
            "",
            "### Tips para Obsidian",
            "",
            "- Usa **Ctrl+G** para abrir el Graph View y navegar visualmente",
            "- Usa **Ctrl+P** → *Switcher* para buscar cualquier nota por nombre",
            "- Filtra el grafo por carpeta (p.ej. solo `tables/`) para ver un sub-grafo",
            "- Instala el plugin **Dataview** para hacer queries SQL sobre las notas",
            "- Con Dataview puedes listar todos los formularios que usan una tabla",
            ""
        ].join("\n");

        registerNote("_index", indexContent.split("\n"), ["_overview", "_entrypoints", "_critical-objects", "_health", "_known-issues", "_dependencies", "_reading-order", "_ai-prompt", "_guide"]);
        registerNote("_overview", overviewContent.split("\n"), ["_index", "_entrypoints", "_critical-objects", "_known-issues", "_dependencies", "_reading-order", "_ai-prompt", "_guide", ...mocGroups.keys()]);
        registerNote("_entrypoints", entrypointsContent.split("\n"), ["_index", "_overview", "_critical-objects", "_dependencies", "_reading-order"]);
        registerNote("_critical-objects", criticalObjectsContent.split("\n"), ["_index", "_overview", "_known-issues", "_dependencies", ...criticalObjects.map((item) => item.notePath)]);
        registerNote("_health", healthContent.split("\n"), ["_index", "_overview"]);
        registerNote("_dependencies", depsLines, depsOutgoing);
        registerNote("_known-issues", knownIssuesContent.split("\n"), ["_index", "_overview", "_health", "_critical-objects", ...partialObjects.map(([notePath]) => notePath)]);
        registerNote("_reading-order", readingOrderContent.split("\n"), ["_index", "_overview", "_entrypoints", "_dependencies", "_critical-objects", "_known-issues"]);
        registerNote("_ai-prompt", aiPromptContent.split("\n"), ["_index", "_overview", "_dependencies", "_critical-objects", "_known-issues", "_reading-order", "_guide"]);
        registerNote("_guide", guideContent.split("\n"), ["_index", "_overview", "_entrypoints", "_critical-objects", "_health", "_known-issues", "_dependencies", "_reading-order", "_ai-prompt"]);

        const aiIndexStartedAt = Date.now();
        const aiIndex = {
            database: metadata.database,
            generated: metadata.generated,
            scope: this.scopeLabel(scope),
            keyDocuments: [
                "_index.md",
                "_overview.md",
                "_entrypoints.md",
                "_critical-objects.md",
                "_health.md",
                "_known-issues.md",
                "_dependencies.md",
                "_reading-order.md",
                "_ai-prompt.md",
                "_guide.md"
            ],
            stats,
            warnings: metadata.warnings,
            criticalObjects,
            objects: [
                ...metadata.tables.map((table) => buildAiIndexEntry({
                    kind: "table",
                    name: String(table.table_name ?? ""),
                    notePath: tableTargets.get(String(table.table_name ?? "")) ?? `tables/${sanitize(String(table.table_name ?? ""))}`,
                    outgoing: noteDrafts.get(tableTargets.get(String(table.table_name ?? "")) ?? `tables/${sanitize(String(table.table_name ?? ""))}`)?.outgoing ?? new Set<string>(),
                    flags: deriveObjectFlags(warningsByObject.get(tableTargets.get(String(table.table_name ?? "")) ?? `tables/${sanitize(String(table.table_name ?? ""))}`), []),
                    recordSource: undefined,
                    procedures: 0
                })),
                ...metadata.queries.map((query) => buildAiIndexEntry({
                    kind: "query",
                    name: String(query.name ?? ""),
                    notePath: queryTargets.get(String(query.name ?? "")) ?? `queries/${sanitize(String(query.name ?? ""))}`,
                    outgoing: noteDrafts.get(queryTargets.get(String(query.name ?? "")) ?? `queries/${sanitize(String(query.name ?? ""))}`)?.outgoing ?? new Set<string>(),
                    flags: deriveObjectFlags(warningsByObject.get(queryTargets.get(String(query.name ?? "")) ?? `queries/${sanitize(String(query.name ?? ""))}`), !String(query.sql ?? "").trim() ? ["sql-unavailable"] : []),
                    recordSource: undefined,
                    procedures: 0
                })),
                ...metadata.forms.map((form) => buildAiIndexEntry({
                    kind: "form",
                    name: String(form.name ?? ""),
                    notePath: formTargets.get(String(form.name ?? "")) ?? `forms/${sanitize(String(form.name ?? ""))}`,
                    outgoing: noteDrafts.get(formTargets.get(String(form.name ?? "")) ?? `forms/${sanitize(String(form.name ?? ""))}`)?.outgoing ?? new Set<string>(),
                    flags: deriveObjectFlags(warningsByObject.get(formTargets.get(String(form.name ?? "")) ?? `forms/${sanitize(String(form.name ?? ""))}`), !String(form.record_source ?? "").trim() ? ["recordsource-missing"] : []),
                    recordSource: String(form.record_source ?? "") || undefined,
                    procedures: Array.isArray(form.procedures) ? form.procedures.length : 0
                })),
                ...metadata.reports.map((report) => buildAiIndexEntry({
                    kind: "report",
                    name: String(report.name ?? ""),
                    notePath: reportTargets.get(String(report.name ?? "")) ?? `reports/${sanitize(String(report.name ?? ""))}`,
                    outgoing: noteDrafts.get(reportTargets.get(String(report.name ?? "")) ?? `reports/${sanitize(String(report.name ?? ""))}`)?.outgoing ?? new Set<string>(),
                    flags: deriveObjectFlags(warningsByObject.get(reportTargets.get(String(report.name ?? "")) ?? `reports/${sanitize(String(report.name ?? ""))}`), []),
                    recordSource: String(report.record_source ?? "") || undefined,
                    procedures: Array.isArray(report.procedures) ? report.procedures.length : 0
                })),
                ...metadata.modules.map((module) => buildAiIndexEntry({
                    kind: "module",
                    name: String(module.name ?? ""),
                    notePath: moduleTargets.get(String(module.name ?? "")) ?? `modules/${sanitize(String(module.name ?? ""))}`,
                    outgoing: noteDrafts.get(moduleTargets.get(String(module.name ?? "")) ?? `modules/${sanitize(String(module.name ?? ""))}`)?.outgoing ?? new Set<string>(),
                    flags: deriveObjectFlags(warningsByObject.get(moduleTargets.get(String(module.name ?? "")) ?? `modules/${sanitize(String(module.name ?? ""))}`), !(Array.isArray(module.procedures) ? module.procedures : []).length ? ["procedures-missing"] : []),
                    recordSource: undefined,
                    procedures: Array.isArray(module.procedures) ? module.procedures.length : 0
                }))
            ]
        };
        await fs.writeFile(path.join(vaultDir, "ai-index.json"), JSON.stringify(aiIndex, null, 2), "utf-8");
        await this.reportProgress(options, {
            phase: "write",
            message: `ai-index.json escrito en ${formatDuration(Date.now() - aiIndexStartedAt)}`
        });

        const bodyWriteStartedAt = Date.now();
        await Promise.all(bodyWrites);
        await this.reportProgress(options, {
            phase: "write",
            message: `Cuerpos de notas escritos (${noteDrafts.size}) en ${formatDuration(Date.now() - bodyWriteStartedAt)}`
        });

        const backlinksStartedAt = Date.now();
        const incomingByTarget = new Map<string, Set<string>>();
        for (const [sourcePath, draft] of noteDrafts.entries()) {
            for (const target of draft.outgoing) {
                if (!incomingByTarget.has(target)) {
                    incomingByTarget.set(target, new Set<string>());
                }
                incomingByTarget.get(target)?.add(sourcePath);
            }
        }

        await this.reportProgress(options, {
            phase: "write",
            message: `Backlinks calculados en ${formatDuration(Date.now() - backlinksStartedAt)}`
        });

        const linksWriteStartedAt = Date.now();
        for (const [notePath, draft] of noteDrafts.entries()) {
            const outgoingSorted = Array.from(draft.outgoing).sort();
            const incomingSorted = Array.from(incomingByTarget.get(notePath) ?? []).sort();

            const linksSection = [
                "",
                "## Links",
                "",
                ...(outgoingSorted.length > 0 ? outgoingSorted.map((target) => `- ${wiki(target)}`) : ["- None"]),
                "",
                "## Referenciado por",
                "",
                ...(incomingSorted.length > 0 ? incomingSorted.map((source) => `- ${wiki(source)}`) : ["- None"]),
                ""
            ].join("\n");

            await fs.appendFile(path.join(vaultDir, `${notePath}.md`), `\n${linksSection}`, "utf-8");
        }

        await this.reportProgress(options, {
            phase: "write",
            message: `Secciones de enlaces escritas en ${formatDuration(Date.now() - linksWriteStartedAt)}; total vault ${formatDuration(Date.now() - vaultStartedAt)}`
        });
    }

    private async reportProgress(options: SecondBrainExportOptions | undefined, event: SecondBrainProgressEvent): Promise<void> {
        await options?.onProgress?.(event);
    }

    private buildUiNote(kind: "Form" | "Report", item: Record<string, unknown>, dependencies: Set<string>): string {
        const controls = Array.isArray(item.controls) ? item.controls : [];
        const procs = Array.isArray(item.procedures)
            ? (item.procedures as Array<{ kind: string; name: string; visibility: string }>)
            : [];
        const sortedDependencies = Array.from(dependencies).sort();

        const propLines: string[] = [];
        if (item.allow_edits !== undefined) { propLines.push(`- AllowEdits: ${item.allow_edits}`); }
        if (item.allow_additions !== undefined) { propLines.push(`- AllowAdditions: ${item.allow_additions}`); }
        if (item.allow_deletions !== undefined) { propLines.push(`- AllowDeletions: ${item.allow_deletions}`); }
        if (item.default_view !== undefined) { propLines.push(`- DefaultView: ${item.default_view}`); }
        if (item.modal !== undefined) { propLines.push(`- Modal: ${item.modal}`); }
        if (item.pop_up !== undefined) { propLines.push(`- PopUp: ${item.pop_up}`); }

        const header = [
            `# ${kind}: ${item.name ?? ""}`,
            "",
            `- RecordSource: ${item.record_source ?? ""}`,
            `- Controls: ${controls.length}`,
            ...propLines,
            ""
        ];

        if (procs.length > 0) {
            header.push(
                `## Procedures (${procs.length})`,
                "",
                "| Kind | Name | Visibility |",
                "|---|---|---|",
                ...procs.map((p) => `| ${p.kind} | ${p.name} | ${p.visibility} |`),
                ""
            );
        }

        header.push(
            "## Depends On",
            "",
            ...(sortedDependencies.length > 0 ? sortedDependencies.map((target) => `- ${wiki(target)}`) : ["- None"]),
            "",
            "## Controls",
            "",
            "| Name | Type | Source | Caption | SubForm |",
            "|---|---|---|---|---|"
        );

        for (const control of controls as Array<Record<string, unknown>>) {
            header.push(
                `| ${control.name ?? ""} | ${control.control_type ?? ""} | ${control.control_source ?? ""} | ${control.caption ?? ""} | ${control.source_object ?? ""} |`
            );
        }

        header.push("", "## Code", "", "```vb", String(item.code ?? ""), "```", "");
        return header.join("\n");
    }

    private collectUiDependencies(
        item: Record<string, unknown>,
        tableTargets: Map<string, string>,
        queryTargets: Map<string, string>
    ): Set<string> {
        const dependencies = new Set<string>();

        const recordSource = String(item.record_source ?? "");
        const recordSourceName = normalizeIdentifier(recordSource);
        const tableTarget = tableTargets.get(recordSourceName);
        if (tableTarget) {
            dependencies.add(tableTarget);
        }
        const queryTarget = queryTargets.get(recordSourceName);
        if (queryTarget) {
            dependencies.add(queryTarget);
        }

        const controls = Array.isArray(item.controls) ? item.controls : [];
        for (const control of controls as Array<Record<string, unknown>>) {
            const source = String(control.control_source ?? "");
            for (const candidate of extractBracketedIdentifiers(source)) {
                const sourceTable = tableTargets.get(candidate);
                if (sourceTable) {
                    dependencies.add(sourceTable);
                }
                const sourceQuery = queryTargets.get(candidate);
                if (sourceQuery) {
                    dependencies.add(sourceQuery);
                }
            }
        }

        return dependencies;
    }

    private computeStats(metadata: SecondBrainMetadata): Record<string, number> {
        return {
            tables: metadata.tables.length,
            columns: metadata.columns.length,
            primary_keys: metadata.primary_keys.length,
            foreign_keys: metadata.foreign_keys.length,
            indexes: metadata.indexes.length,
            queries: metadata.queries.length,
            relationships: metadata.relationships.length,
            forms: metadata.forms.length,
            reports: metadata.reports.length,
            macros: metadata.macros.length,
            modules: metadata.modules.length,
            linked_tables: metadata.linked_tables.length,
            startup_options: metadata.startup_options.length,
            references: metadata.references.length
        };
    }

    private scopeLabel(scope: SecondBrainScope): string {
        if (scope.mode === "full") {
            return "full";
        }
        if (scope.mode === "category") {
            return `category - ${scope.categoryKey}`;
        }
        return `object - ${scope.categoryKey} - ${sanitize(scope.objectInfo.name)}`;
    }
}

function toolTypeForCategory(categoryKey: AccessCategoryKey): string | undefined {
    switch (categoryKey) {
        case "tables":
            return "table";
        case "queries":
            return "query";
        case "forms":
            return "form";
        case "reports":
            return "report";
        case "macros":
            return "macro";
        case "modules":
            return "module";
        default:
            return undefined;
    }
}

function inferQueryType(sql: string): string {
    const trimmed = sql.trim().toLowerCase();
    if (trimmed.startsWith("select")) {
        return "select";
    }
    if (trimmed.startsWith("insert")) {
        return "insert";
    }
    if (trimmed.startsWith("update")) {
        return "update";
    }
    if (trimmed.startsWith("delete")) {
        return "delete";
    }
    if (trimmed.startsWith("transform")) {
        return "crosstab";
    }
    if (trimmed.startsWith("exec")) {
        return "pass-through";
    }
    return "unknown";
}

function safeJson(content: string): Record<string, any> {
    try {
        return JSON.parse(content);
    } catch {
        return {};
    }
}

function groupBy<T>(items: T[], keyGetter: (item: T) => string): Map<string, T[]> {
    const map = new Map<string, T[]>();
    for (const item of items) {
        const key = keyGetter(item);
        if (!map.has(key)) {
            map.set(key, []);
        }
        map.get(key)?.push(item);
    }
    return map;
}

function addToSetMap<K, V>(map: Map<K, Set<V>>, key: K, value: V): void {
    if (!key) {
        return;
    }

    const values = map.get(key) ?? new Set<V>();
    values.add(value);
    map.set(key, values);
}

function formatDuration(durationMs: number): string {
    if (durationMs < 1000) {
        return `${durationMs} ms`;
    }

    if (durationMs < 60000) {
        return `${(durationMs / 1000).toFixed(1)} s`;
    }

    const minutes = Math.floor(durationMs / 60000);
    const seconds = ((durationMs % 60000) / 1000).toFixed(1);
    return `${minutes} min ${seconds} s`;
}

function isLikelyTimeoutError(message: string): boolean {
    const normalized = message.toLowerCase();
    return normalized.includes("timeout") || normalized.includes("timed out") || normalized.includes("supero el timeout");
}

function sanitize(value: string): string {
    return value.replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, "-").replace(/-+/g, "-").trim();
}

function isLikelySubObjectName(name: string): boolean {
    const normalized = name.trim().toLowerCase();
    return normalized.startsWith("sub")
        || normalized.startsWith("sf")
        || normalized.includes(" unterformular")
        || normalized.includes(" unterbericht")
        || normalized.includes("subform")
        || normalized.includes("subreport");
}

function collectWarningMatches<T>(
    warningsByObject: Map<string, string[]>,
    warning: string,
    kind: string,
    items: T[],
    nameGetter: (item: T) => string
): void {
    const prefix = `${kind}:`;
    const lowerWarning = warning.toLowerCase();
    for (const item of items) {
        const name = nameGetter(item);
        if (!name) { continue; }
        if (lowerWarning.includes(`${prefix}${name}`.toLowerCase())) {
            const notePath = `${kind}/${sanitize(name)}`;
            const warnings = warningsByObject.get(notePath) ?? [];
            warnings.push(warning);
            warningsByObject.set(notePath, warnings);
        }
    }
}

function deriveObjectFlags(warnings: string[] | undefined, extraFlags: string[]): string[] {
    const flags = new Set(extraFlags);
    for (const warning of warnings ?? []) {
        const normalized = warning.toLowerCase();
        if (normalized.includes("timeout")) { flags.add("timeout"); }
        if (normalized.includes("no disponible") || normalized.includes("no se pudo")) { flags.add("partial"); }
        if (normalized.includes("controles")) { flags.add("controls-unavailable"); }
        if (normalized.includes("propiedades")) { flags.add("properties-unavailable"); }
        if (normalized.includes("codigo") || normalized.includes("code")) { flags.add("code-unavailable"); }
    }
    return Array.from(flags).sort();
}

function buildAiIndexEntry(input: {
    kind: string;
    name: string;
    notePath: string;
    outgoing: Set<string>;
    flags: string[];
    recordSource?: string;
    procedures: number;
}): Record<string, unknown> {
    return {
        kind: input.kind,
        name: input.name,
        notePath: `${input.notePath}.md`,
        outgoing: Array.from(input.outgoing).sort(),
        flags: input.flags,
        recordSource: input.recordSource,
        procedures: input.procedures,
        status: input.flags.length > 0 ? "partial" : "complete"
    };
}

/** Repara doble encoding UTF-8 (latin1 interpretado como UTF-8). Ej: "Â·" -> "·" */
function fixEncoding(text: string | undefined): string | undefined {
    if (!text) { return text; }
    // Solo aplicar si hay patrón de doble encoding UTF-8 (mojibake):
    // caracteres 0xC0-0xFF seguidos de 0x80-0xBF
    if (!/[\xC0-\xFF][\x80-\xBF]/.test(text)) { return text; }
    try {
        const bytes = Uint8Array.from(text, (c) => c.charCodeAt(0) & 0xff);
        const fixed = new TextDecoder("utf-8", { fatal: false }).decode(bytes);
        // Solo usar si no introdujo caracteres de reemplazo
        return fixed.includes("\uFFFD") ? text : fixed;
    } catch {
        return text;
    }
}

function wiki(target: string): string {
    return `[[${target}]]`;
}

function normalizeIdentifier(value: string): string {
    return value.trim().replace(/^\[|\]$/g, "").replace(/^['"]|['"]$/g, "");
}

function extractBracketedIdentifiers(value: string): string[] {
    const out = new Set<string>();

    const bracketRegex = /\[([^\]]+)\]/g;
    for (const match of value.matchAll(bracketRegex)) {
        const normalized = normalizeIdentifier(match[1] ?? "");
        if (normalized) {
            out.add(normalized);
        }
    }

    const tokenRegex = /\b[A-Za-z_][A-Za-z0-9_]*\b/g;
    for (const match of value.matchAll(tokenRegex)) {
        const normalized = normalizeIdentifier(match[0] ?? "");
        if (normalized) {
            out.add(normalized);
        }
    }

    return Array.from(out);
}

function extractSqlObjectCandidates(sql: string): string[] {
    if (!sql.trim()) {
        return [];
    }

    const out = new Set<string>();
    const patterns = [
        /\bfrom\s+([^\s,;()]+)/gi,
        /\bjoin\s+([^\s,;()]+)/gi,
        /\bupdate\s+([^\s,;()]+)/gi,
        /\binto\s+([^\s,;()]+)/gi,
        /\bdelete\s+from\s+([^\s,;()]+)/gi
    ];

    for (const regex of patterns) {
        for (const match of sql.matchAll(regex)) {
            const normalized = normalizeIdentifier(String(match[1] ?? "").split(".").pop() ?? "");
            if (normalized) {
                out.add(normalized);
            }
        }
    }

    return Array.from(out);
}

function extractVbaObjectLinks(
    code: string,
    formTargets: Map<string, string>,
    reportTargets: Map<string, string>,
    queryTargets: Map<string, string>,
    macroTargets: Map<string, string>,
    tableTargets?: Map<string, string>,
    moduleTargets?: Map<string, string>
): Set<string> {
    const links = new Set<string>();
    const patterns: Array<{ regex: RegExp; map: Map<string, string> }> = [
        { regex: /OpenForm\s*\(?\s*"([^"]+)"/gi, map: formTargets },
        { regex: /\bForms\s*[!(]\s*"?([A-Za-z_][A-Za-z0-9_\s]*)"?[!)]/gi, map: formTargets },
        { regex: /OpenReport\s*\(?\s*"([^"]+)"/gi, map: reportTargets },
        { regex: /OpenQuery\s*\(?\s*"([^"]+)"/gi, map: queryTargets },
        { regex: /QueryDefs\s*\(\s*"([^"]+)"/gi, map: queryTargets },
        { regex: /RunMacro\s*\(?\s*"([^"]+)"/gi, map: macroTargets },
        ...(tableTargets ? [
            { regex: /\bD(?:Lookup|Count|Sum|Avg|Max|Min|First|Last)\s*\([^,]+,\s*"([^"]+)"/gi, map: tableTargets },
            { regex: /\bOpenRecordset\s*\(\s*"([^"]+)"/gi, map: tableTargets },
            { regex: /\bTableDefs\s*\(\s*"([^"]+)"/gi, map: tableTargets }
        ] : []),
        ...(moduleTargets ? [
            { regex: /\b([A-Za-z_][A-Za-z0-9_]+)\s*\.\s*[A-Za-z_]/gi, map: moduleTargets }
        ] : [])
    ];

    for (const pattern of patterns) {
        for (const match of code.matchAll(pattern.regex)) {
            const targetName = normalizeIdentifier(match[1] ?? "");
            const target = pattern.map.get(targetName);
            if (target) {
                links.add(target);
            }
        }
    }

    return links;
}

function extractVbaProcedures(code: string): Array<{ kind: string; name: string; visibility: string }> {
    const procedures: Array<{ kind: string; name: string; visibility: string }> = [];
    const regex = /^[ \t]*(?:(Public|Private|Friend)\s+)?(?:Static\s+)?(?:(Function|Sub)|(Property\s+(?:Get|Let|Set)))\s+(\w+)/gim;
    for (const match of code.matchAll(regex)) {
        const visibility = match[1] ?? "Public";
        const kind = match[2] ?? (match[3] ?? "").trim();
        const name = match[4];
        if (name && kind) {
            procedures.push({ kind, name, visibility });
        }
    }
    return procedures;
}

function countVbaCodeLines(code: string): number {
    return code
        .split(/\r?\n/)
        .filter((line) => {
            const trimmed = line.trim();
            return trimmed.length > 0 && !trimmed.startsWith("'");
        }).length;
}

function inferDomainKey(name: string): string {
    const normalized = normalizeIdentifier(name).trim();
    if (!normalized) {
        return "";
    }

    const tokens = normalized
        .split(/[\s_\-\.]+/)
        .map((token) => token.trim())
        .filter(Boolean);

    if (tokens.length === 0) {
        return "";
    }

    const skipTokens = new Set(["tbl", "table", "qry", "query", "frm", "form", "rpt", "report", "mod", "module", "mcr", "macro"]);

    for (const token of tokens) {
        const lower = token.toLowerCase();
        if (!skipTokens.has(lower) && lower.length >= 3) {
            return lower;
        }
    }

    return tokens[0].toLowerCase();
}
