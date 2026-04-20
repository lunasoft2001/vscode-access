import * as fs from "node:fs/promises";
import * as path from "node:path";
import { AccessCategoryKey, AccessConnection, AccessObjectInfo } from "../models/types";
import { McpAccessClient } from "../mcp/mcpAccessClient";
import { BulkExportService } from "./bulkExportService";

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
    lines: string[];
    outgoing: Set<string>;
}

export class SecondBrainService {
    private static readonly INVENTORY_TIMEOUT_MS = 300000;
    private static readonly OBJECT_TIMEOUT_MS = 180000;
    private static readonly UI_TIMEOUT_MS = 240000;
    private static readonly QUERY_TIMEOUT_MS = 15000;

    private readonly bulkExportService: BulkExportService;

    constructor(private readonly mcpClient: McpAccessClient, globalStoragePath: string) {
        this.bulkExportService = new BulkExportService(mcpClient, globalStoragePath);
    }

    async exportSecondBrain(
        connection: AccessConnection,
        baseOutputDir: string,
        scope: SecondBrainScope,
        options?: SecondBrainExportOptions
    ): Promise<SecondBrainExportResult> {
        const timestamp = new Date().toISOString().replace(/[:]/g, "-").replace(/\..+$/, "");
        const scopeLabel = this.scopeLabel(scope);
        const rootDir = path.join(baseOutputDir, `secondbrain-${sanitize(connection.name)}-${scopeLabel}-${timestamp}`);
        const vaultDir = path.join(rootDir, "db-second-brain");

        await fs.mkdir(vaultDir, { recursive: true });

        await this.reportProgress(options, {
            phase: "inventory",
            message: `Preparando exportacion en ${rootDir}`
        });

        const metadata = await this.buildMetadata(connection, scope, options);
        const metadataPath = path.join(rootDir, "metadata.json");
        await fs.writeFile(metadataPath, JSON.stringify(metadata, null, 2), "utf-8");

        await this.reportProgress(options, {
            phase: "write",
            message: "Escribiendo metadata.json"
        });

        await this.generateVault(vaultDir, metadata, scope, options);

        await this.reportProgress(options, {
            phase: "done",
            message: "SecondBrain generado correctamente"
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
        // Single object: no hay ventaja de bulk, usar método secuencial directo
        if (scope.mode === "object") {
            return this.buildMetadataSequential(connection, scope, options);
        }

        const bulkMode = scope.mode === "full" ? "full" : scope.categoryKey;

        await this.reportProgress(options, {
            phase: "inventory",
            message: "Ejecutando exportacion masiva VBA (una sola llamada COM)..."
        });

        try {
            const raw = await this.bulkExportService.runJsonExport(connection, bulkMode);
            return this.rawToMetadata(raw, connection);
        } catch (bulkError) {
            const errMsg = bulkError instanceof Error ? bulkError.message : String(bulkError);
            await this.reportProgress(options, {
                phase: "inventory",
                message: `Exportacion VBA fallo (${errMsg}). Usando metodo secuencial...`
            });
            return this.buildMetadataSequential(connection, scope, options);
        }
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
            await this.collectCategory(connection, metadata, "tables", options);
            await this.collectCategory(connection, metadata, "queries", options);
            await this.collectCategory(connection, metadata, "forms", options);
            await this.collectCategory(connection, metadata, "reports", options);
            await this.collectCategory(connection, metadata, "macros", options);
            await this.collectCategory(connection, metadata, "modules", options);
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
        options?: SecondBrainExportOptions
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
        const objects = await this.mcpClient.listObjects(connection, toolType, SecondBrainService.INVENTORY_TIMEOUT_MS);
        const total = objects.length;
        let completed = 0;
        for (const objectInfo of objects) {
            completed += 1;
            await this.reportProgress(options, {
                phase: "object",
                message: `${categoryKey}: ${objectInfo.name}`,
                completed,
                total
            });
            try {
                await this.collectSingleObject(connection, metadata, categoryKey, objectInfo, options);
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                metadata.warnings.push(`No se pudo exportar ${categoryKey}:${objectInfo.name} -> ${message}`);
            }
        }
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
                    column_name: field?.name,
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
                sql = await this.mcpClient.getQuerySql(connection, objectName, SecondBrainService.QUERY_TIMEOUT_MS);
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
            const controls = await this.mcpClient.getControls(
                connection,
                objectType,
                objectName,
                SecondBrainService.UI_TIMEOUT_MS
            );
            const properties = await this.mcpClient.getFormReportProperties(
                connection,
                objectType,
                objectName,
                SecondBrainService.UI_TIMEOUT_MS
            );
            const recordSource = properties.find((item) => item.name.toLowerCase() === "recordsource")?.value ?? "";
            const doc = await this.mcpClient.getObjectDocument(
                connection,
                categoryKey,
                objectName,
                objectInfo.metadata,
                SecondBrainService.UI_TIMEOUT_MS
            );

            const uiObj = {
                name: objectName,
                record_source: recordSource,
                controls: controls.map((ctrl) => ({
                    name: ctrl.name,
                    control_type: ctrl.type_name,
                    control_source: ctrl.control_source,
                    caption: ctrl.caption
                })),
                code: doc.content
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
                code: doc.content
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

        const columnsByTable = groupBy(metadata.columns, (item) => String(item.table_name ?? ""));
        const pksByTable = groupBy(metadata.primary_keys, (item) => String(item.table_name ?? ""));
        const idxByTable = groupBy(metadata.indexes, (item) => String(item.table_name ?? ""));

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
            "_health"
        ]);

        const noteDrafts = new Map<string, NoteDraft>();

        const registerNote = (notePath: string, lines: string[], outgoing: Iterable<string> = []): void => {
            const validOutgoing = new Set<string>();
            for (const target of outgoing) {
                if (target && target !== notePath && allNoteTargets.has(target)) {
                    validOutgoing.add(target);
                }
            }
            noteDrafts.set(notePath, { notePath, lines, outgoing: validOutgoing });
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

        for (const table of metadata.tables) {
            const name = String(table.table_name ?? "");
            const cols = columnsByTable.get(name) ?? [];
            const pks = (pksByTable.get(name) ?? []).map((item) => String(item.column_name ?? ""));
            const idx = idxByTable.get(name) ?? [];
            const tableNotePath = `tables/${sanitize(name)}`;

            const outgoing = new Set<string>();
            for (const rel of metadata.relationships) {
                if (String(rel.table ?? "") === name || String(rel.foreign_table ?? "") === name) {
                    outgoing.add(`relationships/${sanitize(String(rel.name ?? "(sin nombre)"))}`);
                }
            }
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
                macroTargets
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
                macroTargets
            );
            const moduleNotePath = `modules/${sanitize(name)}`;
            const mocPath = mocByNotePath.get(moduleNotePath);
            if (mocPath) {
                outgoing.add(mocPath);
            }
            const content = [
                `# Module: ${name}`,
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
        const indexContent = [
            `# Access Second Brain: ${metadata.database}`,
            "",
            `Generated: ${metadata.generated}`,
            "",
            `Scope: ${this.scopeLabel(scope)}`,
            "",
            "## Stats",
            "",
            ...Object.entries(stats).map(([key, value]) => `- ${key}: ${value}`),
            "",
            "## Main Notes",
            "",
            "- [[_overview]]",
            "- [[_health]]",
            ""
        ].join("\n");

        const overviewContent = [
            `# Overview: ${metadata.database}`,
            "",
            `Este paquete SecondBrain fue generado con alcance: ${this.scopeLabel(scope)}.`,
            "",
            "## Resumen",
            "",
            ...Object.entries(stats).map(([key, value]) => `- ${key}: ${value}`),
            "",
            "## MOCs",
            "",
            ...(mocGroups.size > 0 ? Array.from(mocGroups.keys()).sort().map((mocPath) => `- ${wiki(mocPath)}`) : ["- None"]),
            ""
        ].join("\n");

        const healthContent = [
            "# Health",
            "",
            metadata.warnings.length > 0 ? "## Warnings" : "## Warnings",
            "",
            ...(metadata.warnings.length > 0
                ? metadata.warnings.map((warning) => `- ${warning}`)
                : ["- None"]),
            ""
        ].join("\n");

        registerNote("_index", indexContent.split("\n"), ["_overview", "_health"]);
        registerNote("_overview", overviewContent.split("\n"), ["_index", "_health", ...mocGroups.keys()]);
        registerNote("_health", healthContent.split("\n"), ["_index", "_overview"]);

        const incomingByTarget = new Map<string, Set<string>>();
        for (const [sourcePath, draft] of noteDrafts.entries()) {
            for (const target of draft.outgoing) {
                if (!incomingByTarget.has(target)) {
                    incomingByTarget.set(target, new Set<string>());
                }
                incomingByTarget.get(target)?.add(sourcePath);
            }
        }

        for (const [notePath, draft] of noteDrafts.entries()) {
            const outgoingSorted = Array.from(draft.outgoing).sort();
            const incomingSorted = Array.from(incomingByTarget.get(notePath) ?? []).sort();

            draft.lines.push(
                "",
                "## Links",
                "",
                ...(outgoingSorted.length > 0 ? outgoingSorted.map((target) => `- ${wiki(target)}`) : ["- None"]),
                "",
                "## Referenciado por",
                "",
                ...(incomingSorted.length > 0 ? incomingSorted.map((source) => `- ${wiki(source)}`) : ["- None"]),
                ""
            );

            await fs.writeFile(path.join(vaultDir, `${notePath}.md`), draft.lines.join("\n"), "utf-8");
        }
    }

    private async reportProgress(options: SecondBrainExportOptions | undefined, event: SecondBrainProgressEvent): Promise<void> {
        await options?.onProgress?.(event);
    }

    private buildUiNote(kind: "Form" | "Report", item: Record<string, unknown>, dependencies: Set<string>): string {
        const controls = Array.isArray(item.controls) ? item.controls : [];
        const sortedDependencies = Array.from(dependencies).sort();
        const header = [
            `# ${kind}: ${item.name ?? ""}`,
            "",
            `- RecordSource: ${item.record_source ?? ""}`,
            `- Controls: ${controls.length}`,
            "",
            "## Depends On",
            "",
            ...(sortedDependencies.length > 0 ? sortedDependencies.map((target) => `- ${wiki(target)}`) : ["- None"]),
            "",
            "## Controls",
            "",
            "| Name | Type | Source | Caption |",
            "|---|---|---|---|"
        ];

        for (const control of controls as Array<Record<string, unknown>>) {
            header.push(
                `| ${control.name ?? ""} | ${control.control_type ?? ""} | ${control.control_source ?? ""} | ${control.caption ?? ""} | `
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

function sanitize(value: string): string {
    return value.replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, "-").replace(/-+/g, "-").trim();
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
    macroTargets: Map<string, string>
): Set<string> {
    const links = new Set<string>();
    const patterns: Array<{ regex: RegExp; map: Map<string, string> }> = [
        { regex: /OpenForm\s*\(?\s*"([^"]+)"/gi, map: formTargets },
        { regex: /OpenReport\s*\(?\s*"([^"]+)"/gi, map: reportTargets },
        { regex: /OpenQuery\s*\(?\s*"([^"]+)"/gi, map: queryTargets },
        { regex: /QueryDefs\s*\(\s*"([^"]+)"/gi, map: queryTargets },
        { regex: /RunMacro\s*\(?\s*"([^"]+)"/gi, map: macroTargets }
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
