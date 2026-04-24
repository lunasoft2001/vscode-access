import * as vscode from "vscode";
import * as fs from "node:fs";
import * as path from "node:path";
import { CategoryNode, DetailNode, ObjectNode } from "./models/treeNodes";
import { McpAccessClient } from "./mcp/mcpAccessClient";
import { ACCESS_CATEGORIES } from "./models/types";
import { AccessTreeProvider } from "./providers/accessTreeProvider";
import { ConnectionStore } from "./services/connectionStore";
import { SecondBrainService } from "./services/secondBrainService";
import { BulkExportService } from "./services/bulkExportService";
import { ExportObjectsService } from "./services/exportObjectsService";
import { offerAccessRestart, restartAccessProcesses } from "./utils/accessRecovery";
import { rt } from "./utils/runtimeL10n";

export function activate(context: vscode.ExtensionContext): void {
    const configuration = () => vscode.workspace.getConfiguration("accessExplorer");
    const connectionStore = new ConnectionStore(context);
    const mcpClient = new McpAccessClient(configuration, context);
    const secondBrainService = new SecondBrainService(mcpClient, context.globalStorageUri.fsPath);
    const bulkExportService = new BulkExportService(mcpClient, context.globalStorageUri.fsPath);
    const exportObjectsService = new ExportObjectsService(bulkExportService);
    const treeProvider = new AccessTreeProvider(connectionStore, mcpClient);
    const secondBrainOutput = vscode.window.createOutputChannel("Access Explorer SecondBrain");
    const secondBrainStatusBar = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 99);
    secondBrainStatusBar.tooltip = rt("secondBrain.status.tooltip");

    context.subscriptions.push(mcpClient);
    context.subscriptions.push(secondBrainOutput, secondBrainStatusBar);

    // Tracks opened Access code documents so they can be saved back
    interface AccessCodeMeta {
        connection: import("./models/types").AccessConnection;
        objectType: "module" | "form" | "report";
        objectName: string;
        procedureName?: string;
        replaceStartLine?: number;
        replaceCount?: number;
        isNew?: boolean;
    }
    interface AccessQueryMeta {
        connection: import("./models/types").AccessConnection;
        queryName?: string;
        isNew?: boolean;
    }
    interface TableDesignerFieldDraft {
        originalName?: string;
        name: string;
        type: string;
        size?: number;
        required: boolean;
        existing?: boolean;
    }
    const codeDocuments = new Map<string, AccessCodeMeta>();
    const queryDocuments = new Map<string, AccessQueryMeta>();
    context.subscriptions.push(
        vscode.workspace.onDidCloseTextDocument((doc) => {
            codeDocuments.delete(doc.uri.toString());
            queryDocuments.delete(doc.uri.toString());
            void updateEditorActionContexts();
        })
    );

    context.subscriptions.push(mcpClient);

    // Status bar item: shows the active connection for the SQL editor
    let activeSqlConnection: import("./models/types").AccessConnection | undefined;
    const sqlStatusBar = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 100);
    sqlStatusBar.command = "accessExplorer.pickSqlConnection";
    sqlStatusBar.tooltip = rt("sql.status.tooltip");

    function updateSqlStatusBar(): void {
        const editor = vscode.window.activeTextEditor;
        const isSql = editor?.document.languageId === "sql";
        if (isSql) {
            sqlStatusBar.text = activeSqlConnection
                ? `$(database) ${activeSqlConnection.name}`
                : `$(database) Conectar Access...`;
            sqlStatusBar.show();
        } else {
            sqlStatusBar.hide();
        }
    }

    async function updateEditorActionContexts(): Promise<void> {
        const editor = vscode.window.activeTextEditor;
        const key = editor?.document.uri.toString();
        const codeMeta = key ? codeDocuments.get(key) : undefined;
        const canSaveCode = !!codeMeta;
        const canSaveModule = codeMeta?.objectType === "module";

        await vscode.commands.executeCommand("setContext", "accessExplorer.canSaveCode", canSaveCode);
        await vscode.commands.executeCommand("setContext", "accessExplorer.canSaveModule", canSaveModule);
        await vscode.commands.executeCommand("setContext", "accessExplorer.canSaveQuery", !!key && queryDocuments.has(key));
        updateSqlStatusBar();
    }

    function trackCodeDocument(
        document: vscode.TextDocument,
        meta: AccessCodeMeta
    ): void {
        codeDocuments.set(document.uri.toString(), meta);
        void updateEditorActionContexts();
    }

    function trackQueryDocument(
        document: vscode.TextDocument,
        meta: AccessQueryMeta
    ): void {
        queryDocuments.set(document.uri.toString(), meta);
        activeSqlConnection = meta.connection;
        void updateEditorActionContexts();
    }

    function getActiveSqlMeta(editor: vscode.TextEditor | undefined): {
        document: vscode.TextDocument;
        meta?: AccessQueryMeta;
    } | undefined {
        if (!editor || editor.document.languageId !== "sql") {
            return undefined;
        }

        return {
            document: editor.document,
            meta: queryDocuments.get(editor.document.uri.toString())
        };
    }

    context.subscriptions.push(sqlStatusBar);
    context.subscriptions.push(vscode.window.onDidChangeActiveTextEditor(() => {
        void updateEditorActionContexts();
    }));
    void updateEditorActionContexts();

    const treeView = vscode.window.createTreeView("accessExplorerView", {
        treeDataProvider: treeProvider,
        showCollapseAll: true
    });

    context.subscriptions.push(treeView);

    void vscode.window.withProgress(
        {
            location: vscode.ProgressLocation.Notification,
            title: "Comprobando requisitos de Access Explorer...",
            cancellable: false
        },
        async () => {
            try {
                await mcpClient.ensurePrerequisites();
                treeProvider.refresh();
                await registerMcpServerSilently(context, mcpClient);
            } catch {
                // The client already shows diagnostics and recovery guidance.
            }
        }
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.registerMcpServer", async () => {
            try {
                const changed = await registerMcpServerSilently(context, mcpClient);
                if (changed) {
                    vscode.window.showInformationMessage(
                        "Servidor MCP de Access registrado en mcp.json de usuario. Recarga VS Code para que Copilot lo detecte.",
                        "Recargar"
                    ).then(action => {
                        if (action === "Recargar") {
                            void vscode.commands.executeCommand("workbench.action.reloadWindow");
                        }
                    });
                } else {
                    vscode.window.showInformationMessage(rt("sql.connection.alreadyRegisteredUpdated"));
                }
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`No se pudo registrar el servidor MCP: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.refresh", async () => {
            treeProvider.refresh();
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.reconnect", async () => {
            try {
                await mcpClient.reconnect();
                treeProvider.refresh();
                vscode.window.showInformationMessage("MCP-Access reconectado.");
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                const recovered = await offerAccessRestart(message);
                if (recovered) {
                    await mcpClient.reconnect();
                    treeProvider.refresh();
                    vscode.window.showInformationMessage("MCP-Access reconectado tras reiniciar Access.");
                    return;
                }

                vscode.window.showErrorMessage(`No se pudo reconectar MCP-Access: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.restartAccessProcesses", async () => {
            try {
                await restartAccessProcesses();
                await mcpClient.disconnect();
                treeProvider.refresh();
                vscode.window.showInformationMessage("Se reiniciaron procesos de Access.");
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`No se pudo reiniciar Access: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.showMcpRuntime", async () => {
            try {
                const info = await vscode.window.withProgress(
                    {
                        location: vscode.ProgressLocation.Notification,
                        title: "Resolviendo runtime MCP...",
                        cancellable: false
                    },
                    async () => await mcpClient.getMcpRuntimeInfo()
                );

                const summary = [
                    `Python: ${info.pythonCommand}`,
                    `Script MCP: ${info.resolvedServerScriptPath}`,
                    `Runtime gestionado: ${info.managedBaseDir}`
                ].join("\n");

                const action = await vscode.window.showInformationMessage(
                    "Runtime MCP de Access Explorer resuelto.",
                    "Copiar bloque mcp.json",
                    "Abrir carpeta runtime",
                    "Ver detalle"
                );

                if (action === "Copiar bloque mcp.json") {
                    await vscode.env.clipboard.writeText(info.mcpJsonSnippet);
                    vscode.window.showInformationMessage("Bloque mcp.json copiado al portapapeles.");
                    return;
                }

                if (action === "Abrir carpeta runtime") {
                    await vscode.commands.executeCommand("revealFileInOS", vscode.Uri.file(info.managedBaseDir));
                    return;
                }

                if (action === "Ver detalle") {
                    const doc = await vscode.workspace.openTextDocument({
                        language: "json",
                        content: JSON.stringify(
                            {
                                runtime: {
                                    pythonCommand: info.pythonCommand,
                                    resolvedServerScriptPath: info.resolvedServerScriptPath,
                                    managedBaseDir: info.managedBaseDir,
                                    managedServerScriptPath: info.managedServerScriptPath
                                },
                                mcpJsonSnippet: JSON.parse(info.mcpJsonSnippet)
                            },
                            null,
                            2
                        )
                    });
                    await vscode.window.showTextDocument(doc, { preview: false });
                    vscode.window.showInformationMessage(summary);
                }
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`No se pudo resolver el runtime MCP: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.searchObjects", async () => {
            const connections = connectionStore.getAll();
            if (connections.length === 0) {
                vscode.window.showInformationMessage("No hay conexiones Access configuradas.");
                return;
            }

            try {
                const picks = await vscode.window.withProgress(
                    {
                        location: vscode.ProgressLocation.Notification,
                        title: "Buscando objetos Access",
                        cancellable: false
                    },
                    async () => {
                        const items: Array<{
                            label: string;
                            description: string;
                            detail: string;
                            node: ObjectNode;
                        }> = [];

                        for (const connection of connections) {
                            for (const category of ACCESS_CATEGORIES) {
                                if (category.key === "relationships") {
                                    const objects = await mcpClient.listRelationships(connection);
                                    for (const object of objects) {
                                        items.push({
                                            label: object.name,
                                            description: `${connection.name}${rt("object.descriptionSeparator")}${category.label}`,
                                            detail: connection.dbPath,
                                            node: new ObjectNode(connection, category.key, object)
                                        });
                                    }
                                    continue;
                                }

                                if (category.key === "references") {
                                    const objects = await mcpClient.listReferences(connection);
                                    for (const object of objects) {
                                        items.push({
                                            label: object.name,
                                            description: `${connection.name}${rt("object.descriptionSeparator")}${category.label}`,
                                            detail: connection.dbPath,
                                            node: new ObjectNode(connection, category.key, object)
                                        });
                                    }
                                    continue;
                                }

                                if (!category.toolObjectType) {
                                    continue;
                                }

                                const objects = await mcpClient.listObjects(connection, category.toolObjectType);
                                for (const object of objects) {
                                    items.push({
                                        label: object.name,
                                        description: `${connection.name}${rt("object.descriptionSeparator")}${category.label}`,
                                        detail: connection.dbPath,
                                        node: new ObjectNode(connection, category.key, object)
                                    });
                                }
                            }
                        }

                        return items;
                    }
                );

                const selected = await vscode.window.showQuickPick(picks, {
                    title: "Buscar objeto Access",
                    matchOnDescription: true,
                    matchOnDetail: true,
                    placeHolder: "Escribe parte del nombre del objeto"
                });

                if (!selected) {
                    return;
                }

                await vscode.commands.executeCommand("accessExplorer.showDetails", selected.node);
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                const recovered = await offerAccessRestart(message);
                if (recovered) {
                    vscode.window.showInformationMessage("Repite la busqueda tras reiniciar Access.");
                    return;
                }

                vscode.window.showErrorMessage(`No se pudo buscar objetos: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.addConnection", async () => {
            const picked = await vscode.window.showOpenDialog({
                canSelectFiles: true,
                canSelectFolders: false,
                canSelectMany: false,
                openLabel: "Seleccionar base Access",
                filters: {
                    "Access Database": ["accdb", "mdb"]
                }
            });

            if (!picked || picked.length === 0) {
                return;
            }

            const dbPath = picked[0].fsPath;
            const defaultName = dbPath.split(/[\\/]/).pop() ?? "Access DB";

            const name = await vscode.window.showInputBox({
                prompt: "Nombre de la conexion",
                value: defaultName,
                validateInput: (value) => (value.trim() ? undefined : "El nombre es obligatorio")
            });

            if (!name) {
                return;
            }

            await connectionStore.add(name.trim(), dbPath);
            treeProvider.refresh();
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.removeConnection", async (node?: any) => {
            const connections = connectionStore.getAll();
            if (connections.length === 0) {
                return;
            }

            let selected = node?.connection;
            if (!selected) {
                const pick = await vscode.window.showQuickPick(
                    connections.map((connection) => ({
                        label: connection.name,
                        description: connection.dbPath,
                        connection
                    })),
                    { title: "Eliminar conexion Access" }
                );
                selected = pick?.connection;
            }

            if (!selected) {
                return;
            }

            const confirm = await vscode.window.showWarningMessage(
                `Eliminar la conexion ${selected.name}?`,
                { modal: true },
                "Eliminar"
            );

            if (confirm !== "Eliminar") {
                return;
            }

            await connectionStore.remove(selected.id);
            treeProvider.refresh();
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.showDetails", async (node?: ObjectNode | DetailNode) => {
            if (!node) {
                return;
            }

            try {
                const objectDoc = await vscode.window.withProgress(
                    {
                        location: vscode.ProgressLocation.Notification,
                        title: `Abriendo ${node.objectInfo.name}`,
                        cancellable: false
                    },
                    async () => {
                        if (node instanceof DetailNode) {
                            return await openDetailNode(mcpClient, node, treeView, treeProvider);
                        }

                        return await mcpClient.getObjectDocument(
                            node.connection,
                            node.categoryKey,
                            node.objectInfo.name,
                            node.objectInfo.metadata
                        );
                    }
                );

                if (!objectDoc) {
                    return;
                }

                const doc = await vscode.workspace.openTextDocument({
                    content: objectDoc.content,
                    language: objectDoc.language
                });

                // Track for "Save to Access"
                if (objectDoc.language === "vb" && "codeMeta" in objectDoc) {
                    const meta = { ...((objectDoc as any).codeMeta as AccessCodeMeta) };
                    if (
                        (meta.objectType === "form" || meta.objectType === "report")
                        && typeof meta.replaceStartLine !== "number"
                    ) {
                        meta.replaceStartLine = 1;
                        meta.replaceCount = doc.lineCount;
                    }
                    trackCodeDocument(doc, meta);
                }

                if (objectDoc.language === "sql" && node.categoryKey === "queries") {
                    trackQueryDocument(doc, {
                        connection: node.connection,
                        queryName: node.objectInfo.name
                    });
                }

                await vscode.window.showTextDocument(doc, { preview: true });
                await updateEditorActionContexts();
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                const recovered = await offerAccessRestart(message);
                if (recovered) {
                    vscode.window.showInformationMessage("Vuelve a hacer clic en el objeto para reintentar.");
                    return;
                }

                vscode.window.showErrorMessage(`No se pudo abrir el objeto: ${message}`);
            }
        })
    );

    // Helper to pick a connection
    async function pickConnection(title?: string): Promise<import("./models/types").AccessConnection | undefined> {
        const connections = connectionStore.getAll();
        if (connections.length === 0) {
            vscode.window.showInformationMessage("No hay conexiones Access configuradas.");
            return undefined;
        }
        if (connections.length === 1) {
            return connections[0];
        }
        const pick = await vscode.window.showQuickPick(
            connections.map((c) => ({ label: c.name, description: c.dbPath, connection: c })),
            { title: title ?? "Seleccionar base de datos Access" }
        );
        return pick?.connection;
    }

    async function pickObjectFromCategory(
        connection: import("./models/types").AccessConnection,
        objectType: "module" | "query" | "table",
        title: string
    ): Promise<string | undefined> {
        const objects = await mcpClient.listObjects(connection, objectType);
        if (objects.length === 0) {
            const objectLabel = objectType === "module"
                ? "modulos"
                : objectType === "query"
                    ? "consultas"
                    : "tablas";
            vscode.window.showInformationMessage(`No hay ${objectLabel} en ${connection.name}.`);
            return undefined;
        }

        const pick = await vscode.window.showQuickPick(
            objects
                .slice()
                .sort((left, right) => left.name.localeCompare(right.name, "es"))
                .map((object) => ({
                    label: object.name,
                    objectName: object.name
                })),
            { title }
        );

        return pick?.objectName;
    }

    function escapeVbaStringLiteral(value: string): string {
        return value.replaceAll("\"", "\"\"");
    }

    function toVbaStringExpression(value: string): string {
        const lines = value.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
        if (lines.length === 0) {
            return "\"\"";
        }

        return lines
            .map((line) => `"${escapeVbaStringLiteral(line)}"`)
            .join(" & vbCrLf & ");
    }

    async function saveQueryDocumentToAccess(
        editor: vscode.TextEditor,
        explicitQueryName?: string
    ): Promise<void> {
        const activeSql = getActiveSqlMeta(editor);
        if (!activeSql) {
            vscode.window.showInformationMessage("No hay un editor SQL activo.");
            return;
        }

        const trackedMeta = activeSql.meta;
        const connection = trackedMeta?.connection
            ?? activeSqlConnection
            ?? await pickConnection("Seleccionar base de datos Access para guardar la consulta");

        if (!connection) {
            return;
        }

        activeSqlConnection = connection;
        updateSqlStatusBar();

        const queryName = explicitQueryName
            ?? await vscode.window.showInputBox({
                prompt: "Nombre de la consulta en Access",
                value: trackedMeta?.queryName ?? "",
                validateInput: (value) => (value.trim() ? undefined : "El nombre es obligatorio")
            });

        if (!queryName?.trim()) {
            return;
        }

        const finalQueryName = queryName.trim();
        const sql = editor.document.getText();
        if (!sql.trim()) {
            vscode.window.showWarningMessage(rt("query.sql.empty"));
            return;
        }

        const escapedQueryName = escapeVbaStringLiteral(finalQueryName);
        const sqlExpression = toVbaStringExpression(sql);
        const expression = [
            "Dim db As DAO.Database",
            "Dim qdf As DAO.QueryDef",
            "Set db = CurrentDb()",
            "On Error Resume Next",
            `Set qdf = db.QueryDefs("${escapedQueryName}")`,
            "If Err.Number <> 0 Then",
            "    Err.Clear",
            `    Set qdf = db.CreateQueryDef("${escapedQueryName}", ${sqlExpression})`,
            "Else",
            `    qdf.SQL = ${sqlExpression}`,
            "End If",
            "On Error GoTo 0"
        ].join("\n");

        await vscode.window.withProgress(
            {
                location: vscode.ProgressLocation.Notification,
                title: `Guardando consulta ${finalQueryName}...`,
                cancellable: false
            },
            () => mcpClient.evalVba(connection, expression, 30000)
        );

        trackQueryDocument(editor.document, {
            connection,
            queryName: finalQueryName
        });

        vscode.window.showInformationMessage(`Consulta guardada en Access: ${finalQueryName}`);
        treeProvider.refresh();
    }

    async function openNewQueryEditor(
        connection: import("./models/types").AccessConnection,
        initialSql = "",
        queryName?: string
    ): Promise<void> {
        const doc = await vscode.workspace.openTextDocument({
            language: "sql",
            content: initialSql
        });
        trackQueryDocument(doc, { connection, queryName, isNew: !queryName });
        await vscode.window.showTextDocument(doc, { preview: false });
    }

    function quoteSqlIdentifier(value: string): string {
        return `[${value.replaceAll("]", "]]")}]`;
    }

    function toAccessFieldTypeSql(field: import("./models/types").AccessTableFieldInfo): string {
        const rawType = String(field.type ?? "TEXT").trim();
        const normalized = rawType.toUpperCase();
        const supportsSize = /^(TEXT|CHAR|VARCHAR|VARBINARY|BINARY)$/i.test(normalized);
        const sizeSuffix = supportsSize && typeof field.size === "number" && field.size > 0
            ? `(${field.size})`
            : "";
        return `${rawType}${sizeSuffix}`;
    }

    function getNextAvailableFieldName(fields: import("./models/types").AccessTableFieldInfo[]): string {
        const existingNames = new Set(
            fields
                .map((field) => String(field.name ?? "").trim().toLowerCase())
                .filter(Boolean)
        );
        const baseName = "NuevoCampo";
        let counter = 1;
        let candidate = baseName;
        while (existingNames.has(candidate.toLowerCase())) {
            counter += 1;
            candidate = `${baseName}${counter}`;
        }
        return candidate;
    }

    function buildCreateTableDdlTemplate(tableName: string): string {
        const quotedName = quoteSqlIdentifier(tableName || "NuevaTabla");
        return [
            `CREATE TABLE ${quotedName} (`,
            "    [Id] AUTOINCREMENT PRIMARY KEY,",
            "    [Nombre] TEXT(255) NOT NULL",
            ");"
        ].join("\n");
    }

    function buildEditTableDdlTemplate(
        tableName: string,
        fields: import("./models/types").AccessTableFieldInfo[]
    ): string {
        return buildAddColumnDdlTemplate(tableName, fields);
    }

    function buildAddColumnDdlTemplate(
        tableName: string,
        fields: import("./models/types").AccessTableFieldInfo[]
    ): string {
        const quotedTableName = quoteSqlIdentifier(tableName);
        const candidate = getNextAvailableFieldName(fields);

        return [
            `ALTER TABLE ${quotedTableName}`,
            `ADD COLUMN ${quoteSqlIdentifier(candidate)} TEXT(100);`
        ].join("\n");
    }

    function buildAlterColumnDdlTemplate(
        tableName: string,
        field: import("./models/types").AccessTableFieldInfo
    ): string {
        const quotedTableName = quoteSqlIdentifier(tableName);
        const quotedFieldName = quoteSqlIdentifier(field.name);
        return [
            `ALTER TABLE ${quotedTableName}`,
            `ALTER COLUMN ${quotedFieldName} ${toAccessFieldTypeSql(field)};`
        ].join("\n");
    }

    function buildDropColumnDdlTemplate(tableName: string, fieldName: string): string {
        const quotedTableName = quoteSqlIdentifier(tableName);
        return [
            `ALTER TABLE ${quotedTableName}`,
            `DROP COLUMN ${quoteSqlIdentifier(fieldName)};`
        ].join("\n");
    }

    async function pickTableField(
        fields: import("./models/types").AccessTableFieldInfo[],
        title: string
    ): Promise<import("./models/types").AccessTableFieldInfo | undefined> {
        if (fields.length === 0) {
            vscode.window.showWarningMessage("La tabla no tiene campos disponibles para esta operacion.");
            return undefined;
        }

        const pick = await vscode.window.showQuickPick(
            fields.map((field) => ({
                label: field.name,
                description: [
                    field.type ?? "",
                    typeof field.size === "number" ? `(${field.size})` : "",
                    field.required ? "required" : "optional"
                ].filter(Boolean).join(" "),
                field
            })),
            { title }
        );

        return pick?.field;
    }

    function normalizeFieldType(value: string): string {
        return value.trim().toUpperCase();
    }

    function buildFieldTypeClause(field: TableDesignerFieldDraft): string {
        const type = normalizeFieldType(field.type || "TEXT");
        const supportsSize = /^(TEXT|CHAR|VARCHAR|VARBINARY|BINARY)$/i.test(type);
        const size = supportsSize && typeof field.size === "number" && field.size > 0
            ? `(${field.size})`
            : "";
        return `${type}${size}`;
    }

    function buildCreateTableFieldDefinition(field: TableDesignerFieldDraft): string {
        const type = normalizeFieldType(field.type || "TEXT");
        const nullability = field.required ? " NOT NULL" : "";
        if (type === "AUTOINCREMENT") {
            return `${quoteSqlIdentifier(field.name)} AUTOINCREMENT${nullability}`;
        }
        return `${quoteSqlIdentifier(field.name)} ${buildFieldTypeClause(field)}${nullability}`;
    }

    function sanitizeTableDesignerFields(fields: TableDesignerFieldDraft[]): TableDesignerFieldDraft[] {
        return fields
            .map((field) => ({
                ...field,
                name: String(field.name ?? "").trim(),
                type: normalizeFieldType(String(field.type ?? "TEXT")),
                size: typeof field.size === "number" && Number.isFinite(field.size) && field.size > 0
                    ? Math.trunc(field.size)
                    : undefined,
                required: Boolean(field.required)
            }))
            .filter((field) => field.name);
    }

    function validateTableDesignerInput(
        tableName: string,
        fields: TableDesignerFieldDraft[]
    ): { tableName: string; fields: TableDesignerFieldDraft[] } {
        const finalTableName = tableName.trim();
        if (!finalTableName) {
            throw new Error("El nombre de la tabla es obligatorio.");
        }

        const finalFields = sanitizeTableDesignerFields(fields);
        if (finalFields.length === 0) {
            throw new Error("Debes definir al menos un campo.");
        }

        const names = new Set<string>();
        for (const field of finalFields) {
            const normalizedName = field.name.toLowerCase();
            if (names.has(normalizedName)) {
                throw new Error(`El campo "${field.name}" está duplicado.`);
            }
            names.add(normalizedName);
        }

        return { tableName: finalTableName, fields: finalFields };
    }

    function buildCreateTableStatements(
        tableName: string,
        fields: TableDesignerFieldDraft[]
    ): string[] {
        return [[
            `CREATE TABLE ${quoteSqlIdentifier(tableName)} (`,
            fields.map((field) => `    ${buildCreateTableFieldDefinition(field)}`).join(",\n"),
            ");"
        ].join("\n")];
    }

    function buildAlterFieldStatement(tableName: string, field: TableDesignerFieldDraft): string {
        const nullability = field.required ? " NOT NULL" : "";
        return [
            `ALTER TABLE ${quoteSqlIdentifier(tableName)}`,
            `ALTER COLUMN ${quoteSqlIdentifier(field.name)} ${buildFieldTypeClause(field)}${nullability};`
        ].join("\n");
    }

    function buildEditTableStatements(
        tableName: string,
        originalFields: import("./models/types").AccessTableFieldInfo[],
        nextFields: TableDesignerFieldDraft[]
    ): string[] {
        const originalByName = new Map(
            originalFields.map((field) => [field.name.toLowerCase(), field])
        );
        const statements: string[] = [];

        const desiredExistingNames = new Set(
            nextFields
                .filter((field) => field.existing)
                .map((field) => field.name.toLowerCase())
        );

        for (const originalField of originalFields) {
            if (!desiredExistingNames.has(originalField.name.toLowerCase())) {
                statements.push(buildDropColumnDdlTemplate(tableName, originalField.name));
            }
        }

        for (const field of nextFields) {
            if (!field.existing) {
                statements.push([
                    `ALTER TABLE ${quoteSqlIdentifier(tableName)}`,
                    `ADD COLUMN ${buildCreateTableFieldDefinition(field)};`
                ].join("\n"));
                continue;
            }

            const original = originalByName.get(field.name.toLowerCase());
            if (!original) {
                statements.push([
                    `ALTER TABLE ${quoteSqlIdentifier(tableName)}`,
                    `ADD COLUMN ${buildCreateTableFieldDefinition(field)};`
                ].join("\n"));
                continue;
            }

            const sameType = normalizeFieldType(String(original.type ?? "TEXT")) === normalizeFieldType(field.type);
            const sameSize = (original.size ?? undefined) === (field.size ?? undefined);
            const sameRequired = Boolean(original.required) === Boolean(field.required);
            if (!sameType || !sameSize || !sameRequired) {
                statements.push(buildAlterFieldStatement(tableName, field));
            }
        }

        return statements;
    }

    async function executeTableDesignerStatements(
        connection: import("./models/types").AccessConnection,
        tableName: string,
        statements: string[],
        mode: "create" | "edit"
    ): Promise<void> {
        if (statements.length === 0) {
            vscode.window.showInformationMessage("No hay cambios de estructura para aplicar.");
            return;
        }

        const actionLabel = mode === "create" ? "crear" : "modificar";
        const confirm = await vscode.window.showWarningMessage(
            mode === "create" ? "Crear tabla" : "Aplicar cambios de tabla",
            {
                modal: true,
                detail: `Se van a ejecutar ${statements.length} sentencia(s) para ${actionLabel} "${tableName}" en "${connection.name}".`
            },
            "Aplicar"
        );

        if (confirm !== "Aplicar") {
            return;
        }

        const results: Array<{ sql: string; payload: unknown }> = [];
        await vscode.window.withProgress(
            {
                location: vscode.ProgressLocation.Notification,
                title: mode === "create" ? `Creando tabla ${tableName}...` : `Aplicando cambios en ${tableName}...`,
                cancellable: false
            },
            async (progress) => {
                for (let index = 0; index < statements.length; index += 1) {
                    const sql = statements[index];
                    progress.report({ message: `${index + 1}/${statements.length}` });
                    const result = await mcpClient.executeDml(connection, sql);
                    results.push({ sql, payload: result.payload });
                }
            }
        );

        treeProvider.refresh();
        const doc = await vscode.workspace.openTextDocument({
            content: [
                `# Resultado diseno de tabla`,
                "",
                `- Conexion: ${connection.name}`,
                `- Tabla: ${tableName}`,
                `- Operacion: ${mode === "create" ? "create" : "edit"}`,
                `- Sentencias ejecutadas: ${results.length}`,
                "",
                ...results.flatMap((entry, index) => ([
                    `## Sentencia ${index + 1}`,
                    "",
                    "```sql",
                    entry.sql,
                    "```",
                    "",
                    "```json",
                    JSON.stringify(entry.payload, null, 2),
                    "```",
                    ""
                ]))
            ].join("\n"),
            language: "markdown"
        });
        await vscode.window.showTextDocument(doc, { preview: true, viewColumn: vscode.ViewColumn.Beside });
        vscode.window.showInformationMessage(
            mode === "create"
                ? `Tabla creada: ${tableName}`
                : `Cambios de tabla aplicados: ${tableName}`
        );
    }

    function showTableDesignerWebview(options: {
        connection: import("./models/types").AccessConnection;
        mode: "create" | "edit";
        initialTableName: string;
        initialFields: TableDesignerFieldDraft[];
        onApply: (tableName: string, fields: TableDesignerFieldDraft[]) => Promise<void>;
    }): void {
        const fieldsJson = JSON.stringify(options.initialFields);
        const panel = vscode.window.createWebviewPanel(
            "accessTableDesigner",
            options.mode === "create"
                ? `Diseno de tabla: nueva tabla`
                : `Diseno de tabla: ${options.initialTableName}`,
            vscode.ViewColumn.Beside,
            { enableScripts: true, retainContextWhenHidden: true }
        );

        panel.webview.html = `<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#1e1e1e;color:#d4d4d4;display:flex;flex-direction:column;height:100vh;overflow:hidden;}
  .toolbar{padding:10px 12px;background:#252526;border-bottom:1px solid #3e3e42;display:flex;align-items:center;gap:10px;flex-wrap:wrap;}
  .title{font-size:13px;font-weight:600;min-width:140px;}
  .table-name{display:flex;align-items:center;gap:8px;flex:1;min-width:260px;}
  .table-name input{flex:1;background:#1f1f1f;border:1px solid #555;color:#ddd;padding:6px 8px;border-radius:4px;font-size:12px;}
  .btn{padding:6px 12px;font-size:12px;border:1px solid #4e4e52;background:#3a3a3c;color:#ccc;border-radius:4px;cursor:pointer;}
  .btn:hover{background:#4e4e52;}
  .btn.primary{background:#007acc;border-color:#007acc;color:#fff;}
  .btn.primary:hover{background:#1193f5;}
  .hint{padding:8px 12px;font-size:11px;color:#9d9d9d;background:#202020;border-bottom:1px solid #3e3e42;}
  .table-wrap{flex:1;overflow:auto;padding:10px;}
  table{width:100%;border-collapse:collapse;font-size:12px;}
  th{background:#2d2d30;color:#9cdcfe;font-weight:600;padding:6px 8px;border:1px solid #3e3e42;text-align:left;}
  td{padding:6px;border:1px solid #3e3e42;vertical-align:middle;}
  td input, td select{width:100%;background:#1f1f1f;border:1px solid #555;color:#ddd;padding:5px 6px;border-radius:4px;font-size:12px;}
  td input[type="checkbox"]{width:auto;transform:scale(1.1);}
  tr:nth-child(even) td{background:#252526;}
  .actions{display:flex;gap:6px;justify-content:center;}
  .status{padding:8px 12px;background:#007acc;color:#fff;font-size:11px;}
  .muted{opacity:.65;}
</style>
</head>
<body>
<div class="toolbar">
  <div class="title">${options.mode === "create" ? "Nueva tabla guiada" : "Editar tabla guiada"}</div>
  <div class="table-name">
    <label for="tableName">Tabla</label>
    <input id="tableName" value="${escapeHtml(options.initialTableName)}" ${options.mode === "edit" ? "disabled" : ""}/>
  </div>
  <button class="btn" id="addRow">Agregar campo</button>
  <button class="btn primary" id="apply">Aplicar</button>
</div>
<div class="hint">Modo guiado con MCP. Para renombrar un campo existente, elimínalo y crea otro nuevo.</div>
<div class="table-wrap">
  <table>
    <thead>
      <tr>
        <th style="width:34%">Campo</th>
        <th style="width:24%">Tipo</th>
        <th style="width:12%">Tamaño</th>
        <th style="width:12%">Required</th>
        <th style="width:18%">Acciones</th>
      </tr>
    </thead>
    <tbody id="tbody"></tbody>
  </table>
</div>
<div class="status" id="status">${escapeHtml(options.connection.name)}</div>
<script>
const vscode = acquireVsCodeApi();
const INITIAL_FIELDS = ${fieldsJson};
const TYPE_OPTIONS = ['TEXT','LONGTEXT','BYTE','INTEGER','LONG','SINGLE','DOUBLE','CURRENCY','DATETIME','YESNO','GUID','AUTOINCREMENT'];
let rows = INITIAL_FIELDS.map(f => ({...f}));

function esc(v){return String(v ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\"/g,'&quot;');}

function render() {
  const tbody = document.getElementById('tbody');
  tbody.innerHTML = '';
  rows.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = \`
      <td><input data-key="name" data-index="\${index}" value="\${esc(row.name)}" \${row.existing ? 'disabled' : ''}/></td>
      <td>
        <select data-key="type" data-index="\${index}">
          \${TYPE_OPTIONS.map(type => \`<option value="\${type}" \${String(row.type).toUpperCase()===type?'selected':''}>\${type}</option>\`).join('')}
        </select>
      </td>
      <td><input data-key="size" data-index="\${index}" type="number" min="1" value="\${row.size ?? ''}" /></td>
      <td style="text-align:center"><input data-key="required" data-index="\${index}" type="checkbox" \${row.required ? 'checked' : ''}/></td>
      <td>
        <div class="actions">
          <button class="btn" data-remove="\${index}">Eliminar</button>
        </div>
      </td>
    \`;
    tbody.appendChild(tr);
  });

  tbody.querySelectorAll('input[data-key], select[data-key]').forEach(el => {
    el.addEventListener('input', syncFromDom);
    el.addEventListener('change', syncFromDom);
  });
  tbody.querySelectorAll('button[data-remove]').forEach(btn => {
    btn.addEventListener('click', () => {
      rows.splice(Number(btn.getAttribute('data-remove')), 1);
      render();
    });
  });
  updateStatus();
}

function syncFromDom() {
  document.querySelectorAll('#tbody tr').forEach((tr, index) => {
    const nameEl = tr.querySelector('[data-key="name"]');
    const typeEl = tr.querySelector('[data-key="type"]');
    const sizeEl = tr.querySelector('[data-key="size"]');
    const requiredEl = tr.querySelector('[data-key="required"]');
    rows[index] = {
      ...rows[index],
      name: nameEl ? nameEl.value : rows[index].name,
      type: typeEl ? typeEl.value : rows[index].type,
      size: sizeEl && sizeEl.value ? Number(sizeEl.value) : undefined,
      required: Boolean(requiredEl && requiredEl.checked)
    };
  });
  updateStatus();
}

function updateStatus() {
  document.getElementById('status').textContent = rows.length + ' campo(s) preparados';
}

document.getElementById('addRow').addEventListener('click', () => {
  rows.push({ name: 'NuevoCampo', type: 'TEXT', size: 100, required: false, existing: false });
  render();
});

document.getElementById('apply').addEventListener('click', () => {
  syncFromDom();
  vscode.postMessage({
    command: 'apply',
    tableName: document.getElementById('tableName').value,
    fields: rows
  });
});

window.addEventListener('message', event => {
  if (event.data.command === 'status') {
    document.getElementById('status').textContent = String(event.data.text || '');
  }
});

render();
</script>
</body>
</html>`;

        panel.webview.onDidReceiveMessage(async (message) => {
            if (message.command !== "apply") {
                return;
            }

            try {
                await options.onApply(
                    String(message.tableName ?? options.initialTableName),
                    Array.isArray(message.fields) ? message.fields as TableDesignerFieldDraft[] : []
                );
                panel.webview.postMessage({ command: "status", text: "Cambios aplicados correctamente." });
            } catch (error) {
                const messageText = error instanceof Error ? error.message : String(error);
                panel.webview.postMessage({ command: "status", text: `Error: ${messageText}` });
                vscode.window.showErrorMessage(`Error en el diseno de tabla: ${messageText}`);
            }
        });
    }

    async function openSqlTemplateEditor(
        connection: import("./models/types").AccessConnection,
        sql: string
    ): Promise<void> {
        activeSqlConnection = connection;
        updateSqlStatusBar();
        const doc = await vscode.workspace.openTextDocument({
            language: "sql",
            content: sql
        });
        await vscode.window.showTextDocument(doc, { preview: false });
    }

    function stringifyResult(result: unknown): string {
        if (result === undefined || result === null) {
            return "Sin resultado.";
        }
        if (typeof result === "string") {
            return result;
        }

        try {
            return JSON.stringify(result, null, 2);
        } catch {
            return String(result);
        }
    }

    function findTopLevelModuleIssue(code: string): string | undefined {
        const lines = code.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
        let inProcedure = false;

        for (let index = 0; index < lines.length; index += 1) {
            const rawLine = lines[index];
            const trimmed = rawLine.trim();
            const lineNumber = index + 1;

            if (!trimmed || trimmed.startsWith("'")) {
                continue;
            }

            if (/^(Public|Private|Friend|Static)?\s*(Sub|Function|Property\s+(Get|Let|Set))\b/i.test(trimmed)) {
                inProcedure = true;
                continue;
            }

            if (/^End\s+(Sub|Function|Property)\b/i.test(trimmed)) {
                inProcedure = false;
                continue;
            }

            if (inProcedure) {
                continue;
            }

            if (
                /^(Option|Attribute|Version)\b/i.test(trimmed)
                || /^(Public|Private|Friend|Global)?\s*(Dim|Const|Declare|Enum|Type|Event)\b/i.test(trimmed)
                || /^End\s+(Enum|Type)\b/i.test(trimmed)
                || /^#(If|Else|ElseIf|End)\b/i.test(trimmed)
            ) {
                continue;
            }

            if (
                /^(Set|Call|Debug\.Print|MsgBox|DoCmd\.|If\b|For\b|Do\b|Select\b|With\b)/i.test(trimmed)
                || /^[A-Za-z_][A-Za-z0-9_\.]*\s*=/.test(trimmed)
            ) {
                return `Posible sentencia ejecutable fuera de procedimiento en la l\u00ednea ${lineNumber}: ${trimmed}`;
            }
        }

        return undefined;
    }

    async function getModuleCompileHint(
        connection: import("./models/types").AccessConnection,
        moduleName: string
    ): Promise<string | undefined> {
        const activeEditor = vscode.window.activeTextEditor;
        const activeMeta = activeEditor
            ? codeDocuments.get(activeEditor.document.uri.toString())
            : undefined;

        if (activeEditor && activeMeta?.objectType === "module" && activeMeta.objectName === moduleName) {
            return findTopLevelModuleIssue(activeEditor.document.getText());
        }

        try {
            const doc = await mcpClient.getObjectDocument(connection, "modules", moduleName);
            return findTopLevelModuleIssue(doc.content);
        } catch {
            return undefined;
        }
    }

    function extractCompileLineNumber(message: string): number | undefined {
        const match = message.match(/\b(?:l\u00ednea|linea|line)\s*:?\s*(\d+)\b/i);
        if (!match?.[1]) {
            return undefined;
        }

        const parsed = Number.parseInt(match[1], 10);
        return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
    }

    async function revealModuleCompileLine(moduleName: string, lineNumber?: number): Promise<void> {
        if (!lineNumber) {
            return;
        }

        const editor = vscode.window.activeTextEditor;
        const meta = editor ? codeDocuments.get(editor.document.uri.toString()) : undefined;
        if (!editor || meta?.objectType !== "module" || meta.objectName !== moduleName) {
            return;
        }

        const targetLine = Math.max(0, Math.min(lineNumber - 1, editor.document.lineCount - 1));
        const position = new vscode.Position(targetLine, 0);
        editor.selection = new vscode.Selection(position, position);
        editor.revealRange(new vscode.Range(position, position), vscode.TextEditorRevealType.InCenter);
    }
    async function pickOutputFolder(title: string): Promise<string | undefined> {
        const folder = await vscode.window.showOpenDialog({
            canSelectFiles: false,
            canSelectFolders: true,
            canSelectMany: false,
            title,
            openLabel: "Seleccionar carpeta"
        });

        return folder?.[0]?.fsPath;
    }

    async function pickSecondBrainOutputFolder(
        title: string,
        connection: import("./models/types").AccessConnection
    ): Promise<string | undefined> {
        const memoryKey = `secondBrain.outputDir.${connection.id}`;
        const remembered = context.globalState.get<string>(memoryKey);
        const defaultPath = remembered ?? path.join(path.dirname(connection.dbPath), "secondBrain");

        const folder = await vscode.window.showOpenDialog({
            canSelectFiles: false,
            canSelectFolders: true,
            canSelectMany: false,
            title,
            defaultUri: vscode.Uri.file(defaultPath),
            openLabel: "Seleccionar carpeta"
        });

        const selected = folder?.[0]?.fsPath;
        if (selected) {
            await context.globalState.update(memoryKey, selected);
        }
        return selected;
    }

    async function pickSecondBrainLinkDensity(): Promise<"standard" | "high" | undefined> {
        const pick = await vscode.window.showQuickPick(
            [
                {
                    label: "Normal",
                    description: "Enlaces base + backlinks",
                    detail: "Faster, cleaner graph",
                    value: "standard" as const
                },
                {
                    label: "Alta densidad",
                    description: "Includes automatic domain MOCs",
                    detail: "More connections in the Obsidian graph",
                    value: "high" as const
                }
            ],
            {
                title: "Densidad de enlaces SecondBrain",
                placeHolder: "Choose how you want to generate cross-links"
            }
        );

        return pick?.value;
    }

    async function runSecondBrainExport(
        title: string,
        runner: (options: Parameters<typeof secondBrainService.exportSecondBrain>[3]) => Promise<import("./services/secondBrainService").SecondBrainExportResult>
    ): Promise<import("./services/secondBrainService").SecondBrainExportResult> {
        secondBrainOutput.clear();
        secondBrainOutput.show(true);
        secondBrainStatusBar.text = "$(sync~spin) SecondBrain: iniciando...";
        secondBrainStatusBar.show();

        const executeOnce = async (): Promise<import("./services/secondBrainService").SecondBrainExportResult> => {
            let lastPercent = 0;
            let lastLoggedObjectKey = "";
            let lastUiMessage = "";

            await mcpClient.reconnect();
            secondBrainOutput.appendLine("[inventory] MCP reconnected to start export.");

            return await vscode.window.withProgress(
                {
                    location: vscode.ProgressLocation.Notification,
                    title,
                    cancellable: false
                },
                async (progress) => {
                    return await runner({
                        onProgress: async (event) => {
                            const prefix = `[${event.phase}]`;
                            const shouldLogObjectProgress = event.phase !== "object"
                                || typeof event.completed !== "number"
                                || typeof event.total !== "number"
                                || event.completed === event.total
                                || event.completed === 1
                                || event.completed % 25 === 0
                                || Math.floor((event.completed / Math.max(event.total, 1)) * 100) !== lastPercent;

                            if (shouldLogObjectProgress) {
                                const objectKey = `${event.phase}:${event.message}:${event.completed ?? ""}/${event.total ?? ""}`;
                                if (objectKey !== lastLoggedObjectKey) {
                                    secondBrainOutput.appendLine(`${prefix} ${event.message}`);
                                    lastLoggedObjectKey = objectKey;
                                }
                            }

                            const statusMessage = `$(sync~spin) SecondBrain: ${event.message}`;
                            if (statusMessage !== lastUiMessage) {
                                secondBrainStatusBar.text = statusMessage;
                                lastUiMessage = statusMessage;
                            }

                            if (typeof event.completed === "number" && typeof event.total === "number" && event.total > 0) {
                                const percent = Math.min(100, Math.max(0, Math.floor((event.completed / event.total) * 100)));
                                const increment = Math.max(0, percent - lastPercent);
                                lastPercent = percent;
                                progress.report({ increment, message: shouldLogObjectProgress ? event.message : `${event.completed}/${event.total}` });
                                return;
                            }

                            if (event.message !== lastUiMessage.replace("$(sync~spin) SecondBrain: ", "")) {
                                progress.report({ message: event.message });
                            }
                        }
                    });
                }
            );
        };

        try {
            return await executeOnce();
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            secondBrainOutput.appendLine(`[error] ${message}`);

            const recovered = await offerAccessRestart(message);
            if (!recovered) {
                throw error;
            }

            secondBrainOutput.appendLine("[inventory] Access restarted. Retrying export once...");
            await mcpClient.reconnect();
            return await executeOnce();
        } finally {
            try {
                await mcpClient.disconnect();
            } catch {
                // ignore cleanup errors
            }
            secondBrainStatusBar.hide();
        }
    }

    // Helper to execute SQL and show result as markdown table (SELECT) or DML confirmation
    async function runSqlAndShow(connection: import("./models/types").AccessConnection, sql: string): Promise<void> {
        const trimmed = sql.trim();
        const isDml = /^\s*(INSERT|UPDATE|DELETE|CREATE|DROP|ALTER|TRUNCATE)\b/i.test(trimmed);

        if (isDml) {
            const verb = trimmed.match(/^\s*(\w+)/i)?.[1]?.toUpperCase() ?? "DML";
            const confirm = await vscode.window.showWarningMessage(
                rt("sql.confirm.title"),
                {
                    modal: true,
                    detail: rt("sql.confirm.detail", verb, connection.name)
                },
                "Ejecutar"
            );
            if (confirm !== "Ejecutar") {
                return;
            }

            const result = await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `Ejecutando ${verb}...`, cancellable: false },
                () => mcpClient.executeDml(connection, trimmed)
            );

            const affected = result.rowsAffected;
            const msg = affected !== undefined
                ? `${verb} completado \u00b7 ${affected} fila(s) afectada(s)`
                : `${verb} completado`;

            const doc = await vscode.workspace.openTextDocument({
                content: [
                    `# Resultado ${verb}`,
                    "",
                    `- ${rt("sql.result.connection")}: ${connection.name}`,
                    `- SQL: \`${trimmed}\``,
                    affected !== undefined ? `- Filas afectadas: **${affected}**` : "",
                    "",
                    "```json",
                    JSON.stringify(result.payload, null, 2),
                    "```"
                ].filter((l) => l !== undefined).join("\n"),
                language: "markdown"
            });
            await vscode.window.showTextDocument(doc, { preview: true, viewColumn: vscode.ViewColumn.Beside });
            treeProvider.refresh();
            vscode.window.showInformationMessage(msg);
            return;
        }

        const preview = await vscode.window.withProgress(
            { location: vscode.ProgressLocation.Notification, title: "Ejecutando SQL...", cancellable: false },
            () => mcpClient.executeRawSqlQuery(connection, trimmed)
        );
        showResultsWebview(preview.sql, connection.name, preview.rows, preview.rowCount);
    }

    // Select the active connection for the SQL editor and update the status bar
    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.pickSqlConnection", async () => {
            const connections = connectionStore.getAll();
            if (connections.length === 0) {
                vscode.window.showInformationMessage("No hay conexiones Access configuradas.");
                return;
            }
            const items = connections.map((c) => ({
                label: c.name,
                description: c.dbPath,
                connection: c,
                picked: activeSqlConnection?.id === c.id
            }));
            const pick = await vscode.window.showQuickPick(items, {
                title: rt("sql.pickProfile.title"),
                placeHolder: rt("sql.pickProfile.placeholder")
            });
            if (pick) {
                activeSqlConnection = pick.connection;
                updateSqlStatusBar();
            }
        })
    );

    // Open a new empty SQL editor
    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.openSqlEditor", async () => {
            const doc = await vscode.workspace.openTextDocument({
                language: "sql",
                content: ""
            });
            await vscode.window.showTextDocument(doc);
            if (!activeSqlConnection) {
                activeSqlConnection = await pickConnection(rt("sql.openEditor.pickConnection"));
                updateSqlStatusBar();
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.newQuery", async (node?: any) => {
            const connection = node?.connection ?? activeSqlConnection ?? await pickConnection("Seleccionar base de datos para la nueva consulta");
            if (!connection) {
                return;
            }

            await openNewQueryEditor(connection);
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.createTableDesigner", async (node?: any) => {
            const connection = node?.connection ?? activeSqlConnection ?? await pickConnection("Seleccionar base de datos para la nueva tabla guiada");
            if (!connection) {
                return;
            }

            showTableDesignerWebview({
                connection,
                mode: "create",
                initialTableName: "NuevaTabla",
                initialFields: [
                    { name: "Id", type: "AUTOINCREMENT", required: true },
                    { name: "Nombre", type: "TEXT", size: 255, required: true }
                ],
                onApply: async (tableName, fields) => {
                    const validated = validateTableDesignerInput(tableName, fields);
                    const statements = buildCreateTableStatements(validated.tableName, validated.fields);
                    await executeTableDesignerStatements(connection, validated.tableName, statements, "create");
                }
            });
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.editTableDesigner", async (node?: any) => {
            const connection = node?.connection ?? activeSqlConnection ?? await pickConnection("Seleccionar base de datos para editar la tabla guiada");
            if (!connection) {
                return;
            }

            const tableName = node?.categoryKey === "tables"
                ? String(node.objectInfo?.name ?? "").trim()
                : await pickObjectFromCategory(connection, "table", "Editar tabla guiada");

            if (!tableName) {
                return;
            }

            const originalFields = await mcpClient.getTableFields(connection, tableName);
            showTableDesignerWebview({
                connection,
                mode: "edit",
                initialTableName: tableName,
                initialFields: originalFields.map((field) => ({
                    originalName: field.name,
                    name: field.name,
                    type: String(field.type ?? "TEXT"),
                    size: field.size,
                    required: Boolean(field.required),
                    existing: true
                })),
                onApply: async (submittedTableName, fields) => {
                    const validated = validateTableDesignerInput(submittedTableName, fields);
                    const statements = buildEditTableStatements(tableName, originalFields, validated.fields);
                    await executeTableDesignerStatements(connection, tableName, statements, "edit");
                }
            });
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.createTableDdl", async (node?: any) => {
            const connection = node?.connection ?? activeSqlConnection ?? await pickConnection("Seleccionar base de datos para la nueva tabla");
            if (!connection) {
                return;
            }

            const tableName = await vscode.window.showInputBox({
                prompt: "Nombre de la nueva tabla",
                value: "NuevaTabla",
                validateInput: (value) => (value.trim() ? undefined : "El nombre es obligatorio")
            });

            if (!tableName?.trim()) {
                return;
            }

            await openSqlTemplateEditor(connection, buildCreateTableDdlTemplate(tableName.trim()));
            vscode.window.showInformationMessage("Plantilla DDL abierta. Ajusta el SQL y ejecútalo para crear la tabla.");
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.editTableDdl", async (node?: any) => {
            const connection = node?.connection ?? activeSqlConnection ?? await pickConnection("Seleccionar base de datos para editar la tabla");
            if (!connection) {
                return;
            }

            const tableName = node?.categoryKey === "tables"
                ? String(node.objectInfo?.name ?? "").trim()
                : await pickObjectFromCategory(connection, "table", "Editar tabla DDL");

            if (!tableName) {
                return;
            }

            const fields = await mcpClient.getTableFields(connection, tableName);
            const operation = await vscode.window.showQuickPick(
                [
                    {
                        label: "ADD COLUMN",
                        description: "Agregar un nuevo campo",
                        value: "add" as const
                    },
                    {
                        label: "ALTER COLUMN",
                        description: "Modificar un campo existente",
                        value: "alter" as const
                    },
                    {
                        label: "DROP COLUMN",
                        description: "Eliminar un campo existente",
                        value: "drop" as const
                    }
                ],
                {
                    title: `Editar tabla DDL: ${tableName}`,
                    placeHolder: "Selecciona la operacion DDL"
                }
            );

            if (!operation) {
                return;
            }

            let sql = buildEditTableDdlTemplate(tableName, fields);
            if (operation.value === "alter") {
                const field = await pickTableField(fields, `ALTER COLUMN en ${tableName}`);
                if (!field) {
                    return;
                }
                sql = buildAlterColumnDdlTemplate(tableName, field);
            } else if (operation.value === "drop") {
                const field = await pickTableField(fields, `DROP COLUMN en ${tableName}`);
                if (!field) {
                    return;
                }
                sql = buildDropColumnDdlTemplate(tableName, field.name);
            } else {
                sql = buildAddColumnDdlTemplate(tableName, fields);
            }

            await openSqlTemplateEditor(connection, sql);
            vscode.window.showInformationMessage(`Plantilla ${operation.label} abierta. Ajusta la sentencia y ejecutala para modificar la tabla.`);
        })
    );

    // Run SQL from a quick input box
    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.runSqlQuery", async () => {
            const connection = activeSqlConnection ?? await pickConnection();
            if (!connection) {
                return;
            }
            activeSqlConnection = connection;
            updateSqlStatusBar();

            const sql = await vscode.window.showInputBox({
                prompt: rt("sql.input.prompt", connection.name),
                placeHolder: "SELECT * FROM [MiTabla] WHERE ...",
                ignoreFocusOut: true
            });

            if (!sql?.trim()) {
                return;
            }

            try {
                await runSqlAndShow(connection, sql);
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`Error ejecutando SQL: ${message}`);
            }
        })
    );

    // Ejecuta el texto seleccionado (o todo el documento) del editor activo
    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.executeActiveSql", async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                vscode.window.showInformationMessage("No hay editor activo con SQL.");
                return;
            }

            const selection = editor.selection;
            const sql = selection.isEmpty
                ? editor.document.getText()
                : editor.document.getText(selection);

            if (!sql?.trim()) {
                vscode.window.showInformationMessage(rt("sql.editor.emptySelection"));
                return;
            }

            const connection = activeSqlConnection ?? await pickConnection(rt("sql.execute.pickConnection"));
            if (!connection) {
                return;
            }
            activeSqlConnection = connection;
            updateSqlStatusBar();

            try {
                await runSqlAndShow(connection, sql);
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`Error ejecutando SQL: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.saveQueryToAccess", async (node?: any) => {
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                vscode.window.showInformationMessage("No hay editor activo.");
                return;
            }

            try {
                const explicitQueryName = node?.categoryKey === "queries"
                    ? String(node.objectInfo?.name ?? "").trim() || undefined
                    : undefined;
                await saveQueryDocumentToAccess(editor, explicitQueryName);
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`Error al guardar la consulta: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.deleteQuery", async (node?: any) => {
            const connection = node?.connection ?? activeSqlConnection ?? await pickConnection("Seleccionar base de datos para eliminar la consulta");
            if (!connection) {
                return;
            }

            const queryName = node?.categoryKey === "queries"
                ? node.objectInfo?.name
                : await pickObjectFromCategory(connection, "query", "Eliminar consulta");

            if (!queryName) {
                return;
            }

            const confirm = await vscode.window.showWarningMessage(
                rt("query.delete.title", queryName),
                {
                    modal: true,
                    detail: rt("query.delete.detail", connection.name)
                },
                "Eliminar"
            );

            if (confirm !== "Eliminar") {
                return;
            }

            const escapedQueryName = escapeVbaStringLiteral(String(queryName));
            await vscode.window.withProgress(
                {
                    location: vscode.ProgressLocation.Notification,
                    title: `Eliminando consulta ${queryName}...`,
                    cancellable: false
                },
                () => mcpClient.evalVba(connection, `CurrentDb.QueryDefs.Delete "${escapedQueryName}"`, 15000)
            );

            treeProvider.refresh();
            vscode.window.showInformationMessage(`Consulta eliminada: ${queryName}`);
        })
    );

    // Guardar codigo VBA activo de vuelta al archivo Access
    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.saveCodeToAccess", async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                vscode.window.showInformationMessage("No hay editor activo.");
                return;
            }
            const meta = codeDocuments.get(editor.document.uri.toString());
            if (!meta) {
                vscode.window.showWarningMessage("Este documento no est\u00e1 asociado a un objeto Access. \u00c1brelo desde el explorador.");
                return;
            }
            const isProcedure = !!meta.procedureName;
            const targetName = meta.procedureName
                ? `${meta.objectName}.${meta.procedureName}`
                : meta.objectName;
            const targetLabel = isProcedure ? "procedimiento" : "c\u00f3digo";
            const actionLabel = isProcedure ? "Guardar procedimiento" : "Guardar";
            const confirm = await vscode.window.showWarningMessage(
                isProcedure ? `\u00bfGuardar procedimiento en Access?` : `\u00bfGuardar en Access?`,
                {
                    modal: true,
                    detail: `Se sobrescribir\u00e1 el ${targetLabel} de "${targetName}" (${meta.objectType}) en "${meta.connection.name}".`
                },
                actionLabel
            );
            if (confirm !== actionLabel) {
                return;
            }
            try {
                const code = editor.document.getText();
                await vscode.window.withProgress(
                    {
                        location: vscode.ProgressLocation.Notification,
                        title: isProcedure ? `Guardando procedimiento ${targetName}...` : `Guardando ${targetName}...`,
                        cancellable: false
                    },
                    async () => {
                        if (
                            typeof meta.replaceStartLine === "number"
                            && typeof meta.replaceCount === "number"
                            && !meta.isNew
                        ) {
                            await mcpClient.replaceCodeLines(
                                meta.connection,
                                meta.objectType,
                                meta.objectName,
                                meta.replaceStartLine,
                                meta.replaceCount,
                                code
                            );
                            meta.replaceCount = editor.document.lineCount;
                            codeDocuments.set(editor.document.uri.toString(), meta);
                            return;
                        }

                        await mcpClient.setCode(meta.connection, meta.objectType, meta.objectName, code);
                    }
                );
                treeProvider.refresh();
                vscode.window.showInformationMessage(
                    isProcedure
                        ? `Procedimiento guardado en Access: ${targetName}`
                        : `C\u00f3digo guardado en Access: ${targetName}`
                );
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`Error al guardar: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.saveModuleToAccess", async () => {
            await vscode.commands.executeCommand("accessExplorer.saveCodeToAccess");
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.toggleVbComment", async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor || editor.document.languageId !== "vb") {
                vscode.window.showInformationMessage("No hay un editor VBA activo.");
                return;
            }

            const document = editor.document;
            const selections = editor.selections.length > 0 ? editor.selections : [editor.selection];

            await editor.edit((editBuilder) => {
                for (const selection of selections) {
                    const startLine = selection.start.line;
                    const endLine = selection.isEmpty
                        ? selection.start.line
                        : (selection.end.character === 0 ? Math.max(selection.end.line - 1, selection.start.line) : selection.end.line);

                    const lineNumbers = Array.from(
                        { length: Math.max(0, endLine - startLine) + 1 },
                        (_, index) => startLine + index
                    );

                    const meaningfulLines = lineNumbers.filter((lineNumber) => {
                        const text = document.lineAt(lineNumber).text;
                        return text.trim().length > 0;
                    });

                    if (meaningfulLines.length === 0) {
                        continue;
                    }

                    const allCommented = meaningfulLines.every((lineNumber) => {
                        const text = document.lineAt(lineNumber).text;
                        return /^\s*'/.test(text);
                    });

                    for (const lineNumber of meaningfulLines) {
                        const line = document.lineAt(lineNumber);
                        const text = line.text;
                        const indentMatch = text.match(/^\s*/);
                        const indentLength = indentMatch ? indentMatch[0].length : 0;

                        if (allCommented) {
                            const uncommented = text.replace(/^(\s*)'\s?/, "$1");
                            editBuilder.replace(line.range, uncommented);
                        } else {
                            const commented = `${text.slice(0, indentLength)}'${text.slice(indentLength)}`;
                            editBuilder.replace(line.range, commented);
                        }
                    }
                }
            });
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.createModule", async (node?: any) => {
            const connection = node?.connection ?? await pickConnection("Seleccionar base de datos para el nuevo m\u00f3dulo");
            if (!connection) {
                return;
            }

            const moduleName = await vscode.window.showInputBox({
                prompt: "Nombre del nuevo m\u00f3dulo VBA",
                validateInput: (value) => (value.trim() ? undefined : "El nombre es obligatorio")
            });

            if (!moduleName?.trim()) {
                return;
            }

            const doc = await vscode.workspace.openTextDocument({
                language: "vb",
                content: "Option Compare Database\nOption Explicit\n"
            });
            trackCodeDocument(doc, {
                connection,
                objectType: "module",
                objectName: moduleName.trim(),
                isNew: true
            });
            await vscode.window.showTextDocument(doc, { preview: false });
            await updateEditorActionContexts();
            vscode.window.showInformationMessage(`M\u00f3dulo listo para editar: ${moduleName.trim()}. Usa "Save to Access" para crearlo.`);
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.deleteModule", async (node?: any) => {
            const connection = node?.connection ?? await pickConnection("Seleccionar base de datos para eliminar el m\u00f3dulo");
            if (!connection) {
                return;
            }

            const moduleName = node?.categoryKey === "modules"
                ? node.objectInfo?.name
                : await pickObjectFromCategory(connection, "module", "Eliminar m\u00f3dulo VBA");

            if (!moduleName) {
                return;
            }

            const confirm = await vscode.window.showWarningMessage(
                `\u00bfEliminar el m\u00f3dulo "${moduleName}"?`,
                {
                    modal: true,
                    detail: `Se eliminar\u00e1 de "${connection.name}".`
                },
                "Eliminar"
            );

            if (confirm !== "Eliminar") {
                return;
            }

            await vscode.window.withProgress(
                {
                    location: vscode.ProgressLocation.Notification,
                    title: `Eliminando m\u00f3dulo ${moduleName}...`,
                    cancellable: false
                },
                () => mcpClient.deleteVbaModule(connection, String(moduleName))
            );

            treeProvider.refresh();
            vscode.window.showInformationMessage(`M\u00f3dulo eliminado: ${moduleName}`);
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.compileModule", async (node?: any) => {
            const connection = node?.connection ?? await pickConnection("Seleccionar base de datos para compilar el m\u00f3dulo");
            if (!connection) {
                return;
            }

            const moduleName = node?.categoryKey === "modules"
                ? node.objectInfo?.name
                : codeDocuments.get(vscode.window.activeTextEditor?.document.uri.toString() ?? "")?.objectType === "module"
                    ? codeDocuments.get(vscode.window.activeTextEditor?.document.uri.toString() ?? "")?.objectName
                    : await pickObjectFromCategory(connection, "module", "Compilar m\u00f3dulo VBA");

            if (!moduleName) {
                return;
            }

            try {
                const result = await vscode.window.withProgress(
                    {
                        location: vscode.ProgressLocation.Notification,
                        title: `Compilando m\u00f3dulo ${moduleName}...`,
                        cancellable: false
                    },
                    () => mcpClient.compileModule(connection, String(moduleName), 30000)
                );

                vscode.window.showInformationMessage(`M\u00f3dulo compilado correctamente: ${moduleName}. ${result}`);
            } catch (error) {
                let message = error instanceof Error ? error.message : String(error);
                const topLevelIssue = await getModuleCompileHint(connection, String(moduleName));
                if (topLevelIssue && !message.includes(topLevelIssue)) {
                    message = `${message}\n\n${topLevelIssue}`;
                }
                const lineNumber = extractCompileLineNumber(message);
                const doc = await vscode.workspace.openTextDocument({
                    content: `Compilaci\u00f3n del m\u00f3dulo \u2014 ${moduleName}\n${"=".repeat(60)}\n\n${message}`,
                    language: "plaintext"
                });
                await vscode.window.showTextDocument(doc, { preview: false, viewColumn: vscode.ViewColumn.Beside });
                await revealModuleCompileLine(String(moduleName), lineNumber);
                const suffix = typeof lineNumber === "number" ? ` L\u00ednea: ${lineNumber}.` : "";
                vscode.window.showErrorMessage(`Error compilando el m\u00f3dulo ${moduleName}.${suffix} Revisa el documento abierto.`);
            }
        })
    );
    // Compilar VBA
    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.compileVba", async (node?: any) => {
            const connection = node?.connection ?? await pickConnection("Seleccionar base de datos para compilar VBA");
            if (!connection) {
                return;
            }
            try {
                const result = await vscode.window.withProgress(
                    { location: vscode.ProgressLocation.Notification, title: `Compilando VBA en ${connection.name}...`, cancellable: false },
                    () => mcpClient.compileVba(connection)
                );

                // Detect errors: any mention of "error", line numbers, or non-success text
                const hasErrors = /error|line\s+\d+|compile\s+error/i.test(result)
                    && !/success|ok|compiled successfully|no error/i.test(result);

                if (hasErrors) {
                    const doc = await vscode.workspace.openTextDocument({
                        content: `${rt("compileVba.docTitle", connection.name)}\n${"=".repeat(60)}\n\n${result}`,
                        language: "plaintext"
                    });
                    await vscode.window.showTextDocument(doc, { preview: false });
                    vscode.window.showErrorMessage(rt("compileVba.errorMessage", connection.name));
                } else {
                    vscode.window.showInformationMessage(rt("compileVba.successMessage", result));
                }
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                // Also show the full error text as a document so no info is lost
                const doc = await vscode.workspace.openTextDocument({
                    content: `${rt("compileVba.errorDocTitle", connection.name)}\n${"=".repeat(60)}\n\n${message}`,
                    language: "plaintext"
                });
                await vscode.window.showTextDocument(doc, { preview: false });
                vscode.window.showErrorMessage(`Error compilando VBA: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.openVbaConsole", async (node?: any) => {
            const connection = node?.connection ?? await pickConnection("Seleccionar base de datos para la consola VBA");
            if (!connection) {
                return;
            }

            showVbaConsoleWebview(connection, mcpClient);
        })
    );

    // Compact & Repair
    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.compactRepair", async (node?: any) => {
            const connection = node?.connection ?? await pickConnection("Seleccionar base de datos para compactar");
            if (!connection) {
                return;
            }
            const confirm = await vscode.window.showWarningMessage(
                rt("compactRepair.title", connection.name),
                { modal: true, detail: rt("compactRepair.detail", connection.dbPath) },
                "Compactar"
            );
            if (confirm !== "Compactar") {
                return;
            }
            try {
                const result = await vscode.window.withProgress(
                    { location: vscode.ProgressLocation.Notification, title: `Compactando ${connection.name}...`, cancellable: false },
                    () => mcpClient.compactRepair(connection)
                );
                vscode.window.showInformationMessage(`Compact & Repair: ${result}`);
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`Error en Compact & Repair: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.createDatabase", async () => {
            const target = await vscode.window.showSaveDialog({
                title: "Nueva base de datos Access",
                saveLabel: "Crear base de datos",
                filters: {
                    "Access Database": ["accdb"]
                }
            });

            if (!target) {
                return;
            }

            const defaultName = path.basename(target.fsPath);
            const connectionName = await vscode.window.showInputBox({
                prompt: "Nombre de la nueva conexion",
                value: defaultName,
                validateInput: (value) => (value.trim() ? undefined : "El nombre es obligatorio")
            });

            if (!connectionName?.trim()) {
                return;
            }

            await vscode.window.withProgress(
                {
                    location: vscode.ProgressLocation.Notification,
                    title: "Creando base de datos Access...",
                    cancellable: false
                },
                async () => {
                    await mcpClient.createDatabase(target.fsPath, 30000);

                    // create_database may leave the new file open in Access.
                    try {
                        await mcpClient.closeAccess(10000);
                    } catch {
                        // Ignore: there may be no active Access instance to close.
                    }

                    await mcpClient.disconnect();
                }
            );

            await connectionStore.upsert(connectionName.trim(), target.fsPath);
            treeProvider.refresh();
            vscode.window.showInformationMessage(`Base de datos creada y agregada: ${connectionName.trim()}`);
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.closeAccess", async () => {
            const confirm = await vscode.window.showWarningMessage(
                rt("closeAccess.title"),
                {
                    modal: true,
                    detail: rt("closeAccess.detail")
                },
                "Cerrar Access"
            );

            if (confirm !== "Cerrar Access") {
                return;
            }

            await vscode.window.withProgress(
                {
                    location: vscode.ProgressLocation.Notification,
                    title: "Cerrando Access...",
                    cancellable: false
                },
                async () => {
                    await mcpClient.closeAccess(10000);
                    await mcpClient.disconnect();
                }
            );

            treeProvider.refresh();
            vscode.window.showInformationMessage(rt("closeAccess.success"));
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.secondBrain.full", async (node?: any) => {
            const connection = node?.connection ?? await pickConnection("Seleccionar base de datos para SecondBrain (completo)");
            if (!connection) {
                return;
            }

            const linkDensity = await pickSecondBrainLinkDensity();
            if (!linkDensity) {
                return;
            }

            const outputDir = await pickSecondBrainOutputFolder("Seleccionar carpeta de salida para SecondBrain completo", connection);
            if (!outputDir) {
                return;
            }

            try {
                const result = await runSecondBrainExport(
                    `Generando SecondBrain completo de ${connection.name}...`,
                    (options) => secondBrainService.exportSecondBrain(connection, outputDir, { mode: "full" }, {
                        ...options,
                        linkDensity
                    })
                );

                const action = await vscode.window.showInformationMessage(
                    `SecondBrain completo generado (${result.stats.tables} tablas, ${result.stats.queries} consultas).`,
                    "Abrir carpeta",
                    rt("openIndex")
                );

                if (action === "Abrir carpeta") {
                    await vscode.commands.executeCommand("revealFileInOS", vscode.Uri.file(result.outputDir));
                }

                if (action === rt("openIndex")) {
                    const doc = await vscode.workspace.openTextDocument(vscode.Uri.file(`${result.vaultDir}\\_index.md`));
                    await vscode.window.showTextDocument(doc, { preview: false });
                }
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`No se pudo generar SecondBrain completo: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.secondBrain.category", async (node?: CategoryNode) => {
            let connection = node?.connection;
            let categoryKey = node?.categoryKey;

            if (!connection) {
                connection = await pickConnection("Seleccionar base de datos para SecondBrain por tipo");
            }

            if (!connection) {
                return;
            }

            if (!categoryKey) {
                const pick = await vscode.window.showQuickPick(
                    ACCESS_CATEGORIES.map((category) => ({
                        label: category.label,
                        detail: category.key,
                        categoryKey: category.key
                    })),
                    { title: "Seleccionar tipo de objeto para SecondBrain" }
                );
                categoryKey = pick?.categoryKey;
            }

            if (!categoryKey) {
                return;
            }

            const linkDensity = await pickSecondBrainLinkDensity();
            if (!linkDensity) {
                return;
            }

            const outputDir = await pickSecondBrainOutputFolder(`Seleccionar carpeta de salida para ${categoryKey}`, connection);
            if (!outputDir) {
                return;
            }

            try {
                const result = await runSecondBrainExport(
                    `Generando SecondBrain (${categoryKey}) de ${connection.name}...`,
                    (options) => secondBrainService.exportSecondBrain(connection, outputDir, { mode: "category", categoryKey }, {
                        ...options,
                        linkDensity
                    })
                );

                const action = await vscode.window.showInformationMessage(
                    `SecondBrain por tipo generado (${categoryKey}).`,
                    "Abrir carpeta",
                    rt("openIndex")
                );

                if (action === "Abrir carpeta") {
                    await vscode.commands.executeCommand("revealFileInOS", vscode.Uri.file(result.outputDir));
                }

                if (action === rt("openIndex")) {
                    const doc = await vscode.workspace.openTextDocument(vscode.Uri.file(`${result.vaultDir}\\_index.md`));
                    await vscode.window.showTextDocument(doc, { preview: false });
                }
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`No se pudo generar SecondBrain por tipo: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.secondBrain.object", async (node?: ObjectNode) => {
            if (!node) {
                vscode.window.showInformationMessage(rt("secondBrain.selectObject"));
                return;
            }

            const linkDensity = await pickSecondBrainLinkDensity();
            if (!linkDensity) {
                return;
            }

            const outputDir = await pickSecondBrainOutputFolder(`Seleccionar carpeta de salida para ${node.objectInfo.name}`, node.connection);
            if (!outputDir) {
                return;
            }

            try {
                const result = await runSecondBrainExport(
                    `Generando SecondBrain del objeto ${node.objectInfo.name}...`,
                    (options) => secondBrainService.exportSecondBrain(node.connection, outputDir, {
                        mode: "object",
                        categoryKey: node.categoryKey,
                        objectInfo: node.objectInfo
                    }, {
                        ...options,
                        linkDensity
                    })
                );

                const action = await vscode.window.showInformationMessage(
                    `SecondBrain individual generado para ${node.objectInfo.name}.`,
                    "Abrir carpeta",
                    rt("openIndex")
                );

                if (action === "Abrir carpeta") {
                    await vscode.commands.executeCommand("revealFileInOS", vscode.Uri.file(result.outputDir));
                }

                if (action === rt("openIndex")) {
                    const doc = await vscode.workspace.openTextDocument(vscode.Uri.file(`${result.vaultDir}\\_index.md`));
                    await vscode.window.showTextDocument(doc, { preview: false });
                }
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`No se pudo generar SecondBrain individual: ${message}`);
            }
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.exportObjects.full", async (node?: any) => {
            const connection = node?.connection ?? await pickConnection("Seleccionar base de datos para exportar objetos (completo)");
            if (!connection) {
                return;
            }

            const outputDir = await pickOutputFolder("Seleccionar carpeta de salida para exportar objetos");
            if (!outputDir) {
                return;
            }

            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `Exportando objetos de ${connection.name}...`, cancellable: false },
                async () => {
                    try {
                        const result = await exportObjectsService.exportObjects(
                            connection,
                            outputDir,
                            { mode: "full" }
                        );
                        const action = await vscode.window.showInformationMessage(
                            `Objetos exportados en ${result.outputDir}`,
                            "Abrir carpeta"
                        );
                        if (action === "Abrir carpeta") {
                            await vscode.commands.executeCommand("revealFileInOS", vscode.Uri.file(result.outputDir));
                        }
                    } catch (error) {
                        const message = error instanceof Error ? error.message : String(error);
                        vscode.window.showErrorMessage(`No se pudo exportar objetos: ${message}`);
                    }
                }
            );
        })
    );

    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.exportObjects.category", async (node?: CategoryNode) => {
            let connection = node?.connection;
            let categoryKey = node?.categoryKey;

            if (!connection) {
                connection = await pickConnection("Seleccionar base de datos para exportar por tipo");
            }
            if (!connection) {
                return;
            }

            if (!categoryKey) {
                const pick = await vscode.window.showQuickPick(
                    ACCESS_CATEGORIES.map((category) => ({
                        label: category.label,
                        detail: category.key,
                        categoryKey: category.key
                    })),
                    { title: "Seleccionar tipo de objeto a exportar" }
                );
                categoryKey = pick?.categoryKey;
            }
            if (!categoryKey) {
                return;
            }

            const outputDir = await pickOutputFolder(`Seleccionar carpeta de salida para exportar ${categoryKey}`);
            if (!outputDir) {
                return;
            }

            await vscode.window.withProgress(
                { location: vscode.ProgressLocation.Notification, title: `Exportando ${categoryKey} de ${connection.name}...`, cancellable: false },
                async () => {
                    try {
                        const result = await exportObjectsService.exportObjects(
                            connection,
                            outputDir,
                            { mode: "category", categoryKey }
                        );
                        const action = await vscode.window.showInformationMessage(
                            `Tipo "${categoryKey}" exportado en ${result.outputDir}`,
                            "Abrir carpeta"
                        );
                        if (action === "Abrir carpeta") {
                            await vscode.commands.executeCommand("revealFileInOS", vscode.Uri.file(result.outputDir));
                        }
                    } catch (error) {
                        const message = error instanceof Error ? error.message : String(error);
                        vscode.window.showErrorMessage(`No se pudo exportar por tipo: ${message}`);
                    }
                }
            );
        })
    );
}

export function deactivate(): void {
    // VS Code disposes subscriptions and closes the MCP transport via client.dispose.
}

/**
 * Registers the MCP-Access server in the VS Code user-level mcp.json so that
 * Copilot and other MCP-aware tools can discover and use it automatically.
 *
 * The user-level mcp.json lives at:
 *   <userDataDir>/mcp.json   (e.g. %APPDATA%\Code\User\mcp.json on Windows)
 *
 * We derive the path from context.globalStorageUri which is always inside
 * <userDataDir>/globalStorage/<extensionId>.
 *
 * @returns true if the file was written/changed, false if already up-to-date.
 */
async function registerMcpServerSilently(
    context: vscode.ExtensionContext,
    mcpClient: McpAccessClient
): Promise<boolean> {
    const info = await mcpClient.getMcpRuntimeInfo();

    // globalStorageUri = .../Code/User/globalStorage/<ext-id>
    // Two dirs up -> .../Code/User
    const userDataDir = path.resolve(context.globalStorageUri.fsPath, "..", "..");
    const mcpJsonPath = path.join(userDataDir, "mcp.json");

    // Read existing mcp.json or start with an empty object
    let existing: { servers?: Record<string, unknown> } = {};
    try {
        const raw = fs.readFileSync(mcpJsonPath, "utf-8");
        existing = JSON.parse(raw) as typeof existing;
    } catch {
        // File missing or invalid JSON - we will create/overwrite it
    }

    if (!existing.servers) {
        existing.servers = {};
    }

    // Extract our server entry from the snippet the client already builds
    const snippet = JSON.parse(info.mcpJsonSnippet) as { servers: Record<string, unknown> };
    const serverKey = "access-explorer-local";
    const newEntry = snippet.servers[serverKey];

    // Check whether the entry is already identical (avoid unnecessary disk writes)
    const existing_entry_json = JSON.stringify(existing.servers[serverKey]);
    const new_entry_json = JSON.stringify(newEntry);
    if (existing_entry_json === new_entry_json) {
        return false;
    }

    existing.servers[serverKey] = newEntry;

    // Ensure the directory exists (it always should, but be safe)
    fs.mkdirSync(userDataDir, { recursive: true });
    fs.writeFileSync(mcpJsonPath, JSON.stringify(existing, null, 2) + "\n", "utf-8");

    return true;
}

async function openDetailNode(
    mcpClient: McpAccessClient,
    node: DetailNode,
    treeView: vscode.TreeView<any>,
    treeProvider: AccessTreeProvider
) {
    if (node.detailKind === "formLayoutAction" || node.detailKind === "reportLayoutAction") {
        const objectType = node.detailKind === "formLayoutAction" ? "form" : "report";
        const controls = await mcpClient.getControls(node.connection, objectType, node.objectInfo.name);
        showLayoutWebview(node.objectInfo.name, controls, node.connection, objectType, node.objectInfo.name, treeView, treeProvider);
        return undefined;
    }

    if (node.detailKind === "formScreenshotAction" || node.detailKind === "reportScreenshotAction") {
        const objectType = node.detailKind === "formScreenshotAction" ? "form" : "report";
        const screenshot = await mcpClient.getObjectScreenshot(node.connection, objectType, node.objectInfo.name);
        if (screenshot.path && fs.existsSync(screenshot.path)) {
            await vscode.commands.executeCommand("vscode.open", vscode.Uri.file(screenshot.path));
            return undefined;
        }

        return {
            title: `${node.objectInfo.name}.screenshot.json`,
            language: "json",
            content: JSON.stringify(screenshot.metadata ?? screenshot, null, 2)
        };
    }

    if (node.detailKind === "querySqlAction") {
        const sql = await mcpClient.getQuerySql(node.connection, node.objectInfo.name);
        return {
            title: `${node.objectInfo.name}.sql`,
            language: "sql",
            content: sql
        };
    }

    if (node.detailKind === "queryRunJsonAction") {
        const preview = await mcpClient.executeQueryPreview(node.connection, node.objectInfo.name, 200);
        return {
            title: `${node.objectInfo.name}.result.json`,
            language: "json",
            content: JSON.stringify(
                {
                    sql: preview.sql,
                    rowCount: preview.rowCount,
                    rows: preview.rows,
                    payload: preview.payload
                },
                null,
                2
            )
        };
    }

    if (node.detailKind === "queryRunTableAction") {
        const preview = await mcpClient.executeQueryPreview(node.connection, node.objectInfo.name, 200);
        showResultsWebview(preview.sql, node.connection.name, preview.rows, preview.rowCount);
        return undefined;
    }

    if (node.detailKind === "tableDataJsonAction") {
        const preview = await mcpClient.getTableDataPreview(node.connection, node.objectInfo.name, 100);
        return {
            title: `${node.objectInfo.name}.data.json`,
            language: "json",
            content: JSON.stringify(
                {
                    sql: preview.sql,
                    rowCount: preview.rowCount,
                    rows: preview.rows,
                    payload: preview.payload
                },
                null,
                2
            )
        };
    }

    if (node.detailKind === "tableDataTableAction") {
        const preview = await mcpClient.getTableDataPreview(node.connection, node.objectInfo.name, 100);
        showResultsWebview(preview.sql, node.connection.name, preview.rows, preview.rowCount);
        return undefined;
    }

    if (node.detailKind === "moduleCodeAction") {
        return await mcpClient.getObjectDocument(
            node.connection,
            "modules",
            node.objectInfo.name,
            node.objectInfo.metadata
        );
    }

    if (node.detailKind === "formCodeAction") {
        return await mcpClient.getObjectDocument(
            node.connection,
            "forms",
            node.objectInfo.name,
            node.objectInfo.metadata
        );
    }

    if (node.detailKind === "reportCodeAction") {
        return await mcpClient.getObjectDocument(
            node.connection,
            "reports",
            node.objectInfo.name,
            node.objectInfo.metadata
        );
    }

    if (node.detailKind === "macroCodeAction") {
        return await mcpClient.getObjectDocument(
            node.connection,
            "macros",
            node.objectInfo.name,
            node.objectInfo.metadata
        );
    }

    if (node.detailKind === "procedure") {
        const procName = String(node.payload?.name ?? node.label);
        const objectType = String(node.payload?.objectType ?? "module") as "module" | "form" | "report";
        const procedureDoc = await mcpClient.getProcedureDocument(
            node.connection,
            objectType,
            node.objectInfo.name,
            procName
        );
        return {
            ...procedureDoc,
            codeMeta: {
                connection: node.connection,
                objectType,
                objectName: node.objectInfo.name,
                procedureName: procName,
                replaceStartLine: typeof node.payload?.start_line === "number" ? node.payload.start_line : undefined,
                replaceCount: typeof node.payload?.count === "number" ? node.payload.count : undefined
            }
        };
    }

    if (node.detailKind === "controlProcedure") {
        const procName = String(node.payload?.name ?? node.label);
        const objectType = String(node.payload?.objectType ?? "form") as "form" | "report";
        const procedureDoc = await mcpClient.getProcedureDocument(
            node.connection,
            objectType,
            node.objectInfo.name,
            procName
        );
        return {
            ...procedureDoc,
            codeMeta: {
                connection: node.connection,
                objectType,
                objectName: node.objectInfo.name,
                procedureName: procName,
                replaceStartLine: typeof node.payload?.start_line === "number" ? node.payload.start_line : undefined,
                replaceCount: typeof node.payload?.count === "number" ? node.payload.count : undefined
            }
        };
    }

    if (node.detailKind === "controlPropertiesAction") {
        const objectType = String(node.payload?.objectType ?? "form") as "form" | "report";
        const controlName = String(node.payload?.name ?? node.label);
        const props = await mcpClient.getControlRaw(
            node.connection,
            objectType,
            node.objectInfo.name,
            controlName
        );
        showControlPropsWebview(node.connection, objectType, node.objectInfo.name, controlName, props, mcpClient);
        return undefined;
    }

    if (node.detailKind === "tableField" || node.detailKind === "property") {
        return {
            title: `${node.objectInfo.name}.${node.label}.json`,
            language: "json",
            content: JSON.stringify(node.payload ?? {}, null, 2)
        };
    }

    return {
        title: `${node.objectInfo.name}.json`,
        language: "json",
        content: JSON.stringify(node.payload ?? node.objectInfo.metadata ?? {}, null, 2)
    };
}

function showLayoutWebview(
    objectName: string,
    controls: Array<{
        name: string;
        type_name?: string;
        left?: number;
        top?: number;
        width?: number;
        height?: number;
    }>,
    connection: any,
    objectType: string,
    parentObjectName: string,
    treeView: vscode.TreeView<any>,
    treeProvider: AccessTreeProvider
) {
    const maxRight = controls.reduce((acc, c) => Math.max(acc, (c.left ?? 0) + (c.width ?? 0)), 0);
    const maxBottom = controls.reduce((acc, c) => Math.max(acc, (c.top ?? 0) + (c.height ?? 0)), 0);
    const safeMaxRight = Math.max(1, maxRight);
    const safeMaxBottom = Math.max(1, maxBottom);
    const targetWidth = 1200;
    const scale = Math.min(1, targetWidth / safeMaxRight);
    const canvasWidth = Math.ceil(safeMaxRight * scale);
    const canvasHeight = Math.ceil(safeMaxBottom * scale);

    // Color scheme per control type
    const colorMap: Record<string, { bg: string; border: string; text: string }> = {
        CommandButton: { bg: "#ef553b1f", border: "#ef553b", text: "#fecaca" },
        TextBox: { bg: "#3b821f1f", border: "#3b821f", text: "#a3e635" },
        ComboBox: { bg: "#f59e0b1f", border: "#f59e0b", text: "#fde047" },
        ListBox: { bg: "#f59e0b1f", border: "#f59e0b", text: "#fde047" },
        CheckBox: { bg: "#8b5cf61f", border: "#8b5cf6", text: "#e9d5ff" },
        OptionButton: { bg: "#8b5cf61f", border: "#8b5cf6", text: "#e9d5ff" },
        ToggleButton: { bg: "#8b5cf61f", border: "#8b5cf6", text: "#e9d5ff" },
        Label: { bg: "#60a5fa1f", border: "#60a5fa", text: "#bfdbfe" },
        SubForm: { bg: "#a16207ff", border: "#92400e", text: "#fed7aa" },
        Image: { bg: "#759c3e1f", border: "#759c3e", text: "#dcfce7" },
        TabControl: { bg: "#7c3aed1f", border: "#7c3aed", text: "#ddd6fe" },
    };

    const getColorForType = (typeName?: string): { bg: string; border: string; text: string } => {
        if (!typeName) return { bg: "#64748b1f", border: "#64748b", text: "#cbd5e1" };
        return colorMap[typeName] || { bg: "#64748b1f", border: "#64748b", text: "#cbd5e1" };
    };

    const boxes = controls.map((c) => {
        const left = Math.round((c.left ?? 0) * scale);
        const top = Math.round((c.top ?? 0) * scale);
        const width = Math.max(8, Math.round((c.width ?? 60) * scale));
        const height = Math.max(8, Math.round((c.height ?? 30) * scale));
        const label = `${c.name}${c.type_name ? ` (${c.type_name})` : ""}`;
        const colors = getColorForType(c.type_name);
        return { left, top, width, height, label, type: c.type_name || "otros", name: c.name, ...colors };
    });

    // Count types
    const typeCounts: Record<string, number> = {};
    controls.forEach((c) => {
        const type = c.type_name || "otros";
        typeCounts[type] = (typeCounts[type] || 0) + 1;
    });

    const panel = vscode.window.createWebviewPanel(
        "accessExplorerLayout",
        `Layout: ${objectName}`,
        vscode.ViewColumn.Beside,
        {
            enableFindWidget: true,
            enableScripts: true
        }
    );

    const legendHtml = Object.entries(colorMap)
        .map(([type, colors]) => {
            const count = typeCounts[type] || 0;
            return count > 0
                ? `<div class="legend-item"><div class="legend-box" style="background: ${colors.bg}; border: 2px solid ${colors.border};"></div><span>${type} (${count})</span></div>`
                : "";
        })
        .join("");

    panel.webview.html = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <style>
    * { box-sizing: border-box; }
    body { 
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; 
      margin: 0; 
      padding: 16px; 
      color: #e6edf3; 
      background: linear-gradient(135deg, #0b1f3a 0%, #132f4c 100%);
      overflow-x: auto;
    }
    .container { max-width: 100%; }
    .header { 
      margin-bottom: 20px; 
      padding-bottom: 12px; 
      border-bottom: 1px solid #3b5b8a;
    }
    .title { font-size: 18px; font-weight: 600; margin: 0 0 8px 0; }
    .info { 
      display: flex; 
      gap: 20px; 
      font-size: 12px; 
      opacity: 0.85;
      flex-wrap: wrap;
    }
    .info-item { display: flex; gap: 4px; }
    .legend {
      margin-bottom: 16px;
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
      gap: 8px;
    }
    .legend-item {
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 11px;
      padding: 4px 8px;
      background: rgba(60, 91, 138, 0.3);
      border-radius: 4px;
    }
    .legend-box { width: 12px; height: 12px; border-radius: 2px; flex-shrink: 0; }
    .canvas-wrapper {
      border: 2px solid #3b5b8a;
      border-radius: 6px;
      background: #102a4d;
      overflow: auto;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.4);
    }
    .canvas { 
      position: relative;
      width: ${canvasWidth}px; 
      height: ${canvasHeight}px;
      background: linear-gradient(135deg, #0f2744 0%, #1a3a52 100%);
      padding: 0;
    }
    .box { 
      position: absolute; 
      border: 2px solid; 
      border-radius: 3px;
      color: #fff; 
      font-size: 10px; 
      padding: 3px 4px; 
      overflow: hidden; 
      white-space: nowrap; 
      text-overflow: ellipsis;
      font-weight: 500;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.3);
      transition: all 0.2s ease;
      cursor: pointer;
    }
    .box:hover {
      z-index: 10;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.8) !important;
      transform: scale(1.08);
      border-width: 3px;
    }
    .box.selected {
      box-shadow: 0 0 16px rgba(255, 255, 255, 0.6) !important;
      border-width: 3px;
      filter: brightness(1.2);
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="title">${rt("layout.title")} ${escapeHtml(objectName)}</div>
      <div class="info">
        <div class="info-item"><strong>${rt("layout.controls")}:</strong> ${controls.length}</div>
        <div class="info-item"><strong>${rt("layout.scale")}:</strong> ${scale.toFixed(2)}x</div>
        <div class="info-item"><strong>${rt("layout.size")}:</strong> ${safeMaxRight}\u00d7${safeMaxBottom}px</div>
        <div class="info-item" style="font-style: italic; opacity: 0.7;">${rt("layout.clickHint")}</div>
      </div>
    </div>
    <div class="legend">${legendHtml}</div>
    <div class="canvas-wrapper">
      <div class="canvas">
        ${boxes
            .map(
                (b) =>
                    `<div class="box" data-control-name="${escapeHtml(b.name)}" title="${escapeHtml(b.label)}" style="left:${b.left}px;top:${b.top}px;width:${b.width}px;height:${b.height}px;background:${b.bg};border-color:${b.border};color:${b.text};">${escapeHtml(b.label)}</div>`
            )
            .join("\n")}
      </div>
    </div>
  </div>
    <script>
        const vscode = acquireVsCodeApi();
        document.querySelectorAll('.box').forEach(box => {
            box.addEventListener('click', () => {
                document.querySelectorAll('.box').forEach(b => b.classList.remove('selected'));
                box.classList.add('selected');
                const controlName = box.getAttribute('data-control-name');
                vscode.postMessage({ command: 'revealControl', controlName: controlName });
            });
        });
    </script>
</body>
</html>`;

    // Handle messages from WebView
    panel.webview.onDidReceiveMessage(async (message) => {
        if (message.command === "revealControl") {
            const controlName = String(message.controlName ?? "").trim();
            if (!controlName) {
                return;
            }

            try {
                const objType = objectType as "form" | "report";
                const controlNode = await treeProvider.findControlNode(connection, objType, parentObjectName, controlName);

                if (controlNode) {
                    try {
                        await treeView.reveal(controlNode, { select: true, focus: true, expand: true });
                    } catch {
                        // reveal can fail if the node instance is not currently materialized in the UI tree.
                    }

                    await vscode.commands.executeCommand("accessExplorer.showDetails", controlNode);

                    const associatedProcedures = Array.isArray(controlNode.payload?.associatedProcedures)
                        ? (controlNode.payload?.associatedProcedures as Array<Record<string, unknown>>)
                        : [];
                    const firstProcedure = associatedProcedures[0];

                    if (firstProcedure) {
                        const procName = String(firstProcedure.name ?? "").trim();
                        if (procName) {
                            const controlProcedureNode = new DetailNode(
                                controlNode.connection,
                                controlNode.categoryKey,
                                controlNode.objectInfo,
                                "controlProcedure",
                                procName,
                                {
                                    ...firstProcedure,
                                    objectType: objType,
                                    controlName: controlName
                                }
                            );
                            await vscode.commands.executeCommand("accessExplorer.showDetails", controlProcedureNode);
                        }
                    }
                } else {
                    vscode.window.showWarningMessage(`No se encontro el control '${controlName}' en el arbol.`);
                }
            } catch (error) {
                console.error("Error revealing control:", error);
            }
        }
    });
}

function showVbaConsoleWebview(
    connection: import("./models/types").AccessConnection,
    mcpClient: McpAccessClient
): void {
    const panel = vscode.window.createWebviewPanel(
        "accessVbaConsole",
        `Consola VBA \u00b7 ${connection.name}`,
        vscode.ViewColumn.Beside,
        { enableScripts: true, retainContextWhenHidden: true }
    );

    panel.webview.html = `<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#1e1e1e;color:#d4d4d4;display:flex;flex-direction:column;height:100vh;}
  .toolbar{display:flex;gap:8px;align-items:center;padding:10px 12px;background:#252526;border-bottom:1px solid #3e3e42;}
  .toolbar label{font-size:12px;color:#9cdcfe;}
  .toolbar select,.toolbar input,.toolbar textarea{background:#1f1f1f;color:#ddd;border:1px solid #555;border-radius:4px;padding:6px 8px;font-size:12px;}
  .toolbar select,.toolbar input{height:32px;}
  .toolbar button{height:32px;padding:0 12px;border:1px solid #007acc;background:#007acc;color:#fff;border-radius:4px;cursor:pointer;}
  .toolbar button:hover{background:#1193f5;}
  .toolbar .secondary{background:#3a3a3c;border-color:#4e4e52;color:#ccc;}
  .body{display:grid;grid-template-columns:1.2fr 1fr;gap:0;flex:1;min-height:0;}
  .pane{display:flex;flex-direction:column;min-height:0;}
  .pane h2{padding:10px 12px;font-size:12px;font-weight:600;background:#2d2d30;border-bottom:1px solid #3e3e42;}
  textarea{width:100%;resize:none;min-height:120px;line-height:1.4;background:#1f1f1f;color:#ddd;border:0;padding:12px;font-family:Consolas, monospace;}
  #code{flex:1;}
  #args{height:72px;border-top:1px solid #3e3e42;}
  #history,#output{flex:1;overflow:auto;padding:12px;}
  .entry{border:1px solid #3e3e42;border-radius:6px;padding:10px;margin-bottom:10px;background:#252526;}
  .entry .meta{font-size:11px;color:#9cdcfe;margin-bottom:6px;}
  .entry pre{white-space:pre-wrap;word-break:break-word;font-family:Consolas, monospace;font-size:12px;}
  .status{padding:8px 12px;font-size:11px;background:#007acc;color:#fff;}
  .hint{font-size:11px;color:#9d9d9d;padding:8px 12px;border-top:1px solid #3e3e42;background:#202020;}
</style>
</head>
<body>
<div class="toolbar">
  <label for="mode">Modo</label>
  <select id="mode">
    <option value="eval">evalVba</option>
    <option value="run">runVba</option>
  </select>
  <input id="proc" type="text" placeholder="Procedimiento para runVba" style="flex:1;min-width:220px;display:none;"/>
  <button id="execute">Ejecutar</button>
  <button id="clear" class="secondary">Limpiar salida</button>
</div>
<div class="body">
  <div class="pane">
    <h2>C\u00f3digo / expresi\u00f3n</h2>
    <textarea id="code" spellcheck="false" placeholder="Debug.Print 1 + 1"></textarea>
    <h2>Argumentos JSON para runVba</h2>
    <textarea id="args" spellcheck="false" placeholder='[1, "texto", true]'></textarea>
  </div>
  <div class="pane">
    <h2>Salida</h2>
    <div id="output"></div>
    <h2>Historial</h2>
    <div id="history"></div>
  </div>
</div>
<div class="status" id="status">${escapeHtml(connection.name)}</div>
<div class="hint">evalVba acepta bloques VBA. runVba espera un procedimiento p\u00fablico y argumentos en formato JSON array.</div>
<script>
const vscode = acquireVsCodeApi();
const modeEl = document.getElementById('mode');
const procEl = document.getElementById('proc');
const codeEl = document.getElementById('code');
const argsEl = document.getElementById('args');
const outputEl = document.getElementById('output');
const historyEl = document.getElementById('history');
const statusEl = document.getElementById('status');
const history = [];

function syncMode() {
  const isRun = modeEl.value === 'run';
  procEl.style.display = isRun ? '' : 'none';
  argsEl.style.display = isRun ? '' : 'none';
}

function addHistory(entry) {
  history.unshift(entry);
  historyEl.innerHTML = history.map(item => (
    '<div class="entry"><div class="meta">' + esc(item.mode) + ' \u00b7 ' + esc(item.when) + '</div><pre>'
    + esc(item.input)
    + '</pre></div>'
  )).join('');
}

function setOutput(title, text) {
  outputEl.innerHTML = '<div class="entry"><div class="meta">' + esc(title) + '</div><pre>' + esc(text) + '</pre></div>' + outputEl.innerHTML;
}

function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}

modeEl.addEventListener('change', syncMode);
document.getElementById('clear').addEventListener('click', () => {
  outputEl.innerHTML = '';
  statusEl.textContent = 'Salida limpiada';
});
document.getElementById('execute').addEventListener('click', () => {
  const mode = modeEl.value;
  const code = codeEl.value;
  const procedure = procEl.value;
  const args = argsEl.value;
  vscode.postMessage({ command: 'execute', mode, code, procedure, args });
});

window.addEventListener('message', event => {
  const msg = event.data;
  if (msg.command === 'result') {
    setOutput(msg.title, msg.output);
    addHistory({ mode: msg.mode, when: msg.when, input: msg.input });
    statusEl.textContent = msg.status;
  }
  if (msg.command === 'error') {
    setOutput('Error', msg.output);
    statusEl.textContent = msg.status;
  }
});

syncMode();
</script>
</body>
</html>`;

    panel.webview.onDidReceiveMessage(async (message) => {
        if (message.command !== "execute") {
            return;
        }

        const mode = String(message.mode ?? "eval");
        const code = String(message.code ?? "").trim();

        if (mode === "eval" && !code) {
            vscode.window.showWarningMessage("Introduce una expresi\u00f3n o bloque VBA.");
            return;
        }

        if (mode === "run") {
            const procedure = String(message.procedure ?? "").trim();
            if (!procedure) {
                vscode.window.showWarningMessage("Indica el nombre del procedimiento para runVba.");
                return;
            }
        }

        try {
            const result = await vscode.window.withProgress(
                {
                    location: vscode.ProgressLocation.Notification,
                    title: mode === "run" ? "Ejecutando procedimiento VBA..." : "Evaluando VBA...",
                    cancellable: false
                },
                async () => {
                    if (mode === "run") {
                        const argsText = String(message.args ?? "").trim();
                        const args = argsText ? JSON.parse(argsText) : [];
                        if (!Array.isArray(args)) {
                            throw new Error("Los argumentos de runVba deben ser un array JSON.");
                        }

                        return await mcpClient.runVba(connection, String(message.procedure), args, 30000);
                    }

                    return await mcpClient.evalVba(connection, code, 30000);
                }
            );

            panel.webview.postMessage({
                command: "result",
                title: mode === "run" ? `runVba \u00b7 ${String(message.procedure ?? "")}` : "evalVba",
                output: stringifyConsoleResult(result),
                input: mode === "run"
                    ? `${String(message.procedure ?? "").trim()}(${String(message.args ?? "").trim() || "[]"})`
                    : code,
                mode,
                when: new Date().toLocaleTimeString(),
                status: "\u00daltima ejecuci\u00f3n completada"
            });
        } catch (error) {
            const messageText = error instanceof Error ? error.message : String(error);
            panel.webview.postMessage({
                command: "error",
                output: messageText,
                status: "La ejecuci\u00f3n devolvi\u00f3 un error"
            });
            vscode.window.showErrorMessage(`Error en la consola VBA: ${messageText}`);
        }
    });
}

function stringifyConsoleResult(result: unknown): string {
    if (result === undefined || result === null) {
        return "Sin resultado.";
    }
    if (typeof result === "string") {
        return result;
    }

    try {
        return JSON.stringify(result, null, 2);
    } catch {
        return String(result);
    }
}

function escapeHtml(input: string): string {
    return input
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function renderPreviewAsMarkdownTable(
    sql: string,
    rowCount: number | undefined,
    rows: Array<Record<string, unknown>> | undefined
): string {
    const safeRows = rows ?? [];

    const header = [
        "# Vista de datos",
        "",
        `- SQL: ${sql}`,
        `- Filas devueltas: ${rowCount ?? safeRows.length}`,
        ""
    ].join("\n");

    if (safeRows.length === 0) {
        return `${header}Sin filas para mostrar.`;
    }

    const columns = Array.from(
        safeRows.reduce((set, row) => {
            Object.keys(row ?? {}).forEach((key) => set.add(key));
            return set;
        }, new Set<string>())
    );

    if (columns.length === 0) {
        return `${header}Sin columnas para mostrar.`;
    }

    const markdownHeader = `| ${columns.map(escapeMarkdownCell).join(" | ")} |`;
    const markdownSeparator = `| ${columns.map(() => "---").join(" | ")} |`;
    const markdownRows = safeRows
        .map((row) => `| ${columns.map((col) => escapeMarkdownCell(formatCellValue(row[col]))).join(" | ")} |`)
        .join("\n");

    return [header, markdownHeader, markdownSeparator, markdownRows].join("\n");
}

function formatCellValue(value: unknown): string {
    if (value === null || value === undefined) {
        return "";
    }
    if (typeof value === "object") {
        try {
            return JSON.stringify(value);
        } catch {
            return String(value);
        }
    }
    return String(value);
}

function escapeMarkdownCell(value: string): string {
    return value
        .replaceAll("\\", "\\\\")
        .replaceAll("|", "\\|")
        .replaceAll("\n", "<br>")
        .replaceAll("\r", "");
}

function showControlPropsWebview(
    connection: import("./models/types").AccessConnection,
    objectType: "form" | "report",
    objectName: string,
    controlName: string,
    props: Record<string, unknown>,
    mcpClient: McpAccessClient
): void {
    const propsJson = JSON.stringify(props);
    const panel = vscode.window.createWebviewPanel(
        "accessControlProps",
        `Propiedades: ${controlName}`,
        vscode.ViewColumn.Beside,
        { enableScripts: true, retainContextWhenHidden: true }
    );

    panel.webview.html = `<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#1e1e1e;color:#d4d4d4;display:flex;flex-direction:column;height:100vh;overflow:hidden;}
  .toolbar{padding:7px 12px;background:#252526;border-bottom:1px solid #3e3e42;display:flex;align-items:center;gap:8px;}
  .title{font-size:13px;font-weight:600;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
  .btn{padding:4px 14px;font-size:12px;border:1px solid #4e4e52;background:#3a3a3c;color:#ccc;border-radius:3px;cursor:pointer;}
  .btn:hover{background:#4e4e52;}
  .btn.primary{background:#007acc;border-color:#007acc;color:#fff;}
  .btn.primary:hover{background:#0090f1;}
  .status{font-size:11px;padding:0 8px;color:#9d9d9d;}
  .table-wrap{flex:1;overflow:auto;padding:8px;}
  table{width:100%;border-collapse:collapse;font-size:12px;}
  th{background:#2d2d30;color:#9cdcfe;font-weight:600;padding:5px 8px;border:1px solid #3e3e42;text-align:left;}
  td{padding:4px 6px;border:1px solid #3e3e42;vertical-align:middle;}
  td.key-cell{color:#ce9178;white-space:nowrap;width:40%;font-family:monospace;}
  td.val-cell{width:60%;}
  td.val-cell input{width:100%;background:#3c3c3c;border:1px solid #555;color:#ddd;padding:3px 6px;border-radius:3px;font-size:12px;outline:none;}
  td.val-cell input:focus{border-color:#007acc;background:#1e3a5f;}
  td.val-cell input.modified{border-color:#f59e0b;background:#2a1f00;}
  tr:nth-child(even) td{background:#252526;}
  .footer{padding:6px 12px;background:#007acc;color:#fff;font-size:11px;flex-shrink:0;}
</style>
</head>
<body>
<div class="toolbar">
  <div class="title">${rt("layout.title")} ${escapeHtml(controlName)}${rt("controlProps.titleSeparator")}${escapeHtml(objectName)} (${escapeHtml(objectType)})</div>
  <span class="status" id="status"></span>
  <button class="btn" onclick="resetAll()">${rt("controlProps.reset")}</button>
  <button class="btn primary" onclick="saveProps()">${rt("controlProps.save")}</button>
</div>
<div class="table-wrap">
  <table id="grid"><thead><tr><th>Propiedad</th><th>Valor</th></tr></thead><tbody id="tbody"></tbody></table>
</div>
<div class="footer" id="footer">${escapeHtml(connection.name)}${rt("controlProps.titleSeparator")}${escapeHtml(objectName)}.${escapeHtml(controlName)}</div>
<script>
const vscode = acquireVsCodeApi();
const ORIGINAL = ${propsJson};
const current = Object.assign({}, ORIGINAL);
let modifiedKeys = new Set();

function build(){
  const tbody = document.getElementById('tbody');
  Object.entries(ORIGINAL).forEach(([k,v])=>{
    const tr = document.createElement('tr');
    tr.innerHTML = '<td class="key-cell">'+escH(k)+'</td><td class="val-cell"><input data-key="'+escH(k)+'" value="'+escH(String(v??''))+'"/></td>';
    tbody.appendChild(tr);
  });
  document.querySelectorAll('input[data-key]').forEach(inp=>{
    inp.addEventListener('input',function(){
      const k=this.getAttribute('data-key');
      if(this.value===String(ORIGINAL[k]??'')){modifiedKeys.delete(k);this.classList.remove('modified');}
      else{modifiedKeys.add(k);this.classList.add('modified');}
      updateStatus();
    });
  });
}

function updateStatus(){
  document.getElementById('status').textContent = modifiedKeys.size>0 ? modifiedKeys.size+' propiedad(es) modificada(s)' : '';
}

function resetAll(){
  document.querySelectorAll('input[data-key]').forEach(inp=>{
    const k=inp.getAttribute('data-key');
    inp.value=String(ORIGINAL[k]??'');
    inp.classList.remove('modified');
  });
  modifiedKeys.clear();
  updateStatus();
}

function saveProps(){
  const changed={};
  document.querySelectorAll('input[data-key]').forEach(inp=>{
    const k=inp.getAttribute('data-key');
    if(modifiedKeys.has(k)) changed[k]=inp.value;
  });
  if(!Object.keys(changed).length){ vscode.postMessage({command:'info',text:'Sin cambios.'}); return; }
  vscode.postMessage({command:'save',props:changed});
}

function escH(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}

window.addEventListener('message', e => {
  if(e.data.command==='saved'){ modifiedKeys.clear(); document.querySelectorAll('input').forEach(i=>i.classList.remove('modified')); updateStatus(); }
});

build();
</script>
</body>
</html>`;

    panel.webview.onDidReceiveMessage(async (msg) => {
        if (msg.command === "info") {
            vscode.window.showInformationMessage(String(msg.text ?? ""));
        }
        if (msg.command === "save") {
            const changed = msg.props as Record<string, unknown>;
            try {
                await vscode.window.withProgress(
                    { location: vscode.ProgressLocation.Notification, title: `Guardando propiedades de ${controlName}...`, cancellable: false },
                    () => mcpClient.setControlProps(connection, objectType, objectName, controlName, changed)
                );
                vscode.window.showInformationMessage(`Propiedades guardadas: ${Object.keys(changed).join(", ")}`);
                panel.webview.postMessage({ command: "saved" });
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`Error al guardar propiedades: ${message}`);
            }
        }
    });
}

function showResultsWebview(
    sql: string,
    connectionName: string,
    rows: Array<Record<string, unknown>> | undefined,
    rowCount: number | undefined
): void {
    const safeRows = rows ?? [];
    const columns = Array.from(
        safeRows.reduce((set, row) => {
            Object.keys(row ?? {}).forEach((key) => set.add(key));
            return set;
        }, new Set<string>())
    );

    const rowsJson = JSON.stringify(safeRows);
    const columnsJson = JSON.stringify(columns);
    const displayCount = rowCount ?? safeRows.length;
    const shortSql = sql.length > 120 ? sql.slice(0, 117) + "..." : sql;

    const panel = vscode.window.createWebviewPanel(
        "accessSqlResults",
        `Resultados SQL`,
        vscode.ViewColumn.Beside,
        { enableScripts: true, retainContextWhenHidden: true }
    );

    panel.webview.html = `<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#1e1e1e;color:#d4d4d4;display:flex;flex-direction:column;height:100vh;overflow:hidden;}
  .toolbar{padding:6px 10px;background:#252526;border-bottom:1px solid #3e3e42;display:flex;align-items:center;gap:10px;flex-shrink:0;flex-wrap:wrap;}
  .toolbar-info{font-size:11px;color:#9d9d9d;flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
  .toolbar-info strong{color:#ccc;}
  .btn{padding:3px 10px;font-size:11px;border:1px solid #4e4e52;background:#3a3a3c;color:#ccc;border-radius:3px;cursor:pointer;}
  .btn:hover{background:#4e4e52;}
  #filter-box{padding:3px 8px;font-size:11px;background:#3c3c3c;border:1px solid #555;color:#ddd;border-radius:3px;width:180px;outline:none;}
  #filter-box:focus{border-color:#007acc;}
  .table-wrap{flex:1;overflow:auto;}
  table{width:max-content;min-width:100%;border-collapse:collapse;font-size:12px;}
  thead{position:sticky;top:0;z-index:2;}
  th{background:#2d2d30;color:#9cdcfe;font-weight:600;padding:5px 8px;border:1px solid #3e3e42;white-space:nowrap;user-select:none;cursor:pointer;}
  th:hover{background:#37373a;}
  th .sort-icon{margin-left:4px;opacity:0.5;font-size:10px;}
  th.asc .sort-icon::after{content:'\\25B2';}
  th.desc .sort-icon::after{content:'\\25BC';}
  th:not(.asc):not(.desc) .sort-icon::after{content:'\\21C5';}
  td{padding:4px 8px;border:1px solid #3e3e42;white-space:pre;max-width:400px;overflow:hidden;text-overflow:ellipsis;vertical-align:top;}
  td.null-cell{color:#6a6a6a;font-style:italic;}
  tr:nth-child(even) td{background:#252526;}
  tr:hover td{background:#2a2d2e!important;}
  tr.selected-row td{background:#094771!important;}
  td[contenteditable="true"]{outline:none;}
  td[contenteditable="true"]:focus{background:#1e3a5f!important;box-shadow:inset 0 0 0 2px #007acc;}
  .modified{background:#1e3a5f!important;}
  .status-bar{padding:3px 10px;background:#007acc;color:#fff;font-size:11px;flex-shrink:0;}
  .no-rows{padding:20px;color:#6a6a6a;font-style:italic;text-align:center;}
  .row-num{color:#6a6a6a;font-size:10px;user-select:none;padding:4px 4px;border:1px solid #3e3e42;background:#1e1e1e;text-align:right;min-width:36px;}
</style>
</head>
<body>
<div class="toolbar">
  <div class="toolbar-info" title="${escapeHtml(sql)}">
    <strong>${escapeHtml(connectionName)}</strong> &nbsp;${rt("results.sep")}&nbsp; ${escapeHtml(shortSql)} &nbsp;${rt("results.sep")}&nbsp; <strong>${displayCount}</strong> fila(s)
  </div>
  <input id="filter-box" placeholder="${escapeHtml(rt("results.filter.placeholder"))}" oninput="applyFilter(this.value)"/>
  <button class="btn" onclick="exportCsv()">${rt("results.exportCsv")}</button>
  <button class="btn" onclick="copySelected()">${rt("results.copySelection")}</button>
</div>
<div class="table-wrap" id="table-wrap">
  <div class="no-rows" id="no-rows" style="display:none">Sin resultados</div>
  <table id="grid">
    <thead id="thead"></thead>
    <tbody id="tbody"></tbody>
  </table>
</div>
<div class="status-bar" id="status-bar">Listo</div>

<script>
const vscode = acquireVsCodeApi();
const ALL_ROWS = ${rowsJson};
const COLUMNS = ${columnsJson};
let displayedRows = ALL_ROWS.slice();
let sortCol = null;
let sortDir = 1;

function cellVal(v){
  if(v===null||v===undefined) return null;
  if(typeof v==='object') { try{return JSON.stringify(v);}catch{return String(v);} }
  return String(v);
}

function buildHeader(){
  const tr = document.createElement('tr');
  const th0 = document.createElement('th');
  th0.textContent = '#';
  th0.style.minWidth='36px';
  tr.appendChild(th0);
  COLUMNS.forEach((col,i)=>{
    const th = document.createElement('th');
    th.dataset.col = col;
    th.innerHTML = escH(col) + '<span class="sort-icon"></span>';
    th.addEventListener('click',()=>sortByCol(col,th));
    tr.appendChild(th);
  });
  document.getElementById('thead').appendChild(tr);
}

function buildBody(rows){
  const tbody = document.getElementById('tbody');
  tbody.innerHTML='';
  if(!rows.length){
    document.getElementById('no-rows').style.display='block';
    document.getElementById('grid').style.display='none';
    return;
  }
  document.getElementById('no-rows').style.display='none';
  document.getElementById('grid').style.display='';
  rows.forEach((row,i)=>{
    const tr = document.createElement('tr');
    tr.dataset.idx = i;
    tr.addEventListener('click',()=>{
      document.querySelectorAll('.selected-row').forEach(r=>r.classList.remove('selected-row'));
      tr.classList.add('selected-row');
      updateStatus(1+' fila seleccionada');
    });
    const tdNum = document.createElement('td');
    tdNum.className='row-num';
    tdNum.textContent = i+1;
    tr.appendChild(tdNum);
    COLUMNS.forEach(col=>{
      const val = cellVal(row[col]);
      const td = document.createElement('td');
      if(val===null){td.className='null-cell';td.textContent='NULL';}
      else{td.textContent=val;}
      td.title = val??'NULL';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  updateStatus(rows.length + ' fila(s)' + (rows.length<ALL_ROWS.length?' (filtrado de '+ALL_ROWS.length+')':''));
}

function sortByCol(col, thEl){
  if(sortCol===col){ sortDir*=-1; } else { sortCol=col; sortDir=1; }
  document.querySelectorAll('th').forEach(t=>t.classList.remove('asc','desc'));
  thEl.classList.add(sortDir===1?'asc':'desc');
  displayedRows.sort((a,b)=>{
    const av=cellVal(a[col])||''; const bv=cellVal(b[col])||'';
    const an=parseFloat(av); const bn=parseFloat(bv);
    if(!isNaN(an)&&!isNaN(bn)) return (an-bn)*sortDir;
    return av.localeCompare(bv)*sortDir;
  });
  buildBody(displayedRows);
}

function applyFilter(q){
  const lq=q.toLowerCase();
  if(!lq){displayedRows=ALL_ROWS.slice();}
  else{
    displayedRows=ALL_ROWS.filter(row=>
      COLUMNS.some(col=>{const v=cellVal(row[col]);return v&&v.toLowerCase().includes(lq);})
    );
  }
  buildBody(displayedRows);
}

function exportCsv(){
  const header=COLUMNS.join(',');
  const lines=displayedRows.map(row=>COLUMNS.map(col=>{
    const v=cellVal(row[col])??'';
    return '"'+v.replaceAll('"','""')+'"';
  }).join(','));
  const csv=[header,...lines].join('\\n');
  vscode.postMessage({command:'exportCsv',csv,filename:'resultados.csv'});
}

function copySelected(){
  const sel=document.querySelectorAll('.selected-row');
  if(!sel.length){copyAll();return;}
  const lines=[];
  sel.forEach(tr=>{
    const tds=[...tr.querySelectorAll('td')].slice(1);
    lines.push(tds.map(td=>td.textContent??'').join('\\t'));
  });
  vscode.postMessage({command:'copyText',text:lines.join('\\n')});
}

function copyAll(){
  const header=COLUMNS.join('\\t');
  const lines=displayedRows.map(row=>COLUMNS.map(col=>cellVal(row[col])??'').join('\\t'));
  vscode.postMessage({command:'copyText',text:[header,...lines].join('\\n')});
}

function updateStatus(msg){ document.getElementById('status-bar').textContent=msg; }

function escH(s){return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}

buildHeader();
buildBody(ALL_ROWS);
</script>
</body>
</html>`;

    panel.webview.onDidReceiveMessage(async (msg) => {
        if (msg.command === "copyText") {
            await vscode.env.clipboard.writeText(String(msg.text ?? ""));
            vscode.window.showInformationMessage("Copiado al portapapeles.");
        }
        if (msg.command === "exportCsv") {
            const uri = await vscode.window.showSaveDialog({
                defaultUri: vscode.Uri.file(String(msg.filename ?? "resultados.csv")),
                filters: { "CSV": ["csv"] }
            });
            if (uri) {
                await vscode.workspace.fs.writeFile(uri, Buffer.from(String(msg.csv ?? ""), "utf-8"));
                vscode.window.showInformationMessage(`CSV guardado: ${uri.fsPath}`);
            }
        }
    });
}

