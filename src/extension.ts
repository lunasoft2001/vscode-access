import * as vscode from "vscode";
import * as fs from "node:fs";
import { CategoryNode, DetailNode, ObjectNode } from "./models/treeNodes";
import { McpAccessClient } from "./mcp/mcpAccessClient";
import { ACCESS_CATEGORIES } from "./models/types";
import { AccessTreeProvider } from "./providers/accessTreeProvider";
import { ConnectionStore } from "./services/connectionStore";
import { SecondBrainService } from "./services/secondBrainService";
import { BulkExportService } from "./services/bulkExportService";
import { ExportObjectsService } from "./services/exportObjectsService";
import { offerAccessRestart, restartAccessProcesses } from "./utils/accessRecovery";

export function activate(context: vscode.ExtensionContext): void {
    const configuration = () => vscode.workspace.getConfiguration("accessExplorer");
    const connectionStore = new ConnectionStore(context);
    const mcpClient = new McpAccessClient(configuration, context);
    const secondBrainService = new SecondBrainService(mcpClient);
    const bulkExportService = new BulkExportService(mcpClient);
    const exportObjectsService = new ExportObjectsService(bulkExportService);
    const treeProvider = new AccessTreeProvider(connectionStore, mcpClient);
    const secondBrainOutput = vscode.window.createOutputChannel("Access Explorer SecondBrain");
    const secondBrainStatusBar = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 99);
    secondBrainStatusBar.tooltip = "Estado de generación de SecondBrain";

    context.subscriptions.push(mcpClient);
    context.subscriptions.push(secondBrainOutput, secondBrainStatusBar);

    // Tracks opened Access code documents so they can be saved back
    interface AccessCodeMeta {
        connection: import("./models/types").AccessConnection;
        objectType: "module" | "form" | "report";
        objectName: string;
    }
    const codeDocuments = new Map<string, AccessCodeMeta>();
    context.subscriptions.push(
        vscode.workspace.onDidCloseTextDocument((doc) => {
            codeDocuments.delete(doc.uri.toString());
        })
    );

    context.subscriptions.push(mcpClient);

    // Status bar item: muestra la conexión activa para el editor SQL
    let activeSqlConnection: import("./models/types").AccessConnection | undefined;
    const sqlStatusBar = vscode.window.createStatusBarItem(vscode.StatusBarAlignment.Left, 100);
    sqlStatusBar.command = "accessExplorer.pickSqlConnection";
    sqlStatusBar.tooltip = "Selecciona la conexión Access para el editor SQL";

    function updateSqlStatusBar(): void {
        const editor = vscode.window.activeTextEditor;
        const isSql = editor?.document.languageId === "sql";
        if (isSql) {
            sqlStatusBar.text = activeSqlConnection
                ? `$(database) ${activeSqlConnection.name}`
                : `$(database) Conectar Access…`;
            sqlStatusBar.show();
        } else {
            sqlStatusBar.hide();
        }
    }

    context.subscriptions.push(sqlStatusBar);
    context.subscriptions.push(vscode.window.onDidChangeActiveTextEditor(() => updateSqlStatusBar()));
    updateSqlStatusBar();

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
            } catch {
                // El cliente ya muestra diagnóstico y pasos de corrección.
            }
        }
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
                                            description: `${connection.name} · ${category.label}`,
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
                                            description: `${connection.name} · ${category.label}`,
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
                                        description: `${connection.name} · ${category.label}`,
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
                    const meta = (objectDoc as any).codeMeta as AccessCodeMeta;
                    codeDocuments.set(doc.uri.toString(), meta);
                }

                await vscode.window.showTextDocument(doc, { preview: true });
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

    async function pickSecondBrainLinkDensity(): Promise<"standard" | "high" | undefined> {
        const pick = await vscode.window.showQuickPick(
            [
                {
                    label: "Normal",
                    description: "Enlaces base + backlinks",
                    detail: "Más rápido, grafo limpio",
                    value: "standard" as const
                },
                {
                    label: "Alta densidad",
                    description: "Incluye MOCs automáticos por dominio",
                    detail: "Más conexiones en el grafo de Obsidian",
                    value: "high" as const
                }
            ],
            {
                title: "Densidad de enlaces SecondBrain",
                placeHolder: "Selecciona cómo quieres generar las interconexiones"
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

            await mcpClient.reconnect();
            secondBrainOutput.appendLine("[inventory] MCP reconectado para iniciar exportación.");

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
                            secondBrainOutput.appendLine(`${prefix} ${event.message}`);
                            secondBrainStatusBar.text = `$(sync~spin) SecondBrain: ${event.message}`;

                            if (typeof event.completed === "number" && typeof event.total === "number" && event.total > 0) {
                                const percent = Math.min(100, Math.max(0, Math.floor((event.completed / event.total) * 100)));
                                const increment = Math.max(0, percent - lastPercent);
                                lastPercent = percent;
                                progress.report({ increment, message: event.message });
                                return;
                            }

                            progress.report({ message: event.message });
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

            secondBrainOutput.appendLine("[inventory] Access reiniciado. Reintentando exportación una vez...");
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
                `Confirmación requerida`,
                {
                    modal: true,
                    detail: `Vas a ejecutar una sentencia ${verb} en "${connection.name}".\n\nEsta operación puede modificar o eliminar datos. ¿Continuar?`
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
                ? `${verb} completado · ${affected} fila(s) afectada(s)`
                : `${verb} completado`;

            const doc = await vscode.workspace.openTextDocument({
                content: [
                    `# Resultado ${verb}`,
                    "",
                    `- Conexión: ${connection.name}`,
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
            vscode.window.showInformationMessage(msg);
            return;
        }

        const preview = await vscode.window.withProgress(
            { location: vscode.ProgressLocation.Notification, title: "Ejecutando SQL...", cancellable: false },
            () => mcpClient.executeRawSqlQuery(connection, trimmed)
        );
        showResultsWebview(preview.sql, connection.name, preview.rows, preview.rowCount);
    }

    // Selecciona la conexión activa para el editor SQL y actualiza el status bar
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
                title: "Elegir un perfil de conexión Access",
                placeHolder: "Elegir un perfil de conexión de la lista siguiente"
            });
            if (pick) {
                activeSqlConnection = pick.connection;
                updateSqlStatusBar();
            }
        })
    );

    // Abre un nuevo editor SQL vacío
    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.openSqlEditor", async () => {
            const doc = await vscode.workspace.openTextDocument({
                language: "sql",
                content: ""
            });
            await vscode.window.showTextDocument(doc);
            if (!activeSqlConnection) {
                activeSqlConnection = await pickConnection("Elige la conexión Access para este editor SQL");
                updateSqlStatusBar();
            }
        })
    );

    // Ejecuta la query de un input box rápido
    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.runSqlQuery", async () => {
            const connection = activeSqlConnection ?? await pickConnection();
            if (!connection) {
                return;
            }
            activeSqlConnection = connection;
            updateSqlStatusBar();

            const sql = await vscode.window.showInputBox({
                prompt: `SQL · ${connection.name}`,
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
                vscode.window.showInformationMessage("El editor está vacío o no hay texto seleccionado.");
                return;
            }

            const connection = activeSqlConnection ?? await pickConnection("Elige la conexión Access para ejecutar el SQL");
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

    // Guardar código VBA activo de vuelta al archivo Access
    context.subscriptions.push(
        vscode.commands.registerCommand("accessExplorer.saveCodeToAccess", async () => {
            const editor = vscode.window.activeTextEditor;
            if (!editor) {
                vscode.window.showInformationMessage("No hay editor activo.");
                return;
            }
            const meta = codeDocuments.get(editor.document.uri.toString());
            if (!meta) {
                vscode.window.showWarningMessage("Este documento no está asociado a un objeto Access. Ábrelo desde el explorador.");
                return;
            }
            const confirm = await vscode.window.showWarningMessage(
                `¿Guardar en Access?`,
                {
                    modal: true,
                    detail: `Se sobreescribirá el código de "${meta.objectName}" (${meta.objectType}) en "${meta.connection.name}".`
                },
                "Guardar"
            );
            if (confirm !== "Guardar") {
                return;
            }
            try {
                await vscode.window.withProgress(
                    { location: vscode.ProgressLocation.Notification, title: `Guardando ${meta.objectName}…`, cancellable: false },
                    () => mcpClient.setCode(meta.connection, meta.objectType, meta.objectName, editor.document.getText())
                );
                vscode.window.showInformationMessage(`Código guardado en Access: ${meta.objectName}`);
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                vscode.window.showErrorMessage(`Error al guardar: ${message}`);
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
                    { location: vscode.ProgressLocation.Notification, title: `Compilando VBA en ${connection.name}…`, cancellable: false },
                    () => mcpClient.compileVba(connection)
                );

                // Detect errors: any mention of "error", line numbers, or non-success text
                const hasErrors = /error|line\s+\d+|compile\s+error/i.test(result)
                    && !/success|ok|compiled successfully|no error/i.test(result);

                if (hasErrors) {
                    const doc = await vscode.workspace.openTextDocument({
                        content: `Compilación VBA — ${connection.name}\n${"=".repeat(60)}\n\n${result}`,
                        language: "plaintext"
                    });
                    await vscode.window.showTextDocument(doc, { preview: false });
                    vscode.window.showErrorMessage(`Compilación VBA con errores en "${connection.name}". Revisa el documento abierto.`);
                } else {
                    vscode.window.showInformationMessage(`Compilación VBA correcta: ${result}`);
                }
            } catch (error) {
                const message = error instanceof Error ? error.message : String(error);
                // Also show the full error text as a document so no info is lost
                const doc = await vscode.workspace.openTextDocument({
                    content: `Error de compilación VBA — ${connection.name}\n${"=".repeat(60)}\n\n${message}`,
                    language: "plaintext"
                });
                await vscode.window.showTextDocument(doc, { preview: false });
                vscode.window.showErrorMessage(`Error compilando VBA: ${message}`);
            }
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
                `¿Compact & Repair "${connection.name}"?`,
                { modal: true, detail: `Esto compactará y reparará "${connection.dbPath}". Access debe estar cerrado o el archivo no bloqueado.` },
                "Compactar"
            );
            if (confirm !== "Compactar") {
                return;
            }
            try {
                const result = await vscode.window.withProgress(
                    { location: vscode.ProgressLocation.Notification, title: `Compactando ${connection.name}…`, cancellable: false },
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
        vscode.commands.registerCommand("accessExplorer.secondBrain.full", async (node?: any) => {
            const connection = node?.connection ?? await pickConnection("Seleccionar base de datos para SecondBrain (completo)");
            if (!connection) {
                return;
            }

            const linkDensity = await pickSecondBrainLinkDensity();
            if (!linkDensity) {
                return;
            }

            const outputDir = await pickOutputFolder("Seleccionar carpeta de salida para SecondBrain completo");
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
                    "Abrir índice"
                );

                if (action === "Abrir carpeta") {
                    await vscode.commands.executeCommand("revealFileInOS", vscode.Uri.file(result.outputDir));
                }

                if (action === "Abrir índice") {
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

            const outputDir = await pickOutputFolder(`Seleccionar carpeta de salida para ${categoryKey}`);
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
                    "Abrir índice"
                );

                if (action === "Abrir carpeta") {
                    await vscode.commands.executeCommand("revealFileInOS", vscode.Uri.file(result.outputDir));
                }

                if (action === "Abrir índice") {
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
                vscode.window.showInformationMessage("Selecciona un objeto en el árbol para generar su SecondBrain individual.");
                return;
            }

            const linkDensity = await pickSecondBrainLinkDensity();
            if (!linkDensity) {
                return;
            }

            const outputDir = await pickOutputFolder(`Seleccionar carpeta de salida para ${node.objectInfo.name}`);
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
                    "Abrir índice"
                );

                if (action === "Abrir carpeta") {
                    await vscode.commands.executeCommand("revealFileInOS", vscode.Uri.file(result.outputDir));
                }

                if (action === "Abrir índice") {
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
        return await mcpClient.getProcedureDocument(
            node.connection,
            objectType,
            node.objectInfo.name,
            procName
        );
    }

    if (node.detailKind === "controlProcedure") {
        const procName = String(node.payload?.name ?? node.label);
        const objectType = String(node.payload?.objectType ?? "form") as "form" | "report";
        return await mcpClient.getProcedureDocument(
            node.connection,
            objectType,
            node.objectInfo.name,
            procName
        );
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
      <div class="title">📐 ${escapeHtml(objectName)}</div>
      <div class="info">
        <div class="info-item">📊 <strong>Controles:</strong> ${controls.length}</div>
        <div class="info-item">🔍 <strong>Escala:</strong> ${scale.toFixed(2)}x</div>
        <div class="info-item">📏 <strong>Tamaño:</strong> ${safeMaxRight}×${safeMaxBottom}px</div>
        <div class="info-item" style="font-style: italic; opacity: 0.7;">💡 Haz clic en un control para ir a él en el árbol</div>
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
  <div class="title">🎛 ${escapeHtml(controlName)} · ${escapeHtml(objectName)} (${escapeHtml(objectType)})</div>
  <span class="status" id="status"></span>
  <button class="btn" onclick="resetAll()">↺ Restaurar</button>
  <button class="btn primary" onclick="saveProps()">💾 Guardar en Access</button>
</div>
<div class="table-wrap">
  <table id="grid"><thead><tr><th>Propiedad</th><th>Valor</th></tr></thead><tbody id="tbody"></tbody></table>
</div>
<div class="footer" id="footer">${escapeHtml(connection.name)} · ${escapeHtml(objectName)}.${escapeHtml(controlName)}</div>
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
                    { location: vscode.ProgressLocation.Notification, title: `Guardando propiedades de ${controlName}…`, cancellable: false },
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
    const shortSql = sql.length > 120 ? sql.slice(0, 117) + "…" : sql;

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
  th.asc .sort-icon::after{content:'▲';}
  th.desc .sort-icon::after{content:'▼';}
  th:not(.asc):not(.desc) .sort-icon::after{content:'⇅';}
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
    <strong>${escapeHtml(connectionName)}</strong> &nbsp;·&nbsp; ${escapeHtml(shortSql)} &nbsp;·&nbsp; <strong>${displayCount}</strong> fila(s)
  </div>
  <input id="filter-box" placeholder="Filtrar…" oninput="applyFilter(this.value)"/>
  <button class="btn" onclick="exportCsv()">⬇ CSV</button>
  <button class="btn" onclick="copySelected()">📋 Copiar selección</button>
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
