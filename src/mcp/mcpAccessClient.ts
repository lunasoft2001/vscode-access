import * as vscode from "vscode";
import { clearTimeout, setTimeout } from "node:timers";
import * as fs from "node:fs";
import * as path from "node:path";
import { spawn } from "node:child_process";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import {
    AccessCategoryKey,
    AccessConnection,
    AccessControlInfo,
    AccessObjectDocument,
    AccessObjectInfo,
    AccessPropertyInfo,
    AccessProcedureInfo,
    AccessQueryPreview,
    AccessScreenshotInfo,
    AccessTableFieldInfo
} from "../models/types";

interface AccessReference {
    name: string;
    guid?: string;
    path?: string;
}

interface AccessRelationship {
    name: string;
    table?: string;
    foreign_table?: string;
}

interface ShellResult {
    exitCode: number;
    stdout: string;
    stderr: string;
}

interface PythonBootstrapCommand {
    command: string;
    args: string[];
}

export class McpAccessClient {
    private client: Client | undefined;
    private transport: StdioClientTransport | undefined;
    private readonly output: vscode.OutputChannel;
    private installPromptShown = false;

    constructor(private readonly getConfig: () => vscode.WorkspaceConfiguration) {
        this.output = vscode.window.createOutputChannel("Access Explorer");
    }

    dispose(): void {
        this.output.dispose();
    }

    async reconnect(): Promise<void> {
        await this.disconnect();
        await this.connect();
    }

    async disconnect(): Promise<void> {
        if (this.client) {
            try {
                await this.client.close();
            } catch {
                // ignore close failures and continue cleanup
            }
        }
        this.client = undefined;
        this.transport = undefined;
    }

    async listObjects(connection: AccessConnection, objectType: string): Promise<AccessObjectInfo[]> {
        const payload = await this.callTool("list_objects", {
            db_path: connection.dbPath,
            object_type: objectType
        });

        const raw = this.extractObjectArray(payload, objectType);

        return raw.map((item) => {
            if (typeof item === "string") {
                return {
                    name: item,
                    objectType,
                    metadata: { name: item }
                };
            }

            return {
                name: String(item.name ?? item.object_name ?? "(sin nombre)"),
                objectType,
                metadata: item
            };
        });
    }

    async listRelationships(connection: AccessConnection): Promise<AccessObjectInfo[]> {
        const payload = await this.callTool("list_relationships", { db_path: connection.dbPath });
        const relationships: AccessRelationship[] = Array.isArray(payload)
            ? payload
            : payload.relationships ?? [];

        return relationships.map((relationship) => ({
            name: relationship.name,
            objectType: "relationship",
            metadata: relationship as unknown as Record<string, unknown>
        }));
    }

    async listReferences(connection: AccessConnection): Promise<AccessObjectInfo[]> {
        const payload = await this.callTool("list_references", { db_path: connection.dbPath });
        const references: AccessReference[] = Array.isArray(payload) ? payload : payload.references ?? [];

        return references.map((reference) => ({
            name: reference.name,
            objectType: "reference",
            metadata: reference as unknown as Record<string, unknown>
        }));
    }

    async getObjectDocument(
        connection: AccessConnection,
        categoryKey: AccessCategoryKey,
        objectName: string,
        metadata?: Record<string, unknown>
    ): Promise<AccessObjectDocument> {
        if (categoryKey === "queries") {
            const payload = await this.callTool("manage_query", {
                db_path: connection.dbPath,
                action: "get_sql",
                query_name: objectName
            });

            const sql = this.extractSql(payload);
            return {
                title: `${objectName}.sql`,
                language: "sql",
                content: sql
            };
        }

        if (categoryKey === "tables") {
            const payload = await this.callTool("table_info", {
                db_path: connection.dbPath,
                table_name: objectName
            });

            return {
                title: `${objectName}.table.json`,
                language: "json",
                content: this.stringifyPayload(payload)
            };
        }

        if (categoryKey === "modules" || categoryKey === "forms" || categoryKey === "reports") {
            const objectType = categoryKey === "modules"
                ? "module"
                : categoryKey === "forms"
                    ? "form"
                    : "report";

            const payload = await this.callTool("get_code", {
                db_path: connection.dbPath,
                object_type: objectType,
                object_name: objectName
            });

            const rawCode = this.extractCode(payload);
            const content = categoryKey === "forms" || categoryKey === "reports"
                ? this.extractCodeBehindForm(rawCode)
                : rawCode;

            return {
                title: `${objectName}.bas`,
                language: "vb",
                content,
                codeMeta: { connection, objectType, objectName }
            };
        }

        if (categoryKey === "macros") {
            try {
                const payload = await this.callTool("get_code", {
                    db_path: connection.dbPath,
                    object_type: "macro",
                    object_name: objectName
                });

                return {
                    title: `${objectName}.macro.txt`,
                    language: "plaintext",
                    content: this.extractCode(payload)
                };
            } catch {
                return {
                    title: `${objectName}.macro.json`,
                    language: "json",
                    content: JSON.stringify(metadata ?? { name: objectName }, null, 2)
                };
            }
        }

        return {
            title: `${objectName}.json`,
            language: "json",
            content: JSON.stringify(metadata ?? { name: objectName }, null, 2)
        };
    }

    async getTableFields(connection: AccessConnection, tableName: string): Promise<AccessTableFieldInfo[]> {
        const payload = await this.callTool("table_info", {
            db_path: connection.dbPath,
            table_name: tableName
        });

        const fields = Array.isArray(payload?.fields) ? payload.fields : [];
        return fields.map((field: any) => ({
            name: String(field.name ?? "(sin nombre)"),
            type: field.type ? String(field.type) : undefined,
            size: typeof field.size === "number" ? field.size : undefined,
            required: typeof field.required === "boolean" ? field.required : undefined
        }));
    }

    async getTableDataPreview(
        connection: AccessConnection,
        tableName: string,
        limit = 100
    ): Promise<AccessQueryPreview> {
        const safeLimit = Number.isFinite(limit) && limit > 0 ? Math.floor(limit) : 100;
        const attempts = Array.from(new Set([safeLimit, 50, 20, 10].filter((n) => n > 0)));
        let lastError: unknown;

        for (const attemptLimit of attempts) {
            const sql = `SELECT TOP ${attemptLimit} * FROM [${escapeSqlIdentifier(tableName)}]`;
            try {
                const payload = await this.callTool(
                    "execute_sql",
                    {
                        db_path: connection.dbPath,
                        sql,
                        limit: attemptLimit
                    },
                    this.getSqlTimeout()
                );

                return {
                    sql,
                    rowCount: this.extractRowCount(payload),
                    rows: this.extractRows(payload),
                    payload
                };
            } catch (error) {
                lastError = error;
                const message = error instanceof Error ? error.message : String(error);
                if (!isTimeoutError(message) || attemptLimit === attempts[attempts.length - 1]) {
                    break;
                }
            }
        }

        const finalMessage = lastError instanceof Error ? lastError.message : String(lastError ?? "");
        if (isTimeoutError(finalMessage)) {
            throw new Error(
                `No se pudieron cargar los datos de la tabla '${tableName}' por timeout. `
                + "Prueba con menos filas o aumenta accessExplorer.mcp.sqlQueryTimeoutMs."
            );
        }

        throw (lastError instanceof Error ? lastError : new Error(finalMessage || "Error al cargar datos de tabla."));
    }

    async getQuerySql(connection: AccessConnection, queryName: string): Promise<string> {
        const payload = await this.callTool("manage_query", {
            db_path: connection.dbPath,
            action: "get_sql",
            query_name: queryName
        });
        return this.extractSql(payload);
    }

    private getSqlTimeout(): number {
        return this.getConfig().get<number>("mcp.sqlQueryTimeoutMs", 600000);
    }

    async executeQueryPreview(
        connection: AccessConnection,
        queryName: string,
        limit = 200
    ): Promise<AccessQueryPreview> {
        const sql = `SELECT TOP ${limit} * FROM [${queryName}]`;
        const payload = await this.callTool("execute_sql", {
            db_path: connection.dbPath,
            sql,
            limit
        }, this.getSqlTimeout());

        return {
            sql,
            rowCount: this.extractRowCount(payload),
            rows: this.extractRows(payload),
            payload
        };
    }

    async executeRawSqlQuery(
        connection: AccessConnection,
        sql: string,
        limit = 500
    ): Promise<AccessQueryPreview> {
        const payload = await this.callTool("execute_sql", {
            db_path: connection.dbPath,
            sql,
            limit
        }, this.getSqlTimeout());

        return {
            sql,
            rowCount: this.extractRowCount(payload),
            rows: this.extractRows(payload),
            payload
        };
    }

    async executeDml(
        connection: AccessConnection,
        sql: string
    ): Promise<AccessQueryPreview> {
        await this.suppressAccessAlerts(connection);
        try {
            const payload = await this.callTool("execute_sql", {
                db_path: connection.dbPath,
                sql
            }, this.getSqlTimeout());

            // El servidor MCP puede devolver rows_affected o similar
            const affected = this.extractRowsAffected(payload);
            return {
                sql,
                rowsAffected: affected,
                rows: this.extractRows(payload),
                payload
            };
        } finally {
            await this.restoreAccessAlerts(connection);
        }
    }

    async getModuleProcedures(
        connection: AccessConnection,
        objectType: "module" | "form" | "report",
        objectName: string
    ): Promise<AccessProcedureInfo[]> {
        const payload = await this.callTool("vbe_module_info", {
            db_path: connection.dbPath,
            object_type: objectType,
            object_name: objectName
        });

        const procs = Array.isArray(payload?.procs) ? payload.procs : [];
        return procs.map((proc: any) => ({
            name: String(proc.name ?? "(sin nombre)"),
            start_line: typeof proc.start_line === "number" ? proc.start_line : undefined,
            count: typeof proc.count === "number" ? proc.count : undefined
        }));
    }

    async getProcedureDocument(
        connection: AccessConnection,
        objectType: "module" | "form" | "report",
        objectName: string,
        procName: string
    ): Promise<AccessObjectDocument> {
        const payload = await this.callTool("vbe_get_proc", {
            db_path: connection.dbPath,
            object_type: objectType,
            object_name: objectName,
            proc_name: procName
        });

        const content = this.extractCode(payload);
        return {
            title: `${objectName}.${procName}.bas`,
            language: "vb",
            content
        };
    }

    async getControls(
        connection: AccessConnection,
        objectType: "form" | "report",
        objectName: string
    ): Promise<AccessControlInfo[]> {
        const payload = await this.callTool("list_controls", {
            db_path: connection.dbPath,
            object_type: objectType,
            object_name: objectName
        });

        const controls = Array.isArray(payload?.controls) ? payload.controls : [];
        return controls.map((ctrl: any) => ({
            name: String(ctrl.name ?? "(sin nombre)"),
            type_name: ctrl.type_name ? String(ctrl.type_name) : undefined,
            control_source: ctrl.control_source ? String(ctrl.control_source) : undefined,
            caption: ctrl.caption ? String(ctrl.caption) : undefined,
            left: toNumber(ctrl.left),
            top: toNumber(ctrl.top),
            width: toNumber(ctrl.width),
            height: toNumber(ctrl.height)
        }));
    }

    async getObjectScreenshot(
        connection: AccessConnection,
        objectType: "form" | "report",
        objectName: string
    ): Promise<AccessScreenshotInfo> {
        const payload = await this.callTool("screenshot", {
            db_path: connection.dbPath,
            object_type: objectType,
            object_name: objectName,
            max_width: 1920,
            wait_ms: 1000,
            open_timeout_sec: 120
        }, 180000);

        return {
            path: typeof payload?.path === "string"
                ? payload.path
                : typeof payload?.image_path === "string"
                    ? payload.image_path
                    : undefined,
            width: toNumber(payload?.width),
            height: toNumber(payload?.height),
            metadata: payload && typeof payload === "object" ? payload : undefined
        };
    }

    async getControlDefinition(
        connection: AccessConnection,
        objectType: "form" | "report",
        objectName: string,
        controlName: string
    ): Promise<AccessObjectDocument> {
        const payload = await this.callTool("get_control", {
            db_path: connection.dbPath,
            object_type: objectType,
            object_name: objectName,
            control_name: controlName
        });

        const content = this.extractCode(payload);
        return {
            title: `${objectName}.${controlName}.control.txt`,
            language: "plaintext",
            content
        };
    }

    async getControlAssociatedProcedures(
        connection: AccessConnection,
        objectType: "form" | "report",
        objectName: string,
        controlName: string
    ): Promise<AccessProcedureInfo[]> {
        const procedures = await this.getModuleProcedures(connection, objectType, objectName);
        const normalizedControlName = controlName.toLowerCase();

        return procedures.filter((proc) => {
            const lower = proc.name.toLowerCase();
            return lower.startsWith(`${normalizedControlName}_`);
        });
    }

    async getFormReportProperties(
        connection: AccessConnection,
        objectType: "form" | "report",
        objectName: string
    ): Promise<AccessPropertyInfo[]> {
        const payload = await this.callTool("get_code", {
            db_path: connection.dbPath,
            object_type: objectType,
            object_name: objectName
        });

        return this.extractTopLevelProperties(this.extractCode(payload), objectType);
    }

    async getControlRaw(
        connection: AccessConnection,
        objectType: "form" | "report",
        objectName: string,
        controlName: string
    ): Promise<Record<string, unknown>> {
        const payload = await this.callTool("get_control", {
            db_path: connection.dbPath,
            object_type: objectType,
            object_name: objectName,
            control_name: controlName
        });
        // Return structured payload if available, else parse text
        if (typeof payload === "object" && payload !== null && !("text" in payload)) {
            return payload as Record<string, unknown>;
        }
        const text = this.extractCode(payload);
        const result: Record<string, unknown> = {};
        for (const line of text.split(/\r?\n/)) {
            const sep = line.indexOf("=");
            if (sep > 0) {
                result[line.slice(0, sep).trim()] = line.slice(sep + 1).trim();
            }
        }
        return result;
    }

    async setControlProps(
        connection: AccessConnection,
        objectType: "form" | "report",
        objectName: string,
        controlName: string,
        props: Record<string, unknown>
    ): Promise<void> {
        await this.suppressAccessAlerts(connection);
        try {
            await this.callTool("set_control_props", {
                db_path: connection.dbPath,
                object_type: objectType,
                object_name: objectName,
                control_name: controlName,
                props
            });
        } finally {
            await this.restoreAccessAlerts(connection);
        }
    }

    /**
     * Silencia los diálogos modales de Access (MsgBox, confirmaciones) para que
     * no bloqueen el servidor MCP. Se ignoran errores porque Access puede no estar abierto.
     */
    private async suppressAccessAlerts(connection: AccessConnection): Promise<void> {
        try {
            await this.callTool("eval_vba", {
                db_path: connection.dbPath,
                code: "Application.DisplayAlerts = False\nDoCmd.SetWarnings False"
            }, 5000);
        } catch {
            // Non-fatal: Access may not be running interactively
        }
    }

    private async restoreAccessAlerts(connection: AccessConnection): Promise<void> {
        try {
            await this.callTool("eval_vba", {
                db_path: connection.dbPath,
                code: "Application.DisplayAlerts = True\nDoCmd.SetWarnings True"
            }, 5000);
        } catch {
            // Non-fatal
        }
    }

    async compileVba(connection: AccessConnection): Promise<string> {
        await this.suppressAccessAlerts(connection);
        try {
            const payload = await this.callTool("compile_vba", {
                db_path: connection.dbPath
            }, 120000);
            return this.extractText(payload) || JSON.stringify(payload);
        } finally {
            await this.restoreAccessAlerts(connection);
        }
    }

    async compactRepair(connection: AccessConnection): Promise<string> {
        await this.suppressAccessAlerts(connection);
        try {
            const payload = await this.callTool("compact_repair", {
                db_path: connection.dbPath
            }, 300000);
            return this.extractText(payload) || JSON.stringify(payload);
        } finally {
            await this.restoreAccessAlerts(connection);
        }
    }

    async setCode(
        connection: AccessConnection,
        objectType: "module" | "form" | "report",
        objectName: string,
        code: string
    ): Promise<void> {
        await this.suppressAccessAlerts(connection);
        try {
            await this.callTool("set_code", {
                db_path: connection.dbPath,
                object_type: objectType,
                object_name: objectName,
                code
            });
        } finally {
            await this.restoreAccessAlerts(connection);
        }
    }

    private async connect(): Promise<void> {
        if (this.client) {
            return;
        }

        const cfg = this.getConfig();
        const serverScriptPath = await this.resolveServerScriptPathWithInstaller(cfg);
        const pythonCommand = this.resolvePythonCommand(cfg, serverScriptPath);
        await this.ensurePythonCommandAvailable(pythonCommand);

        const transport = new StdioClientTransport({
            command: pythonCommand,
            args: [serverScriptPath],
            stderr: "pipe"
        });

        const client = new Client(
            { name: "access-explorer", version: "0.0.1" },
            { capabilities: {} }
        );

        try {
            this.output.appendLine(`MCP connect -> python: ${pythonCommand}`);
            this.output.appendLine(`MCP connect -> script: ${serverScriptPath}`);

            await client.connect(transport);

            // Smoke test after initialize to fail early with a clear message.
            const tools = await client.listTools();
            this.output.appendLine(`MCP connected. Tools: ${tools.tools?.length ?? 0}`);

            this.transport = transport;
            this.client = client;
        } catch (error) {
            try {
                await client.close();
            } catch {
                // ignore close failures during failed connect cleanup
            }

            const reason = error instanceof Error ? error.message : String(error);
            throw new Error(
                `No se pudo conectar con MCP-Access. ${reason} | python=${pythonCommand} | script=${serverScriptPath}`
            );
        }
    }

    private async resolveServerScriptPathWithInstaller(cfg: vscode.WorkspaceConfiguration): Promise<string> {
        try {
            return this.resolveServerScriptPath(cfg);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);
            if (!this.isMissingServerScriptError(message)) {
                throw error;
            }

            if (this.installPromptShown) {
                throw error;
            }
            this.installPromptShown = true;

            const action = await vscode.window.showWarningMessage(
                "No se encontró MCP-Access (access_mcp_server.py). ¿Quieres instalarlo automáticamente?",
                "Instalar automáticamente",
                "Abrir configuración"
            );

            if (action === "Abrir configuración") {
                await vscode.commands.executeCommand(
                    "workbench.action.openSettings",
                    "accessExplorer.mcp.serverScriptPath"
                );
                throw new Error(
                    "Configura accessExplorer.mcp.serverScriptPath con la ruta de access_mcp_server.py."
                );
            }

            if (action !== "Instalar automáticamente") {
                throw new Error("MCP-Access no está instalado o configurado.");
            }

            try {
                await vscode.window.withProgress(
                    {
                        location: vscode.ProgressLocation.Notification,
                        title: "Instalando MCP-Access...",
                        cancellable: false
                    },
                    async () => {
                        await this.installMcpAccessDefault();
                    }
                );
            } catch (installError) {
                const details = installError instanceof Error ? installError.message : String(installError);
                this.output.appendLine(`Error instalando MCP-Access: ${details}`);
                this.output.show(true);

                await this.showInstallFailureDiagnostics(details);

                throw new Error(
                    `No se pudo instalar MCP-Access automaticamente. ${details}`
                );
            }

            const scriptPath = this.resolveServerScriptPath(cfg);
            vscode.window.showInformationMessage("MCP-Access instalado correctamente.");
            return scriptPath;
        }
    }

    private resolveServerScriptPath(cfg: vscode.WorkspaceConfiguration): string {
        const configuredPath = cfg.get<string>("mcp.serverScriptPath", "").trim();
        if (configuredPath) {
            if (fs.existsSync(configuredPath)) {
                return configuredPath;
            }
            throw new Error(`No existe access_mcp_server.py en: ${configuredPath}`);
        }

        const userHome = process.env.USERPROFILE ?? process.env.HOME ?? "";
        const candidates = [
            path.join(userHome, "mcp-servers", "MCP-Access", "access_mcp_server.py"),
            path.join(userHome, "mcp-servers", "access_mcp_server.py")
        ];

        const found = candidates.find((candidate) => fs.existsSync(candidate));
        if (found) {
            this.output.appendLine(`MCP script detectado automaticamente: ${found}`);
            return found;
        }

        throw new Error(
            "Configura accessExplorer.mcp.serverScriptPath con la ruta absoluta de access_mcp_server.py."
        );
    }

    private isMissingServerScriptError(message: string): boolean {
        return message.includes("No existe access_mcp_server.py")
            || message.includes("Configura accessExplorer.mcp.serverScriptPath");
    }

    private async installMcpAccessDefault(): Promise<void> {
        const userHome = process.env.USERPROFILE ?? process.env.HOME ?? "";
        if (!userHome) {
            throw new Error("No se pudo determinar la carpeta del usuario para instalar MCP-Access.");
        }

        const baseDir = path.join(userHome, "mcp-servers");
        const repoDir = path.join(baseDir, "MCP-Access");
        const venvDir = path.join(baseDir, ".venv");
        const serverScriptPath = path.join(repoDir, "access_mcp_server.py");

        await fs.promises.mkdir(baseDir, { recursive: true });

        const gitAvailable = await this.commandAvailable("git", ["--version"]);
        if (!gitAvailable) {
            throw new Error("No se encontró Git en PATH. Instálalo para poder descargar MCP-Access automáticamente.");
        }

        if (!fs.existsSync(repoDir)) {
            await this.runCommand("git", ["clone", "https://github.com/unmateria/MCP-Access.git", repoDir]);
        } else {
            const pull = await this.runCommand("git", ["-C", repoDir, "pull", "--ff-only"], undefined, true);
            if (pull.exitCode !== 0) {
                this.output.appendLine(`Aviso: no se pudo actualizar MCP-Access con git pull. ${pull.stderr}`);
            }
        }

        const bootstrapPython = await this.detectPythonBootstrapCommand();
        await this.runCommand(bootstrapPython.command, [...bootstrapPython.args, "-m", "venv", venvDir]);

        const venvPython = path.join(venvDir, "Scripts", "python.exe");
        if (!fs.existsSync(venvPython)) {
            throw new Error(`No se encontró Python del entorno virtual en: ${venvPython}`);
        }

        await this.runCommand(venvPython, ["-m", "pip", "install", "--upgrade", "pip"]);

        // Algunos forks de MCP-Access no son instalables con -e (sin setup.py/pyproject.toml).
        // En ese caso intentamos instalar dependencias desde requirements*.txt y continuamos.
        const editableInstall = await this.runCommand(
            venvPython,
            ["-m", "pip", "install", "-e", repoDir],
            undefined,
            true
        );

        if (editableInstall.exitCode !== 0) {
            this.output.appendLine(
                "Aviso: pip install -e falló. Se intentará instalación por requirements*.txt."
            );
            this.output.appendLine(`Detalle: ${editableInstall.stderr || editableInstall.stdout}`);

            const requirementCandidates = [
                path.join(repoDir, "requirements.txt"),
                path.join(repoDir, "requirements-dev.txt"),
                path.join(repoDir, "requirements_mcp.txt")
            ];

            let requirementsInstalled = false;
            for (const reqFile of requirementCandidates) {
                if (!fs.existsSync(reqFile)) {
                    continue;
                }

                const reqInstall = await this.runCommand(
                    venvPython,
                    ["-m", "pip", "install", "-r", reqFile],
                    undefined,
                    true
                );

                if (reqInstall.exitCode === 0) {
                    requirementsInstalled = true;
                    this.output.appendLine(`Dependencias instaladas desde: ${reqFile}`);
                    break;
                }

                this.output.appendLine(
                    `Aviso: instalación fallida desde ${reqFile}: ${reqInstall.stderr || reqInstall.stdout}`
                );
            }

            if (!requirementsInstalled) {
                this.output.appendLine(
                    "No se encontró requirements válido; se continuará si el script del servidor existe."
                );
            }
        }

        if (!fs.existsSync(serverScriptPath)) {
            throw new Error(`Instalación incompleta: no existe ${serverScriptPath}`);
        }
    }

    private async detectPythonBootstrapCommand(): Promise<PythonBootstrapCommand> {
        const candidates: PythonBootstrapCommand[] = [
            { command: "py", args: ["-3"] },
            { command: "python", args: [] }
        ];

        for (const candidate of candidates) {
            if (await this.commandAvailable(candidate.command, [...candidate.args, "--version"])) {
                return candidate;
            }
        }

        throw new Error(
            "No se encontró un comando Python válido (py o python). Instala Python 3.9+ y vuelve a intentarlo."
        );
    }

    private async ensurePythonCommandAvailable(pythonCommand: string): Promise<void> {
        const available = await this.commandAvailable(pythonCommand, ["--version"]);
        if (!available) {
            const action = await vscode.window.showWarningMessage(
                `No se puede ejecutar Python con '${pythonCommand}'.`,
                "Descargar Python",
                "Abrir configuración"
            );

            if (action === "Descargar Python") {
                await vscode.env.openExternal(vscode.Uri.parse("https://www.python.org/downloads/windows/"));
            }

            if (action === "Abrir configuración") {
                await vscode.commands.executeCommand(
                    "workbench.action.openSettings",
                    "accessExplorer.mcp.pythonCommand"
                );
            }

            throw new Error(
                `No se puede ejecutar Python con '${pythonCommand}'. Revisa accessExplorer.mcp.pythonCommand.`
            );
        }
    }

    private async showInstallFailureDiagnostics(details: string): Promise<void> {
        const userHome = process.env.USERPROFILE ?? process.env.HOME ?? "";
        const baseDir = userHome ? path.join(userHome, "mcp-servers") : "(no disponible)";
        const repoDir = userHome ? path.join(baseDir, "MCP-Access") : "(no disponible)";
        const venvPython = userHome
            ? path.join(baseDir, ".venv", "Scripts", "python.exe")
            : "(no disponible)";

        const content = [
            "# Error al instalar MCP-Access",
            "",
            "## Error técnico",
            "",
            "```text",
            details,
            "```",
            "",
            "## Comprobaciones sugeridas",
            "",
            "1. Verificar que Git esté instalado y en PATH (comando: git --version).",
            "2. Verificar que Python 3.9+ esté instalado (comando: py -3 --version o python --version).",
            "3. Comprobar permisos de escritura en la carpeta de instalación.",
            "",
            "## Rutas usadas por la instalación",
            "",
            `- Base: ${baseDir}`,
            `- Repositorio MCP-Access: ${repoDir}`,
            `- Python del entorno virtual: ${venvPython}`,
            "",
            "## Instalación manual alternativa",
            "",
            "```powershell",
            "git clone https://github.com/unmateria/MCP-Access.git",
            "cd MCP-Access",
            "py -3 -m venv .venv",
            ".\\.venv\\Scripts\\python.exe -m pip install --upgrade pip",
            ".\\.venv\\Scripts\\python.exe -m pip install -e .   # si existe setup.py/pyproject.toml",
            ".\\.venv\\Scripts\\python.exe -m pip install -r requirements.txt   # alternativa",
            "```",
            "",
            "Después, configura accessExplorer.mcp.serverScriptPath con la ruta de access_mcp_server.py si fuera necesario."
        ].join("\n");

        const doc = await vscode.workspace.openTextDocument({
            content,
            language: "markdown"
        });

        await vscode.window.showTextDocument(doc, { preview: false, viewColumn: vscode.ViewColumn.Beside });
    }

    private async commandAvailable(command: string, args: string[]): Promise<boolean> {
        const result = await this.runCommand(command, args, undefined, true);
        return result.exitCode === 0;
    }

    private async runCommand(
        command: string,
        args: string[],
        cwd?: string,
        allowFailure = false
    ): Promise<ShellResult> {
        return new Promise((resolve, reject) => {
            const child = spawn(command, args, {
                cwd,
                windowsHide: true,
                shell: false
            });

            let stdout = "";
            let stderr = "";

            child.stdout.on("data", (chunk: Buffer) => {
                stdout += chunk.toString();
            });

            child.stderr.on("data", (chunk: Buffer) => {
                stderr += chunk.toString();
            });

            child.on("error", (spawnError) => {
                if (allowFailure) {
                    resolve({ exitCode: -1, stdout, stderr: spawnError.message || stderr });
                    return;
                }
                reject(spawnError);
            });

            child.on("close", (code) => {
                const exitCode = code ?? -1;
                const result: ShellResult = { exitCode, stdout: stdout.trim(), stderr: stderr.trim() };

                if (!allowFailure && exitCode !== 0) {
                    const details = result.stderr || result.stdout || `Código de salida ${exitCode}`;
                    reject(new Error(`${command} ${args.join(" ")} falló: ${details}`));
                    return;
                }

                resolve(result);
            });
        });
    }

    private resolvePythonCommand(
        cfg: vscode.WorkspaceConfiguration,
        serverScriptPath: string
    ): string {
        const configured = cfg.get<string>("mcp.pythonCommand", "").trim();

        const scriptDir = path.dirname(serverScriptPath);
        const userHome = process.env.USERPROFILE ?? process.env.HOME ?? "";
        const candidates = [
            path.join(scriptDir, ".venv", "Scripts", "python.exe"),
            path.join(path.dirname(scriptDir), ".venv", "Scripts", "python.exe"),
            path.join(userHome, "mcp-servers", ".venv", "Scripts", "python.exe")
        ];

        const found = candidates.find((candidate) => fs.existsSync(candidate));
        if (found) {
            this.output.appendLine(`Python detectado automaticamente: ${found}`);
            return found;
        }

        if (configured) {
            return configured;
        }

        return "python";
    }

    private async callTool(
        toolSuffix: string,
        args: Record<string, unknown>,
        timeoutOverrideMs?: number
    ): Promise<any> {
        await this.connect();
        if (!this.client) {
            throw new Error("No fue posible inicializar el cliente MCP.");
        }

        const cfg = this.getConfig();
        const timeoutMs = timeoutOverrideMs ?? cfg.get<number>("mcp.requestTimeoutMs", 30000);
        const toolPrefix = cfg.get<string>("mcp.toolPrefix", "access").trim();
        const name = toolPrefix ? `${toolPrefix}_${toolSuffix}` : toolSuffix;

        this.output.appendLine(`MCP call: ${name}`);

        const call = this.client.callTool({ name, arguments: args });
        const result = await withTimeout(call, timeoutMs);

        if (result.isError) {
            const message = this.extractText(result.content) || "Error MCP desconocido";
            throw new Error(message);
        }

        if ((result as any).structuredContent !== undefined) {
            return (result as any).structuredContent;
        }

        const text = this.extractText(result.content);
        if (!text) {
            return {};
        }

        const textError = this.extractToolTextError(text);
        if (textError) {
            throw new Error(textError);
        }

        try {
            return JSON.parse(text);
        } catch {
            return { text };
        }
    }

    private extractText(content: unknown): string {
        if (!Array.isArray(content)) {
            return "";
        }

        const textChunks = content
            .map((item: any) => {
                if (item?.type === "text") {
                    return String(item.text ?? "");
                }
                return "";
            })
            .filter(Boolean);

        return textChunks.join("\n").trim();
    }

    private extractToolTextError(text: string): string | undefined {
        if (!text.startsWith("ERROR in tool")) {
            return undefined;
        }

        if (text.includes("La base de datos ya está abierta")) {
            return "La base de datos ya esta abierta en Microsoft Access. Cierra Access (todas las ventanas), luego ejecuta Access: Reconnect MCP y Access: Refresh.";
        }

        const messageMatch = text.match(/Message:\s*([^\n]+)/i);
        if (messageMatch?.[1]) {
            return messageMatch[1].trim();
        }

        return text.split("\n").slice(0, 3).join(" ").trim();
    }

    private extractObjectArray(
        payload: any,
        objectType: string
    ): Array<string | Record<string, unknown>> {
        if (Array.isArray(payload)) {
            return payload;
        }

        if (!payload || typeof payload !== "object") {
            return [];
        }

        const byType = payload[objectType];
        if (Array.isArray(byType)) {
            return byType;
        }

        if (Array.isArray(payload.objects)) {
            return payload.objects;
        }

        if (Array.isArray(payload.items)) {
            return payload.items;
        }

        return [];
    }

    private extractSql(payload: any): string {
        if (typeof payload === "string") {
            return payload;
        }

        if (payload && typeof payload === "object") {
            if (typeof payload.sql === "string") {
                return payload.sql;
            }
            if (typeof payload.query_sql === "string") {
                return payload.query_sql;
            }
            if (typeof payload.text === "string") {
                return payload.text;
            }
        }

        return this.stringifyPayload(payload);
    }

    private extractCode(payload: any): string {
        if (typeof payload === "string") {
            return payload;
        }

        if (payload && typeof payload === "object") {
            if (typeof payload.code === "string") {
                return payload.code;
            }
            if (typeof payload.text === "string") {
                return payload.text;
            }
        }

        return this.stringifyPayload(payload);
    }

    private stringifyPayload(payload: any): string {
        if (payload === undefined || payload === null) {
            return "";
        }

        if (typeof payload === "string") {
            return payload;
        }

        return JSON.stringify(payload, null, 2);
    }

    private extractCodeBehindForm(raw: string): string {
        const markerMatch = /^CodeBehindForm\s*$/m.exec(raw);
        if (!markerMatch || markerMatch.index === undefined) {
            return raw;
        }

        return raw.slice(markerMatch.index + markerMatch[0].length).trimStart();
    }

    private extractTopLevelProperties(
        raw: string,
        objectType: "form" | "report"
    ): AccessPropertyInfo[] {
        const expectedRoot = objectType === "form" ? "Begin Form" : "Begin Report";
        const lines = raw.split(/\r?\n/);
        const properties: AccessPropertyInfo[] = [];
        let inRoot = false;
        let depth = 0;

        for (const line of lines) {
            const trimmed = line.trim();
            if (!inRoot) {
                if (trimmed === expectedRoot) {
                    inRoot = true;
                    depth = 1;
                }
                continue;
            }

            if (trimmed === "CodeBehindForm") {
                break;
            }

            if (trimmed.startsWith("Begin ")) {
                depth += 1;
                continue;
            }

            if (trimmed === "End") {
                depth -= 1;
                if (depth <= 0) {
                    break;
                }
                continue;
            }

            if (depth !== 1 || !trimmed.includes("=")) {
                continue;
            }

            const index = trimmed.indexOf("=");
            const name = trimmed.slice(0, index).trim();
            const value = trimmed.slice(index + 1).trim();
            if (!name) {
                continue;
            }

            properties.push({ name, value });
        }

        return properties;
    }

    private extractRows(payload: any): Record<string, unknown>[] | undefined {
        if (Array.isArray(payload?.rows)) {
            return payload.rows as Record<string, unknown>[];
        }
        if (Array.isArray(payload?.records)) {
            return payload.records as Record<string, unknown>[];
        }
        if (Array.isArray(payload?.data)) {
            return payload.data as Record<string, unknown>[];
        }
        return undefined;
    }

    private extractRowCount(payload: any): number | undefined {
        const candidates = [payload?.row_count, payload?.count, payload?.record_count];
        for (const candidate of candidates) {
            if (typeof candidate === "number") {
                return candidate;
            }
        }

        const rows = this.extractRows(payload);
        if (rows) {
            return rows.length;
        }

        return undefined;
    }

    private extractRowsAffected(payload: any): number | undefined {
        const candidates = [
            payload?.rows_affected,
            payload?.rowsAffected,
            payload?.affected_rows,
            payload?.row_count,
            payload?.count
        ];
        for (const candidate of candidates) {
            if (typeof candidate === "number") {
                return candidate;
            }
        }
        // Intentar parsear desde texto si el servidor devuelve algo como "1 row(s) affected"
        const text = typeof payload === "string" ? payload : String(payload?.message ?? payload?.text ?? "");
        const match = text.match(/(\d+)\s+row/i);
        if (match) {
            return parseInt(match[1], 10);
        }
        return undefined;
    }
}

function toNumber(value: unknown): number | undefined {
    if (typeof value === "number" && Number.isFinite(value)) {
        return value;
    }
    if (typeof value === "string" && value.trim() !== "") {
        const n = Number(value);
        if (Number.isFinite(n)) {
            return n;
        }
    }
    return undefined;
}

function escapeSqlIdentifier(identifier: string): string {
    return identifier.replaceAll("]", "]]");
}

function isTimeoutError(message: string): boolean {
    const normalized = message.toLowerCase();
    return normalized.includes("timeout") || normalized.includes("timed out");
}

async function withTimeout<T>(promise: Promise<T>, timeoutMs: number): Promise<T> {
    return await new Promise<T>((resolve, reject) => {
        const timer = setTimeout(() => {
            reject(new Error(`La llamada MCP supero el timeout de ${timeoutMs} ms.`));
        }, timeoutMs);

        promise
            .then((value) => {
                clearTimeout(timer);
                resolve(value);
            })
            .catch((err) => {
                clearTimeout(timer);
                reject(err);
            });
    });
}
