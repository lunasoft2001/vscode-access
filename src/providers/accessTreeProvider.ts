import * as vscode from "vscode";
import {
    AccessTreeNode,
    CategoryNode,
    ConnectionNode,
    DetailNode,
    MessageNode,
    ObjectNode
} from "../models/treeNodes";
import { ACCESS_CATEGORIES, AccessConnection } from "../models/types";
import { McpAccessClient } from "../mcp/mcpAccessClient";
import { ConnectionStore } from "../services/connectionStore";
import { isAccessDatabaseOpenError, offerAccessRestart } from "../utils/accessRecovery";

export class AccessTreeProvider implements vscode.TreeDataProvider<AccessTreeNode> {
    private readonly onDidChangeEmitter = new vscode.EventEmitter<AccessTreeNode | undefined>();
    readonly onDidChangeTreeData = this.onDidChangeEmitter.event;

    constructor(
        private readonly connectionStore: ConnectionStore,
        private readonly mcpClient: McpAccessClient
    ) { }

    refresh(): void {
        this.onDidChangeEmitter.fire(undefined);
    }

    getTreeItem(element: AccessTreeNode): vscode.TreeItem {
        return element;
    }

    async getChildren(element?: AccessTreeNode): Promise<AccessTreeNode[]> {
        if (!element) {
            return this.getConnectionNodes();
        }

        if (element instanceof ConnectionNode) {
            return ACCESS_CATEGORIES.map(
                (category) => new CategoryNode(element.connection, category.key, category.label)
            );
        }

        if (element instanceof CategoryNode) {
            return await this.getCategoryObjects(element.connection, element.categoryKey);
        }

        if (element instanceof ObjectNode) {
            return this.getObjectBranches(element);
        }

        if (element instanceof DetailNode) {
            return await this.getDetailChildren(element);
        }

        return [];
    }

    private getConnectionNodes(): AccessTreeNode[] {
        const connections = this.connectionStore.getAll();
        if (connections.length === 0) {
            return [new MessageNode("Sin conexiones. Usa Access: Add Connection.")];
        }

        return connections.map((connection) => new ConnectionNode(connection));
    }

    private async getCategoryObjects(
        connection: AccessConnection,
        categoryKey: CategoryNode["categoryKey"],
        attempt = 0
    ): Promise<AccessTreeNode[]> {
        try {
            if (categoryKey === "relationships") {
                const relationships = await this.mcpClient.listRelationships(connection);
                return this.mapObjects(connection, categoryKey, relationships);
            }

            if (categoryKey === "references") {
                const references = await this.mcpClient.listReferences(connection);
                return this.mapObjects(connection, categoryKey, references);
            }

            const category = ACCESS_CATEGORIES.find((item) => item.key === categoryKey);
            if (!category?.toolObjectType) {
                return [new MessageNode("Categoria no soportada.")];
            }

            const objects = await this.mcpClient.listObjects(connection, category.toolObjectType);
            return this.mapObjects(connection, categoryKey, objects);
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);

            if (attempt === 0 && isAccessDatabaseOpenError(message)) {
                const recovered = await offerAccessRestart(message);
                if (recovered) {
                    try {
                        await this.mcpClient.reconnect();
                    } catch {
                        // Ignore reconnect errors here; getCategoryObjects will return a message node below.
                    }
                    return await this.getCategoryObjects(connection, categoryKey, 1);
                }
            }

            return [new MessageNode(`Error: ${message}`)];
        }
    }

    private mapObjects(
        connection: AccessConnection,
        categoryKey: CategoryNode["categoryKey"],
        objects: Array<{ name: string; objectType: string; metadata?: Record<string, unknown> }>
    ): AccessTreeNode[] {
        if (objects.length === 0) {
            return [new MessageNode("Sin elementos.")];
        }

        return objects
            .slice()
            .sort((a, b) => a.name.localeCompare(b.name, "es"))
            .map((object) => new ObjectNode(connection, categoryKey, object));
    }

    private getObjectBranches(node: ObjectNode): AccessTreeNode[] {
        const { connection, categoryKey, objectInfo } = node;

        if (categoryKey === "tables") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "tableFieldsBranch", "Campos"),
                new DetailNode(connection, categoryKey, objectInfo, "tableDataTableAction", "Datos (tabla, TOP 100)"),
                new DetailNode(connection, categoryKey, objectInfo, "tableDataJsonAction", "Datos (JSON, TOP 100)")
            ];
        }

        if (categoryKey === "queries") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "querySqlAction", "SQL"),
                new DetailNode(connection, categoryKey, objectInfo, "queryRunTableAction", "Ejecutar (tabla, TOP 200)"),
                new DetailNode(connection, categoryKey, objectInfo, "queryRunJsonAction", "Ejecutar (JSON, TOP 200)")
            ];
        }

        if (categoryKey === "modules") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "moduleProceduresBranch", "Procedimientos"),
                new DetailNode(connection, categoryKey, objectInfo, "moduleCodeAction", "Codigo completo")
            ];
        }

        if (categoryKey === "forms") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "formPropertiesBranch", "Propiedades"),
                new DetailNode(connection, categoryKey, objectInfo, "formControlsBranch", "Controles"),
                new DetailNode(connection, categoryKey, objectInfo, "formLayoutAction", "Layout (posiciones)"),
                new DetailNode(connection, categoryKey, objectInfo, "formScreenshotAction", "Captura"),
                new DetailNode(connection, categoryKey, objectInfo, "formProceduresBranch", "Procedimientos VBA"),
                new DetailNode(connection, categoryKey, objectInfo, "formCodeAction", "Codigo VBA")
            ];
        }

        if (categoryKey === "reports") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "reportPropertiesBranch", "Propiedades"),
                new DetailNode(connection, categoryKey, objectInfo, "reportControlsBranch", "Controles"),
                new DetailNode(connection, categoryKey, objectInfo, "reportLayoutAction", "Layout (posiciones)"),
                new DetailNode(connection, categoryKey, objectInfo, "reportScreenshotAction", "Captura"),
                new DetailNode(connection, categoryKey, objectInfo, "reportProceduresBranch", "Procedimientos VBA"),
                new DetailNode(connection, categoryKey, objectInfo, "reportCodeAction", "Codigo VBA")
            ];
        }

        if (categoryKey === "macros") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "macroCodeAction", "Definicion")
            ];
        }

        return [];
    }

    private async getDetailChildren(node: DetailNode, attempt = 0): Promise<AccessTreeNode[]> {
        try {
            if (node.detailKind === "tableFieldsBranch") {
                const fields = await this.mcpClient.getTableFields(node.connection, node.objectInfo.name);
                if (fields.length === 0) {
                    return [new MessageNode("Sin campos.")];
                }

                return fields.map((field) => {
                    const parts = [field.type ?? "", field.size ? `(${field.size})` : ""]
                        .filter(Boolean)
                        .join(" ");
                    const description = [parts, field.required ? "required" : "optional"]
                        .filter(Boolean)
                        .join(" | ");

                    return new DetailNode(
                        node.connection,
                        node.categoryKey,
                        node.objectInfo,
                        "tableField",
                        field.name,
                        field as unknown as Record<string, unknown>,
                        description
                    );
                });
            }

            if (node.detailKind === "moduleProceduresBranch") {
                const procedures = await this.mcpClient.getModuleProcedures(
                    node.connection,
                    "module",
                    node.objectInfo.name
                );

                if (procedures.length === 0) {
                    return [new MessageNode("Sin procedimientos.")];
                }

                return procedures.map((proc) => new DetailNode(
                    node.connection,
                    node.categoryKey,
                    node.objectInfo,
                    "procedure",
                    proc.name,
                    {
                        ...(proc as unknown as Record<string, unknown>),
                        objectType: "module"
                    },
                    describeProcedure(proc.start_line, proc.count)
                ));
            }

            if (node.detailKind === "formProceduresBranch" || node.detailKind === "reportProceduresBranch") {
                const objectType = node.detailKind === "formProceduresBranch" ? "form" : "report";
                const procedures = await this.mcpClient.getModuleProcedures(
                    node.connection,
                    objectType,
                    node.objectInfo.name
                );

                if (procedures.length === 0) {
                    return [new MessageNode("Sin procedimientos.")];
                }

                return procedures.map((proc) => new DetailNode(
                    node.connection,
                    node.categoryKey,
                    node.objectInfo,
                    "procedure",
                    proc.name,
                    {
                        ...(proc as unknown as Record<string, unknown>),
                        objectType
                    },
                    describeProcedure(proc.start_line, proc.count)
                ));
            }

            if (node.detailKind === "formPropertiesBranch" || node.detailKind === "reportPropertiesBranch") {
                const objectType = node.detailKind === "formPropertiesBranch" ? "form" : "report";
                const properties = await this.mcpClient.getFormReportProperties(
                    node.connection,
                    objectType,
                    node.objectInfo.name
                );

                if (properties.length === 0) {
                    return [new MessageNode("Sin propiedades.")];
                }

                return properties.map((prop) => new DetailNode(
                    node.connection,
                    node.categoryKey,
                    node.objectInfo,
                    "property",
                    prop.name,
                    prop as unknown as Record<string, unknown>,
                    prop.value
                ));
            }

            if (node.detailKind === "formControlsBranch") {
                const controls = await this.mcpClient.getControls(
                    node.connection,
                    "form",
                    node.objectInfo.name
                );
                const procedures = await this.mcpClient.getModuleProcedures(
                    node.connection,
                    "form",
                    node.objectInfo.name
                );
                const proceduresByControl = indexProceduresByControl(procedures);

                if (controls.length === 0) {
                    return [new MessageNode("Sin controles.")];
                }

                return controls.map((ctrl) => {
                    const controlProcedures = proceduresByControl.get(normalizeName(ctrl.name)) ?? [];
                    const description = [ctrl.type_name, ctrl.control_source ? `source: ${ctrl.control_source}` : undefined]
                        .concat(controlProcedures.length > 0 ? [`eventos: ${controlProcedures.length}`] : [])
                        .filter(Boolean)
                        .join(" | ");
                    return new DetailNode(
                        node.connection,
                        node.categoryKey,
                        node.objectInfo,
                        "control",
                        ctrl.name,
                        {
                            ...(ctrl as unknown as Record<string, unknown>),
                            objectType: "form",
                            associatedProcedures: controlProcedures
                        },
                        description
                    );
                });
            }

            if (node.detailKind === "reportControlsBranch") {
                const controls = await this.mcpClient.getControls(
                    node.connection,
                    "report",
                    node.objectInfo.name
                );
                const procedures = await this.mcpClient.getModuleProcedures(
                    node.connection,
                    "report",
                    node.objectInfo.name
                );
                const proceduresByControl = indexProceduresByControl(procedures);

                if (controls.length === 0) {
                    return [new MessageNode("Sin controles.")];
                }

                return controls.map((ctrl) => {
                    const controlProcedures = proceduresByControl.get(normalizeName(ctrl.name)) ?? [];
                    const description = [ctrl.type_name, ctrl.control_source ? `source: ${ctrl.control_source}` : undefined]
                        .concat(controlProcedures.length > 0 ? [`eventos: ${controlProcedures.length}`] : [])
                        .filter(Boolean)
                        .join(" | ");
                    return new DetailNode(
                        node.connection,
                        node.categoryKey,
                        node.objectInfo,
                        "control",
                        ctrl.name,
                        {
                            ...(ctrl as unknown as Record<string, unknown>),
                            objectType: "report",
                            associatedProcedures: controlProcedures
                        },
                        description
                    );
                });
            }

            if (node.detailKind === "control") {
                return [
                    new DetailNode(
                        node.connection,
                        node.categoryKey,
                        node.objectInfo,
                        "controlPropertiesAction",
                        "Propiedades completas",
                        node.payload
                    ),
                    new DetailNode(
                        node.connection,
                        node.categoryKey,
                        node.objectInfo,
                        "controlCodeBranch",
                        "Codigo asociado",
                        node.payload
                    )
                ];
            }

            if (node.detailKind === "controlCodeBranch") {
                const objectType = String(node.payload?.objectType ?? "form") as "form" | "report";
                const controlName = String(node.payload?.name ?? "");
                if (!controlName) {
                    return [new MessageNode("Control sin nombre.")];
                }

                const embedded = node.payload?.associatedProcedures;
                const procedures = Array.isArray(embedded)
                    ? embedded
                    : await this.mcpClient.getControlAssociatedProcedures(
                        node.connection,
                        objectType,
                        node.objectInfo.name,
                        controlName
                    );

                if (procedures.length === 0) {
                    return [new MessageNode("Sin procedimientos asociados.")];
                }

                return procedures.map((proc) => new DetailNode(
                    node.connection,
                    node.categoryKey,
                    node.objectInfo,
                    "controlProcedure",
                    proc.name,
                    {
                        ...(proc as unknown as Record<string, unknown>),
                        objectType,
                        controlName
                    },
                    describeProcedure(proc.start_line, proc.count)
                ));
            }

            return [];
        } catch (error) {
            const message = error instanceof Error ? error.message : String(error);

            if (attempt === 0 && isAccessDatabaseOpenError(message)) {
                const recovered = await offerAccessRestart(message);
                if (recovered) {
                    try {
                        await this.mcpClient.reconnect();
                    } catch {
                        // Ignore reconnect errors here; return message below if retry fails.
                    }
                    return await this.getDetailChildren(node, 1);
                }
            }

            return [new MessageNode(`Error: ${message}`)];
        }
    }

    async findControlNode(
        connection: AccessConnection,
        objectType: "form" | "report",
        parentObjectName: string,
        controlName: string
    ): Promise<DetailNode | undefined> {
        try {
            // Find the category key based on objectType
            const categoryKey = objectType === "form" ? "forms" : "reports";
            const branchKind = objectType === "form" ? "formControlsBranch" : "reportControlsBranch";

            // Create a DetailNode for the controls branch
            const objectInfo = { name: parentObjectName, objectType, metadata: {} };
            const branchNode = new DetailNode(connection, categoryKey, objectInfo, branchKind, "Controles");

            // Get the children of the controls branch (which are individual controls)
            const controlChildren = await this.getDetailChildren(branchNode);

            // Find the control with matching name
            for (const child of controlChildren) {
                if (child instanceof DetailNode) {
                    if (normalizeName(String(child.label)) === normalizeName(controlName)) {
                        return child;
                    }
                }
            }

            return undefined;
        } catch (error) {
            console.error(`Error finding control node: ${error}`);
            return undefined;
        }
    }
}

function normalizeName(value: string): string {
    return value.trim().toLowerCase();
}

function indexProceduresByControl(
    procedures: Array<{ name: string; start_line?: number; count?: number }>
): Map<string, Array<{ name: string; start_line?: number; count?: number }>> {
    const map = new Map<string, Array<{ name: string; start_line?: number; count?: number }>>();

    for (const proc of procedures) {
        const idx = proc.name.indexOf("_");
        if (idx <= 0) {
            continue;
        }

        const controlName = normalizeName(proc.name.slice(0, idx));
        const list = map.get(controlName) ?? [];
        list.push(proc);
        map.set(controlName, list);
    }

    return map;
}

function describeProcedure(startLine?: number, count?: number): string | undefined {
    const parts: string[] = [];
    if (typeof startLine === "number") {
        parts.push(`linea ${startLine}`);
    }
    if (typeof count === "number") {
        parts.push(`${count} lineas`);
    }

    return parts.length > 0 ? parts.join(" | ") : undefined;
}
