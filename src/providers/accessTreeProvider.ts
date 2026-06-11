import * as vscode from "vscode";
import {
    ActionNode,
    AccessTreeNode,
    CategoryNode,
    ConnectionNode,
    DetailNode,
    MessageNode,
    ObjectNode
} from "../models/treeNodes";
import {
    ACCESS_CATEGORIES,
    ACCESS_CATEGORY_ACTIONS,
    ACCESS_OBJECT_ACTIONS,
    AccessConnection
} from "../models/types";
import { McpAccessClient } from "../mcp/mcpAccessClient";
import { ConnectionStore } from "../services/connectionStore";
import { isAccessDatabaseOpenError, offerAccessRestart } from "../utils/accessRecovery";
import { rt } from "../utils/runtimeL10n";

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

        if (element instanceof ActionNode) {
            return [];
        }

        return [];
    }

    private getConnectionNodes(): AccessTreeNode[] {
        const connections = this.connectionStore.getAll();
        if (connections.length === 0) {
            return [new MessageNode(rt("tree.message.noConnections"))];
        }

        return connections.map((connection) => new ConnectionNode(connection));
    }

    private async getCategoryObjects(
        connection: AccessConnection,
        categoryKey: CategoryNode["categoryKey"],
        attempt = 0
    ): Promise<AccessTreeNode[]> {
        try {
            const actionNodes = this.getCategoryActionNodes(connection, categoryKey);

            if (categoryKey === "relationships") {
                const relationships = await this.mcpClient.listRelationships(connection);
                return [...actionNodes, ...this.mapObjects(connection, categoryKey, relationships)];
            }

            if (categoryKey === "references") {
                const references = await this.mcpClient.listReferences(connection);
                return [...actionNodes, ...this.mapObjects(connection, categoryKey, references)];
            }

            const category = ACCESS_CATEGORIES.find((item) => item.key === categoryKey);
            if (!category?.toolObjectType) {
                return actionNodes.length > 0
                    ? actionNodes
                    : [new MessageNode(rt("tree.message.unsupportedCategory"))];
            }

            const objects = await this.mcpClient.listObjects(connection, category.toolObjectType);
            return [...actionNodes, ...this.mapObjects(connection, categoryKey, objects)];
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

            return [new MessageNode(rt("tree.message.error", message))];
        }
    }

    private mapObjects(
        connection: AccessConnection,
        categoryKey: CategoryNode["categoryKey"],
        objects: Array<{ name: string; objectType: string; metadata?: Record<string, unknown> }>
    ): AccessTreeNode[] {
        if (objects.length === 0) {
            return [new MessageNode(rt("tree.message.noItems"))];
        }

        return objects
            .slice()
            .sort((a, b) => a.name.localeCompare(b.name))
            .map((object) => new ObjectNode(connection, categoryKey, object));
    }

    private getObjectBranches(node: ObjectNode): AccessTreeNode[] {
        const { connection, categoryKey, objectInfo } = node;

        if (categoryKey === "tables") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "tableFieldsBranch", rt("tree.label.fields")),
                new DetailNode(connection, categoryKey, objectInfo, "tableDataTableAction", rt("tree.label.dataTableTop100")),
                new DetailNode(connection, categoryKey, objectInfo, "tableDataJsonAction", rt("tree.label.dataJsonTop100")),
                new DetailNode(connection, categoryKey, objectInfo, "tableManagementBranch", rt("tree.label.management"))
            ];
        }

        if (categoryKey === "queries") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "querySqlAction", "SQL"),
                new DetailNode(connection, categoryKey, objectInfo, "queryRunTableAction", rt("tree.label.runTableTop200")),
                new DetailNode(connection, categoryKey, objectInfo, "queryRunJsonAction", rt("tree.label.runJsonTop200")),
                new DetailNode(connection, categoryKey, objectInfo, "queryManagementBranch", rt("tree.label.management"))
            ];
        }

        if (categoryKey === "modules") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "moduleProceduresBranch", rt("tree.label.procedures")),
                new DetailNode(connection, categoryKey, objectInfo, "moduleCodeAction", rt("tree.label.fullCode")),
                new DetailNode(connection, categoryKey, objectInfo, "moduleManagementBranch", rt("tree.label.management"))
            ];
        }

        if (categoryKey === "forms") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "formPropertiesBranch", rt("tree.label.properties")),
                new DetailNode(connection, categoryKey, objectInfo, "formControlsBranch", rt("tree.label.controls")),
                new DetailNode(connection, categoryKey, objectInfo, "formLayoutAction", rt("tree.label.layoutPositions")),
                new DetailNode(connection, categoryKey, objectInfo, "formScreenshotAction", rt("tree.label.screenshot")),
                new DetailNode(connection, categoryKey, objectInfo, "formProceduresBranch", rt("tree.label.vbaProcedures")),
                new DetailNode(connection, categoryKey, objectInfo, "formCodeAction", rt("tree.label.vbaCode"))
            ];
        }

        if (categoryKey === "reports") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "reportPropertiesBranch", rt("tree.label.properties")),
                new DetailNode(connection, categoryKey, objectInfo, "reportControlsBranch", rt("tree.label.controls")),
                new DetailNode(connection, categoryKey, objectInfo, "reportLayoutAction", rt("tree.label.layoutPositions")),
                new DetailNode(connection, categoryKey, objectInfo, "reportScreenshotAction", rt("tree.label.screenshot")),
                new DetailNode(connection, categoryKey, objectInfo, "reportProceduresBranch", rt("tree.label.vbaProcedures")),
                new DetailNode(connection, categoryKey, objectInfo, "reportCodeAction", rt("tree.label.vbaCode"))
            ];
        }

        if (categoryKey === "macros") {
            return [
                new DetailNode(connection, categoryKey, objectInfo, "macroCodeAction", rt("tree.label.definition"))
            ];
        }

        return [];
    }

    private async getDetailChildren(node: DetailNode, attempt = 0): Promise<AccessTreeNode[]> {
        try {
            if (node.detailKind === "tableFieldsBranch") {
                const fields = await this.mcpClient.getTableFields(node.connection, node.objectInfo.name);
                if (fields.length === 0) {
                    return [new MessageNode(rt("tree.message.noFields"))];
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

            if (node.detailKind === "queryManagementBranch" || node.detailKind === "moduleManagementBranch") {
                return this.getObjectActionNodes(node.connection, node.categoryKey, node.objectInfo);
            }

            if (node.detailKind === "tableManagementBranch") {
                return this.getObjectActionNodes(node.connection, node.categoryKey, node.objectInfo);
            }

            if (node.detailKind === "moduleProceduresBranch") {
                const procedures = await this.mcpClient.getModuleProcedures(
                    node.connection,
                    "module",
                    node.objectInfo.name
                );

                if (procedures.length === 0) {
                    return [new MessageNode(rt("tree.message.noProcedures"))];
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
                    return [new MessageNode(rt("tree.message.noProcedures"))];
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
                    return [new MessageNode(rt("tree.message.noProperties"))];
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
                    return [new MessageNode(rt("tree.message.noControls"))];
                }

                return controls.map((ctrl) => {
                    const controlProcedures = proceduresByControl.get(normalizeName(ctrl.name)) ?? [];
                    const description = [ctrl.type_name, ctrl.control_source ? `source: ${ctrl.control_source}` : undefined]
                        .concat(controlProcedures.length > 0 ? [rt("tree.label.eventCount", controlProcedures.length)] : [])
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
                    return [new MessageNode(rt("tree.message.noControls"))];
                }

                return controls.map((ctrl) => {
                    const controlProcedures = proceduresByControl.get(normalizeName(ctrl.name)) ?? [];
                    const description = [ctrl.type_name, ctrl.control_source ? `source: ${ctrl.control_source}` : undefined]
                        .concat(controlProcedures.length > 0 ? [rt("tree.label.eventCount", controlProcedures.length)] : [])
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
                        rt("tree.label.fullProperties"),
                        node.payload
                    ),
                    new DetailNode(
                        node.connection,
                        node.categoryKey,
                        node.objectInfo,
                        "controlCodeBranch",
                        rt("tree.label.associatedCode"),
                        node.payload
                    )
                ];
            }

            if (node.detailKind === "controlCodeBranch") {
                const objectType = String(node.payload?.objectType ?? "form") as "form" | "report";
                const controlName = String(node.payload?.name ?? "");
                if (!controlName) {
                    return [new MessageNode(rt("tree.message.controlNoName"))];
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
                    return [new MessageNode(rt("tree.message.noAssociatedProcedures"))];
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

            return [new MessageNode(rt("tree.message.error", message))];
        }
    }

    private getCategoryActionNodes(
        connection: AccessConnection,
        categoryKey: CategoryNode["categoryKey"]
    ): ActionNode[] {
        const actions = ACCESS_CATEGORY_ACTIONS[categoryKey] ?? [];
        return actions.map((action) => new ActionNode(connection, categoryKey, action));
    }

    private getObjectActionNodes(
        connection: AccessConnection,
        categoryKey: CategoryNode["categoryKey"],
        objectInfo: ObjectNode["objectInfo"]
    ): AccessTreeNode[] {
        const actions = ACCESS_OBJECT_ACTIONS[categoryKey] ?? [];
        if (actions.length === 0) {
            return [new MessageNode(rt("tree.message.noActionsAvailable"))];
        }

        return actions.map((action) => new ActionNode(connection, categoryKey, action, objectInfo));
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
            const branchNode = new DetailNode(connection, categoryKey, objectInfo, branchKind, rt("tree.label.controls"));

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
        parts.push(rt("tree.label.line", startLine));
    }
    if (typeof count === "number") {
        parts.push(rt("tree.label.lines", count));
    }

    return parts.length > 0 ? parts.join(" | ") : undefined;
}
