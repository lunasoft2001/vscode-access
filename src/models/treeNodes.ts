import * as vscode from "vscode";
import { AccessCategoryKey, AccessConnection, AccessObjectInfo } from "./types";

export type AccessTreeNode = ConnectionNode | CategoryNode | ObjectNode | DetailNode | MessageNode;

export type AccessDetailKind =
    | "tableFieldsBranch"
    | "tableDataJsonAction"
    | "tableDataTableAction"
    | "tableField"
    | "querySqlAction"
    | "queryRunJsonAction"
    | "queryRunTableAction"
    | "moduleProceduresBranch"
    | "formProceduresBranch"
    | "reportProceduresBranch"
    | "procedure"
    | "moduleCodeAction"
    | "formPropertiesBranch"
    | "formControlsBranch"
    | "formLayoutAction"
    | "formScreenshotAction"
    | "formCodeAction"
    | "reportPropertiesBranch"
    | "reportControlsBranch"
    | "reportLayoutAction"
    | "reportScreenshotAction"
    | "reportCodeAction"
    | "property"
    | "control"
    | "controlPropertiesAction"
    | "controlCodeBranch"
    | "controlProcedure"
    | "macroCodeAction";

export class ConnectionNode extends vscode.TreeItem {
    constructor(public readonly connection: AccessConnection) {
        super(connection.name, vscode.TreeItemCollapsibleState.Collapsed);
        this.contextValue = "accessConnection";
        this.description = connection.dbPath;
        this.tooltip = `${connection.name}\n${connection.dbPath}`;
        this.iconPath = new vscode.ThemeIcon("database");
    }
}

export class CategoryNode extends vscode.TreeItem {
    constructor(
        public readonly connection: AccessConnection,
        public readonly categoryKey: AccessCategoryKey,
        label: string
    ) {
        super(label, vscode.TreeItemCollapsibleState.Collapsed);
        this.contextValue = "accessCategory";
        this.iconPath = new vscode.ThemeIcon(iconForCategory(categoryKey));
    }
}

export class ObjectNode extends vscode.TreeItem {
    constructor(
        public readonly connection: AccessConnection,
        public readonly categoryKey: AccessCategoryKey,
        public readonly objectInfo: AccessObjectInfo
    ) {
        super(
            objectInfo.name,
            isBranchCategory(categoryKey)
                ? vscode.TreeItemCollapsibleState.Collapsed
                : vscode.TreeItemCollapsibleState.None
        );
        this.contextValue = "accessObject";
        this.description = objectDescription(categoryKey, objectInfo);
        this.tooltip = `${objectInfo.name}${this.description ? `\n${this.description}` : ""}`;
        this.iconPath = new vscode.ThemeIcon(iconForObject(categoryKey));

        if (!isBranchCategory(categoryKey)) {
            this.command = {
                command: "accessExplorer.showDetails",
                title: "Show details",
                arguments: [this]
            };
        }
    }
}

export class DetailNode extends vscode.TreeItem {
    constructor(
        public readonly connection: AccessConnection,
        public readonly categoryKey: AccessCategoryKey,
        public readonly objectInfo: AccessObjectInfo,
        public readonly detailKind: AccessDetailKind,
        label: string,
        public readonly payload?: Record<string, unknown>,
        description?: string
    ) {
        super(label, isBranchDetail(detailKind)
            ? vscode.TreeItemCollapsibleState.Collapsed
            : vscode.TreeItemCollapsibleState.None);

        this.contextValue = "accessDetail";
        this.iconPath = new vscode.ThemeIcon(iconForDetail(detailKind));
        this.description = description;
        this.tooltip = description ? `${label}\n${description}` : label;

        if (!isBranchDetail(detailKind)) {
            this.command = {
                command: "accessExplorer.showDetails",
                title: "Open",
                arguments: [this]
            };
        }
    }
}

export class MessageNode extends vscode.TreeItem {
    constructor(message: string) {
        super(message, vscode.TreeItemCollapsibleState.None);
        this.contextValue = "accessMessage";
        this.iconPath = new vscode.ThemeIcon("info");
    }
}

function isBranchCategory(categoryKey: AccessCategoryKey): boolean {
    return (
        categoryKey === "tables"
        || categoryKey === "queries"
        || categoryKey === "forms"
        || categoryKey === "reports"
        || categoryKey === "modules"
        || categoryKey === "macros"
    );
}

function isBranchDetail(detailKind: AccessDetailKind): boolean {
    return (
        detailKind === "tableFieldsBranch"
        || detailKind === "moduleProceduresBranch"
        || detailKind === "formProceduresBranch"
        || detailKind === "reportProceduresBranch"
        || detailKind === "formPropertiesBranch"
        || detailKind === "reportPropertiesBranch"
        || detailKind === "formControlsBranch"
        || detailKind === "reportControlsBranch"
        || detailKind === "control"
        || detailKind === "controlCodeBranch"
    );
}

function iconForCategory(categoryKey: AccessCategoryKey): string {
    switch (categoryKey) {
        case "tables":
            return "table";
        case "queries":
            return "search-view-icon";
        case "forms":
            return "preview";
        case "reports":
            return "graph";
        case "macros":
            return "symbol-event";
        case "modules":
            return "file-code";
        case "relationships":
            return "git-compare";
        case "references":
            return "references";
        default:
            return "folder";
    }
}

function iconForObject(categoryKey: AccessCategoryKey): string {
    switch (categoryKey) {
        case "tables":
            return "table";
        case "queries":
            return "symbol-function";
        case "forms":
            return "preview";
        case "reports":
            return "graph-line";
        case "macros":
            return "symbol-event";
        case "modules":
            return "file-code";
        case "relationships":
            return "git-compare";
        case "references":
            return "references";
        default:
            return "symbol-object";
    }
}

function objectDescription(categoryKey: AccessCategoryKey, objectInfo: AccessObjectInfo): string | undefined {
    if (categoryKey === "references") {
        return "reference";
    }

    if (categoryKey === "relationships") {
        return objectInfo.objectType;
    }

    return undefined;
}

function iconForDetail(detailKind: AccessDetailKind): string {
    switch (detailKind) {
        case "tableFieldsBranch":
            return "list-tree";
        case "tableDataJsonAction":
        case "queryRunJsonAction":
            return "json";
        case "tableDataTableAction":
        case "queryRunTableAction":
            return "table";
        case "tableField":
            return "symbol-field";
        case "querySqlAction":
            return "file-code";
        case "queryRunJsonAction":
        case "queryRunTableAction":
            return "play";
        case "moduleProceduresBranch":
        case "formProceduresBranch":
        case "reportProceduresBranch":
            return "symbol-method";
        case "procedure":
            return "symbol-method";
        case "moduleCodeAction":
            return "file-code";
        case "formPropertiesBranch":
        case "reportPropertiesBranch":
            return "settings-gear";
        case "formControlsBranch":
        case "reportControlsBranch":
            return "list-selection";
        case "formLayoutAction":
        case "reportLayoutAction":
            return "layout";
        case "formScreenshotAction":
        case "reportScreenshotAction":
            return "device-camera";
        case "formCodeAction":
        case "reportCodeAction":
        case "macroCodeAction":
            return "file-code";
        case "control":
            return "symbol-property";
        case "controlPropertiesAction":
            return "settings";
        case "controlCodeBranch":
            return "symbol-event";
        case "controlProcedure":
            return "symbol-method";
        case "property":
            return "settings";
        default:
            return "symbol-misc";
    }
}
