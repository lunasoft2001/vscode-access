import { rt } from "../utils/runtimeL10n";

export type AccessCategoryKey =
    | "tables"
    | "queries"
    | "forms"
    | "reports"
    | "macros"
    | "modules"
    | "relationships"
    | "references";

export interface AccessConnection {
    id: string;
    name: string;
    dbPath: string;
}

export interface AccessObjectInfo {
    name: string;
    objectType: string;
    metadata?: Record<string, unknown>;
}

export interface AccessObjectDocument {
    title: string;
    language: string;
    content: string;
    codeMeta?: {
        connection: AccessConnection;
        objectType: "module" | "form" | "report";
        objectName: string;
        procedureName?: string;
        replaceStartLine?: number;
        replaceCount?: number;
        isNew?: boolean;
    };
}

export interface AccessTableFieldInfo {
    name: string;
    type?: string;
    size?: number;
    required?: boolean;
}

export interface AccessControlInfo {
    name: string;
    type_name?: string;
    control_source?: string;
    caption?: string;
    source_object?: string;
    left?: number;
    top?: number;
    width?: number;
    height?: number;
}

export interface AccessScreenshotInfo {
    path?: string;
    width?: number;
    height?: number;
    metadata?: Record<string, unknown>;
}

export interface AccessPropertyInfo {
    name: string;
    value: string;
}

export interface AccessProcedureInfo {
    name: string;
    start_line?: number;
    count?: number;
}

export interface AccessQueryPreview {
    sql: string;
    rowCount?: number;
    rowsAffected?: number;
    rows?: Record<string, unknown>[];
    payload: unknown;
}

export interface AccessCategory {
    key: AccessCategoryKey;
    label: string;
    toolObjectType?: string;
}

export type AccessTreeActionKind =
    | "createTableDesigner"
    | "editTableDesigner"
    | "createTableDdl"
    | "editTableDdl"
    | "createModule"
    | "deleteModule"
    | "compileModule"
    | "newQuery"
    | "saveQueryToAccess"
    | "deleteQuery";

export interface AccessTreeActionDefinition {
    kind: AccessTreeActionKind;
    label: string;
    command: string;
    description?: string;
}

export const ACCESS_CATEGORY_ACTIONS: Partial<Record<AccessCategoryKey, AccessTreeActionDefinition[]>> = {
    tables: [
        {
            kind: "createTableDesigner",
            label: rt("tree.action.newGuidedTable"),
            command: "accessExplorer.createTableDesigner",
            description: rt("tree.action.newGuidedTable.desc")
        },
        {
            kind: "createTableDdl",
            label: rt("tree.action.newDdlTable"),
            command: "accessExplorer.createTableDdl",
            description: rt("tree.action.newDdlTable.desc")
        }
    ],
    modules: [
        {
            kind: "createModule",
            label: rt("tree.action.newModule"),
            command: "accessExplorer.createModule",
            description: rt("tree.action.newModule.desc")
        }
    ],
    queries: [
        {
            kind: "newQuery",
            label: rt("tree.action.newSavedQuery"),
            command: "accessExplorer.newQuery",
            description: rt("tree.action.newSavedQuery.desc")
        }
    ]
};

export const ACCESS_OBJECT_ACTIONS: Partial<Record<AccessCategoryKey, AccessTreeActionDefinition[]>> = {
    tables: [
        {
            kind: "editTableDesigner",
            label: rt("tree.action.editGuidedTable"),
            command: "accessExplorer.editTableDesigner",
            description: rt("tree.action.editGuidedTable.desc")
        },
        {
            kind: "editTableDdl",
            label: rt("tree.action.editDdlTable"),
            command: "accessExplorer.editTableDdl",
            description: rt("tree.action.editDdlTable.desc")
        }
    ],
    modules: [
        {
            kind: "compileModule",
            label: rt("tree.action.compileModule"),
            command: "accessExplorer.compileModule",
            description: rt("tree.action.compileModule.desc")
        },
        {
            kind: "deleteModule",
            label: rt("tree.action.deleteModule"),
            command: "accessExplorer.deleteModule",
            description: rt("tree.action.deleteModule.desc")
        }
    ],
    queries: [
        {
            kind: "saveQueryToAccess",
            label: rt("tree.action.saveQuery"),
            command: "accessExplorer.saveQueryToAccess",
            description: rt("tree.action.saveQuery.desc")
        },
        {
            kind: "deleteQuery",
            label: rt("tree.action.deleteQuery"),
            command: "accessExplorer.deleteQuery",
            description: rt("tree.action.deleteQuery.desc")
        }
    ]
};

export const ACCESS_CATEGORIES: AccessCategory[] = [
    { key: "tables", label: rt("tree.category.tables"), toolObjectType: "table" },
    { key: "queries", label: rt("tree.category.queries"), toolObjectType: "query" },
    { key: "forms", label: rt("tree.category.forms"), toolObjectType: "form" },
    { key: "reports", label: rt("tree.category.reports"), toolObjectType: "report" },
    { key: "macros", label: rt("tree.category.macros"), toolObjectType: "macro" },
    { key: "modules", label: rt("tree.category.modules"), toolObjectType: "module" },
    { key: "relationships", label: rt("tree.category.relationships") },
    { key: "references", label: rt("tree.category.references") }
];
