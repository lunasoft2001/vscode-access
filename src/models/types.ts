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
            label: "Nueva tabla guiada...",
            command: "accessExplorer.createTableDesigner",
            description: "Abrir un diseñador guiado para crear una tabla"
        },
        {
            kind: "createTableDdl",
            label: "Nueva tabla DDL...",
            command: "accessExplorer.createTableDdl",
            description: "Abrir una plantilla SQL DDL para crear una tabla"
        }
    ],
    modules: [
        {
            kind: "createModule",
            label: "Nuevo modulo...",
            command: "accessExplorer.createModule",
            description: "Crear un modulo VBA"
        }
    ],
    queries: [
        {
            kind: "newQuery",
            label: "Nueva consulta guardada...",
            command: "accessExplorer.newQuery",
            description: "Abrir un editor SQL para una consulta guardada de Access"
        }
    ]
};

export const ACCESS_OBJECT_ACTIONS: Partial<Record<AccessCategoryKey, AccessTreeActionDefinition[]>> = {
    tables: [
        {
            kind: "editTableDesigner",
            label: "Editar tabla guiada...",
            command: "accessExplorer.editTableDesigner",
            description: "Abrir un diseñador guiado para modificar esta tabla"
        },
        {
            kind: "editTableDdl",
            label: "Editar tabla DDL...",
            command: "accessExplorer.editTableDdl",
            description: "Abrir una plantilla SQL DDL para modificar esta tabla"
        }
    ],
    modules: [
        {
            kind: "compileModule",
            label: "Compilar modulo",
            command: "accessExplorer.compileModule",
            description: "Compilar solo este modulo"
        },
        {
            kind: "deleteModule",
            label: "Eliminar modulo",
            command: "accessExplorer.deleteModule",
            description: "Eliminar este modulo VBA"
        }
    ],
    queries: [
        {
            kind: "saveQueryToAccess",
            label: "Guardar consulta en Access",
            command: "accessExplorer.saveQueryToAccess",
            description: "Guardar el SQL editado como QueryDef"
        },
        {
            kind: "deleteQuery",
            label: "Eliminar consulta",
            command: "accessExplorer.deleteQuery",
            description: "Eliminar esta consulta guardada"
        }
    ]
};

export const ACCESS_CATEGORIES: AccessCategory[] = [
    { key: "tables", label: "Tablas", toolObjectType: "table" },
    { key: "queries", label: "Consultas", toolObjectType: "query" },
    { key: "forms", label: "Formularios", toolObjectType: "form" },
    { key: "reports", label: "Informes", toolObjectType: "report" },
    { key: "macros", label: "Macros", toolObjectType: "macro" },
    { key: "modules", label: "Modulos/VBA", toolObjectType: "module" },
    { key: "relationships", label: "Relaciones" },
    { key: "references", label: "Referencias VBA" }
];
