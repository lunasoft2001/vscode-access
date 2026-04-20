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
