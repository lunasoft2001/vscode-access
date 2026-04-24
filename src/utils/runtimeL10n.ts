import * as vscode from "vscode";

type RuntimeLanguage = "en" | "es";

const translations = {
    en: {
        "secondBrain.status.tooltip": "SecondBrain generation status",
        "sql.status.tooltip": "Select the Access connection for the SQL editor",
        "sql.connection.alreadyRegisteredUpdated": "The MCP server was already registered and is up to date.",
        "object.descriptionSeparator": " \u00b7 ",
        "query.sql.empty": "The SQL query is empty.",
        "sql.confirm.title": "Confirmation required",
        "sql.confirm.detail": "You are about to run a {0} statement on \"{1}\".\n\nThis operation may modify or delete data. Continue?",
        "sql.result.connection": "Connection",
        "sql.pickProfile.title": "Choose an Access connection profile",
        "sql.pickProfile.placeholder": "Choose a connection profile from the list below",
        "sql.openEditor.pickConnection": "Choose the Access connection for this SQL editor",
        "sql.input.prompt": "SQL \u00b7 {0}",
        "sql.editor.emptySelection": "The editor is empty or there is no selected text.",
        "sql.execute.pickConnection": "Choose the Access connection to run the SQL",
        "query.delete.title": "Delete query \"{0}\"?",
        "query.delete.detail": "It will be deleted from \"{0}\".",
        "compileVba.docTitle": "VBA compilation \u00b7 {0}",
        "compileVba.errorMessage": "VBA compilation with errors in \"{0}\". Review the opened document.",
        "compileVba.successMessage": "VBA compilation successful: {0}",
        "compileVba.errorDocTitle": "VBA compilation error \u00b7 {0}",
        "compactRepair.title": "Compact & Repair \"{0}\"?",
        "compactRepair.detail": "This will compact and repair \"{0}\". Access must be closed or the file must not be locked.",
        "closeAccess.title": "Close Microsoft Access?",
        "closeAccess.detail": "The extension will try to close the Access instance controlled by MCP.",
        "closeAccess.success": "Microsoft Access was closed successfully.",
        "openIndex": "Open index",
        "secondBrain.selectObject": "Select an object in the tree to generate its individual SecondBrain.",
        "layout.title": "Layout",
        "layout.controls": "Controls",
        "layout.scale": "Scale",
        "layout.size": "Size",
        "layout.clickHint": "Click a control to reveal it in the tree",
        "controlProps.titleSeparator": " \u00b7 ",
        "controlProps.reset": "Reset",
        "controlProps.save": "Save to Access",
        "results.sep": " \u00b7 ",
        "results.filter.placeholder": "Filter...",
        "results.exportCsv": "CSV",
        "results.copySelection": "Copy selection"
    },
    es: {
        "secondBrain.status.tooltip": "Estado de generacion de SecondBrain",
        "sql.status.tooltip": "Selecciona la conexion Access para el editor SQL",
        "sql.connection.alreadyRegisteredUpdated": "El servidor MCP ya estaba registrado y esta actualizado.",
        "object.descriptionSeparator": " \u00b7 ",
        "query.sql.empty": "La consulta SQL esta vacia.",
        "sql.confirm.title": "Confirmacion requerida",
        "sql.confirm.detail": "Vas a ejecutar una sentencia {0} en \"{1}\".\n\nEsta operacion puede modificar o eliminar datos. \u00bfContinuar?",
        "sql.result.connection": "Conexion",
        "sql.pickProfile.title": "Elegir un perfil de conexion Access",
        "sql.pickProfile.placeholder": "Elegir un perfil de conexion de la lista siguiente",
        "sql.openEditor.pickConnection": "Elige la conexion Access para este editor SQL",
        "sql.input.prompt": "SQL \u00b7 {0}",
        "sql.editor.emptySelection": "El editor esta vacio o no hay texto seleccionado.",
        "sql.execute.pickConnection": "Elige la conexion Access para ejecutar el SQL",
        "query.delete.title": "\u00bfEliminar la consulta \"{0}\"?",
        "query.delete.detail": "Se eliminara de \"{0}\".",
        "compileVba.docTitle": "Compilacion VBA \u00b7 {0}",
        "compileVba.errorMessage": "Compilacion VBA con errores en \"{0}\". Revisa el documento abierto.",
        "compileVba.successMessage": "Compilacion VBA correcta: {0}",
        "compileVba.errorDocTitle": "Error de compilacion VBA \u00b7 {0}",
        "compactRepair.title": "\u00bfCompact & Repair \"{0}\"?",
        "compactRepair.detail": "Esto compactara y reparara \"{0}\". Access debe estar cerrado o el archivo no bloqueado.",
        "closeAccess.title": "\u00bfCerrar Microsoft Access?",
        "closeAccess.detail": "Se intentara cerrar la instancia de Access controlada por MCP.",
        "closeAccess.success": "Microsoft Access se cerro correctamente.",
        "openIndex": "Abrir indice",
        "secondBrain.selectObject": "Selecciona un objeto en el arbol para generar su SecondBrain individual.",
        "layout.title": "Layout",
        "layout.controls": "Controles",
        "layout.scale": "Escala",
        "layout.size": "Tamano",
        "layout.clickHint": "Haz clic en un control para ir a el en el arbol",
        "controlProps.titleSeparator": " \u00b7 ",
        "controlProps.reset": "Restaurar",
        "controlProps.save": "Guardar en Access",
        "results.sep": " \u00b7 ",
        "results.filter.placeholder": "Filtrar...",
        "results.exportCsv": "CSV",
        "results.copySelection": "Copiar seleccion"
    }
} as const;

function getRuntimeLanguage(): RuntimeLanguage {
    const language = vscode.env.language.toLowerCase();
    if (language.startsWith("es")) {
        return "es";
    }
    return "en";
}

export function rt(key: keyof typeof translations.en, ...args: Array<string | number>): string {
    const lang = getRuntimeLanguage();
    const template = translations[lang][key] ?? translations.en[key];
    return template.replace(/\{(\d+)\}/g, (_, index) => String(args[Number(index)] ?? ""));
}
