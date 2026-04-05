import * as vscode from "vscode";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
let lastPromptAt = 0;

export function isAccessDatabaseOpenError(message: string): boolean {
    const normalized = message
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase();

    return normalized.includes("base de datos ya esta abierta")
        || normalized.includes("database is already open")
        || normalized.includes("already open");
}

export async function offerAccessRestart(message: string): Promise<boolean> {
    if (!isAccessDatabaseOpenError(message)) {
        return false;
    }

    const now = Date.now();
    if (now - lastPromptAt < 5000) {
        return false;
    }
    lastPromptAt = now;

    const choice = await vscode.window.showWarningMessage(
        "Access indica que la base ya esta abierta. ¿Quieres cerrar todas las instancias de Microsoft Access y reintentar?",
        { modal: true },
        "Reiniciar Access"
    );

    if (choice !== "Reiniciar Access") {
        return false;
    }

    await restartAccessProcesses();
    vscode.window.showInformationMessage("Access reiniciado. Reintentando...");
    return true;
}

export async function restartAccessProcesses(): Promise<void> {
    if (process.platform !== "win32") {
        throw new Error("Reiniciar Access solo esta soportado en Windows.");
    }

    try {
        await execFileAsync("taskkill", ["/IM", "MSACCESS.EXE", "/F"]);
    } catch {
        // If there was no running Access process, taskkill can fail. This is harmless.
    }
}
