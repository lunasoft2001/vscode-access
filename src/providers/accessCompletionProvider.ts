import * as vscode from "vscode";
import type { AccessConnection } from "../models/types";
import type { AccessMetadataService } from "../services/accessMetadataService";
import { getSqlCompletions, getVbaCompletions } from "../intellisense/completionCore";
import type { CompletionEntry, CompletionEntryKind } from "../intellisense/completionCore";
import { getLogicalLinePrefix } from "../intellisense/vbaRecognizer";

export interface AccessCompletionOptions {
    metadata: AccessMetadataService;
    /** Connection bound to a VBA document opened through Access Explorer, or undefined for any other document. */
    resolveVbaConnection: (document: vscode.TextDocument) => AccessConnection | undefined;
    /** Connection bound to a SQL document opened through Access Explorer, or undefined for any other document. */
    resolveSqlConnection: (document: vscode.TextDocument) => AccessConnection | undefined;
}

const ITEM_KINDS: Record<CompletionEntryKind, vscode.CompletionItemKind> = {
    form: vscode.CompletionItemKind.Class,
    control: vscode.CompletionItemKind.Property,
    field: vscode.CompletionItemKind.Field
};

export function registerAccessCompletionProviders(options: AccessCompletionOptions): vscode.Disposable {
    const vbaProvider: vscode.CompletionItemProvider = {
        provideCompletionItems(document, position, token) {
            const connection = options.resolveVbaConnection(document);
            if (!connection) {
                return undefined;
            }

            const line = document.lineAt(position.line).text;
            const linePrefix = getLogicalLinePrefix((index) => document.lineAt(index).text, position.line, position.character);
            return resolveEntries(
                getVbaCompletions(options.metadata, connection, linePrefix, line.slice(position.character)),
                position,
                token
            );
        }
    };

    const sqlProvider: vscode.CompletionItemProvider = {
        provideCompletionItems(document, position, token) {
            const connection = options.resolveSqlConnection(document);
            if (!connection) {
                return undefined;
            }

            return resolveEntries(
                getSqlCompletions(options.metadata, connection, document.getText(), document.offsetAt(position)),
                position,
                token
            );
        }
    };

    return vscode.Disposable.from(
        vscode.languages.registerCompletionItemProvider({ language: "vb" }, vbaProvider, "!"),
        vscode.languages.registerCompletionItemProvider({ language: "sql" }, sqlProvider, ".")
    );
}

async function resolveEntries(
    pending: Promise<CompletionEntry[]>,
    position: vscode.Position,
    token: vscode.CancellationToken
): Promise<vscode.CompletionList | undefined> {
    // On cancellation the load keeps running in the metadata service and its result stays cached.
    const entries = await raceCancellation(pending, token);
    if (!entries || entries.length === 0) {
        return undefined;
    }

    return new vscode.CompletionList(entries.map((entry) => toCompletionItem(entry, position)), false);
}

function toCompletionItem(entry: CompletionEntry, position: vscode.Position): vscode.CompletionItem {
    const item = new vscode.CompletionItem(entry.label, ITEM_KINDS[entry.kind]);
    item.insertText = entry.insertText;
    item.filterText = entry.filterText;
    item.detail = entry.detail;
    item.sortText = entry.sortText;
    item.range = new vscode.Range(
        position.translate(0, -entry.replaceBefore),
        position.translate(0, entry.replaceAfter)
    );
    return item;
}

function raceCancellation<T>(pending: Promise<T>, token: vscode.CancellationToken): Promise<T | undefined> {
    if (token.isCancellationRequested) {
        return Promise.resolve(undefined);
    }

    return new Promise<T | undefined>((resolve, reject) => {
        const subscription = token.onCancellationRequested(() => resolve(undefined));
        pending.then(
            (value) => {
                subscription.dispose();
                resolve(value);
            },
            (error) => {
                subscription.dispose();
                reject(error);
            }
        );
    });
}
