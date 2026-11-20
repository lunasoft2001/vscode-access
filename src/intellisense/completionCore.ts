import type { AccessConnection } from "../models/types";
import type { AccessMetadataService } from "../services/accessMetadataService";
import { recognizeVbaContext } from "./vbaRecognizer";
import { resolveSqlQualifier } from "./sqlAliasScanner";

export type CompletionEntryKind = "form" | "control" | "field";

/** Editor-agnostic completion entry; the vscode adapter turns these into CompletionItems. */
export interface CompletionEntry {
    label: string;
    insertText: string;
    filterText: string;
    detail?: string;
    kind: CompletionEntryKind;
    sortText?: string;
    /** Characters before the cursor that the insertion replaces (the partial the user already typed). */
    replaceBefore: number;
    /** Characters after the cursor that the insertion replaces (a closing "]" that is already there). */
    replaceAfter: number;
}

type MetadataReader = Pick<AccessMetadataService, "getForms" | "getFormControls" | "getTables" | "getTableFields">;

const SIMPLE_IDENTIFIER = /^\p{L}[\p{L}\p{N}_]*$/u;
const SQL_WORDS_NEEDING_BRACKETS = new Set([
    "all", "and", "any", "as", "asc", "between", "by", "desc", "from", "group", "in", "index", "is", "join", "key",
    "like", "not", "null", "on", "or", "order", "percent", "select", "set", "table", "top", "user", "values", "where"
]);

/**
 * VBA: `Forms!<partial>` lists forms, `Forms!<form>!<partial>` lists that form's controls.
 * Never calls the metadata service unless one of those two shapes is recognized.
 */
export async function getVbaCompletions(
    metadata: MetadataReader,
    connection: AccessConnection,
    linePrefix: string,
    lineSuffix: string
): Promise<CompletionEntry[]> {
    const context = recognizeVbaContext(linePrefix);
    if (!context) {
        return [];
    }

    if (context.kind === "forms") {
        const forms = await metadata.getForms(connection);
        return forms.map((form) => buildEntry(form.name, "form", context.partial, lineSuffix, undefined));
    }

    const form = (await metadata.getForms(connection)).find((item) => sameName(item.name, context.formName));
    if (!form) {
        return [];
    }

    const controls = await metadata.getFormControls(connection, form.name);
    return controls.map((control, index) =>
        buildEntry(control.name, "control", context.partial, lineSuffix, control.type, index));
}

/** SQL: `<alias>.<partial>` inside a simple SELECT lists the fields of the table the alias refers to. */
export async function getSqlCompletions(
    metadata: MetadataReader,
    connection: AccessConnection,
    text: string,
    offset: number
): Promise<CompletionEntry[]> {
    const context = resolveSqlQualifier(text, offset);
    if (!context) {
        return [];
    }

    const table = (await metadata.getTables(connection)).find((item) => sameName(item.name, context.source));
    if (!table) {
        return [];
    }

    const lineSuffix = text.slice(offset, offset + 1);
    const fields = await metadata.getTableFields(connection, table.name);
    return fields.map((field, index) => {
        const description = [field.type, field.size ? `(${field.size})` : undefined, field.required ? "required" : undefined]
            .filter(Boolean)
            .join(" ");
        return buildEntry(
            field.name,
            "field",
            text.slice(context.partialStart, offset),
            lineSuffix,
            description || undefined,
            index,
            true
        );
    });
}

function buildEntry(
    name: string,
    kind: CompletionEntryKind,
    partial: string,
    textAfterCursor: string,
    detail: string | undefined,
    order?: number,
    checkSqlWords = false
): CompletionEntry {
    const typedBracket = partial.startsWith("[");
    const needsBracket = typedBracket
        || !SIMPLE_IDENTIFIER.test(name)
        || (checkSqlWords && SQL_WORDS_NEEDING_BRACKETS.has(name.toLowerCase()));
    const bracketed = `[${name}]`;

    return {
        label: name,
        insertText: needsBracket ? bracketed : name,
        filterText: typedBracket ? bracketed : name,
        detail,
        kind,
        sortText: order === undefined ? undefined : String(order).padStart(5, "0"),
        replaceBefore: partial.length,
        replaceAfter: typedBracket && textAfterCursor.startsWith("]") ? 1 : 0
    };
}

function sameName(left: string, right: string): boolean {
    return left.toLowerCase() === right.toLowerCase();
}
