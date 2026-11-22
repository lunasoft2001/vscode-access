/**
 * Small lightweight scanner for the SELECT scope around the cursor. It is intentionally not a SQL parser:
 * it understands FROM/JOIN table references with optional aliases for one simple SELECT statement and
 * returns nothing for syntax it does not support (subqueries, UNION, TRANSFORM/PIVOT, PARAMETERS, ...).
 */

export interface SourceBinding {
    alias: string;
    source: string;
}

export interface SqlQualifierContext {
    /** The text before the dot, e.g. "o" in "o.Ord". */
    qualifier: string;
    /** The table the qualifier refers to. */
    source: string;
    /** What the user already typed after the dot (may start with "["). */
    partial: string;
    /** Document offset where the partial starts (the cursor offset when nothing was typed). */
    partialStart: number;
}

type TokenKind = "word" | "bracket" | "string" | "comment" | "punct";

interface Token {
    kind: TokenKind;
    text: string;
    start: number;
    end: number;
    /** False for an unterminated string, bracket or comment. */
    closed: boolean;
}

const MAX_SCANNED_LENGTH = 200_000;

const NOT_AN_ALIAS = new Set([
    "on", "where", "group", "order", "having", "inner", "left", "right", "outer", "full", "cross", "join",
    "in", "union", "select", "from", "as", "set", "values", "into", "with", "like", "between", "and", "or",
    "not", "is", "using"
]);

const UNSUPPORTED_WORDS = new Set(["union"]);

const TOKEN_PATTERN = /(--[^\r\n]*)|(\/\*[\s\S]*?(?:\*\/|$))|('(?:[^'\r\n]|'')*(?:'|(?=[\r\n]|$)))|("(?:[^"\r\n]|"")*(?:"|(?=[\r\n]|$)))|(\[[^\]\r\n]*\]?)|([\p{L}\p{N}_]+)|(\s+)|([\s\S])/gu;

/** Table references visible in the statement that contains `offset`. Empty when the statement is unsupported. */
export function findSourceBindings(text: string, offset: number): SourceBinding[] {
    if (text.length > MAX_SCANNED_LENGTH) {
        return [];
    }

    const tokens = tokenize(text);
    if (isInsideStringOrComment(tokens, offset)) {
        return [];
    }

    return bindSources(statementAt(tokens, offset))
        .filter((binding): binding is SourceBinding => binding !== undefined);
}

/** Resolves `<alias>.<partial>` at the cursor to the table the alias refers to, or undefined. */
export function resolveSqlQualifier(text: string, offset: number): SqlQualifierContext | undefined {
    if (text.length > MAX_SCANNED_LENGTH) {
        return undefined;
    }

    const tokens = tokenize(text);
    if (isInsideStringOrComment(tokens, offset)) {
        return undefined;
    }

    const statement = statementAt(tokens, offset);
    const located = locateQualifier(text, statement, offset);
    if (!located) {
        return undefined;
    }

    const binding = bindSources(statement).find((item) => item?.alias.toLowerCase() === located.qualifier.toLowerCase());
    return binding ? { ...located, source: binding.source } : undefined;
}

function tokenize(text: string): Token[] {
    const tokens: Token[] = [];
    TOKEN_PATTERN.lastIndex = 0;

    for (let match = TOKEN_PATTERN.exec(text); match; match = TOKEN_PATTERN.exec(text)) {
        const [value, lineComment, blockComment, singleQuoted, doubleQuoted, bracket, word, space] = match;
        if (space) {
            continue;
        }

        const start = match.index;
        const end = start + value.length;
        if (lineComment) {
            tokens.push({ kind: "comment", text: value, start, end, closed: false });
        } else if (blockComment) {
            tokens.push({ kind: "comment", text: value, start, end, closed: value.length >= 4 && value.endsWith("*/") });
        } else if (singleQuoted || doubleQuoted) {
            tokens.push({ kind: "string", text: value, start, end, closed: value.length >= 2 && value.endsWith(value[0]) });
        } else if (bracket) {
            tokens.push({ kind: "bracket", text: value, start, end, closed: value.endsWith("]") });
        } else if (word) {
            tokens.push({ kind: "word", text: value, start, end, closed: true });
        } else {
            tokens.push({ kind: "punct", text: value, start, end, closed: true });
        }
    }

    return tokens;
}

function isInsideStringOrComment(tokens: Token[], offset: number): boolean {
    return tokens.some((token) =>
        (token.kind === "string" || token.kind === "comment")
        && offset > token.start
        && (token.closed ? offset < token.end : offset <= token.end));
}

/** Comment-free tokens of the ';'-delimited statement containing the offset. */
function statementAt(tokens: Token[], offset: number): Token[] {
    let current: Token[] = [];
    for (const token of tokens) {
        if (token.kind === "comment") {
            continue;
        }
        if (token.kind === "punct" && token.text === ";") {
            if (token.start >= offset) {
                return current;
            }
            current = [];
            continue;
        }
        current.push(token);
    }
    return current;
}

function locateQualifier(
    text: string,
    statement: Token[],
    offset: number
): Pick<SqlQualifierContext, "qualifier" | "partial" | "partialStart"> | undefined {
    let index = statement.length - 1;
    while (index >= 0 && statement[index].start >= offset) {
        index--;
    }
    if (index < 0) {
        return undefined;
    }

    let partial = "";
    let partialStart = offset;
    const last = statement[index];
    if ((last.kind === "word" || last.kind === "bracket") && offset <= last.end) {
        partial = text.slice(last.start, offset);
        partialStart = last.start;
        index--;
    }

    const dot = statement[index];
    const qualifier = statement[index - 1];
    if (!dot || dot.kind !== "punct" || dot.text !== "." || dot.end !== partialStart) {
        return undefined;
    }
    if (!qualifier || qualifier.end !== dot.start || !isName(qualifier)) {
        return undefined;
    }

    // "a.b." would be a schema/database qualified name, which is not supported.
    const before = statement[index - 2];
    if (before?.kind === "punct" && before.text === "." && before.end === qualifier.start) {
        return undefined;
    }

    return { qualifier: nameOf(qualifier), partial, partialStart };
}

/** Undefined entries mark ambiguous aliases (declared twice) so they never resolve. */
function bindSources(statement: Token[]): Array<SourceBinding | undefined> {
    if (!isSupportedSelect(statement)) {
        return [];
    }

    const bindings = new Map<string, SourceBinding | undefined>();
    const add = (alias: string, source: string): void => {
        const key = alias.toLowerCase();
        bindings.set(key, bindings.has(key) ? undefined : { alias, source });
    };

    for (let i = 0; i < statement.length; i++) {
        const keyword = statement[i];
        const isFrom = isWord(keyword, "from");
        if (!isFrom && !isWord(keyword, "join")) {
            continue;
        }

        let next = readTableReference(statement, i + 1, add);
        // FROM A a, B b
        while (isFrom && next >= 0 && statement[next]?.kind === "punct" && statement[next].text === ",") {
            next = readTableReference(statement, next + 1, add);
        }
    }

    return [...bindings.values()];
}

/** Reads `[(...] table [[AS] alias]` starting at `index`; returns the index after it, or -1 if unsupported. */
function readTableReference(statement: Token[], index: number, add: (alias: string, source: string) => void): number {
    // Access wraps join groups in parentheses: FROM ((A INNER JOIN B ON ...) INNER JOIN C ON ...)
    while (statement[index]?.kind === "punct" && statement[index].text === "(") {
        index++;
    }

    const table = statement[index];
    if (!table || !isName(table) || (table.kind === "word" && NOT_AN_ALIAS.has(table.text.toLowerCase()))) {
        return -1;
    }

    index++;
    const following = statement[index];
    // schema.table, table IN 'otherdb' and derived tables are not supported.
    if ((following?.kind === "punct" && following.text === ".") || (following && isWord(following, "in"))) {
        return -1;
    }

    if (following && isWord(following, "as")) {
        index++;
    }

    const alias = statement[index];
    if (alias && isName(alias) && !(alias.kind === "word" && NOT_AN_ALIAS.has(alias.text.toLowerCase()))) {
        if (statement[index + 1] && isWord(statement[index + 1], "in")) {
            return -1;
        }
        add(nameOf(alias), nameOf(table));
        return index + 1;
    }

    add(nameOf(table), nameOf(table));
    return index;
}

/** A statement is supported only when it is a plain SELECT; TRANSFORM and PARAMETERS statements fail the first-word check. */
function isSupportedSelect(statement: Token[]): boolean {
    if (statement.length === 0 || !isWord(statement[0], "select")) {
        return false;
    }

    return !statement.some((token, i) =>
        (token.kind === "word" && UNSUPPORTED_WORDS.has(token.text.toLowerCase()))
        || (token.kind === "punct" && token.text === "(" && statement[i + 1] !== undefined && isWord(statement[i + 1], "select")));
}

function isWord(token: Token, word: string): boolean {
    return token.kind === "word" && token.text.toLowerCase() === word;
}

function isName(token: Token): boolean {
    return token.kind === "word" || (token.kind === "bracket" && token.closed);
}

function nameOf(token: Token): string {
    return token.kind === "bracket" ? token.text.slice(1, -1) : token.text;
}
