export type VbaCompletionContext =
    | { kind: "forms"; partial: string }
    | { kind: "formControls"; formName: string; partial: string };

const MAX_CONTINUATION_LINES = 24;

const NAME = "[\\p{L}\\p{N}_]";
// Anchored to the end of the logical line, so only the text right before the cursor is inspected.
const CONTROLS_TAIL = new RegExp(`(?<![\\p{L}\\p{N}_.!\\]])Forms!(${NAME}+|\\[[^\\]]+\\])!(\\[[^\\]]*|${NAME}*)$`, "iu");
const FORMS_TAIL = new RegExp(`(?<![\\p{L}\\p{N}_.!\\]])Forms!(\\[[^\\]]*|${NAME}*)$`, "iu");

/**
 * Recognizes the two supported contexts at the end of a logical line prefix:
 *   Forms!<partial>             -> forms
 *   Forms!<form>!<partial>      -> controls of <form>
 * Anything else (including text inside a string or comment) returns undefined.
 */
export function recognizeVbaContext(linePrefix: string): VbaCompletionContext | undefined {
    if (!isInCode(linePrefix)) {
        return undefined;
    }

    const controls = CONTROLS_TAIL.exec(linePrefix);
    if (controls) {
        return { kind: "formControls", formName: unbracket(controls[1]), partial: controls[2] };
    }

    const forms = FORMS_TAIL.exec(linePrefix);
    if (forms) {
        return { kind: "forms", partial: forms[1] };
    }

    return undefined;
}

/** Joins the lines continued with " _" so a statement split over several lines is seen as one prefix. */
export function getLogicalLinePrefix(getLineText: (line: number) => string, line: number, character: number): string {
    let prefix = getLineText(line).slice(0, character);
    for (let previous = line - 1, count = 0; previous >= 0 && count < MAX_CONTINUATION_LINES; previous--, count++) {
        const text = getLineText(previous);
        if (!/\s_\s*$/.test(text)) {
            break;
        }
        prefix = text.replace(/\s_\s*$/, " ") + prefix;
    }
    return prefix;
}

export function unbracket(name: string): string {
    return name.startsWith("[") && name.endsWith("]") ? name.slice(1, -1) : name;
}

/** True when the end of the prefix is plain code: not inside a string literal and not in a ' or Rem comment. */
function isInCode(prefix: string): boolean {
    let inString = false;
    let inBracket = false;
    let atStatementStart = true;

    for (let i = 0; i < prefix.length; i++) {
        const ch = prefix[i];

        if (inString) {
            if (ch === "\"") {
                inString = false;
            }
            continue;
        }

        if (inBracket) {
            if (ch === "]") {
                inBracket = false;
            }
            continue;
        }

        if (ch === "[") {
            inBracket = true;
            atStatementStart = false;
            continue;
        }
        if (ch === "'") {
            return false;
        }
        if (ch === "\"") {
            inString = true;
            atStatementStart = false;
            continue;
        }
        if (atStatementStart && /^rem(\s|$)/i.test(prefix.slice(i))) {
            return false;
        }
        if (ch === ":") {
            atStatementStart = true;
        } else if (!/\s/.test(ch)) {
            atStatementStart = false;
        }
    }

    return !inString;
}
