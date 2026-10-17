import { normalizeMarkdownNewlines } from "./markdown-normalize.js";

type MarkdownLiteralPart = {
    kind: "text" | "escaped";
    text: string;
    rawText: string;
};
type OrderedListMarkerMatch = {
    digits: string;
    dotIndex: number;
} | null;

function escapeMarkdownLineStartSegment(text: string): string {
    return String(text || "")
        .replace(/^(\s*)([#>])/u, "$1\\$2")
        .replace(/^(\s*)([-+*])(\s+)/u, "$1\\$2$3")
        .replace(/^(\s*)(\d+)\.(\s+)/u, "$1$2\\.$3");
}

/** Escapes line-start markers so text never becomes a heading or list. */
function escapeMarkdownLineStart(text: string): string {
    return normalizeMarkdownNewlines(text)
        .split("\n")
        .map((line) => escapeMarkdownLineStartSegment(line))
        .join("\n");
}

/** Escapes one character for Markdown literals. Guards line-start markers only at line starts. */
function getEscapedMarkdownLiteralText(ch: string, atLineStart: boolean, nextChar: string): string | null {
    if (ch === "\\") return "\\\\";
    if (ch === "&") return "&amp;";
    if (ch === "<") return "&lt;";
    if (ch === ">") return "&gt;";
    if (/[`*_{}\[\]()!|~]/.test(ch)) return `\\${ch}`;
    if (atLineStart && ch === "#") return `\\${ch}`;
    if (atLineStart && /[-+*]/.test(ch) && /\s/u.test(nextChar)) return `\\${ch}`;
    return null;
}

/** Reads an ordered list marker such as `1. `. Returns null without digits, dot, and space. */
function parseOrderedListMarker(source: string, index: number, atLineStart: boolean): OrderedListMarkerMatch {
    if (!atLineStart || !/\d/u.test(source[index] || "")) {
        return null;
    }
    let digits = source[index];
    let cursor = index + 1;
    while (cursor < source.length && /\d/u.test(source[cursor])) {
        digits += source[cursor];
        cursor += 1;
    }
    if (source[cursor] !== "." || !/\s/u.test(source[cursor + 1] || "")) {
        return null;
    }
    return {
        digits,
        dotIndex: cursor,
    };
}

/** Splits text into plain and escaped parts. Keeps the raw character per part. */
function escapeMarkdownLiteralParts(text: string): MarkdownLiteralPart[] {
    const source = String(text || "");
    const parts: MarkdownLiteralPart[] = [];
    let buffer = "";

    function pushTextBuffer() {
        if (!buffer) return;
        parts.push({ kind: "text", text: buffer, rawText: buffer });
        buffer = "";
    }

    function pushEscaped(textValue: string, rawText: string) {
        pushTextBuffer();
        if (!textValue) return;
        parts.push({ kind: "escaped", text: textValue, rawText });
    }

    for (let index = 0; index < source.length; index += 1) {
        const ch = source[index];
        const atLineStart = index === 0;
        const next = source[index + 1] || "";
        const escapedText = getEscapedMarkdownLiteralText(ch, atLineStart, next);
        if (escapedText) {
            pushEscaped(escapedText, ch);
            continue;
        }
        const orderedListMarker = parseOrderedListMarker(source, index, atLineStart);
        if (orderedListMarker) {
            pushTextBuffer();
            parts.push({ kind: "text", text: orderedListMarker.digits, rawText: orderedListMarker.digits });
            parts.push({ kind: "escaped", text: "\\.", rawText: "." });
            index = orderedListMarker.dotIndex;
            continue;
        }
        buffer += ch;
    }

    pushTextBuffer();
    return parts;
}

function escapeMarkdownLiteralText(text: string): string {
    return normalizeMarkdownNewlines(text)
        .split("\n")
        .map((line) =>
            escapeMarkdownLiteralParts(line)
                .map((part) => part.text)
                .join("")
        )
        .join("\n");
}

/**
 * Escapes a name from the file for one heading line, such as a sheet name.
 * Keeps the text readable: backslash-escapes only marks that change a heading
 * line (emphasis, code, links, HTML, strike), "&" before an entity, and a trailing "#".
 */
function escapeMarkdownHeadingText(text: string): string {
    return (
        String(text || "")
            .replace(/\s+/g, " ")
            .trim()
            .replace(/[\\`*_\[\]<>~]/g, "\\$&")
            .replace(/&(?=#?[A-Za-z0-9]+;)/g, "&amp;")
            // A trailing run of "#" would close the heading and vanish.
            .replace(/(#+)$/, (marks) => marks.replace(/#/g, "\\#"))
    );
}

/** Escapes cell text for Markdown output. */
export { escapeMarkdownHeadingText, escapeMarkdownLineStart, escapeMarkdownLiteralParts, escapeMarkdownLiteralText };
