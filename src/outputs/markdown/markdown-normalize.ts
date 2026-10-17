const MARKDOWN_UNSAFE_UNICODE_REGEX =
    /[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F-\u009F\u00AD\u200B-\u200F\u2028\u2029\u202A-\u202E\u2060-\u206F\uFEFF\uFDD0-\uFDEF\uFFFE\uFFFF]/g;
const MARKDOWN_LINE_BREAK_REGEX = /\r\n?|\n/g;

function normalizeMarkdownNewlines(text: string, replacement = "\n"): string {
    return String(text || "").replace(MARKDOWN_LINE_BREAK_REGEX, replacement);
}

/** Collapses line breaks and tabs to spaces. Strips unsafe unicode. */
function normalizeMarkdownText(text: string): string {
    return normalizeMarkdownNewlines(String(text || ""), " ")
        .replace(MARKDOWN_UNSAFE_UNICODE_REGEX, " ")
        .replace(/\t/g, " ");
}

// Leaves pipes after an odd number of backslashes alone. That text is usually already escaped.
function escapeMarkdownPipes(text: string): string {
    return String(text || "").replace(/(\\*)\|/g, (match, slashes: string) =>
        slashes.length % 2 === 1 ? match : `${slashes}\\|`
    );
}

function normalizeMarkdownTableCell(text: string): string {
    return escapeMarkdownPipes(normalizeMarkdownText(text));
}

/** Strips leading heading markers. */
function normalizeMarkdownHeadingText(text: string): string {
    return normalizeMarkdownText(text).replace(/^#+\s*/, "");
}

/** Strips leading list markers. */
function normalizeMarkdownListItemText(text: string): string {
    return normalizeMarkdownText(text).replace(/^([-*+]|\d+\.)\s+/, "");
}

/** Normalizes cell text for Markdown output. */
export {
    normalizeMarkdownNewlines,
    normalizeMarkdownText,
    escapeMarkdownPipes,
    normalizeMarkdownTableCell,
    normalizeMarkdownHeadingText,
    normalizeMarkdownListItemText,
};
