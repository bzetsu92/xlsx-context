import { normalizeMarkdownTableCell } from "./markdown-normalize.js";

function escapeMarkdownTableCell(text: string): string {
    return normalizeMarkdownTableCell(text);
}

/**
 * Escapes cell text for Markdown tables.
 */
export { escapeMarkdownTableCell };
