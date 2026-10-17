import type { MergeRange } from "../../model/index.js";

function colToLetters(col: number): string {
    let current = col;
    let result = "";
    while (current > 0) {
        const remainder = (current - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        current = Math.floor((current - 1) / 26);
    }
    return result;
}

function lettersToCol(letters: string): number {
    let result = 0;
    for (const ch of String(letters || "").toUpperCase()) {
        result = result * 26 + (ch.charCodeAt(0) - 64);
    }
    return result;
}

/** Parses A1 text into 1-based row and column numbers. Strips `$` markers. Returns zeros for invalid addresses. */
function parseCellAddress(address: string): { row: number; col: number } {
    const normalized = String(address || "")
        .trim()
        .replace(/\$/g, "");
    const match = normalized.match(/^([A-Z]+)(\d+)$/i);
    if (!match) {
        return { row: 0, col: 0 };
    }
    return {
        col: lettersToCol(match[1]),
        row: Number(match[2]),
    };
}

function normalizeFormulaAddress(address: string): string {
    return String(address || "")
        .trim()
        .replace(/\$/g, "")
        .toUpperCase();
}

function formatRange(startRow: number, startCol: number, endRow: number, endCol: number): string {
    return `${colToLetters(startCol)}${startRow}-${colToLetters(endCol)}${endRow}`;
}

function parseRangeRef(ref: string): MergeRange {
    const parts = String(ref || "").split(":");
    const start = parseCellAddress(parts[0] || "");
    const end = parseCellAddress(parts[1] || parts[0] || "");
    return {
        startRow: start.row,
        startCol: start.col,
        endRow: end.row,
        endCol: end.col,
        ref,
    };
}

/** Parses `A1:B2` text into normalized endpoint addresses. Returns null for invalid ranges. */
function parseRangeAddress(rawRange: string): { start: string; end: string } | null {
    const match = String(rawRange || "")
        .trim()
        .match(/^(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)$/i);
    if (!match) return null;
    return {
        start: normalizeFormulaAddress(match[1]),
        end: normalizeFormulaAddress(match[2]),
    };
}

/**
 * Parses a sheet-qualified range such as `'My Sheet'!$B$4:$B$7` or `Sheet1!C3`.
 * The sheet name is null when the reference has no sheet part.
 * Returns null for anything else, for example names or multi-area references.
 */
function parseSheetRangeRef(
    ref: string
): { sheetName: string | null; startRow: number; startCol: number; endRow: number; endCol: number } | null {
    const match = String(ref || "")
        .trim()
        .match(/^(?:(?:'((?:[^']|'')+)'|([^'!:()]+))!)?(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?$/i);
    if (!match) return null;
    const start = parseCellAddress(match[3]);
    const end = parseCellAddress(match[4] || match[3]);
    if (start.row === 0 || end.row === 0) return null;
    return {
        sheetName: match[1] !== undefined ? match[1].replace(/''/g, "'") : (match[2] ?? null),
        startRow: Math.min(start.row, end.row),
        startCol: Math.min(start.col, end.col),
        endRow: Math.max(start.row, end.row),
        endCol: Math.max(start.col, end.col),
    };
}

/** Converts cell addresses between A1 text and 1-based row-col numbers. */
export {
    colToLetters,
    lettersToCol,
    parseCellAddress,
    normalizeFormulaAddress,
    formatRange,
    parseRangeRef,
    parseRangeAddress,
    parseSheetRangeRef,
};
