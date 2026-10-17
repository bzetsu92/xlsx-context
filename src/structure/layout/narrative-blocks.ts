/** Narrative text outside tables: row segments, items, and blocks. */
import type { NarrativeBlock, NarrativeItem, ParsedCell, ParsedSheet, TableCandidate } from "../../model/index.js";
import { isIsoDateToken, isWeekdayToken } from "../formatting/date-tokens.js";

type NarrativeRowSegment = {
    startCol: number;
    values: string[];
};

/** Turns one narrative cell into text. Markdown escapes it. The AI IR passes plain text. */
type NarrativeCellFormatter = (cell: ParsedCell) => string;

/** Splits a row into segments on gaps wider than 4 columns. */
function splitNarrativeRowSegments(cells: ParsedCell[], formatCell: NarrativeCellFormatter): NarrativeRowSegment[] {
    const segments: Array<NarrativeRowSegment & { lastCol: number }> = [];
    let current: { startCol: number; values: string[]; lastCol: number } | null = null;
    for (const cell of cells) {
        const value = formatCell(cell).trim();
        if (!value) continue;
        if (!current || cell.col - current.lastCol > 4) {
            current = {
                startCol: cell.col,
                values: [value],
                lastCol: cell.col,
            };
            segments.push(current);
        } else {
            current.values.push(value);
            current.lastCol = cell.col;
        }
    }
    return segments.map((segment) => ({
        startCol: segment.startCol,
        values: segment.values,
    }));
}

function collectNarrativeCellsByRow(sheet: ParsedSheet, tables: TableCandidate[]): Map<number, ParsedCell[]> {
    const rowMap = new Map<number, ParsedCell[]>();
    for (const cell of sheet.cells) {
        if (!cell.outputValue) continue;
        if (isCellInAnyTable(cell.row, cell.col, tables)) continue;
        const entries = rowMap.get(cell.row) || [];
        entries.push(cell);
        rowMap.set(cell.row, entries);
    }
    return rowMap;
}

function createNarrativeItem(rowNumber: number, segment: NarrativeRowSegment): NarrativeItem {
    const text = segment.values.join(" ").trim();
    return {
        row: rowNumber,
        startCol: segment.startCol,
        text,
        cellValues: segment.values,
    };
}

function createNarrativeItemsForRow(
    rowNumber: number,
    cells: ParsedCell[],
    formatCell: NarrativeCellFormatter
): NarrativeItem[] {
    return splitNarrativeRowSegments(cells, formatCell)
        .map((segment) => createNarrativeItem(rowNumber, segment))
        .filter((item) => !!item.text);
}

function buildNarrativeItems(
    sheet: ParsedSheet,
    tables: TableCandidate[],
    formatCell: NarrativeCellFormatter
): NarrativeItem[] {
    const rowMap = collectNarrativeCellsByRow(sheet, tables);
    const rowNumbers = Array.from(rowMap.keys()).sort((a, b) => a - b);
    const items: NarrativeItem[] = [];
    for (const rowNumber of rowNumbers) {
        const cells = (rowMap.get(rowNumber) || []).slice().sort((a, b) => a.col - b.col);
        items.push(...createNarrativeItemsForRow(rowNumber, cells, formatCell));
    }
    return items;
}

/** Starts a new block after a row gap or a column jump wider than 3. */
function shouldStartNarrativeBlock(
    current: NarrativeBlock | null,
    rowNumber: number,
    previousRow: number,
    startCol: number
): boolean {
    return !current || rowNumber - previousRow > 1 || Math.abs(startCol - current.startCol) > 3;
}

function isCalendarDateItem(item: NarrativeItem | null | undefined): boolean {
    if (!item) return false;
    const values = (item.cellValues || []).map((value) => String(value || "").trim()).filter(Boolean);
    return values.length >= 5 && values.every((value) => isIsoDateToken(value) || isWeekdayToken(value));
}

function blockHasCalendarDateItem(block: NarrativeBlock | null | undefined): boolean {
    return !!block?.items?.some((item) => isCalendarDateItem(item));
}

/** Keeps calendar grids together within 4 rows and 24 columns. Date rows may drift wider. */
function shouldAppendToCalendarNarrativeBlock(
    current: NarrativeBlock | null,
    item: NarrativeItem,
    previousRow: number
): boolean {
    if (!current || !blockHasCalendarDateItem(current)) {
        return false;
    }
    const rowGap = item.row - previousRow;
    if (rowGap < 0 || rowGap > 4) {
        return false;
    }
    const startColDelta = Math.abs(item.startCol - current.startCol);
    return startColDelta <= 24 || isCalendarDateItem(item);
}

function createNarrativeBlockFromItem(item: NarrativeItem): NarrativeBlock {
    return {
        startRow: item.row,
        startCol: item.startCol,
        endRow: item.row,
        lines: [item.text],
        items: [item],
    };
}

function appendNarrativeItem(block: NarrativeBlock, item: NarrativeItem): void {
    block.lines.push(item.text);
    block.endRow = item.row;
    block.items.push(item);
}

/** A line right above a table, within its columns, is the table's caption and stands on its own. */
function isTableCaptionItem(item: NarrativeItem, tables: TableCandidate[]): boolean {
    return tables.some(
        (table) => item.row === table.startRow - 1 && item.startCol >= table.startCol && item.startCol <= table.endCol
    );
}

/**
 * Groups the text outside tables into narrative blocks.
 * formatCell turns a cell into text. Markdown passes escaped Markdown. The AI IR passes plain text.
 */
function extractNarrativeBlocks(
    sheet: ParsedSheet,
    tables: TableCandidate[],
    formatCell: NarrativeCellFormatter
): NarrativeBlock[] {
    const items = buildNarrativeItems(sheet, tables, formatCell);
    const blocks: NarrativeBlock[] = [];
    let current: NarrativeBlock | null = null;
    let previousRow = -100;

    for (const item of items) {
        if (current && shouldAppendToCalendarNarrativeBlock(current, item, previousRow)) {
            appendNarrativeItem(current, item);
        } else if (
            !current ||
            shouldStartNarrativeBlock(current, item.row, previousRow, item.startCol) ||
            isTableCaptionItem(item, tables)
        ) {
            current = createNarrativeBlockFromItem(item);
            blocks.push(current);
        } else {
            appendNarrativeItem(current, item);
        }
        previousRow = item.row;
    }

    return blocks;
}

function isCellInAnyTable(row: number, col: number, tables: TableCandidate[]): boolean {
    return tables.some(
        (table) => row >= table.startRow && row <= table.endRow && col >= table.startCol && col <= table.endCol
    );
}

export {
    blockHasCalendarDateItem,
    buildNarrativeItems,
    collectNarrativeCellsByRow,
    extractNarrativeBlocks,
    isCellInAnyTable,
    splitNarrativeRowSegments,
};
export type { NarrativeCellFormatter };
