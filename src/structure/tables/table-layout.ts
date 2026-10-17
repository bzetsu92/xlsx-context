/** Display matrices and kept row-column layouts for table candidates. */
import type { MergeRange, TableCandidate } from "../../model/index.js";
import type { CellLike, SheetLike } from "./table-seeds.js";

type TableMatrixOptions = {
    trimText?: boolean;
    removeEmptyRows?: boolean;
    removeEmptyColumns?: boolean;
};

/** Absolute row and column numbers of a candidate that an output keeps. */
type TableLayout = {
    rows: number[];
    cols: number[];
};
/** Builds the display matrix for a candidate through the caller cell formatter. */
function matrixFromCandidate<TCell extends CellLike, TSheet extends SheetLike<TCell> = SheetLike<TCell>>(
    sheet: TSheet,
    candidate: TableCandidate,
    options: TableMatrixOptions,
    buildCellMap: (sheet: TSheet) => Map<string, TCell>,
    formatCellForMarkdown: (cell: TCell | undefined, options: TableMatrixOptions) => string
): string[][] {
    const cellMap = buildCellMap(sheet);
    const rows: string[][] = [];
    for (let row = candidate.startRow; row <= candidate.endRow; row += 1) {
        const currentRow: string[] = [];
        for (let col = candidate.startCol; col <= candidate.endCol; col += 1) {
            const cell = cellMap.get(`${row}:${col}`);
            let value = formatCellForMarkdown(cell, options);
            if (options.trimText !== false) {
                value = value.trim();
            }
            currentRow.push(value);
        }
        rows.push(currentRow);
    }
    applyMergeTokens(rows, sheet.merges, candidate.startRow, candidate.startCol, candidate.endRow, candidate.endCol);

    const layout = planTableLayout(candidate, options, (row, col) =>
        isMeaningfulMarkdownCell(rows[row - candidate.startRow][col - candidate.startCol])
    );
    return layout.rows.map((row) => layout.cols.map((col) => rows[row - candidate.startRow][col - candidate.startCol]));
}

/**
 * Selects the rows and columns of a candidate to keep.
 * Markdown and the AI IR both call this, so they drop the same empty lines
 * and skip the same empty tables. Each output decides what "meaningful" means.
 * Columns are judged over kept rows only.
 */
function planTableLayout(
    candidate: TableCandidate,
    options: TableMatrixOptions,
    isMeaningful: (row: number, col: number) => boolean
): TableLayout {
    let rows: number[] = [];
    for (let row = candidate.startRow; row <= candidate.endRow; row += 1) rows.push(row);
    let cols: number[] = [];
    for (let col = candidate.startCol; col <= candidate.endCol; col += 1) cols.push(col);
    if (options.removeEmptyRows !== false) {
        rows = rows.filter((row) => cols.some((col) => isMeaningful(row, col)));
    }
    if (options.removeEmptyColumns !== false && rows.length > 0) {
        cols = cols.filter((col) => rows.some((row) => isMeaningful(row, col)));
    }
    return { rows, cols };
}

/** Treats blank text and merge continuation tokens as empty. */
function isMeaningfulMarkdownCell(value: string): boolean {
    const text = String(value || "").trim();
    if (!text) return false;
    return text !== "[←M←]" && text !== "[↑M↑]";
}

/** Stamps covered merge cells: `[←M←]` continues the row, `[↑M↑]` continues below the master. */
function applyMergeTokens(
    matrix: string[][],
    merges: MergeRange[],
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
): void {
    for (const merge of merges) {
        if (merge.endRow < startRow || merge.startRow > endRow || merge.endCol < startCol || merge.startCol > endCol) {
            continue;
        }
        for (let row = merge.startRow; row <= merge.endRow; row += 1) {
            for (let col = merge.startCol; col <= merge.endCol; col += 1) {
                if (row === merge.startRow && col === merge.startCol) continue;
                const matrixRow = row - startRow;
                const matrixCol = col - startCol;
                if (!matrix[matrixRow] || typeof matrix[matrixRow][matrixCol] === "undefined") {
                    continue;
                }
                matrix[matrixRow][matrixCol] = row === merge.startRow ? "[←M←]" : "[↑M↑]";
            }
        }
    }
}
export { applyMergeTokens, isMeaningfulMarkdownCell, matrixFromCandidate, planTableLayout };
export type { TableLayout, TableMatrixOptions };
