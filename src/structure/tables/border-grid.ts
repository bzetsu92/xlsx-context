import type { BorderFlags } from "../../model/index.js";

type CellLike = {
    borders: BorderFlags;
    outputValue?: string;
};

type EdgeStats = {
    nonEmptyCount: number;
    borderCount: number;
    rawBorderCount: number;
    topCount: number;
    bottomCount: number;
    maxTextLength: number;
};

function getCellAt(cellMap: Map<string, CellLike>, row: number, col: number): CellLike | undefined {
    return cellMap.get(`${row}:${col}`);
}

function hasNormalizedBorderOnSide(
    cellMap: Map<string, CellLike>,
    row: number,
    col: number,
    side: keyof BorderFlags
): boolean {
    const cell = getCellAt(cellMap, row, col);
    if (side === "top") {
        const above = getCellAt(cellMap, row - 1, col);
        return !!cell?.borders.top || !!above?.borders.bottom;
    }
    if (side === "bottom") {
        const below = getCellAt(cellMap, row + 1, col);
        return !!cell?.borders.bottom || !!below?.borders.top;
    }
    if (side === "left") {
        const left = getCellAt(cellMap, row, col - 1);
        return !!cell?.borders.left || !!left?.borders.right;
    }
    const right = getCellAt(cellMap, row, col + 1);
    return !!cell?.borders.right || !!right?.borders.left;
}

function hasAnyNormalizedBorder(cellMap: Map<string, CellLike>, row: number, col: number): boolean {
    const cell = getCellAt(cellMap, row, col);
    const cellBorders = cell?.borders;
    if (cellBorders && (cellBorders.top || cellBorders.bottom || cellBorders.left || cellBorders.right)) {
        return true;
    }
    return (
        !!getCellAt(cellMap, row - 1, col)?.borders.bottom ||
        !!getCellAt(cellMap, row + 1, col)?.borders.top ||
        !!getCellAt(cellMap, row, col - 1)?.borders.right ||
        !!getCellAt(cellMap, row, col + 1)?.borders.left
    );
}

function collectTableEdgeStats(
    cellMap: Map<string, CellLike>,
    row: number,
    startCol: number,
    endCol: number
): EdgeStats {
    let nonEmptyCount = 0;
    let borderCount = 0;
    let rawBorderCount = 0;
    let topCount = 0;
    let bottomCount = 0;
    let maxTextLength = 0;

    for (let col = startCol; col <= endCol; col += 1) {
        const cell = getCellAt(cellMap, row, col);
        const text = String(cell?.outputValue || "").trim();
        if (text) {
            nonEmptyCount += 1;
            maxTextLength = Math.max(maxTextLength, text.length);
        }
        const hasOwnBorder = !!(
            cell &&
            (cell.borders.top || cell.borders.bottom || cell.borders.left || cell.borders.right)
        );
        if (hasOwnBorder) {
            rawBorderCount += 1;
        }
        const topEdge = !!cell?.borders.top || !!getCellAt(cellMap, row - 1, col)?.borders.bottom;
        const bottomEdge = !!cell?.borders.bottom || !!getCellAt(cellMap, row + 1, col)?.borders.top;
        if (topEdge) {
            topCount += 1;
        }
        if (bottomEdge) {
            bottomCount += 1;
        }
        if (
            hasOwnBorder ||
            topEdge ||
            bottomEdge ||
            !!getCellAt(cellMap, row, col - 1)?.borders.right ||
            !!getCellAt(cellMap, row, col + 1)?.borders.left
        ) {
            borderCount += 1;
        }
    }

    return { nonEmptyCount, borderCount, rawBorderCount, topCount, bottomCount, maxTextLength };
}

function countNormalizedBorderedCells(
    cellMap: Map<string, CellLike>,
    startRow: number,
    startCol: number,
    endRow: number,
    endCol: number
): number {
    let count = 0;
    for (let row = startRow; row <= endRow; row += 1) {
        for (let col = startCol; col <= endCol; col += 1) {
            if (hasAnyNormalizedBorder(cellMap, row, col)) {
                count += 1;
            }
        }
    }
    return count;
}

/** Lists bordered cells plus the neighbors their borders touch. Stays small on sparse grids, so callers count it instead of scanning. */
function collectNormalizedBorderPositions(cellMap: Map<string, CellLike>): Array<{ row: number; col: number }> {
    const seen = new Set<string>();
    const positions: Array<{ row: number; col: number }> = [];
    const add = (row: number, col: number) => {
        const key = `${row}:${col}`;
        if (seen.has(key)) return;
        seen.add(key);
        positions.push({ row, col });
    };
    for (const [key, cell] of cellMap) {
        const borders = cell.borders;
        if (!(borders.top || borders.bottom || borders.left || borders.right)) continue;
        const separator = key.indexOf(":");
        const row = Number(key.slice(0, separator));
        const col = Number(key.slice(separator + 1));
        add(row, col);
        if (borders.bottom) add(row + 1, col);
        if (borders.top) add(row - 1, col);
        if (borders.right) add(row, col + 1);
        if (borders.left) add(row, col - 1);
    }
    return positions;
}

/** Reads shared cell borders from either side of an edge. Excel stores one edge on one cell only. */
export {
    getCellAt,
    hasNormalizedBorderOnSide,
    hasAnyNormalizedBorder,
    collectTableEdgeStats,
    countNormalizedBorderedCells,
    collectNormalizedBorderPositions,
};
