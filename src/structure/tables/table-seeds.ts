import type { BorderFlags, MergeRange } from "../../model/index.js";

type CellLike = {
    row: number;
    col: number;
    outputValue: string;
    borders: BorderFlags;
    valueType?: string;
    textStyle?: { bold: boolean };
    richTextRuns?: Array<{ bold: boolean }> | null;
};

type SheetLike<TCell extends CellLike = CellLike> = {
    cells: TCell[];
    merges: MergeRange[];
    /** Excel tables that the file declares. */
    tables?: Array<{ start: string; end: string; displayName: string; headerRowCount: number }>;
};
function collectTableSeedCells<TCell extends CellLike>(sheet: SheetLike<TCell>): TCell[] {
    return sheet.cells.filter((cell) => {
        const hasValue = !!String(cell.outputValue || "").trim();
        const hasBorder = cell.borders.top || cell.borders.bottom || cell.borders.left || cell.borders.right;
        return hasValue || hasBorder;
    });
}

function collectBorderSeedCells<TCell extends CellLike>(sheet: SheetLike<TCell>): TCell[] {
    return sheet.cells.filter(
        (cell) => cell.borders.top || cell.borders.bottom || cell.borders.left || cell.borders.right
    );
}

function areBorderAdjacent<TCell extends CellLike>(current: TCell, next: TCell): boolean {
    if (current.row === next.row && Math.abs(current.col - next.col) === 1) {
        return (
            (current.borders.top && next.borders.top) ||
            (current.borders.bottom && next.borders.bottom) ||
            (current.col < next.col
                ? current.borders.right && next.borders.left
                : current.borders.left && next.borders.right)
        );
    }
    if (current.col === next.col && Math.abs(current.row - next.row) === 1) {
        return (
            (current.borders.left && next.borders.left) ||
            (current.borders.right && next.borders.right) ||
            (current.row < next.row
                ? current.borders.bottom && next.borders.top
                : current.borders.top && next.borders.bottom)
        );
    }
    return false;
}

const NEIGHBOR_DELTAS: ReadonlyArray<readonly [number, number]> = [
    [1, 0],
    [-1, 0],
    [0, 1],
    [0, -1],
];
// Keys positions as one number. Avoids one string allocation per lookup. Columns stay far below the stride.
const POSITION_ROW_STRIDE = 1 << 21;

/** Groups seed cells into connected components. Border mode links border-sharing cells only. */
function collectConnectedComponents<TCell extends CellLike>(
    seedCells: TCell[],
    adjacencyMode: "grid" | "border" = "grid"
): TCell[][] {
    if (seedCells.length === 0) return [];
    let minRow = Infinity;
    let maxRow = -Infinity;
    let minCol = Infinity;
    let maxCol = -Infinity;
    for (const cell of seedCells) {
        if (cell.row < minRow) minRow = cell.row;
        if (cell.row > maxRow) maxRow = cell.row;
        if (cell.col < minCol) minCol = cell.col;
        if (cell.col > maxCol) maxCol = cell.col;
    }
    const width = maxCol - minCol + 1;
    const height = maxRow - minRow + 1;
    const area = width * height;
    // Scans a dense grid over the bounding box. Falls back to hashing when cells spread thin.
    if (area > Math.max(1 << 16, seedCells.length * 8)) {
        return collectConnectedComponentsSparse(seedCells, adjacencyMode);
    }
    const slotOf = new Int32Array(area);
    for (let index = 0; index < seedCells.length; index += 1) {
        const cell = seedCells[index] as TCell;
        slotOf[(cell.row - minRow) * width + (cell.col - minCol)] = index + 1;
    }
    const visited = new Uint8Array(area);
    const components: TCell[][] = [];

    for (const cell of seedCells) {
        const slot = (cell.row - minRow) * width + (cell.col - minCol);
        if (visited[slot]) continue;
        const queue = [cell];
        const component: TCell[] = [];
        visited[slot] = 1;
        let queueHead = 0;
        while (queueHead < queue.length) {
            const current = queue[queueHead] as TCell;
            queueHead += 1;
            component.push(current);
            const rowOffset = current.row - minRow;
            const colOffset = current.col - minCol;
            for (const [rowDelta, colDelta] of NEIGHBOR_DELTAS) {
                const nextRow = rowOffset + rowDelta;
                const nextCol = colOffset + colDelta;
                if (nextRow < 0 || nextRow >= height || nextCol < 0 || nextCol >= width) continue;
                const nextSlot = nextRow * width + nextCol;
                const seedIndex = slotOf[nextSlot];
                if (seedIndex === 0 || visited[nextSlot]) continue;
                const nextCell = seedCells[seedIndex - 1] as TCell;
                if (adjacencyMode === "border" && !areBorderAdjacent(current, nextCell)) continue;
                visited[nextSlot] = 1;
                queue.push(nextCell);
            }
        }
        components.push(component);
    }

    return components;
}

function collectConnectedComponentsSparse<TCell extends CellLike>(
    seedCells: TCell[],
    adjacencyMode: "grid" | "border" = "grid"
): TCell[][] {
    const positionMap = new Map<number, TCell>();
    for (const cell of seedCells) {
        positionMap.set(cell.row * POSITION_ROW_STRIDE + cell.col, cell);
    }
    const visited = new Set<number>();
    const components: TCell[][] = [];

    for (const cell of seedCells) {
        const key = cell.row * POSITION_ROW_STRIDE + cell.col;
        if (visited.has(key)) continue;
        const queue = [cell];
        const component: TCell[] = [];
        visited.add(key);
        let queueHead = 0;
        while (queueHead < queue.length) {
            const current = queue[queueHead] as TCell;
            queueHead += 1;
            component.push(current);
            for (const [rowDelta, colDelta] of NEIGHBOR_DELTAS) {
                const nextKey = (current.row + rowDelta) * POSITION_ROW_STRIDE + current.col + colDelta;
                const nextCell = positionMap.get(nextKey);
                if (!nextCell || visited.has(nextKey)) continue;
                if (adjacencyMode === "border" && !areBorderAdjacent(current, nextCell)) continue;
                visited.add(nextKey);
                queue.push(nextCell);
            }
        }
        components.push(component);
    }

    return components;
}

function isWithinBounds(
    bounds: { startRow: number; startCol: number; endRow: number; endCol: number },
    candidate: { startRow: number; startCol: number; endRow: number; endCol: number }
): boolean {
    return (
        candidate.startRow >= bounds.startRow &&
        candidate.startCol >= bounds.startCol &&
        candidate.endRow <= bounds.endRow &&
        candidate.endCol <= bounds.endCol
    );
}

function boundsOverlap(
    left: { startRow: number; startCol: number; endRow: number; endCol: number },
    right: { startRow: number; startCol: number; endRow: number; endCol: number }
): boolean {
    return !(
        left.endRow < right.startRow ||
        left.startRow > right.endRow ||
        left.endCol < right.startCol ||
        left.startCol > right.endCol
    );
}

function getBoundsArea(bounds: { startRow: number; startCol: number; endRow: number; endCol: number }): number {
    return Math.max(1, (bounds.endRow - bounds.startRow + 1) * (bounds.endCol - bounds.startCol + 1));
}

function getCombinedCandidateArea(
    candidates: Array<{ startRow: number; startCol: number; endRow: number; endCol: number }>
): number {
    return candidates.reduce((sum, candidate) => sum + getBoundsArea(candidate), 0);
}
export {
    areBorderAdjacent,
    boundsOverlap,
    collectBorderSeedCells,
    collectConnectedComponents,
    collectConnectedComponentsSparse,
    collectTableSeedCells,
    getBoundsArea,
    getCombinedCandidateArea,
    isWithinBounds,
};
export type { CellLike, SheetLike };
