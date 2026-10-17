import type { TableCandidate } from "../../model/index.js";
import { parseCellAddress } from "../addresses/address-utils.js";
import {
    collectNormalizedBorderPositions,
    collectTableEdgeStats,
    countNormalizedBorderedCells,
} from "./border-grid.js";
import {
    boundsOverlap,
    collectBorderSeedCells,
    collectConnectedComponents,
    collectTableSeedCells,
    getBoundsArea,
    getCombinedCandidateArea,
    isWithinBounds,
} from "./table-seeds.js";
import type { CellLike, SheetLike } from "./table-seeds.js";

type TableDetectionMode = "balanced" | "border" | "planner-aware";

type TableScoreWeights = {
    minGrid: number;
    borderPresence: number;
    densityHigh: number;
    densityVeryHigh: number;
    headerish: number;
    mergeHeavyPenalty: number;
    prosePenalty: number;
    threshold: number;
};

const DEFAULT_TABLE_SCORE_WEIGHTS: TableScoreWeights = {
    minGrid: 2,
    borderPresence: 3,
    densityHigh: 2,
    densityVeryHigh: 1,
    headerish: 2,
    mergeHeavyPenalty: -1,
    prosePenalty: -2,
    threshold: 4,
};

/** Drops candidates that smaller contained candidates already cover. */
function pruneRedundantCandidates(candidates: TableCandidate[]): TableCandidate[] {
    return candidates.filter((candidate, candidateIndex) => {
        const candidateArea = getBoundsArea(candidate);
        const hasSingleDominatingContainedCandidate = candidates.some((other, otherIndex) => {
            if (candidateIndex === otherIndex) return false;
            if (!isWithinBounds(candidate, other)) return false;
            const otherArea = getBoundsArea(other);
            if (otherArea < candidateArea * 0.4) return false;
            return candidateArea > otherArea;
        });
        if (hasSingleDominatingContainedCandidate) {
            return false;
        }
        const containedCandidates = candidates.filter((other, otherIndex) => {
            if (candidateIndex === otherIndex) return false;
            if (!isWithinBounds(candidate, other)) return false;
            return getBoundsArea(other) < candidateArea;
        });
        if (containedCandidates.length >= 2 && getCombinedCandidateArea(containedCandidates) >= candidateArea * 0.6) {
            return false;
        }
        return true;
    });
}

/** Drops clusters of narrow column candidates that read as calendar day columns. */
function pruneCalendarLikeColumnCandidates(candidates: TableCandidate[]): TableCandidate[] {
    const dropKeys = new Set<string>();
    const sorted = [...candidates].sort((left, right) => {
        if (left.startRow !== right.startRow) return left.startRow - right.startRow;
        if (left.endRow !== right.endRow) return left.endRow - right.endRow;
        return left.startCol - right.startCol;
    });
    let cluster: TableCandidate[] = [];

    function flushCluster(): void {
        if (cluster.length < 3) {
            cluster = [];
            return;
        }
        for (const candidate of cluster) {
            dropKeys.add(`${candidate.startRow}:${candidate.startCol}:${candidate.endRow}:${candidate.endCol}`);
        }
        cluster = [];
    }

    function isCalendarLikeColumn(candidate: TableCandidate): boolean {
        const rowCount = candidate.endRow - candidate.startRow + 1;
        const colCount = candidate.endCol - candidate.startCol + 1;
        return rowCount >= 4 && colCount <= 3;
    }

    for (const candidate of sorted) {
        if (!isCalendarLikeColumn(candidate)) {
            flushCluster();
            continue;
        }
        const previous = cluster[cluster.length - 1];
        if (!previous) {
            cluster.push(candidate);
            continue;
        }
        const sameBand = candidate.startRow === previous.startRow && candidate.endRow === previous.endRow;
        const horizontalGap = candidate.startCol - previous.endCol;
        if (sameBand && horizontalGap >= 1 && horizontalGap <= 2) {
            cluster.push(candidate);
        } else {
            flushCluster();
            cluster.push(candidate);
        }
    }
    flushCluster();

    return candidates.filter(
        (candidate) =>
            !dropKeys.has(`${candidate.startRow}:${candidate.startCol}:${candidate.endRow}:${candidate.endCol}`)
    );
}

const NUMERIC_VALUE_TYPES = new Set(["", "n", "d", "b", "e", "formula"]);

function isBoldCell(cell: CellLike): boolean {
    if (cell.richTextRuns && cell.richTextRuns.length > 0) return cell.richTextRuns.every((run) => run.bold);
    return !!cell.textStyle?.bold;
}

function collectRowValueCells<TCell extends CellLike>(
    cellMap: Map<string, TCell>,
    row: number,
    startCol: number,
    endCol: number
): TCell[] {
    const cells: TCell[] = [];
    for (let col = startCol; col <= endCol; col += 1) {
        const cell = cellMap.get(`${row}:${col}`);
        if (cell && String(cell.outputValue || "").trim()) cells.push(cell);
    }
    return cells;
}

/**
 * A first row of numbers or dates reads as data (`base1 | 10`).
 * Bold styling apart from the next row keeps it a header (`項目 | 2024 | 2025`).
 */
function detectHeaderRow<TCell extends CellLike>(
    cellMap: Map<string, TCell>,
    bounds: { startRow: number; startCol: number; endRow: number; endCol: number }
): boolean {
    const valueRows: TCell[][] = [];
    for (let row = bounds.startRow; row <= bounds.endRow && valueRows.length < 2; row += 1) {
        const cells = collectRowValueCells(cellMap, row, bounds.startCol, bounds.endCol);
        if (cells.length > 0) valueRows.push(cells);
    }
    const [first, second = []] = valueRows;
    if (!first) return true;
    const hasDataValue = first.some((cell) => cell.valueType !== undefined && NUMERIC_VALUE_TYPES.has(cell.valueType));
    if (!hasDataValue) return true;
    return first.every(isBoldCell) && !second.every(isBoldCell);
}

/**
 * Joins fallback components across one empty spacer column.
 * Joined parts must cover exactly the same rows.
 * The same-rows rule keeps titles and formula clusters out.
 */
function mergeSpacerSeparatedComponents<TCell extends CellLike>(components: TCell[][]): TCell[][] {
    type Entry = { cells: TCell[]; startRow: number; endRow: number; startCol: number; endCol: number };
    const entries: Entry[] = components.map((cells) => {
        const entry: Entry = { cells, startRow: Infinity, endRow: -Infinity, startCol: Infinity, endCol: -Infinity };
        for (const cell of cells) {
            if (cell.row < entry.startRow) entry.startRow = cell.row;
            if (cell.row > entry.endRow) entry.endRow = cell.row;
            if (cell.col < entry.startCol) entry.startCol = cell.col;
            if (cell.col > entry.endCol) entry.endCol = cell.col;
        }
        return entry;
    });
    entries.sort((left, right) => left.startCol - right.startCol);
    const merged: Entry[] = [];
    for (const entry of entries) {
        const target = merged.find(
            (other) =>
                other.endRow - other.startRow >= 1 &&
                other.startRow === entry.startRow &&
                other.endRow === entry.endRow &&
                other.endCol + 2 === entry.startCol
        );
        if (target) {
            target.cells = target.cells.concat(entry.cells);
            target.endCol = entry.endCol;
        } else {
            merged.push({ ...entry });
        }
    }
    return merged.map((entry) => entry.cells);
}

/** Excel tables that the file declares. Their ranges are exact, so they need no score. */
function collectDeclaredTables(sheet: SheetLike, threshold: number): TableCandidate[] {
    return (sheet.tables || []).flatMap((table) => {
        const start = parseCellAddress(table.start);
        const end = parseCellAddress(table.end);
        if (!start.row || !start.col || end.row < start.row || end.col < start.col) return [];
        return [
            {
                startRow: start.row,
                startCol: start.col,
                endRow: end.row,
                endCol: end.col,
                score: threshold,
                reasonSummary: [`Excel table ${table.displayName}`],
                hasHeaderRow: table.headerRowCount > 0,
                name: table.displayName,
            },
        ];
    });
}

/**
 * Scores connected cell groups as table candidates.
 * Declared Excel tables win over every detected candidate that overlaps them.
 * Prefers borders, density, and headers. Prunes prose and calendar grids by mode.
 *
 * @param tableDetectionMode `balanced`, `border`, or `planner-aware`.
 */
function detectTableCandidates<TCell extends CellLike, TSheet extends SheetLike<TCell> = SheetLike<TCell>>(
    sheet: TSheet,
    buildCellMap: (sheet: TSheet) => Map<string, TCell>,
    scoreWeights: TableScoreWeights = DEFAULT_TABLE_SCORE_WEIGHTS,
    tableDetectionMode: TableDetectionMode = "balanced"
): TableCandidate[] {
    const cellMap = buildCellMap(sheet);
    const allSeedCells = collectTableSeedCells(sheet);
    const borderSeedCells = collectBorderSeedCells(sheet);
    const candidates: TableCandidate[] = [];
    const candidateKeys = new Set<string>();
    const fallbackCandidates = new Set<TableCandidate>();

    function countSparseRows(component: TCell[], startRow: number, endRow: number): number {
        const nonEmptyByRow = new Map<number, number>();
        for (const entry of component) {
            if (entry.outputValue.trim()) {
                nonEmptyByRow.set(entry.row, (nonEmptyByRow.get(entry.row) || 0) + 1);
            }
        }
        let sparseRows = 0;
        for (let row = startRow; row <= endRow; row += 1) {
            if ((nonEmptyByRow.get(row) || 0) <= 2) {
                sparseRows += 1;
            }
        }
        return sparseRows;
    }

    const borderPositions = collectNormalizedBorderPositions(cellMap);

    function countBorderedInBounds(
        startRow: number,
        startCol: number,
        endRow: number,
        endCol: number,
        area: number
    ): number {
        if (borderPositions.length > area) {
            return countNormalizedBorderedCells(cellMap, startRow, startCol, endRow, endCol);
        }
        let count = 0;
        for (const position of borderPositions) {
            if (
                position.row >= startRow &&
                position.row <= endRow &&
                position.col >= startCol &&
                position.col <= endCol
            ) {
                count += 1;
            }
        }
        return count;
    }

    function maybePushCandidate(component: TCell[], sourceKind: "border" | "fallback" = "border"): void {
        const rows = component.map((entry) => entry.row);
        const cols = component.map((entry) => entry.col);
        const startRow = rows.reduce((a, b) => (a < b ? a : b), Infinity);
        const endRow = rows.reduce((a, b) => (a > b ? a : b), -Infinity);
        const startCol = cols.reduce((a, b) => (a < b ? a : b), Infinity);
        const endCol = cols.reduce((a, b) => (a > b ? a : b), -Infinity);
        const area = Math.max(1, (endRow - startRow + 1) * (endCol - startCol + 1));
        const density = component.filter((entry) => entry.outputValue.trim()).length / area;
        const rowCount = endRow - startRow + 1;
        const colCount = endCol - startCol + 1;
        const sparseRowCount = countSparseRows(component, startRow, endRow);
        const nonEmptyCells = component.filter((entry) => entry.outputValue.trim());
        if (rowCount < 2 || colCount < 2) {
            return;
        }

        let score = 0;
        const reasons: string[] = [];
        const normalizedBorderedCellCount = countBorderedInBounds(startRow, startCol, endRow, endCol, area);
        if (rowCount >= 2 && colCount >= 2) {
            score += scoreWeights.minGrid;
            reasons.push(`At least 2x2 (+${scoreWeights.minGrid})`);
        }
        if (normalizedBorderedCellCount >= Math.max(2, Math.ceil(component.length * 0.3))) {
            score += scoreWeights.borderPresence;
            reasons.push(`Has borders (+${scoreWeights.borderPresence})`);
        }
        if (density >= 0.55) {
            score += scoreWeights.densityHigh;
            reasons.push(`High density (+${scoreWeights.densityHigh})`);
        }
        if (density >= 0.8) {
            score += scoreWeights.densityVeryHigh;
            reasons.push(`Very high density (+${scoreWeights.densityVeryHigh})`);
        }

        const firstRowCells = component.filter((entry) => entry.row === startRow).sort((a, b) => a.col - b.col);
        const headerishCount = firstRowCells.filter((entry) => {
            const value = entry.outputValue.trim();
            return value.length > 0 && value.length <= 24 && !/^\d+(?:\.\d+)?$/.test(value);
        }).length;
        if (headerishCount >= 2) {
            score += scoreWeights.headerish;
            reasons.push(`Header-like first row (+${scoreWeights.headerish})`);
        }

        const mergedArea = sheet.merges.filter((merge) => {
            return !(
                merge.endRow < startRow ||
                merge.startRow > endRow ||
                merge.endCol < startCol ||
                merge.startCol > endCol
            );
        }).length;
        if (mergedArea >= Math.max(2, Math.ceil(area * 0.08))) {
            score += scoreWeights.mergeHeavyPenalty;
            reasons.push(`Many merged cells (${scoreWeights.mergeHeavyPenalty})`);
        }

        const plannerAwareMode = tableDetectionMode === "planner-aware";

        if (sourceKind === "border") {
            const looksLikeTinyMergedLabelStub =
                rowCount <= 2 && colCount <= 2 && mergedArea >= 1 && nonEmptyCells.length <= 2 && headerishCount <= 1;
            if (looksLikeTinyMergedLabelStub) {
                return;
            }
            if (mergedArea >= 2 && density < 0.25 && headerishCount < 2) {
                return;
            }
            if (plannerAwareMode) {
                const looksLikeWideSparseMergeForm =
                    colCount >= 8 &&
                    rowCount >= 4 &&
                    density < 0.45 &&
                    mergedArea >= Math.max(4, rowCount - 1) &&
                    sparseRowCount >= Math.ceil(rowCount * 0.7);
                if (looksLikeWideSparseMergeForm) {
                    return;
                }
            }
        } else {
            if (mergedArea >= 2 && rowCount <= 6 && colCount >= 10 && density < 0.25) {
                return;
            }
            if (plannerAwareMode) {
                const looksLikeMixedLayoutSheet =
                    rowCount >= 20 && colCount >= 8 && mergedArea >= 4 && sparseRowCount >= 4 && density < 0.8;
                if (looksLikeMixedLayoutSheet) {
                    return;
                }
            }
        }

        const avgTextLength =
            nonEmptyCells.reduce((sum, entry) => sum + entry.outputValue.trim().length, 0) /
            Math.max(1, nonEmptyCells.length);
        if (avgTextLength > 36 && density < 0.7) {
            score += scoreWeights.prosePenalty;
            reasons.push(`Mostly long prose (${scoreWeights.prosePenalty})`);
        }

        if (score >= scoreWeights.threshold) {
            const normalizedBounds = trimTableCandidateBounds(
                cellMap,
                {
                    startRow,
                    startCol,
                    endRow,
                    endCol,
                },
                borderPositions.length > 0
            );
            if (
                normalizedBounds.endRow - normalizedBounds.startRow + 1 < 2 ||
                normalizedBounds.endCol - normalizedBounds.startCol + 1 < 2
            ) {
                return;
            }
            const key = `${normalizedBounds.startRow}:${normalizedBounds.startCol}:${normalizedBounds.endRow}:${normalizedBounds.endCol}`;
            if (candidateKeys.has(key)) {
                return;
            }
            candidateKeys.add(key);
            const candidate: TableCandidate = {
                startRow: normalizedBounds.startRow,
                startCol: normalizedBounds.startCol,
                endRow: normalizedBounds.endRow,
                endCol: normalizedBounds.endCol,
                score,
                reasonSummary: reasons,
                hasHeaderRow: detectHeaderRow(cellMap, normalizedBounds),
            };
            candidates.push(candidate);
            if (sourceKind === "fallback") fallbackCandidates.add(candidate);
        }
    }

    for (const component of collectConnectedComponents(
        borderSeedCells,
        tableDetectionMode === "border" ? "border" : "grid"
    )) {
        maybePushCandidate(component, "border");
    }

    if (tableDetectionMode !== "border") {
        for (const component of mergeSpacerSeparatedComponents(collectConnectedComponents(allSeedCells))) {
            const rows = component.map((entry) => entry.row);
            const cols = component.map((entry) => entry.col);
            const bounds = {
                startRow: rows.reduce((a, b) => (a < b ? a : b), Infinity),
                startCol: cols.reduce((a, b) => (a < b ? a : b), Infinity),
                endRow: rows.reduce((a, b) => (a > b ? a : b), -Infinity),
                endCol: cols.reduce((a, b) => (a > b ? a : b), -Infinity),
            };
            const containingBorderCandidates = candidates.filter((candidate) => isWithinBounds(candidate, bounds));
            const fallbackArea = getBoundsArea(bounds);
            const shadowedByBorderCandidate = containingBorderCandidates.some(
                (candidate) => getBoundsArea(candidate) >= fallbackArea * 0.4
            );
            const shadowedByMultipleBorderCandidates =
                containingBorderCandidates.length >= 2 &&
                getCombinedCandidateArea(containingBorderCandidates) >= fallbackArea * 0.6;
            if (shadowedByBorderCandidate || shadowedByMultipleBorderCandidates) {
                continue;
            }
            maybePushCandidate(component, "fallback");
        }
        clipFallbacksOverlappingBorderCandidates();
    }

    /**
     * A fallback region can partly cover a bordered table without being shadowed by it,
     * which would render the shared rows twice. Drop the overlapping row band from the
     * fallback and re-score the rows above and below it on their own.
     */
    function clipFallbacksOverlappingBorderCandidates(): void {
        const borderCandidates = candidates.filter((candidate) => !fallbackCandidates.has(candidate));
        for (const fallback of [...fallbackCandidates]) {
            const overlapping = borderCandidates.filter((other) => boundsOverlap(fallback, other));
            if (overlapping.length === 0) continue;
            candidates.splice(candidates.indexOf(fallback), 1);
            fallbackCandidates.delete(fallback);
            const bandStart = Math.min(...overlapping.map((other) => other.startRow));
            const bandEnd = Math.max(...overlapping.map((other) => other.endRow));
            const inColumns = (cell: TCell) => cell.col >= fallback.startCol && cell.col <= fallback.endCol;
            const above = allSeedCells.filter(
                (cell) => inColumns(cell) && cell.row >= fallback.startRow && cell.row < bandStart
            );
            const below = allSeedCells.filter(
                (cell) => inColumns(cell) && cell.row > bandEnd && cell.row <= fallback.endRow
            );
            if (above.length > 0) maybePushCandidate(above, "fallback");
            if (below.length > 0) maybePushCandidate(below, "fallback");
        }
    }

    const prunedCandidates = pruneRedundantCandidates(candidates);
    const detectedCandidates =
        tableDetectionMode === "planner-aware" ? pruneCalendarLikeColumnCandidates(prunedCandidates) : prunedCandidates;
    const declared = collectDeclaredTables(sheet, scoreWeights.threshold);
    const finalCandidates = [
        ...declared,
        ...detectedCandidates.filter((candidate) => !declared.some((table) => boundsOverlap(candidate, table))),
    ];

    return finalCandidates.sort((left, right) => {
        if (left.startRow !== right.startRow) return left.startRow - right.startRow;
        return left.startCol - right.startCol;
    });
}

/** Trims sparse outer rows around bordered bands. */
function trimTableCandidateBounds<TCell extends CellLike>(
    cellMap: Map<string, TCell>,
    bounds: { startRow: number; startCol: number; endRow: number; endCol: number },
    hasNormalizedBorders = true
): { startRow: number; startCol: number; endRow: number; endCol: number } {
    let { startRow, startCol, endRow, endCol } = bounds;
    const minBorderedCells = Math.max(2, Math.ceil((endCol - startCol + 1) * 0.5));

    while (hasNormalizedBorders && endRow - startRow + 1 >= 2) {
        const topStats = collectTableEdgeStats(cellMap, startRow, startCol, endCol);
        const nextStats = collectTableEdgeStats(cellMap, startRow + 1, startCol, endCol);
        const shouldTrimTop =
            topStats.nonEmptyCount <= 2 &&
            topStats.rawBorderCount === 0 &&
            nextStats.borderCount >= minBorderedCells &&
            nextStats.nonEmptyCount >= Math.max(2, Math.ceil((endCol - startCol + 1) * 0.5));
        if (!shouldTrimTop) {
            break;
        }
        startRow += 1;
    }

    let previousStats = hasNormalizedBorders ? collectTableEdgeStats(cellMap, startRow, startCol, endCol) : null;
    for (let row = startRow + 1; previousStats && row <= endRow; row += 1) {
        const currentStats = collectTableEdgeStats(cellMap, row, startCol, endCol);
        const shouldBreakAtCurrentRow =
            (previousStats.borderCount >= minBorderedCells ||
                previousStats.bottomCount >= minBorderedCells ||
                currentStats.topCount >= minBorderedCells) &&
            currentStats.rawBorderCount === 0 &&
            currentStats.nonEmptyCount <= 1;
        if (shouldBreakAtCurrentRow) {
            endRow = row - 1;
            break;
        }
        previousStats = currentStats;
    }

    while (endRow - startRow + 1 >= 2) {
        const bottomStats = collectTableEdgeStats(cellMap, endRow, startCol, endCol);
        const previousStats = collectTableEdgeStats(cellMap, endRow - 1, startCol, endCol);
        const shouldTrimBottom =
            ((previousStats.borderCount >= minBorderedCells ||
                previousStats.bottomCount >= minBorderedCells ||
                bottomStats.topCount >= minBorderedCells) &&
                bottomStats.rawBorderCount === 0 &&
                bottomStats.nonEmptyCount <= 1) ||
            (bottomStats.nonEmptyCount <= 1 &&
                bottomStats.rawBorderCount === 0 &&
                bottomStats.maxTextLength >= 12 &&
                previousStats.nonEmptyCount >= Math.max(2, Math.ceil((endCol - startCol + 1) * 0.5)));
        if (!shouldTrimBottom) {
            break;
        }
        endRow -= 1;
    }

    return { startRow, startCol, endRow, endCol };
}

/** Detects table regions from bordered and valued cells. */
export {
    detectHeaderRow,
    detectTableCandidates,
    pruneCalendarLikeColumnCandidates,
    pruneRedundantCandidates,
    trimTableCandidateBounds,
    DEFAULT_TABLE_SCORE_WEIGHTS as defaultTableScoreWeights,
};

// Re-exports for the previous deep path. New code imports table-seeds and table-layout directly.
export { collectBorderSeedCells, collectTableSeedCells } from "./table-seeds.js";
export { applyMergeTokens, isMeaningfulMarkdownCell, matrixFromCandidate, planTableLayout } from "./table-layout.js";
