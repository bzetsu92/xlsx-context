import type { ParsedShapeAsset, SheetGrid, ShapeBlock } from "../../model/worksheet.js";
import { columnAtEmu, columnPositionAt, createUniformGrid, rowAtEmu, rowPositionAt } from "../workbook/sheet-grid.js";
import { resolveShapeConnections } from "./shape-connections.js";

type RegionGapDeps = {
    defaultCellWidthEmu: number;
    defaultCellHeightEmu: number;
    shapeBlockGapXEmu: number;
    shapeBlockGapYEmu: number;
    /** Grid of the sheet. Without it, every cell has the default size. */
    grid?: SheetGrid;
};

type CellBox = ParsedShapeAsset["bbox"];

type RegionSpace = { grid: SheetGrid; gapCols: number; gapRows: number };

/**
 * Picks the grid and the proximity gaps.
 * The configured gaps count default cells. Gaps are measured in cells, not EMU,
 * because sheets with narrow custom columns would otherwise join distant shapes.
 */
function regionSpace(deps: RegionGapDeps): RegionSpace {
    return {
        grid: deps.grid ?? createUniformGrid(deps.defaultCellWidthEmu, deps.defaultCellHeightEmu),
        gapCols: deps.shapeBlockGapXEmu / deps.defaultCellWidthEmu,
        gapRows: deps.shapeBlockGapYEmu / deps.defaultCellHeightEmu,
    };
}

/** Converts a box to fractional cell positions. */
function cellBox(box: ParsedShapeAsset["bbox"], grid: SheetGrid): ParsedShapeAsset["bbox"] {
    return {
        left: columnPositionAt(grid, box.left),
        top: rowPositionAt(grid, box.top),
        right: columnPositionAt(grid, box.right),
        bottom: rowPositionAt(grid, box.bottom),
    };
}

/** Measures the gap between two boxes. Overlapping boxes have zero gap. */
function bboxGap(a: ParsedShapeAsset["bbox"], b: ParsedShapeAsset["bbox"]): { dx: number; dy: number } {
    const dx = a.right < b.left ? b.left - a.right : b.right < a.left ? a.left - b.right : 0;
    const dy = a.bottom < b.top ? b.top - a.bottom : b.bottom < a.top ? a.top - b.bottom : 0;
    return { dx, dy };
}

/** Finds the component root with path compression. */
function findRoot(parent: number[], index: number): number {
    let root = index;
    while (parent[root] !== root) root = parent[root] as number;
    let cursor = index;
    while (parent[cursor] !== root) {
        const next = parent[cursor] as number;
        parent[cursor] = root;
        cursor = next;
    }
    return root;
}

/** Groups each connector with every shape that one of its ends reaches. Unlinked shapes stay alone here. */
function groupConnectedComponents(shapes: ParsedShapeAsset[]): Map<number, number[]> {
    const parent = shapes.map((_shape, index) => index);
    for (const entry of resolveShapeConnections(shapes)) {
        for (const target of [entry.startIndex, entry.endIndex]) {
            if (target === null) continue;
            const rootConnector = findRoot(parent, entry.connectorIndex);
            const rootTarget = findRoot(parent, target);
            if (rootConnector !== rootTarget) parent[rootConnector] = rootTarget;
        }
    }
    const groups = new Map<number, number[]>();
    shapes.forEach((_shape, index) => {
        const root = findRoot(parent, index);
        const members = groups.get(root);
        if (members) members.push(index);
        else groups.set(root, [index]);
    });
    return groups;
}

/**
 * Merges lone shapes by gap thresholds. Lone shapes reach no other shape by a connector.
 * Joined diagrams never merge by distance, because they are never lone.
 */
function mergeLoneShapesByProximity(shapes: ParsedShapeAsset[], lone: number[], space: RegionSpace): number[][] {
    const visited = new Set<number>();
    const groups: number[][] = [];
    const boxes = new Map(lone.map((index) => [index, cellBox(shapes[index].bbox, space.grid)]));
    for (const index of lone) {
        if (visited.has(index)) continue;
        const queue = [index];
        visited.add(index);
        const members: number[] = [];
        for (let head = 0; head < queue.length; head += 1) {
            const current = queue[head] as number;
            members.push(current);
            for (const other of lone) {
                if (visited.has(other)) continue;
                const { dx, dy } = bboxGap(boxes.get(current) as CellBox, boxes.get(other) as CellBox);
                if (dx <= space.gapCols && dy <= space.gapRows) {
                    visited.add(other);
                    queue.push(other);
                }
            }
        }
        groups.push(members);
    }
    return groups;
}

/** Builds one block from member indexes. Members stay sorted for stable output. */
function buildBlock(shapes: ParsedShapeAsset[], members: number[], space: RegionSpace): ShapeBlock {
    let minLeft = Number.POSITIVE_INFINITY;
    let minTop = Number.POSITIVE_INFINITY;
    let maxRight = 0;
    let maxBottom = 0;
    for (const index of members) {
        const box = shapes[index].bbox;
        minLeft = Math.min(minLeft, box.left);
        minTop = Math.min(minTop, box.top);
        maxRight = Math.max(maxRight, box.right);
        maxBottom = Math.max(maxBottom, box.bottom);
    }
    // A box that ends on a cell edge does not reach the next cell.
    return {
        startCol: columnAtEmu(space.grid, minLeft),
        startRow: rowAtEmu(space.grid, minTop),
        endCol: columnAtEmu(space.grid, Math.max(minLeft, maxRight - 1)),
        endRow: rowAtEmu(space.grid, Math.max(minTop, maxBottom - 1)),
        shapeIndexes: [...members].sort((a, b) => a - b),
    };
}

/**
 * Splits shapes into regions without losing members.
 * Connected components keep their declared topology at any distance.
 * Only unlinked non-connector shapes fall back to proximity grouping.
 */
function partitionShapeRegions(shapes: ParsedShapeAsset[], deps: RegionGapDeps): ShapeBlock[] {
    if (shapes.length === 0) return [];
    const space = regionSpace(deps);
    const components = groupConnectedComponents(shapes);
    const blocks: ShapeBlock[] = [];
    const lone: number[] = [];
    for (const members of components.values()) {
        if (members.length > 1) {
            blocks.push(buildBlock(shapes, members, space));
            continue;
        }
        // A connector that reaches no shape is decoration or points at cells. It joins the shapes near it.
        lone.push(members[0] as number);
    }
    for (const members of mergeLoneShapesByProximity(shapes, lone, space)) {
        blocks.push(buildBlock(shapes, members, space));
    }
    return blocks.sort((a, b) => a.startRow - b.startRow || a.startCol - b.startCol);
}

export { partitionShapeRegions };
export type { RegionGapDeps };
