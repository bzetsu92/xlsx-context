/**
 * Finds what callouts, loose line ends, and marker shapes refer to on the sheet.
 * A shape and a cell meet only in space, so each match is geometric and carries the method that made it.
 */
import type { MergeRange, ParsedCell, ParsedShapeAsset, ParsedSheet, SheetGrid } from "../../model/index.js";
import { resolveShapeConnections } from "../../extraction/assets/shape-connections.js";
import {
    columnAtEmu,
    columnLeftEmu,
    createUniformGrid,
    rowAtEmu,
    rowTopEmu,
} from "../../extraction/workbook/sheet-grid.js";

type Box = { left: number; top: number; right: number; bottom: number };
type Point = { x: number; y: number };

/**
 * One block of visible text, in a cell or in a shape.
 * `cell` is the cell, merged range, or shape box. `box` adds the overflow into empty neighbor cells.
 * `shapeIndex` is set for text inside a shape. `address` is then the cell at the shape's top-left corner.
 */
type TextBlock = { address: string; text: string; cell: Box; box: Box; shapeIndex: number | null };

/** Visible text near a point. */
export type NearbyText = { address: string; text: string };

/**
 * What one shape refers to.
 * `cell` is the cell under the point or marker. `textAddress` and `text` name the nearest visible text.
 * `method` tells how the text was found, so readers can judge the match.
 */
export type ShapeTarget = {
    shapeIndex: number;
    role: "callout" | "lineStart" | "lineEnd" | "marker";
    cell: string;
    textAddress: string | null;
    text: string | null;
    /** A shape that holds the point, or the shape whose text a marker labels. */
    targetShapeIndex: number | null;
    /** A picture that holds the point or the marker. Screen mockups are often screenshots. */
    targetImageIndex: number | null;
    /** For a loose line end: the shape at the other end, when that end is attached. */
    connectedShapeIndex: number | null;
    /** "image" means the marker sits on a picture and no text labels it. */
    method: "inside" | "overflow" | "right" | "image" | "nearest" | "none";
    /** Other visible text near a point, nearest first. Helps readers when the point sits between items. */
    nearby: NearbyText[];
};

// Excel default cell: 64 px wide and 15 pt high.
const EXCEL_DEFAULT_CELL = { widthEmu: 609600, heightEmu: 190500 };
// One full-width character at 11 pt. Text width is an estimate, because fonts and alignment vary.
const FULL_WIDTH_EMU = 139700;
// Search radius around a point. About three default rows.
const POINT_RADIUS_EMU = 571500;
// A point this close to a shape edge still points at that shape. About 5 px.
const EDGE_TOLERANCE_EMU = 47625;
// Texts listed for a point: the match and two more.
const NEARBY_LIMIT = 3;
// Markers are small. Larger shapes with short text are labels or buttons, not markers.
const MARKER_MAX_SIZE_EMU = 762000;
// A style group of this many marker-like shapes must show different numbers. Repeated "1" boxes are mock values.
const MARKER_GROUP_MIN = 4;
const MARKER_DISTINCT_RATIO = 0.3;
const MARKER_TEXT = /^[(（[]?(?:[0-9０-９]{1,3}|[①-⑳㉑-㉟㊱-㊿]|[A-Za-zＡ-Ｚａ-ｚ])[)）\].．]?$/;

/** Estimates the drawn width of a text. Wide characters take one em, others half. */
function textWidthEmu(text: string): number {
    let width = 0;
    for (const char of text) {
        width += /[ᄀ-ᅟ⺀-꓏가-힣豈-﫿︰-﹏＀-｠￠-￦]/.test(char) ? FULL_WIDTH_EMU : FULL_WIDTH_EMU / 2;
    }
    return width;
}

function cellBox(grid: SheetGrid, startRow: number, startCol: number, endRow: number, endCol: number): Box {
    return {
        left: columnLeftEmu(grid, startCol),
        top: rowTopEmu(grid, startRow),
        right: columnLeftEmu(grid, endCol + 1),
        bottom: rowTopEmu(grid, endRow + 1),
    };
}

/** Numbers do not overflow. Excel shows ### instead. */
function isNumberCell(cell: ParsedCell): boolean {
    return (cell.valueType === "" || cell.valueType === "n") && Number.isFinite(Number(cell.rawValue));
}

/**
 * Finds where cell text overflows, as Excel draws it.
 * Left and general text grows right, right-aligned text grows left, and centered text grows both ways.
 * Wrapped, shrunk, and numeric text stays in its cell.
 */
function overflowSides(cell: ParsedCell): { left: number; right: number } {
    if (cell.alignment?.contained || isNumberCell(cell)) return { left: 0, right: 0 };
    const horizontal = cell.alignment?.horizontal || "general";
    if (horizontal === "right") return { left: 1, right: 0 };
    if (horizontal === "center" || horizontal === "centerContinuous") return { left: 0.5, right: 0.5 };
    if (horizontal === "left" || horizontal === "general") return { left: 0, right: 1 };
    return { left: 0, right: 0 };
}

/**
 * Visible text blocks, found by row.
 * Blocks taller than a few rows sit in `tall`, so one huge merged range never fills a million row lists.
 */
type TextIndex = { byRow: Map<number, TextBlock[]>; tall: TextBlock[] };

// Blocks over this many rows go to the tall list.
const TALL_BLOCK_ROWS = 64;

/** Finds the first index whose value is at least the target. Returns the length when none is. */
function lowerBound(values: number[], target: number): number {
    let low = 0;
    let high = values.length;
    while (low < high) {
        const mid = (low + high) >> 1;
        if ((values[mid] as number) < target) low = mid + 1;
        else high = mid;
    }
    return low;
}

/**
 * Finds text cells that a merged range hides. Only the top-left cell of a merge shows its text.
 * The search visits the text cells inside each merge, not every covered cell,
 * so a merge over a whole column costs as little as a small one.
 */
function findHiddenCells(merges: MergeRange[], filled: Map<number, number[]>, rows: number[]): Set<string> {
    const hidden = new Set<string>();
    for (const merge of merges) {
        for (let at = lowerBound(rows, merge.startRow); at < rows.length; at += 1) {
            const row = rows[at] as number;
            if (row > merge.endRow) break;
            const cols = filled.get(row) as number[];
            for (let index = lowerBound(cols, merge.startCol); index < cols.length; index += 1) {
                const col = cols[index] as number;
                if (col > merge.endCol) break;
                if (row !== merge.startRow || col !== merge.startCol) hidden.add(`${row}:${col}`);
            }
        }
    }
    return hidden;
}

/**
 * Builds the visible text blocks of a sheet.
 * A merged range shows its text in the whole range. Other text can overflow into empty neighbor cells.
 * Work grows with the number of text cells, not with the size of rows or merged ranges.
 */
function collectTextBlocks(
    sheet: ParsedSheet,
    grid: SheetGrid,
    shapeTexts: number[],
    colToLetters: (col: number) => string
): TextIndex {
    const filled = new Map<number, number[]>();
    const textCells: ParsedCell[] = [];
    for (const cell of sheet.cells) {
        const text = String(cell.outputValue ?? "").trim();
        if (!text) continue;
        textCells.push(cell);
        const cols = filled.get(cell.row);
        if (cols) cols.push(cell.col);
        else filled.set(cell.row, [cell.col]);
    }
    for (const cols of filled.values()) cols.sort((a, b) => a - b);
    const rows = [...filled.keys()].sort((a, b) => a - b);
    const mergeAt = new Map<string, MergeRange>();
    for (const merge of sheet.merges) mergeAt.set(`${merge.startRow}:${merge.startCol}`, merge);
    const hidden = findHiddenCells(sheet.merges, filled, rows);

    const index: TextIndex = { byRow: new Map(), tall: [] };
    const register = (block: TextBlock) => {
        const firstRow = rowAtEmu(grid, block.box.top);
        const lastRow = rowAtEmu(grid, Math.max(block.box.top, block.box.bottom - 1));
        if (lastRow - firstRow >= TALL_BLOCK_ROWS) {
            index.tall.push(block);
            return;
        }
        for (let row = firstRow; row <= lastRow; row += 1) {
            const list = index.byRow.get(row);
            if (list) list.push(block);
            else index.byRow.set(row, [block]);
        }
    };
    // Mockups draw labels and values as text in shapes. That text is an item like cell text.
    for (const shapeIndex of shapeTexts) {
        const shape = (sheet.shapes as ParsedShapeAsset[])[shapeIndex] as ParsedShapeAsset;
        const box = { ...shape.bbox };
        register({
            address: cellAddress(grid, { x: box.left, y: box.top }, colToLetters),
            text: String(shape.text).trim(),
            cell: box,
            box,
            shapeIndex,
        });
    }
    for (const cell of textCells) {
        const key = `${cell.row}:${cell.col}`;
        if (hidden.has(key)) continue;
        const merge = mergeAt.get(key);
        const text = String(cell.outputValue).trim();
        const own = merge
            ? cellBox(grid, merge.startRow, merge.startCol, merge.endRow, merge.endCol)
            : cellBox(grid, cell.row, cell.col, cell.row, cell.col);
        const box = { ...own };
        if (!merge) {
            const sides = overflowSides(cell);
            const extra = Math.max(0, textWidthEmu(text) - (own.right - own.left));
            const cols = filled.get(cell.row) as number[];
            // The cell itself is in the list. Its neighbors sit right before and after it.
            const at = lowerBound(cols, cell.col);
            const next = cols[at + 1];
            const previous = at > 0 ? cols[at - 1] : undefined;
            const rightLimit = next !== undefined ? columnLeftEmu(grid, next) : Number.POSITIVE_INFINITY;
            const leftLimit = previous !== undefined ? columnLeftEmu(grid, previous + 1) : 0;
            box.right = Math.max(own.right, Math.min(rightLimit, own.right + extra * sides.right));
            box.left = Math.min(own.left, Math.max(leftLimit, own.left - extra * sides.left));
        }
        register({ address: cell.address, text, cell: own, box, shapeIndex: null });
    }
    return index;
}

function distanceToBox(point: Point, box: Box): number {
    const dx = Math.max(box.left - point.x, 0, point.x - box.right);
    const dy = Math.max(box.top - point.y, 0, point.y - box.bottom);
    return Math.hypot(dx, dy);
}

function distanceToCenter(point: Point, box: Box): number {
    return Math.hypot(point.x - (box.left + box.right) / 2, point.y - (box.top + box.bottom) / 2);
}

/** Gap between two boxes. Zero when they touch or overlap. */
function distanceBetween(a: Box, b: Box): number {
    const dx = Math.max(0, b.left - a.right, a.left - b.right);
    const dy = Math.max(0, b.top - a.bottom, a.top - b.bottom);
    return Math.hypot(dx, dy);
}

function contains(box: Box, point: Point): boolean {
    return point.x >= box.left && point.x <= box.right && point.y >= box.top && point.y <= box.bottom;
}

function area(box: Box): number {
    return (box.right - box.left) * (box.bottom - box.top);
}

/** Collects text blocks on the rows that a vertical span touches. Each block appears once. */
function blocksInRows(texts: TextIndex, grid: SheetGrid, top: number, bottom: number): TextBlock[] {
    const found = new Set<TextBlock>();
    for (let row = rowAtEmu(grid, Math.max(0, top)); row <= rowAtEmu(grid, Math.max(0, bottom)); row += 1) {
        for (const block of texts.byRow.get(row) || []) found.add(block);
    }
    for (const block of texts.tall) {
        if (block.box.bottom >= top && block.box.top <= bottom) found.add(block);
    }
    return [...found];
}

function cellAddress(grid: SheetGrid, point: Point, colToLetters: (col: number) => string): string {
    return `${colToLetters(columnAtEmu(grid, point.x))}${rowAtEmu(grid, point.y)}`;
}

/**
 * True for a small shape with a short label that a reader can see as a badge.
 * A shape without fill and without line is plain text, for example a mock value in a screen layout.
 */
function looksLikeMarker(shape: ParsedShapeAsset): boolean {
    if (!shape?.bbox || shape.connection || shape.calloutTip) return false;
    const width = shape.bbox.right - shape.bbox.left;
    const height = shape.bbox.bottom - shape.bbox.top;
    if (width > MARKER_MAX_SIZE_EMU || height > MARKER_MAX_SIZE_EMU) return false;
    if (shape.look && shape.look.fill === "none" && shape.look.line === "none") return false;
    return MARKER_TEXT.test(String(shape.text || "").trim());
}

/**
 * Picks the marker shapes of a sheet.
 * Writers number markers 1, 2, 3, and so on. A large group of one look that repeats one label
 * is a set of mock values, so the group is left out.
 */
function selectMarkers(shapes: ParsedShapeAsset[]): Set<number> {
    const groups = new Map<string, number[]>();
    shapes.forEach((shape, index) => {
        if (!looksLikeMarker(shape)) return;
        const key = JSON.stringify([shape.kind, shape.look?.fill ?? null, shape.look?.line ?? null]);
        const group = groups.get(key);
        if (group) group.push(index);
        else groups.set(key, [index]);
    });
    const markers = new Set<number>();
    for (const group of groups.values()) {
        const labels = new Set(group.map((index) => String(shapes[index]?.text || "").trim()));
        if (group.length >= MARKER_GROUP_MIN && labels.size / group.length < MARKER_DISTINCT_RATIO) continue;
        for (const index of group) markers.add(index);
    }
    return markers;
}

/** Pictures with a box. The smallest one that holds the point wins, because it is the most specific. */
function imageAt(sheet: ParsedSheet, point: Point): number | null {
    let best: number | null = null;
    let bestArea = Number.POSITIVE_INFINITY;
    (sheet.images || []).forEach((image, index) => {
        if (!image.bbox || !contains(image.bbox, point)) return;
        const size = area(image.bbox);
        if (size < bestArea) {
            best = index;
            bestArea = size;
        }
    });
    return best;
}

/** Same drawing part. Only shapes of one drawing share a space. */
function drawingPartOf(shape: ParsedShapeAsset): string {
    const id = shape.stableId || "";
    const hash = id.indexOf("#");
    return hash >= 0 ? id.slice(0, hash) : "";
}

/**
 * Resolves the targets of every callout, loose line end, and marker on one sheet.
 *
 * @param colToLetters Formats a 1-based column as letters.
 * @param fallbackCell Default cell size when the sheet carries no grid.
 */
function resolveShapeTargets(
    sheet: ParsedSheet,
    colToLetters: (col: number) => string,
    fallbackCell: { widthEmu: number; heightEmu: number } = EXCEL_DEFAULT_CELL
): ShapeTarget[] {
    const shapes = sheet.shapes || [];
    // Shapes built by hand can lack a box. They have no place to compare, so they stay out.
    if (!shapes.some((shape) => shape?.bbox)) return [];
    const grid = sheet.grid ?? createUniformGrid(fallbackCell.widthEmu, fallbackCell.heightEmu);
    const markers = selectMarkers(shapes);
    // Callout text is a note about an item, never the item itself.
    const shapeTexts = shapes
        .map((shape, index) => index)
        .filter((index) => {
            const shape = shapes[index] as ParsedShapeAsset;
            return (
                !!shape?.bbox &&
                !shape.connection &&
                !shape.calloutTip &&
                !markers.has(index) &&
                !!String(shape.text || "").trim()
            );
        });
    const texts = collectTextBlocks(sheet, grid, shapeTexts, colToLetters);
    const targets: ShapeTarget[] = [];

    /**
     * The shape that holds the point, with or without text. The smallest one wins.
     * Writers draw a highlight box around the item a callout means, and the tip often stops on its edge.
     */
    const shapeAt = (point: Point, self: number): number | null => {
        let best: number | null = null;
        shapes.forEach((shape, index) => {
            if (index === self || !shape?.bbox || shape.connection || markers.has(index)) return;
            if (drawingPartOf(shape) !== drawingPartOf(shapes[self] as ParsedShapeAsset)) return;
            if (distanceToBox(point, shape.bbox) > EDGE_TOLERANCE_EMU) return;
            if (best === null || area(shape.bbox) < area((shapes[best] as ParsedShapeAsset).bbox)) best = index;
        });
        return best;
    };

    const pointTarget = (
        shapeIndex: number,
        role: ShapeTarget["role"],
        point: Point,
        connectedShapeIndex: number | null = null
    ): ShapeTarget => {
        const found = blocksInRows(texts, grid, point.y - POINT_RADIUS_EMU, point.y + POINT_RADIUS_EMU)
            .filter((block) => block.shapeIndex !== shapeIndex)
            .map((block) => ({ block, distance: distanceToBox(point, block.box) }))
            .filter((entry) => entry.distance <= POINT_RADIUS_EMU)
            .sort(
                (a, b) =>
                    a.distance - b.distance ||
                    distanceToCenter(point, a.block.box) - distanceToCenter(point, b.block.box)
            )
            .slice(0, NEARBY_LIMIT);
        const first = found[0];
        return {
            shapeIndex,
            role,
            cell: cellAddress(grid, point, colToLetters),
            textAddress: first?.block.address ?? null,
            text: first?.block.text ?? null,
            targetShapeIndex: shapeAt(point, shapeIndex),
            targetImageIndex: imageAt(sheet, point),
            connectedShapeIndex,
            method: !first
                ? "none"
                : first.distance > 0
                  ? "nearest"
                  : contains(first.block.cell, point)
                    ? "inside"
                    : "overflow",
            nearby: found.slice(1).map(({ block }) => ({ address: block.address, text: block.text })),
        };
    };

    shapes.forEach((shape, index) => {
        if (shape.calloutTip && shape.bbox) targets.push(pointTarget(index, "callout", shape.calloutTip));
    });

    for (const entry of resolveShapeConnections(shapes)) {
        const connection = (shapes[entry.connectorIndex] as ParsedShapeAsset).connection;
        if (entry.startIndex === null && connection?.startPoint) {
            targets.push(pointTarget(entry.connectorIndex, "lineStart", connection.startPoint, entry.endIndex));
        }
        if (entry.endIndex === null && connection?.endPoint) {
            targets.push(pointTarget(entry.connectorIndex, "lineEnd", connection.endPoint, entry.startIndex));
        }
    }

    shapes.forEach((shape, index) => {
        if (markers.has(index)) targets.push(markerTarget(index, shape, sheet, grid, texts, colToLetters));
    });
    return targets;
}

/**
 * Finds the item that a marker labels.
 * Writers put markers on the item or just left of it. The order is:
 * text whose own cell or shape holds the marker center, then text that starts right after the marker,
 * then a picture under the marker, then overflowing text under the marker,
 * then the nearest text or picture.
 */
function markerTarget(
    shapeIndex: number,
    shape: ParsedShapeAsset,
    sheet: ParsedSheet,
    grid: SheetGrid,
    texts: TextIndex,
    colToLetters: (col: number) => string
): ShapeTarget {
    const box = shape.bbox;
    const width = box.right - box.left;
    const height = box.bottom - box.top;
    const center = { x: (box.left + box.right) / 2, y: (box.top + box.bottom) / 2 };
    const target: ShapeTarget = {
        shapeIndex,
        role: "marker",
        cell: cellAddress(grid, center, colToLetters),
        textAddress: null,
        text: null,
        targetShapeIndex: null,
        targetImageIndex: null,
        connectedShapeIndex: null,
        method: "none",
        nearby: [],
    };
    const others = (skip: TextBlock | null) =>
        candidates
            .filter((entry) => entry.block !== skip)
            .slice(0, NEARBY_LIMIT - 1)
            .map((entry) => ({ address: entry.block.address, text: entry.block.text }));
    const pick = (block: TextBlock, method: ShapeTarget["method"]): ShapeTarget => ({
        ...target,
        textAddress: block.address,
        text: block.text,
        targetShapeIndex: block.shapeIndex,
        method,
        nearby: candidates
            .filter((entry) => entry.block !== block)
            .slice(0, NEARBY_LIMIT - 1)
            .map((entry) => ({ address: entry.block.address, text: entry.block.text })),
    });
    // Ties go to the block whose center is closer to the marker center.
    const closest = (blocks: TextBlock[], score: (block: TextBlock) => number): TextBlock | null => {
        let best: TextBlock | null = null;
        let bestScore = Number.POSITIVE_INFINITY;
        let bestCenter = Number.POSITIVE_INFINITY;
        for (const block of blocks) {
            const value = score(block);
            const toCenter = distanceToCenter(center, block.box);
            if (value < bestScore || (value === bestScore && toCenter < bestCenter)) {
                best = block;
                bestScore = value;
                bestCenter = toCenter;
            }
        }
        return best;
    };

    const nearby = blocksInRows(texts, grid, box.top - POINT_RADIUS_EMU, box.bottom + POINT_RADIUS_EMU).filter(
        (block) => block.shapeIndex !== shapeIndex
    );
    // Every text within reach, nearest first. Other matches stay visible, because markers can sit between items.
    const candidates = nearby
        .map((block) => ({ block, distance: distanceBetween(box, block.box) }))
        .filter((entry) => entry.distance <= POINT_RADIUS_EMU)
        .sort(
            (a, b) =>
                a.distance - b.distance || distanceToCenter(center, a.block.box) - distanceToCenter(center, b.block.box)
        );
    const inside = closest(
        nearby.filter((block) => contains(block.cell, center)),
        (block) => area(block.cell)
    );
    if (inside) return pick(inside, "inside");

    // Same rows: the block overlaps the middle half of the marker height. It starts within one marker width.
    const bandTop = box.top + height / 4;
    const bandBottom = box.bottom - height / 4;
    const right = closest(
        nearby.filter(
            (block) =>
                block.box.bottom >= bandTop &&
                block.box.top <= bandBottom &&
                block.box.left >= center.x &&
                block.box.left - box.right <= width
        ),
        (block) => Math.max(0, block.box.left - box.right)
    );
    if (right) return pick(right, "right");

    const image = imageAt(sheet, center);
    if (image !== null) return { ...target, targetImageIndex: image, method: "image", nearby: others(null) };

    const overflow = closest(
        nearby.filter((block) => contains(block.box, center)),
        (block) => area(block.box)
    );
    if (overflow) return pick(overflow, "overflow");

    const nearest = closest(
        nearby.filter((block) => distanceBetween(box, block.box) <= POINT_RADIUS_EMU),
        (block) => distanceBetween(box, block.box)
    );
    // A picture closer than any text wins, for example a logo next to its marker.
    let nearestImage: number | null = null;
    let imageGap = nearest ? distanceBetween(box, nearest.box) : POINT_RADIUS_EMU;
    (sheet.images || []).forEach((picture, index) => {
        if (!picture.bbox) return;
        const gap = distanceBetween(box, picture.bbox);
        if (gap < imageGap) {
            nearestImage = index;
            imageGap = gap;
        }
    });
    if (nearestImage !== null)
        return { ...target, targetImageIndex: nearestImage, method: "nearest", nearby: others(null) };
    return nearest ? pick(nearest, "nearest") : target;
}

export { resolveShapeTargets };
