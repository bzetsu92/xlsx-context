import { describe, expect, it } from "vitest";
import { extractShapeBlocks } from "../../../../../src/extraction/assets/sheet-shapes.ts";
import { buildShapeGraphForBlock } from "../../../../../src/structure/layout/shape-graph.ts";

const CELL_W = 609600;
const CELL_H = 190500;
const DEPS = {
    defaultCellWidthEmu: CELL_W,
    defaultCellHeightEmu: CELL_H,
    shapeBlockGapXEmu: CELL_W * 4,
    shapeBlockGapYEmu: CELL_H * 6,
};

function box(col, row, cols = 2, rows = 1) {
    const left = col * CELL_W;
    const top = row * CELL_H;
    return { left, top, right: left + cols * CELL_W, bottom: top + rows * CELL_H };
}

function shape(index, drawingId, bbox, connection = null) {
    return {
        sheetName: "Sheet1",
        anchor: "A1",
        drawingId,
        name: `Shape ${drawingId}`,
        kind: connection ? "Connector" : "Rectangle",
        text: "",
        connection,
        widthEmu: null,
        heightEmu: null,
        elementName: connection ? "xdr:cxnSp" : "xdr:sp",
        anchorElementName: "xdr:twoCellAnchor",
        rawEntries: [],
        bbox,
        svgFilename: null,
        svgPath: null,
        svgData: null,
    };
}

function allIndexes(blocks) {
    return blocks.flatMap((b) => b.shapeIndexes);
}

describe("extractShapeBlocks() partition invariants", () => {
    it("keeps every shape exactly once across blocks", () => {
        const shapes = [
            shape(0, "1", box(0, 0)),
            shape(1, "2", box(3, 0)),
            shape(2, "3", box(0, 1), { startId: "1", endId: "2", arrow: "end" }),
            shape(3, "4", box(0, 3)),
            shape(4, "5", box(3, 3)),
            shape(5, "6", box(0, 4), { startId: "4", endId: "5", arrow: "end" }),
        ];
        const blocks = extractShapeBlocks(shapes, DEPS);
        const flat = allIndexes(blocks).sort((a, b) => a - b);
        expect(flat).toEqual([0, 1, 2, 3, 4, 5]);
    });

    it("does not merge two close but unconnected diagrams by distance alone", () => {
        const shapes = [
            shape(0, "1", box(0, 0)),
            shape(1, "2", box(3, 0)),
            shape(2, "3", box(0, 1), { startId: "1", endId: "2", arrow: "end" }),
            shape(3, "4", box(0, 3)),
            shape(4, "5", box(3, 3)),
            shape(5, "6", box(0, 4), { startId: "4", endId: "5", arrow: "end" }),
        ];
        const blocks = extractShapeBlocks(shapes, DEPS);
        expect(blocks.length).toBe(2);
    });

    it("does not lose a CONFIRMED edge when endpoints span an empty area", () => {
        const shapes = [
            shape(0, "1", box(0, 0)),
            shape(1, "2", box(20, 0)),
            shape(2, "3", box(10, 0), { startId: "1", endId: "2", arrow: "end" }),
        ];
        const full = buildShapeGraphForBlock(shapes, [0, 1, 2]);
        expect(full.edges.length).toBe(1);

        const blocks = extractShapeBlocks(shapes, DEPS);
        const perBlockEdges = blocks.map((b) => buildShapeGraphForBlock(shapes, b.shapeIndexes).edges.length);
        const kept = perBlockEdges.reduce((sum, n) => sum + n, 0);
        expect(kept).toBe(full.edges.length);
    });

    it("topology is stable when gap thresholds change", () => {
        const shapes = [
            shape(0, "1", box(0, 0)),
            shape(1, "2", box(3, 0)),
            shape(2, "3", box(0, 1), { startId: "1", endId: "2", arrow: "end" }),
        ];
        const full = buildShapeGraphForBlock(shapes, [0, 1, 2]);
        const tight = extractShapeBlocks(shapes, { ...DEPS, shapeBlockGapXEmu: 1, shapeBlockGapYEmu: 1 });
        const kept = tight
            .map((b) => buildShapeGraphForBlock(shapes, b.shapeIndexes).edges.length)
            .reduce((sum, n) => sum + n, 0);
        expect(kept).toBe(full.edges.length);
    });
});
