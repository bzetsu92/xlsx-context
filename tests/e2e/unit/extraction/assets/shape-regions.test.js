import { describe, expect, it } from "vitest";
import { partitionShapeRegions } from "../../../../../src/extraction/assets/shape-regions.ts";
import { columnLeftEmu } from "../../../../../src/extraction/workbook/sheet-grid.ts";

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

function shape(drawingId, bbox, connection = null) {
    return {
        sheetName: "Sheet1",
        anchor: "A1",
        drawingId,
        name: drawingId,
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

describe("shape regions", () => {
    it("covers every shape exactly once", () => {
        const shapes = [
            shape("1", box(0, 0)),
            shape("2", box(3, 0)),
            shape("3", box(0, 1), { startId: "1", endId: "2", arrow: "end" }),
            shape("4", box(30, 30)),
        ];
        const regions = partitionShapeRegions(shapes, DEPS);
        const flat = regions.flatMap((region) => region.shapeIndexes).sort((a, b) => a - b);
        expect(flat).toEqual([0, 1, 2, 3]);
    });

    it("keeps connected members together across empty space", () => {
        const shapes = [
            shape("1", box(0, 0)),
            shape("2", box(20, 0)),
            shape("3", box(10, 0), { startId: "1", endId: "2", arrow: "end" }),
        ];
        const regions = partitionShapeRegions(shapes, DEPS);
        expect(regions.length).toBe(1);
        expect([...regions[0].shapeIndexes].sort((a, b) => a - b)).toEqual([0, 1, 2]);
    });

    it("keeps unresolved connectors without dropping them", () => {
        const shapes = [shape("1", box(0, 0)), shape("2", box(2, 0), { startId: "1", endId: null, arrow: "end" })];
        const regions = partitionShapeRegions(shapes, DEPS);
        const flat = regions.flatMap((region) => region.shapeIndexes).sort((a, b) => a - b);
        expect(flat).toEqual([0, 1]);
    });

    it("groups lone shapes by proximity only", () => {
        const near = [shape("1", box(0, 0)), shape("2", box(1, 0))];
        expect(partitionShapeRegions(near, DEPS).length).toBe(1);
        const far = [shape("1", box(0, 0)), shape("2", box(30, 30))];
        expect(partitionShapeRegions(far, DEPS).length).toBe(2);
    });

    it("keeps a self link with its shape", () => {
        const shapes = [shape("1", box(0, 0)), shape("2", box(0, 1), { startId: "1", endId: "1", arrow: "end" })];
        const regions = partitionShapeRegions(shapes, DEPS);
        expect(regions.length).toBe(1);
    });

    it("measures proximity in cells on a grid with narrow columns", () => {
        // The default column is wide, but columns 1-40 are narrow. A gap of 10 narrow columns is far.
        const grid = {
            defaultColWidthEmu: CELL_W,
            defaultRowHeightEmu: CELL_H,
            colSpans: [{ start: 1, end: 40, sizeEmu: CELL_W / 4 }],
            rowSpans: [],
            confidence: "EXACT",
        };
        const at = (col) => ({
            left: columnLeftEmu(grid, col),
            top: 0,
            right: columnLeftEmu(grid, col + 1),
            bottom: CELL_H,
        });
        const regions = partitionShapeRegions([shape("1", at(1)), shape("2", at(12))], { ...DEPS, grid });
        expect(regions).toHaveLength(2);
        expect(regions.map((region) => [region.startCol, region.endCol])).toEqual([
            [1, 1],
            [12, 12],
        ]);
    });
});
