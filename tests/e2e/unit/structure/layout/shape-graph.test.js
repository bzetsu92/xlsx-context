import { describe, expect, it } from "vitest";
import {
    buildShapeGraph,
    buildShapeGraphForBlock,
    resolveConnectorEnds,
} from "../../../../../src/structure/layout/shape-graph.ts";

function box(col, row) {
    return { left: col * 10, top: row * 10, right: col * 10 + 5, bottom: row * 10 + 5 };
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

describe("shape graph", () => {
    it("resolves links across the full list", () => {
        const shapes = [
            shape("1", box(0, 0)),
            shape("2", box(50, 50)),
            shape("3", box(25, 25), { startId: "1", endId: "2", arrow: "end" }),
        ];
        const graph = buildShapeGraph(shapes);
        expect(graph.edges.length).toBe(1);
        expect(graph.edges[0]).toMatchObject({ from: 0, to: 1, connector: 2, arrow: "end" });
        expect(graph.looseConnectors).toEqual([]);
    });

    it("points start-only arrows toward their target", () => {
        const shapes = [
            shape("1", box(0, 0)),
            shape("2", box(50, 50)),
            shape("3", box(25, 25), { startId: "1", endId: "2", arrow: "start" }),
        ];
        const graph = buildShapeGraph(shapes);
        expect(graph.edges[0]).toMatchObject({ from: 1, to: 0, arrow: "end" });
    });

    it("keeps parallel edges and lists loose connectors", () => {
        const shapes = [
            shape("1", box(0, 0)),
            shape("2", box(50, 50)),
            shape("3", box(1, 1), { startId: "1", endId: "2", arrow: "end" }),
            shape("4", box(2, 2), { startId: "1", endId: "2", arrow: "none" }),
            shape("5", box(3, 3), { startId: "1", endId: null, arrow: "end" }),
        ];
        const graph = buildShapeGraph(shapes);
        expect(graph.edges.length).toBe(2);
        expect(graph.looseConnectors).toEqual([4]);
    });

    it("keeps drawing-wide reasons for block slices", () => {
        const shapes = [shape("1", box(0, 0)), shape("2", box(50, 50), { startId: "1", endId: "9", arrow: "end" })];
        const graph = buildShapeGraphForBlock(shapes, [1]);
        expect(graph.edges.length).toBe(0);
        expect(graph.looseConnectors).toEqual([1]);
        expect(graph.unresolved[0].reason).toContain("9");
        expect(graph.unresolved[0].reason).not.toContain("'1'");
    });

    it("orients connector ends for every connector and keeps loose ends null", () => {
        const shapes = [
            shape("1", box(0, 0)),
            shape("2", box(50, 50)),
            shape("3", box(25, 25), { startId: "2", endId: "1", arrow: "start" }),
            shape("4", box(30, 30), { startId: "1", endId: null, arrow: "both" }),
        ];
        const ends = resolveConnectorEnds(shapes);
        expect(ends.get(2)).toEqual({ from: 0, to: 1, arrow: "end" });
        expect(ends.get(3)).toEqual({ from: 0, to: null, arrow: "both" });
        expect(ends.has(0)).toBe(false);
    });
});

describe("buildShapeGraphForBlock() with loose connectors", () => {
    it("keeps a half-resolved connector loose instead of inventing an edge", () => {
        const shapes = [shape("1", box(0, 0)), shape("2", box(2, 0), { startId: "1", endId: null, arrow: "end" })];
        const graph = buildShapeGraphForBlock(shapes, [0, 1]);
        expect(graph.looseConnectors).toContain(1);
        expect(graph.edges.length).toBe(0);
    });

    it("keeps arrow=none without inventing a flow direction", () => {
        const shapes = [
            shape("1", box(0, 0)),
            shape("2", box(3, 0)),
            shape("3", box(0, 1), { startId: "1", endId: "2", arrow: "none" }),
        ];
        const graph = buildShapeGraphForBlock(shapes, [0, 1, 2]);
        expect(graph.edges.length).toBe(1);
        expect(graph.edges[0].arrow).toBe("none");
    });
});
