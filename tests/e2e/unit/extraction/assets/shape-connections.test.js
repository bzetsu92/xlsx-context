import { describe, expect, it } from "vitest";
import { resolveShapeConnections } from "../../../../../src/extraction/assets/shape-connections.ts";

function shape(drawingId, bbox, connection = null) {
    return {
        sheetName: "Sheet1",
        anchor: "A1",
        drawingId,
        stableId: `xl/drawings/drawing1.xml#${drawingId}`,
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

const BOX = { left: 0, top: 0, right: 10, bottom: 10 };

describe("shape connections", () => {
    it("confirms fully declared links and infers flow from one arrowhead", () => {
        const shapes = [shape("1", BOX), shape("2", BOX), shape("3", BOX, { startId: "1", endId: "2", arrow: "end" })];
        const [entry] = resolveShapeConnections(shapes);
        expect(entry.resolution).toBe("CONFIRMED");
        expect(entry.startIndex).toBe(0);
        expect(entry.endIndex).toBe(1);
        expect(entry.arrowAtStart).toBe(false);
        expect(entry.arrowAtEnd).toBe(true);
        expect(entry.flowFrom).toBe(0);
        expect(entry.flowTo).toBe(1);
        expect(entry.flowEvidence).toBe("INFERRED");
    });

    it("separates arrowhead placement from business flow", () => {
        const shapes = [
            shape("1", BOX),
            shape("2", BOX),
            shape("3", BOX, { startId: "1", endId: "2", arrow: "start" }),
        ];
        const [entry] = resolveShapeConnections(shapes);
        expect(entry.arrowAtStart).toBe(true);
        expect(entry.arrowAtEnd).toBe(false);
        expect(entry.flowFrom).toBe(1);
        expect(entry.flowTo).toBe(0);
        expect(entry.flowEvidence).toBe("INFERRED");
    });

    it("leaves double-headed and plain connectors without flow", () => {
        const both = [shape("1", BOX), shape("2", BOX), shape("3", BOX, { startId: "1", endId: "2", arrow: "both" })];
        const [bothEntry] = resolveShapeConnections(both);
        expect(bothEntry.resolution).toBe("CONFIRMED");
        expect(bothEntry.flowFrom).toBe(null);
        expect(bothEntry.flowEvidence).toBe("UNRESOLVED");

        const none = [shape("1", BOX), shape("2", BOX), shape("3", BOX, { startId: "1", endId: "2", arrow: "none" })];
        const [noneEntry] = resolveShapeConnections(none);
        expect(noneEntry.resolution).toBe("CONFIRMED");
        expect(noneEntry.flowEvidence).toBe("UNRESOLVED");
    });

    it("marks unknown targets unresolved with reasons", () => {
        const shapes = [shape("1", BOX), shape("2", BOX, { startId: "1", endId: "9", arrow: "end" })];
        const [entry] = resolveShapeConnections(shapes);
        expect(entry.resolution).toBe("UNRESOLVED");
        expect(entry.startIndex).toBe(0);
        expect(entry.endIndex).toBe(null);
        expect(entry.reason).toContain("9");
    });

    it("infers loose connector ends from line-end geometry", () => {
        const loose = { startId: null, endId: null, arrow: "end" };
        const shapes = [
            shape("1", { left: 0, top: 0, right: 100, bottom: 100 }),
            shape("2", { left: 0, top: 300000, right: 100, bottom: 300100 }),
            // A large frame holds both boxes. The nearest box side wins over the frame.
            shape("3", { left: -50000, top: -50000, right: 500000, bottom: 500000 }),
            shape(
                "4",
                { left: 50, top: 100, right: 50, bottom: 300000 },
                {
                    ...loose,
                    startPoint: { x: 50, y: 150 },
                    endPoint: { x: 50, y: 299000 },
                }
            ),
        ];
        const [entry] = resolveShapeConnections(shapes);
        expect(entry.resolution).toBe("INFERRED");
        expect(entry.startIndex).toBe(0);
        expect(entry.endIndex).toBe(1);
        expect(entry.flowFrom).toBe(0);
        expect(entry.flowTo).toBe(1);
    });

    it("keeps a far line end unresolved and ignores other drawing parts", () => {
        const far = shape("1", { left: 0, top: 0, right: 100, bottom: 100 });
        const otherPart = {
            ...shape("2", { left: 0, top: 900000, right: 100, bottom: 900100 }),
            stableId: "xl/drawings/drawing2.xml#2",
        };
        const connector = shape(
            "3",
            { left: 50, top: 100, right: 50, bottom: 900000 },
            {
                startId: null,
                endId: null,
                arrow: "end",
                startPoint: { x: 50, y: 100 },
                endPoint: { x: 50, y: 899000 },
            }
        );
        const [entry] = resolveShapeConnections([far, otherPart, connector]);
        expect(entry.resolution).toBe("UNRESOLVED");
        expect(entry.startIndex).toBe(0);
        expect(entry.endIndex).toBe(null);
        expect(entry.reason).toContain("missing endCxn id");
    });
});
