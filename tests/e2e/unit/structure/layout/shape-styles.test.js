import { describe, expect, it } from "vitest";
import { groupShapeStyles } from "../../../../../src/structure/layout/shape-styles.ts";

function shape(kind, look, connection = null) {
    return { kind, look, connection, text: "", bbox: { left: 0, top: 0, right: 1, bottom: 1 } };
}

const BLUE = { fill: "#0070C0", line: "#000000", dash: null, lineWidthEmu: 15875 };

describe("shape styles", () => {
    it("groups equal looks and ignores small line width changes", () => {
        const styles = groupShapeStyles([
            shape("Shape (wedgeRectCallout)", BLUE),
            shape("Shape (wedgeRectCallout)", { ...BLUE, lineWidthEmu: 9525 }),
            shape("Shape (wedgeRectCallout)", { ...BLUE, fill: "#D7E4BD" }),
        ]);
        expect(styles.map((style) => [style.id, style.presets, style.shapeIndexes])).toEqual([
            ["S1", ["wedgeRectCallout"], [0, 1]],
            ["S2", ["wedgeRectCallout"], [2]],
        ]);
    });

    it("separates dashed and thick looks, but not shape kinds", () => {
        const styles = groupShapeStyles([
            shape("Rectangle", BLUE),
            shape("Rectangle", { ...BLUE, dash: "dash" }),
            shape("Rectangle", { ...BLUE, lineWidthEmu: 38100 }),
            shape("Shape (roundRect)", BLUE),
            shape("Rectangle", { ...BLUE, dash: "solid" }),
        ]);
        expect(styles.map((style) => style.shapeIndexes)).toEqual([[0, 3, 4], [1], [2]]);
        expect(styles[0].presets).toEqual(["rect", "roundRect"]);
        expect(styles[2].thick).toBe(true);
    });

    it("ignores fill on lines", () => {
        const line = { startId: null, endId: null, arrow: "end" };
        const styles = groupShapeStyles([
            shape("Straight Arrow Connector", { ...BLUE, fill: null }, line),
            shape("Straight Arrow Connector", { ...BLUE, fill: "#FFFFFF" }, line),
        ]);
        expect(styles).toHaveLength(1);
        expect(styles[0].isLine).toBe(true);
    });
});
