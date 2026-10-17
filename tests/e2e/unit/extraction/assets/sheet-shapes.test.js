import { describe, expect, it } from "vitest";
import * as sheetShapes from "../../../../../src/extraction/assets/sheet-shapes.ts";

describe("xlsx-context sheet shapes", () => {
    it("groups nearby shapes into shape blocks", () => {
        const api = sheetShapes;
        const blocks = api.extractShapeBlocks(
            [
                { bbox: { left: 0, top: 0, right: 100, bottom: 20 } },
                { bbox: { left: 140, top: 0, right: 240, bottom: 20 } },
                { bbox: { left: 1000, top: 400, right: 1100, bottom: 420 } },
            ],
            {
                defaultCellWidthEmu: 100,
                defaultCellHeightEmu: 20,
                shapeBlockGapXEmu: 100,
                shapeBlockGapYEmu: 40,
            }
        );

        expect(blocks).toHaveLength(2);
        expect(blocks[0].shapeIndexes).toEqual([0, 1]);
        expect(blocks[1].shapeIndexes).toEqual([2]);
    });
});
