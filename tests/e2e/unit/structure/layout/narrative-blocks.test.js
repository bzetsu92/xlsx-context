import { describe, expect, it } from "vitest";
import * as narrativeBlocks from "../../../../../src/structure/layout/narrative-blocks.ts";

const formatCell = (cell) => cell.outputValue;

function createCell(row, col, outputValue) {
    return { row, col, outputValue };
}

function createTable(startRow, startCol, endRow, endCol) {
    return { startRow, startCol, endRow, endCol, score: 1, reasonSummary: [] };
}

describe("xlsx-context narrative blocks", () => {
    it("splits a row into segments on gaps wider than 4 columns", () => {
        const api = narrativeBlocks;
        const cells = [createCell(1, 1, "A"), createCell(1, 2, "B"), createCell(1, 8, "C"), createCell(1, 9, " ")];

        expect(api.splitNarrativeRowSegments(cells, formatCell)).toEqual([
            { startCol: 1, values: ["A", "B"] },
            { startCol: 8, values: ["C"] },
        ]);
    });

    it("skips cells inside tables", () => {
        const api = narrativeBlocks;
        const table = createTable(2, 1, 3, 2);

        expect(api.isCellInAnyTable(2, 2, [table])).toBe(true);
        expect(api.isCellInAnyTable(4, 1, [table])).toBe(false);

        const sheet = { cells: [createCell(1, 1, "Title"), createCell(2, 1, "inside"), createCell(4, 1, "after")] };
        const items = api.buildNarrativeItems(sheet, [table], formatCell);

        expect(items.map((item) => item.text)).toEqual(["Title", "after"]);
    });

    it("groups adjacent rows and splits on row gaps and table captions", () => {
        const api = narrativeBlocks;
        const table = createTable(4, 1, 6, 3);
        const sheet = {
            cells: [
                createCell(1, 1, "Title"),
                createCell(2, 1, "Note"),
                createCell(3, 1, "Caption"),
                createCell(5, 2, "in table"),
                createCell(10, 1, "Footer"),
            ],
        };

        const blocks = api.extractNarrativeBlocks(sheet, [table], formatCell);

        expect(blocks.map((block) => block.lines)).toEqual([["Title", "Note"], ["Caption"], ["Footer"]]);
        expect(blocks[0]).toMatchObject({ startRow: 1, endRow: 2, startCol: 1 });
    });

    it("detects calendar date rows", () => {
        const api = narrativeBlocks;
        const dates = ["2024-01-01", "2024-01-02", "2024-01-03", "2024-01-04", "2024-01-05"];
        const calendar = { items: [{ row: 1, startCol: 1, text: dates.join(" "), cellValues: dates }] };
        const prose = { items: [{ row: 1, startCol: 1, text: "Hello", cellValues: ["Hello"] }] };

        expect(api.blockHasCalendarDateItem(calendar)).toBe(true);
        expect(api.blockHasCalendarDateItem(prose)).toBe(false);
        expect(api.blockHasCalendarDateItem(null)).toBe(false);
    });
});
