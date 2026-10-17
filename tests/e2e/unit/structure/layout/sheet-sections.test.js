import { describe, expect, it } from "vitest";
import * as sheetSections from "../../../../../src/structure/layout/sheet-sections.ts";

const EMPTY_SHEET = { cells: [], images: [], charts: [] };

function createTable(startRow, startCol, endRow, endCol) {
    return { startRow, startCol, endRow, endCol, score: 1, reasonSummary: [] };
}

function createBlock(row, col, values) {
    return {
        startRow: row,
        startCol: col,
        endRow: row,
        lines: [values.join(" ")],
        items: [{ row, startCol: col, text: values.join(" "), cellValues: values }],
    };
}

function createEntry(block) {
    return { sortRow: block.startRow, sortCol: block.startCol, kind: "narrative", narrativeBlock: block };
}

describe("xlsx-context sheet sections", () => {
    it("joins nearby anchors into one section and starts a new one after a gap", () => {
        const api = sheetSections;
        const title = createBlock(1, 1, ["Title"]);
        const near = createTable(3, 1, 5, 3);
        const far = createTable(20, 1, 22, 3);

        const blocks = api.extractSectionBlocks(EMPTY_SHEET, [near, far], [title]);

        expect(blocks).toEqual([
            { startRow: 1, startCol: 1, endRow: 5, endCol: 3 },
            { startRow: 20, startCol: 1, endRow: 22, endCol: 3 },
        ]);
    });

    it("plans reading order per section and orders tables by it", () => {
        const api = sheetSections;
        const title = createBlock(1, 1, ["Title"]);
        const near = createTable(3, 1, 5, 3);
        const far = createTable(20, 1, 22, 3);

        const plan = api.planSheetReadingOrder(EMPTY_SHEET, [far, near], [title]);

        expect(plan.map((section) => section.map((entry) => entry.kind))).toEqual([["narrative", "table"], ["table"]]);
        expect(plan[0][1].table).toBe(near);
        expect(api.orderTablesForReading(EMPTY_SHEET, [far, near], [title])).toEqual([near, far]);
    });

    it("groups each entry into one section only", () => {
        const api = sheetSections;
        const sectionBlocks = [
            { startRow: 1, startCol: 1, endRow: 10, endCol: 10 },
            { startRow: 5, startCol: 1, endRow: 15, endCol: 10 },
        ];
        const entries = [createEntry(createBlock(6, 2, ["A"])), createEntry(createBlock(12, 2, ["B"]))];

        const groups = api.createGroupedSections(sectionBlocks, entries);

        expect(groups.map((group) => group.entries)).toEqual([[entries[0]], [entries[1]]]);
    });

    it("moves side calendars after the main calendar behind a sidebar marker", () => {
        const api = sheetSections;
        const dates = ["2024-01-01", "2024-01-02", "2024-01-03", "2024-01-04", "2024-01-05"];
        const other = createEntry(createBlock(1, 1, ["Notes"]));
        const sidebar = createEntry(createBlock(3, 12, dates));
        const main = createEntry(createBlock(3, 1, dates));
        const createSidebarEntry = (sortRow, sortCol) => ({ sortRow, sortCol, kind: "narrative" });

        const ordered = api.createCalendarAwareSectionEntries([other, sidebar, main], createSidebarEntry);

        expect(ordered).toEqual([main, { sortRow: 2.9, sortCol: 12, kind: "narrative" }, sidebar, other]);
    });
    it("joins an out-of-block asset to the nearest section and keeps a far one as its own group", () => {
        const title = createBlock(1, 7, ["Flow title"]);
        const table = createTable(3, 2, 6, 5);
        const blocks = [
            { startRow: 1, startCol: 7, endRow: 1, endCol: 7 },
            { startRow: 3, startCol: 2, endRow: 6, endCol: 5 },
        ];
        const titleEntry = createEntry(title);
        const tableEntry = { sortRow: 3, sortCol: 2, kind: "table", table };
        const nearAsset = { sortRow: 3, sortCol: 8, kind: "asset" };
        const farAsset = { sortRow: 40, sortCol: 2, kind: "asset" };
        const membership = new Map([
            [title, 0],
            [table, 1],
        ]);

        const groups = sheetSections.createGroupedSections(
            blocks,
            [titleEntry, tableEntry, nearAsset, farAsset],
            membership
        );

        // The asset at H3 is one column from the title block, closer than the table block (three columns).
        expect(groups.map((group) => group.entries)).toEqual([[titleEntry, nearAsset], [tableEntry], [farAsset]]);
    });

    it("inserts a far asset group before later groups by row", () => {
        const late = createBlock(50, 1, ["Late"]);
        const blocks = [{ startRow: 50, startCol: 1, endRow: 50, endCol: 1 }];
        const lateEntry = createEntry(late);
        const asset = { sortRow: 10, sortCol: 30, kind: "asset" };

        const groups = sheetSections.createGroupedSections(blocks, [lateEntry, asset], new Map([[late, 0]]));

        expect(groups.map((group) => group.entries)).toEqual([[asset], [lateEntry]]);
    });
});
