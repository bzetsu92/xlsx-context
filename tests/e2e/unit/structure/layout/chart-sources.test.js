import { describe, expect, it } from "vitest";
import { findChartSourceTables } from "../../../../../src/structure/layout/chart-sources.ts";

function createTable(startRow, startCol, endRow, endCol) {
    return { startRow, startCol, endRow, endCol, score: 1, reasonSummary: [] };
}

function createChart(valueRefs) {
    return {
        anchor: "A1",
        title: "",
        chartType: "bar",
        series: valueRefs.map((valuesRef) => ({ name: "s", valuesRef })),
    };
}

describe("xlsx-context chart sources", () => {
    const left = createTable(3, 2, 8, 5);
    const right = createTable(3, 8, 8, 10);

    it("links every series to the table that holds its values", () => {
        const link = findChartSourceTables(
            "Sales Data",
            createChart(["'Sales Data'!$C$4:$C$8", "'sales data'!$I$4:$I$8", "$D$4:$D$8"]),
            [left, right]
        );

        expect(link.seriesTables).toEqual([left, right, left]);
        expect(link.tables).toEqual([left, right]);
        expect(link.complete).toBe(true);
    });

    it("marks the link incomplete for other sheets, ranges past a table edge, and series without a range", () => {
        const link = findChartSourceTables(
            "Sales Data",
            createChart(["Other!$C$4:$C$8", "'Sales Data'!$C$4:$C$9", "", "'Sales Data'!$C$4:$C$8"]),
            [left]
        );

        expect(link.seriesTables).toEqual([null, null, null, left]);
        expect(link.tables).toEqual([left]);
        expect(link.complete).toBe(false);
    });

    it("treats a chart without series as complete with no tables", () => {
        expect(findChartSourceTables("S", createChart([]), [left])).toEqual({
            seriesTables: [],
            tables: [],
            complete: true,
        });
    });
});
