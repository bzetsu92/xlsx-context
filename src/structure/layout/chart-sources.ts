/** Links chart series to the detected tables that hold their values. */
import type { ParsedChartAsset, TableCandidate } from "../../model/index.js";
import { parseSheetRangeRef } from "../addresses/address-utils.js";

export type ChartSourceLink = {
    /** Table that holds each series' values, in series order. Null when no table on this sheet holds them. */
    seriesTables: (TableCandidate | null)[];
    /** Tables that hold series values, in series order without repeats. */
    tables: TableCandidate[];
    /** True when every series points into one of the tables. Then the tables already show the chart data. */
    complete: boolean;
};

function isRangeInsideTable(
    range: { startRow: number; startCol: number; endRow: number; endCol: number },
    table: TableCandidate
): boolean {
    return (
        range.startRow >= table.startRow &&
        range.endRow <= table.endRow &&
        range.startCol >= table.startCol &&
        range.endCol <= table.endCol
    );
}

/**
 * Finds the tables on this sheet that contain each series value range.
 * A series without a value range, or with a range on another sheet, makes the link incomplete.
 */
function findChartSourceTables(sheetName: string, chart: ParsedChartAsset, tables: TableCandidate[]): ChartSourceLink {
    const found: TableCandidate[] = [];
    const seriesTables: (TableCandidate | null)[] = [];
    let complete = true;
    for (const series of chart.series || []) {
        const range = parseSheetRangeRef(series.valuesRef);
        // Excel compares sheet names without case.
        const onThisSheet =
            !!range && (range.sheetName === null || range.sheetName.toLowerCase() === sheetName.toLowerCase());
        const table =
            range && onThisSheet ? tables.find((candidate) => isRangeInsideTable(range, candidate)) : undefined;
        seriesTables.push(table || null);
        if (!table) {
            complete = false;
            continue;
        }
        if (!found.includes(table)) found.push(table);
    }
    return { seriesTables, tables: found, complete };
}

export { findChartSourceTables };
