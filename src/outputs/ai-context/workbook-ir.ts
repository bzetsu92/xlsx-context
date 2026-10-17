import type {
    CellHyperlink,
    MergeRange,
    NarrativeBlock,
    ParsedCell,
    ParsedSheet,
    ParsedWorkbook,
    TableCandidate,
} from "../../model/index.js";
import type {
    AiIrOptions,
    IrCell,
    IrChart,
    IrHyperlink,
    IrImage,
    IrNarrative,
    IrSectionEntry,
    IrShape,
    IrShapeTarget,
    IrSheet,
    IrTable,
    WorkbookIR,
} from "./types.js";
import { colToLetters, formatRange, parseCellAddress } from "../../structure/addresses/address-utils.js";
import { findChartSourceTables } from "../../structure/layout/chart-sources.js";
import { toMarkdownLinkTarget } from "../../structure/formatting/link-target.js";
import { extractNarrativeBlocks } from "../../structure/layout/narrative-blocks.js";
import { resolveConnectorEnds } from "../../structure/layout/shape-graph.js";
import { resolveShapeTargets } from "../../structure/layout/shape-targets.js";
import { orderTablesForReading, planSheetReadingOrder } from "../../structure/layout/sheet-sections.js";
import { detectTableCandidates } from "../../structure/tables/table-detector.js";
import { planTableLayout } from "../../structure/tables/table-layout.js";

/** Builds an address-to-cell map for one sheet. */
function buildCellMap(cells: ParsedCell[]): Map<string, ParsedCell> {
    const map = new Map<string, ParsedCell>();
    for (const cell of cells) {
        map.set(`${cell.row}:${cell.col}`, cell);
    }
    return map;
}

/** Returns the merge range covering the given position, or null. */
function findMerge(merges: MergeRange[], row: number, col: number): MergeRange | null {
    for (const merge of merges) {
        if (row >= merge.startRow && row <= merge.endRow && col >= merge.startCol && col <= merge.endCol) {
            return merge;
        }
    }
    return null;
}

/**
 * Counts a position with a raw or display value. Skips covered merge cells.
 * Raw counts too, so the IR never drops a raw-only value.
 */
function hasValue(cellMap: Map<string, ParsedCell>, merges: MergeRange[], row: number, col: number): boolean {
    const merge = findMerge(merges, row, col);
    if (merge && (merge.startRow !== row || merge.startCol !== col)) return false;
    const cell = cellMap.get(`${row}:${col}`);
    if (!cell) return false;
    return String(cell.outputValue || "").trim() !== "" || String(cell.rawValue || "").trim() !== "";
}

/** Converts one grid position to an IR cell. Missing cells stay empty. */
function toIrCell(cellMap: Map<string, ParsedCell>, merges: MergeRange[], row: number, col: number): IrCell {
    const cell = cellMap.get(`${row}:${col}`);
    const address = cell ? cell.address : `${colToLetters(col)}${row}`;
    const merged = findMerge(merges, row, col) !== null;
    if (!cell) {
        return { address, row, col, raw: "", display: "", valueType: "", merged, formula: null, hyperlink: null };
    }
    return {
        address,
        row,
        col,
        raw: cell.rawValue,
        display: cell.outputValue,
        valueType: cell.valueType,
        merged,
        formula:
            cell.formulaText && cell.formulaText.length > 0
                ? { text: cell.formulaText, status: cell.resolutionStatus, source: cell.resolutionSource }
                : null,
        hyperlink: toIrHyperlink(cell.hyperlink),
    };
}

/** Internal links keep their location (Sheet2!A1). External links keep their URL. */
function toIrHyperlink(hyperlink: CellHyperlink | null): IrHyperlink | null {
    if (!hyperlink) return null;
    const target = hyperlink.kind === "internal" ? hyperlink.location || hyperlink.target : hyperlink.target;
    if (!String(target || "").trim()) return null;
    return { kind: hyperlink.kind, target: target.trim(), tooltip: hyperlink.tooltip || "" };
}

/**
 * Renders narrative cell text without Markdown escaping. Links match compact table cells.
 * External links use [label](url). Internal links use label (Sheet2!A1).
 */
function toNarrativeCellText(cell: ParsedCell): string {
    const label = String(cell.outputValue || "")
        .replace(/\s+/g, " ")
        .trim();
    const link = toIrHyperlink(cell.hyperlink);
    if (!link || !label) return label;
    const linkTarget = link.kind === "external" ? toMarkdownLinkTarget(link.target) : null;
    return linkTarget ? `[${label}](${linkTarget})` : `${label} (${link.target})`;
}

/** Asset ids follow array order, like the Markdown Chart/Image/Shape headings. */
function assetId(prefix: string, index: number): string {
    return `${prefix}${String(index + 1).padStart(3, "0")}`;
}

/** tableIds holds the emitted tables only, so a source table always has an id. */
function toIrCharts(sheet: ParsedSheet, tableIds: Map<TableCandidate, string>): IrChart[] {
    const tables = [...tableIds.keys()];
    return (sheet.charts || []).map((chart, index) => {
        const link = findChartSourceTables(sheet.name, chart, tables);
        return {
            id: assetId("c", index),
            anchor: chart.anchor,
            title: chart.title,
            type: chart.chartType,
            series: chart.series.map((series, seriesIndex) => {
                const table = link.seriesTables[seriesIndex];
                return {
                    name: series.name,
                    categories: series.categoriesRef,
                    values: series.valuesRef,
                    cachedCategories: [...(series.cachedCategories || [])],
                    cachedValues: [...(series.cachedValues || [])],
                    sourceTable: table ? tableIds.get(table) || null : null,
                    axis: series.axis,
                };
            }),
        };
    });
}

function toIrImages(sheet: ParsedSheet): IrImage[] {
    return (sheet.images || []).map((image, index) => ({
        id: assetId("i", index),
        anchor: image.anchor,
        filename: image.filename,
        path: image.path,
    }));
}

function toIrShapes(sheet: ParsedSheet): IrShape[] {
    const shapes = sheet.shapes || [];
    // Same resolver as the Markdown diagram, so both outputs report the same connections.
    const connectorEnds = resolveConnectorEnds(shapes);
    const shapeId = (index: number | null) => (index === null ? null : assetId("sh", index));
    const targets = new Map<number, IrShapeTarget[]>();
    for (const target of resolveShapeTargets(sheet, colToLetters)) {
        const entry: IrShapeTarget = {
            role: target.role,
            cell: target.cell,
            textAddress: target.textAddress,
            text: target.text,
            method: target.method,
            ...(target.targetShapeIndex !== null ? { shape: assetId("sh", target.targetShapeIndex) } : {}),
            ...(target.targetImageIndex !== null ? { image: assetId("i", target.targetImageIndex) } : {}),
        };
        const list = targets.get(target.shapeIndex);
        if (list) list.push(entry);
        else targets.set(target.shapeIndex, [entry]);
    }
    return shapes.map((shape, index) => {
        const ends = connectorEnds.get(index);
        const shapeTargets = targets.get(index);
        return {
            id: assetId("sh", index),
            anchor: shape.anchor,
            kind: shape.kind,
            name: shape.name,
            text: String(shape.text || "").trim(),
            ...(shape.markedText ? { markedText: shape.markedText } : {}),
            ...(shape.look ? { look: shape.look } : {}),
            connection: ends ? { from: shapeId(ends.from), to: shapeId(ends.to), arrow: ends.arrow } : null,
            ...(shapeTargets ? { targets: shapeTargets } : {}),
        };
    });
}

/**
 * Converts one table candidate to an IR table with its kept cell grid.
 * Returns null for a table without values, which Markdown also skips.
 */
function toIrTable(
    sheet: ParsedSheet,
    cellMap: Map<string, ParsedCell>,
    candidate: TableCandidate,
    tableNumber: number,
    options: AiIrOptions
): IrTable | null {
    const layout = planTableLayout(candidate, options, (row, col) => hasValue(cellMap, sheet.merges, row, col));
    if (layout.rows.length === 0 || layout.cols.length === 0) {
        return null;
    }
    const rows = layout.rows.map((row) => layout.cols.map((col) => toIrCell(cellMap, sheet.merges, row, col)));
    const range = formatRange(candidate.startRow, candidate.startCol, candidate.endRow, candidate.endCol);
    const hasHeader = candidate.hasHeaderRow !== false;
    return {
        id: `t${String(tableNumber).padStart(3, "0")}`,
        range,
        ...(candidate.name ? { name: candidate.name } : {}),
        hasHeader,
        headers: hasHeader ? rows[0].map((cell) => cell.display) : [],
        rows,
        source: { sheet: sheet.name, range, anchor: "" },
    };
}

/** Converts one narrative block to its IR form. */
function toIrNarrative(block: NarrativeBlock): IrNarrative {
    return {
        startRow: block.startRow,
        lines: block.lines.slice(),
        items: block.items.map((item) => ({ row: item.row, text: item.text })),
    };
}

/** Converts one sheet to its IR form in a single pass. */
function toIrSheet(sheet: ParsedSheet, options: AiIrOptions = {}): IrSheet {
    const candidates = detectTableCandidates(
        sheet,
        (entry) => buildCellMap(entry.cells),
        undefined,
        options.tableDetectionMode || "balanced"
    );
    const cellMap = buildCellMap(sheet.cells);
    const narrativeBlocks = extractNarrativeBlocks(sheet, candidates, toNarrativeCellText);
    // Ids follow reading order, matching the Markdown "Table: NNN" numbering.
    const tables: IrTable[] = [];
    const tableIds = new Map<TableCandidate, string>();
    for (const candidate of orderTablesForReading(sheet, candidates, narrativeBlocks)) {
        const table = toIrTable(sheet, cellMap, candidate, tables.length + 1, options);
        if (!table) continue;
        tables.push(table);
        tableIds.set(candidate, table.id);
    }
    const narratives = narrativeBlocks.map((block) => toIrNarrative(block));
    const sections = planSheetReadingOrder(sheet, candidates, narrativeBlocks)
        .map((section) =>
            section.flatMap((entry): IrSectionEntry[] => {
                if (entry.kind === "narrative")
                    return [{ kind: "narrative", index: narrativeBlocks.indexOf(entry.block) }];
                const id = tableIds.get(entry.table);
                return id ? [{ kind: "table", id }] : [];
            })
        )
        .filter((section) => section.length > 0);
    const comments = sheet.comments
        .slice()
        .sort((left, right) => {
            const a = parseCellAddress(left.address);
            const b = parseCellAddress(right.address);
            return a.row - b.row || a.col - b.col;
        })
        .map((comment) => ({
            address: comment.address,
            kind: comment.kind,
            author: comment.author,
            text: comment.text,
        }));
    return {
        name: sheet.name,
        index: sheet.index,
        tables,
        narratives,
        sections,
        comments,
        charts: toIrCharts(sheet, tableIds),
        images: toIrImages(sheet),
        shapes: toIrShapes(sheet),
        state: sheet.state || "visible",
        hiddenRows: sheet.hiddenRows || [],
        hiddenCols: sheet.hiddenCols || [],
    };
}

/**
 * Builds the workbook IR from parsed structures.
 * Reads each sheet once and derives tables, narratives, and comments.
 * Never parses rendered Markdown back into structure.
 */
function buildWorkbookIR(workbook: ParsedWorkbook, options: AiIrOptions = {}): WorkbookIR {
    return {
        workbook: workbook.name,
        sheets: workbook.sheets.map((sheet) => toIrSheet(sheet, options)),
    };
}

/** Builds workbook IR from parsed sheets and tables. */
export { buildWorkbookIR };
