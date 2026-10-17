import type {
    MarkdownFile,
    MarkdownOptions,
    NarrativeBlock,
    ParsedCell,
    ParsedChartAsset,
    ParsedShapeAsset,
    ParsedSheet,
    ParsedWorkbook,
    SectionBlock,
    SheetGrid,
    TableCandidate,
} from "../../model/index.js";
import { resolveMarkdownOptions } from "./markdown-options.js";
import type { ChartSource, ShapeBlock } from "./asset-sections.js";
import { findChartSourceTables } from "../../structure/layout/chart-sources.js";
import { resolveShapeTargets } from "../../structure/layout/shape-targets.js";
import { collectNarrativeCellsByRow, isCellInAnyTable } from "../../structure/layout/narrative-blocks.js";
import { describeHiddenContent } from "../../structure/layout/sheet-visibility.js";
import {
    extractSectionBlocks,
    groupLayoutEntries,
    orderTablesForReading,
} from "../../structure/layout/sheet-sections.js";
import { createAssetSections } from "./asset-sections.js";
import { createCellMarkdown } from "./cell-markdown.js";
import { createContentSections, type ContentSection, type GroupedSection } from "./table-sections.js";

type SheetRenderState = {
    resolvedOptions: ReturnType<typeof resolveMarkdownOptions>;
    charts: ParsedChartAsset[];
    shapes: ParsedShapeAsset[];
    shapeBlocks: ShapeBlock[];
    tables: TableCandidate[];
    narrativeBlocks: NarrativeBlock[];
    sectionBlocks: SectionBlock[];
    formulaDiagnostics: MarkdownFile["summary"]["formulaDiagnostics"];
    sections: ContentSection[];
    groupedSections: GroupedSection[];
    body: string;
    commentSection: string;
};

export type SheetMarkdownDeps = {
    renderNarrativeBlock: (block: NarrativeBlock) => string;
    detectTableCandidates: (
        sheet: ParsedSheet,
        buildCellMap: (sheet: ParsedSheet) => Map<string, ParsedCell>,
        tableDetectionMode?: "balanced" | "border" | "planner-aware"
    ) => TableCandidate[];
    matrixFromCandidate: (
        sheet: ParsedSheet,
        candidate: TableCandidate,
        options: MarkdownOptions,
        buildCellMap: (sheet: ParsedSheet) => Map<string, ParsedCell>,
        formatCellForMarkdown: (cell: ParsedCell | undefined, options: MarkdownOptions) => string
    ) => string[][];
    renderMarkdownTable: (rows: string[][], treatFirstRowAsHeader: boolean) => string;
    createOutputFileName: (workbookName: string, sheetIndex: number, sheetName: string) => string;
    extractShapeBlocks: (
        shapes: ParsedShapeAsset[],
        options: {
            defaultCellWidthEmu: number;
            defaultCellHeightEmu: number;
            shapeBlockGapXEmu: number;
            shapeBlockGapYEmu: number;
            grid?: SheetGrid;
        }
    ) => ShapeBlock[];
    renderHierarchicalRawEntries: (entries: { key: string; value: string }[]) => string[];
    parseCellAddress: (address: string) => { row: number; col: number };
    formatRange: (startRow: number, startCol: number, endRow: number, endCol: number) => string;
    colToLetters: (col: number) => string;
    normalizeMarkdownText?: (text: string) => string;
    defaultCellWidthEmu: number;
    defaultCellHeightEmu: number;
    shapeBlockGapXEmu: number;
    shapeBlockGapYEmu: number;
};

function createSheetMarkdown(deps: SheetMarkdownDeps) {
    const cells = createCellMarkdown({ normalizeMarkdownText: deps.normalizeMarkdownText });
    const content = createContentSections({
        ...deps,
        formatCellForMarkdown: cells.formatCellForMarkdown,
        createMarkdownNarrativeFormatter: cells.createMarkdownNarrativeFormatter,
        buildCellMap: cells.buildCellMap,
    });
    const assets = createAssetSections(deps);

    /** Links each chart to the rendered tables that hold its series, by their Markdown labels. */
    function createChartSources(
        sheet: ParsedSheet,
        charts: ParsedChartAsset[],
        tableSections: ContentSection[]
    ): ChartSource[] {
        const labels = new Map<TableCandidate, string>();
        tableSections.forEach((section, index) => {
            const table = section.table;
            if (!table) return;
            const range = deps.formatRange(table.startRow, table.startCol, table.endRow, table.endCol);
            labels.set(table, `Table ${String(index + 1).padStart(3, "0")} (${range})`);
        });
        return charts.map((chart) => {
            const link = findChartSourceTables(sheet.name, chart, [...labels.keys()]);
            return { tableLabels: link.tables.map((table) => labels.get(table) || ""), complete: link.complete };
        });
    }

    function collectSheetRenderState(
        workbook: ParsedWorkbook,
        sheet: ParsedSheet,
        options: MarkdownOptions = {}
    ): SheetRenderState {
        const resolvedOptions = cells.resolveCachedMarkdownOptions(options);
        // Builds one cell map per render. Detection and every table matrix share it.
        const cellMap = cells.buildCellMap(sheet);
        const sheetCellMap = () => cellMap;
        const charts = sheet.charts || [];
        const shapes = sheet.shapes || [];
        const shapeBlocks = deps.extractShapeBlocks(shapes, {
            defaultCellWidthEmu: deps.defaultCellWidthEmu,
            defaultCellHeightEmu: deps.defaultCellHeightEmu,
            shapeBlockGapXEmu: deps.shapeBlockGapXEmu,
            shapeBlockGapYEmu: deps.shapeBlockGapYEmu,
            grid: sheet.grid,
        });
        const treatFirstRowAsHeader = resolvedOptions.treatFirstRowAsHeader;
        const tableDetectionMode = resolvedOptions.tableDetectionMode;
        const tables = deps.detectTableCandidates(sheet, sheetCellMap, tableDetectionMode);
        const narrativeBlocks = content.extractNarrativeBlocks(workbook, sheet, tables, resolvedOptions);
        const formulaDiagnostics = assets.createFormulaDiagnostics(sheet);
        // Tables are numbered in reading order, as the AI IR numbers them.
        const tablesInReadingOrder = orderTablesForReading(sheet, tables, narrativeBlocks);
        const tableSections = content.createTableSections(
            workbook,
            sheet,
            tablesInReadingOrder,
            resolvedOptions,
            treatFirstRowAsHeader,
            sheetCellMap
        );
        const chartSources = createChartSources(sheet, charts, tableSections);
        const shapeTargets = resolveShapeTargets(sheet, deps.colToLetters, {
            widthEmu: deps.defaultCellWidthEmu,
            heightEmu: deps.defaultCellHeightEmu,
        });
        const assetEntries = assets.createAssetEntries(sheet, charts, shapes, shapeBlocks, {
            includeShapeDetails: resolvedOptions.includeShapeDetails,
            chartSources,
            shapeTargets,
        });
        const sections: ContentSection[] = [
            ...content.createNarrativeSections(narrativeBlocks),
            ...tableSections,
            ...assetEntries,
        ];
        const { sectionBlocks, groupedSections } = groupLayoutEntries(sheet, tables, narrativeBlocks, sections);
        const body = content.renderGroupedSectionBody(groupedSections);
        const commentSection = assets.renderCommentSection(sheet.comments || []);
        return {
            resolvedOptions,
            charts,
            shapes,
            shapeBlocks,
            tables,
            narrativeBlocks,
            sectionBlocks,
            formulaDiagnostics,
            sections,
            groupedSections,
            body,
            commentSection,
        };
    }

    function createSheetMarkdownText(workbook: ParsedWorkbook, sheet: ParsedSheet, state: SheetRenderState): string {
        const hiddenNote = describeHiddenContent(sheet);
        const markdown = [
            `# Book: ${workbook.name}`,
            "",
            `## Sheet: ${sheet.name}`,
            "",
            ...(hiddenNote ? [`> ${hiddenNote}`, ""] : []),
            state.body || "_No extractable body content was found._",
            state.commentSection,
        ].join("\n");
        return markdown;
    }

    function createSheetSummary(sheet: ParsedSheet, state: SheetRenderState): MarkdownFile["summary"] {
        return {
            outputMode: state.resolvedOptions.outputMode,
            formattingMode: state.resolvedOptions.formattingMode,
            tableDetectionMode: state.resolvedOptions.tableDetectionMode,
            sections: state.sectionBlocks.length,
            tables: state.tables.length,
            narrativeBlocks: state.narrativeBlocks.length,
            merges: sheet.merges.length,
            images: sheet.images.length,
            charts: state.charts.length,
            comments: (sheet.comments || []).length,
            cells: sheet.cells.length,
            tableScores: state.tables.map((table) => ({
                range: deps.formatRange(table.startRow, table.startCol, table.endRow, table.endCol),
                score: table.score,
                reasons: [...table.reasonSummary],
            })),
            formulaDiagnostics: state.formulaDiagnostics,
        };
    }

    /** Converts one parsed sheet to a Markdown file with summary. */
    function convertSheetToMarkdown(
        workbook: ParsedWorkbook,
        sheet: ParsedSheet,
        options: MarkdownOptions = {}
    ): MarkdownFile {
        const state = collectSheetRenderState(workbook, sheet, options);
        const fileName = deps.createOutputFileName(workbook.name, sheet.index, sheet.name);
        return {
            fileName,
            sheetName: sheet.name,
            markdown: createSheetMarkdownText(workbook, sheet, state),
            summary: createSheetSummary(sheet, state),
        };
    }

    /** Converts every sheet in workbook order. */
    function convertWorkbookToMarkdownFiles(workbook: ParsedWorkbook, options: MarkdownOptions = {}): MarkdownFile[] {
        return workbook.sheets.map((sheet) => convertSheetToMarkdown(workbook, sheet, options));
    }

    return {
        buildCellMap: cells.buildCellMap,
        formatCellForMarkdown: cells.formatCellForMarkdown,
        isCellInAnyTable,
        splitNarrativeRowSegments: content.splitNarrativeRowSegments,
        collectNarrativeCellsByRow,
        buildNarrativeItems: content.buildNarrativeItems,
        extractNarrativeBlocks: content.extractNarrativeBlocks,
        extractSectionBlocks,
        sortContentSections: content.sortContentSections,
        createGroupedSections: content.createGroupedSections,
        createCalendarAwareSectionEntries: content.createCalendarAwareSectionEntries,
        renderGroupedSectionBody: content.renderGroupedSectionBody,
        collectSheetRenderState,
        createSheetMarkdownText,
        createSheetSummary,
        createAssetEntries: assets.createAssetEntries,
        renderCommentSection: assets.renderCommentSection,
        convertSheetToMarkdown,
        convertWorkbookToMarkdownFiles,
    };
}

/** Converts parsed sheets to Markdown files. */
export { createSheetMarkdown };
