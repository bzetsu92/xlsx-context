import type {
    MarkdownOptions,
    NarrativeBlock,
    NarrativeItem,
    ParsedCell,
    ParsedSheet,
    ParsedWorkbook,
    SectionBlock,
    TableCandidate,
} from "../../model/index.js";
import type { GroupedEntries, LayoutEntry } from "../../structure/layout/sheet-layout.js";
import type { NarrativeCellFormatter } from "../../structure/layout/narrative-blocks.js";
import {
    buildNarrativeItems as layoutBuildNarrativeItems,
    extractNarrativeBlocks as layoutExtractNarrativeBlocks,
    splitNarrativeRowSegments as layoutSplitNarrativeRowSegments,
} from "../../structure/layout/narrative-blocks.js";
import {
    createCalendarAwareSectionEntries as layoutCreateCalendarAwareSectionEntries,
    createGroupedSections as layoutCreateGroupedSections,
    sortLayoutEntries,
} from "../../structure/layout/sheet-sections.js";
import type { SheetMarkdownDeps } from "./sheet-markdown.js";

export type ContentSection = LayoutEntry & {
    markdown: string;
};

export type GroupedSection = GroupedEntries<ContentSection>;

/** Content section dependencies. Cell rendering and table helpers come from the caller. */
type TableSectionDeps = Pick<
    SheetMarkdownDeps,
    "renderNarrativeBlock" | "matrixFromCandidate" | "renderMarkdownTable" | "formatRange"
> & {
    formatCellForMarkdown: (
        cell: ParsedCell | undefined,
        options: MarkdownOptions,
        workbook: ParsedWorkbook | null,
        sheet: ParsedSheet | null
    ) => string;
    createMarkdownNarrativeFormatter: (
        options: MarkdownOptions,
        workbook: ParsedWorkbook | null,
        sheet: ParsedSheet | null
    ) => NarrativeCellFormatter;
    buildCellMap: (sheet: ParsedSheet) => Map<string, ParsedCell>;
};

/** Builds narrative and table content sections with reading order. */
export function createContentSections(deps: TableSectionDeps) {
    const { formatCellForMarkdown, createMarkdownNarrativeFormatter, buildCellMap } = deps;

    function splitNarrativeRowSegments(
        cells: ParsedCell[],
        options: MarkdownOptions,
        workbook: ParsedWorkbook | null = null,
        sheet: ParsedSheet | null = null
    ) {
        return layoutSplitNarrativeRowSegments(cells, createMarkdownNarrativeFormatter(options, workbook, sheet));
    }

    function buildNarrativeItems(
        workbook: ParsedWorkbook,
        sheet: ParsedSheet,
        tables: TableCandidate[],
        options: MarkdownOptions = {}
    ): NarrativeItem[] {
        return layoutBuildNarrativeItems(sheet, tables, createMarkdownNarrativeFormatter(options, workbook, sheet));
    }

    function extractNarrativeBlocks(
        workbook: ParsedWorkbook,
        sheet: ParsedSheet,
        tables: TableCandidate[],
        options: MarkdownOptions = {}
    ): NarrativeBlock[] {
        return layoutExtractNarrativeBlocks(sheet, tables, createMarkdownNarrativeFormatter(options, workbook, sheet));
    }

    const sortContentSections = (sections: ContentSection[]): ContentSection[] => sortLayoutEntries(sections);

    const createGroupedSections = (sectionBlocks: SectionBlock[], sections: ContentSection[]): GroupedSection[] =>
        layoutCreateGroupedSections(sectionBlocks, sections);

    function renderGroupedSectionEntries(entries: ContentSection[]): string {
        return createCalendarAwareSectionEntries(entries)
            .map((section) => section.markdown.trimEnd())
            .join("\n\n")
            .trim();
    }

    function createSidebarSection(sortRow: number, sortCol: number): ContentSection {
        return { sortRow, sortCol, markdown: "### Sidebar\n", kind: "narrative" };
    }

    function createCalendarAwareSectionEntries(entries: ContentSection[]): ContentSection[] {
        return layoutCreateCalendarAwareSectionEntries(entries, createSidebarSection);
    }

    function renderGroupedSectionBody(groupedSections: GroupedSection[]): string {
        return groupedSections
            .map((group) => renderGroupedSectionEntries(group.entries))
            .filter(Boolean)
            .join("\n\n---\n\n")
            .trim();
    }

    function createNarrativeSections(narrativeBlocks: NarrativeBlock[]): ContentSection[] {
        return narrativeBlocks.map((block) => ({
            sortRow: block.startRow,
            sortCol: block.startCol,
            markdown: `${deps.renderNarrativeBlock(block)}\n`,
            kind: "narrative",
            narrativeBlock: block,
        }));
    }

    function createTableSections(
        workbook: ParsedWorkbook,
        sheet: ParsedSheet,
        tables: TableCandidate[],
        options: MarkdownOptions,
        treatFirstRowAsHeader: boolean,
        sheetCellMap: (sheet: ParsedSheet) => Map<string, ParsedCell> = buildCellMap
    ): ContentSection[] {
        const sections: ContentSection[] = [];
        let tableCounter = 1;
        for (const table of tables) {
            const rows = deps.matrixFromCandidate(sheet, table, options, sheetCellMap, (cell, tableOptions) =>
                formatCellForMarkdown(cell, tableOptions, workbook, sheet)
            );
            if (rows.length === 0 || rows[0]?.length === 0) continue;
            const tableMarkdown = deps.renderMarkdownTable(rows, treatFirstRowAsHeader && table.hasHeaderRow !== false);
            sections.push({
                sortRow: table.startRow,
                sortCol: table.startCol,
                markdown: `### Table: ${String(tableCounter).padStart(3, "0")} (${deps.formatRange(table.startRow, table.startCol, table.endRow, table.endCol)})\n\n${table.name ? `- Excel table: ${table.name}\n\n` : ""}${tableMarkdown}\n`,
                kind: "table",
                table,
            });
            tableCounter += 1;
        }
        return sections;
    }

    return {
        splitNarrativeRowSegments,
        buildNarrativeItems,
        extractNarrativeBlocks,
        sortContentSections,
        createGroupedSections,
        createCalendarAwareSectionEntries,
        renderGroupedSectionBody,
        createNarrativeSections,
        createTableSections,
    };
}
