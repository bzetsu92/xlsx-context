/** Sheet sections and reading order from anchors, blocks, and tables. */
import type { NarrativeBlock, ParsedSheet, SectionBlock, TableCandidate } from "../../model/index.js";
import { parseCellAddress } from "../addresses/address-utils.js";
import { isMonthTitleToken } from "../formatting/calendar-patterns.js";
import { isIsoDateToken, isWeekdayToken } from "../formatting/date-tokens.js";
import type { GroupedEntries, LayoutEntry, ReadingOrderEntry } from "./sheet-layout.js";
import { blockHasCalendarDateItem } from "./narrative-blocks.js";

type SectionAnchor = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    calendarLike?: boolean;
    /** Narrative block or table the anchor stands for. Images and charts have none. */
    source?: NarrativeBlock | TableCandidate;
};

/** Section index per narrative block or table, decided while the sections are built. */
type SectionMembership = Map<NarrativeBlock | TableCandidate, number>;

function createNarrativeSectionAnchors(narrativeBlocks: NarrativeBlock[]): SectionAnchor[] {
    return narrativeBlocks.map((block) => ({
        startRow: block.startRow,
        startCol: block.startCol,
        endRow: block.endRow,
        endCol: Math.max(block.startCol, ...block.items.map((item) => item.startCol)),
        calendarLike: blockHasCalendarDateItem(block),
        source: block,
    }));
}

function createTableSectionAnchors(tables: TableCandidate[]): SectionAnchor[] {
    return tables.map((table) => ({
        startRow: table.startRow,
        startCol: table.startCol,
        endRow: table.endRow,
        endCol: table.endCol,
        source: table,
    }));
}

function createPointSectionAnchor(address: string): SectionAnchor | null {
    const anchor = parseCellAddress(address);
    if (anchor.row > 0 && anchor.col > 0) {
        return { startRow: anchor.row, startCol: anchor.col, endRow: anchor.row, endCol: anchor.col };
    }
    return null;
}

function createImageSectionAnchors(sheet: ParsedSheet): SectionAnchor[] {
    return sheet.images
        .map((image) => createPointSectionAnchor(image.anchor))
        .filter((anchor): anchor is SectionAnchor => !!anchor);
}

function createChartSectionAnchors(charts: NonNullable<ParsedSheet["charts"]>): SectionAnchor[] {
    return charts
        .map((chart) => createPointSectionAnchor(chart.anchor))
        .filter((anchor): anchor is SectionAnchor => !!anchor);
}

function createSectionAnchors(
    sheet: ParsedSheet,
    tables: TableCandidate[],
    narrativeBlocks: NarrativeBlock[]
): SectionAnchor[] {
    const charts = sheet.charts || [];
    return [
        ...createNarrativeSectionAnchors(narrativeBlocks),
        ...createTableSectionAnchors(tables),
        ...createImageSectionAnchors(sheet),
        ...createChartSectionAnchors(charts),
    ];
}

function sortSectionAnchors(anchors: SectionAnchor[]): SectionAnchor[] {
    return anchors.sort((left, right) => {
        if (left.startRow !== right.startRow) return left.startRow - right.startRow;
        return left.startCol - right.startCol;
    });
}

function createSectionBlockFromAnchor(anchor: SectionAnchor): SectionBlock {
    return {
        startRow: anchor.startRow,
        startCol: anchor.startCol,
        endRow: anchor.endRow,
        endCol: anchor.endCol,
    };
}

type OpenSection = { block: SectionBlock; lastAnchor: SectionAnchor };

/**
 * Measures the join distance from a section to an anchor.
 * Returns null when the anchor starts a new section.
 * Every open section is a candidate, so side-by-side tables keep their own captions.
 */
function sectionJoinDistance(section: OpenSection, anchor: SectionAnchor): number | null {
    const verticalGapThreshold = 4;
    const horizontalGapThreshold = 3;
    const verticalGap = anchor.startRow - section.block.endRow;
    const horizontalGap = Math.max(0, anchor.startCol - section.block.endCol, section.block.startCol - anchor.endCol);
    if (anchor.calendarLike && section.lastAnchor.calendarLike && verticalGap <= 8 && horizontalGap <= 24) {
        return horizontalGap;
    }
    if (verticalGap > verticalGapThreshold || horizontalGap > horizontalGapThreshold) {
        return null;
    }
    return horizontalGap;
}

function extendSectionBlock(section: SectionBlock, anchor: SectionAnchor): void {
    section.startRow = Math.min(section.startRow, anchor.startRow);
    section.startCol = Math.min(section.startCol, anchor.startCol);
    section.endRow = Math.max(section.endRow, anchor.endRow);
    section.endCol = Math.max(section.endCol, anchor.endCol);
}

function assignSectionBlocks(
    sheet: ParsedSheet,
    tables: TableCandidate[],
    narrativeBlocks: NarrativeBlock[]
): { blocks: SectionBlock[]; membership: SectionMembership } {
    const anchors = sortSectionAnchors(createSectionAnchors(sheet, tables, narrativeBlocks));
    const sections: OpenSection[] = [];
    const membership: SectionMembership = new Map();

    for (const anchor of anchors) {
        let bestIndex = -1;
        let bestDistance = Infinity;
        let bestVerticalGap = Infinity;
        sections.forEach((section, index) => {
            const distance = sectionJoinDistance(section, anchor);
            if (distance === null) return;
            const verticalGap = Math.max(0, anchor.startRow - section.block.endRow);
            if (distance < bestDistance || (distance === bestDistance && verticalGap < bestVerticalGap)) {
                bestIndex = index;
                bestDistance = distance;
                bestVerticalGap = verticalGap;
            }
        });
        if (bestIndex < 0) {
            sections.push({ block: createSectionBlockFromAnchor(anchor), lastAnchor: anchor });
            bestIndex = sections.length - 1;
        } else {
            extendSectionBlock(sections[bestIndex].block, anchor);
            sections[bestIndex].lastAnchor = anchor;
        }
        if (anchor.source) membership.set(anchor.source, bestIndex);
    }

    return { blocks: sections.map((section) => section.block), membership };
}

function extractSectionBlocks(
    sheet: ParsedSheet,
    tables: TableCandidate[],
    narrativeBlocks: NarrativeBlock[]
): SectionBlock[] {
    return assignSectionBlocks(sheet, tables, narrativeBlocks).blocks;
}

function sortLayoutEntries<T extends LayoutEntry>(sections: T[]): T[] {
    return sections.sort((left, right) => {
        if (left.sortRow !== right.sortRow) return left.sortRow - right.sortRow;
        return left.sortCol - right.sortCol;
    });
}

/** Catches all sections when the sheet has no section blocks. */
function createFallbackSectionBlock(): SectionBlock {
    return {
        startRow: -1,
        startCol: -1,
        endRow: Number.MAX_SAFE_INTEGER,
        endCol: Number.MAX_SAFE_INTEGER,
    };
}

function isSectionInsideBlock(section: LayoutEntry, block: SectionBlock): boolean {
    return (
        section.sortRow >= block.startRow &&
        section.sortRow <= block.endRow &&
        section.sortCol >= block.startCol &&
        section.sortCol <= block.endCol
    );
}

function createGroupedSections<T extends LayoutEntry>(
    sectionBlocks: SectionBlock[],
    sections: T[],
    membership?: SectionMembership
): GroupedEntries<T>[] {
    const blocks: SectionBlock[] = sectionBlocks.length > 0 ? sectionBlocks : [createFallbackSectionBlock()];
    const memberIndex = (section: T): number | undefined => {
        const source = section.narrativeBlock || section.table;
        return source && membership ? membership.get(source) : undefined;
    };
    // Blocks can overlap after extension. Each section joins one block only, else it renders twice.
    // Known membership wins. The first containing block is the fallback.
    const assigned = new Set<T>();
    const groups = blocks
        .map((block, blockIndex) => {
            const entries = sections.filter((section) => {
                if (assigned.has(section)) return false;
                const index = memberIndex(section);
                return index === undefined ? isSectionInsideBlock(section, block) : index === blockIndex;
            });
            entries.forEach((section) => assigned.add(section));
            return { block, entries };
        })
        .filter((group) => group.entries.length > 0);
    // An entry outside every block, for example a shape block, joins the nearest group by the anchor rule.
    // With no group near, it gets its own group at its row. It never drops out.
    for (const section of sections) {
        if (assigned.has(section)) continue;
        const point = {
            startRow: section.sortRow,
            startCol: section.sortCol,
            endRow: section.sortRow,
            endCol: section.sortCol,
        };
        let nearest: GroupedEntries<T> | null = null;
        let nearestDistance = Infinity;
        for (const group of groups) {
            const distance = sectionJoinDistance({ block: group.block, lastAnchor: point }, point);
            if (distance !== null && distance < nearestDistance) {
                nearest = group;
                nearestDistance = distance;
            }
        }
        if (nearest) {
            nearest.entries.push(section);
            sortLayoutEntries(nearest.entries);
            continue;
        }
        const group = {
            block: {
                startRow: section.sortRow,
                startCol: section.sortCol,
                endRow: section.sortRow,
                endCol: section.sortCol,
            },
            entries: [section],
        };
        const insertAt = groups.findIndex((other) => other.block.startRow > section.sortRow);
        groups.splice(insertAt < 0 ? groups.length : insertAt, 0, group);
    }
    return groups;
}

/** Sorts the entries and groups them into the sections of the sheet. */
function groupLayoutEntries<T extends LayoutEntry>(
    sheet: ParsedSheet,
    tables: TableCandidate[],
    narrativeBlocks: NarrativeBlock[],
    sections: T[]
): { sectionBlocks: SectionBlock[]; groupedSections: GroupedEntries<T>[] } {
    const { blocks, membership } = assignSectionBlocks(sheet, tables, narrativeBlocks);
    sortLayoutEntries(sections);
    return { sectionBlocks: blocks, groupedSections: createGroupedSections(blocks, sections, membership) };
}

/**
 * Plans the reading order shared by Markdown and the AI IR.
 * Lists sections in sheet order. Orders blocks and tables as Markdown renders them.
 */
function planSheetReadingOrder(
    sheet: ParsedSheet,
    tables: TableCandidate[],
    narrativeBlocks: NarrativeBlock[]
): ReadingOrderEntry[][] {
    const sections: LayoutEntry[] = [
        ...narrativeBlocks.map((block): LayoutEntry => ({
            sortRow: block.startRow,
            sortCol: block.startCol,
            kind: "narrative",
            narrativeBlock: block,
        })),
        ...tables.map((table): LayoutEntry => ({
            sortRow: table.startRow,
            sortCol: table.startCol,
            kind: "table",
            table,
        })),
    ];
    // The sidebar marker has no source, so it drops out of the reading order.
    const createSidebarEntry = (sortRow: number, sortCol: number): LayoutEntry => ({
        sortRow,
        sortCol,
        kind: "narrative",
    });
    return groupLayoutEntries(sheet, tables, narrativeBlocks, sections).groupedSections.map((group) =>
        createCalendarAwareSectionEntries(group.entries, createSidebarEntry).flatMap((entry): ReadingOrderEntry[] => {
            if (entry.table) return [{ kind: "table", table: entry.table }];
            if (entry.narrativeBlock) return [{ kind: "narrative", block: entry.narrativeBlock }];
            return [];
        })
    );
}

/** Lists tables in reading order. Appends tables the plan misses in detection order. */
function orderTablesForReading(
    sheet: ParsedSheet,
    tables: TableCandidate[],
    narrativeBlocks: NarrativeBlock[]
): TableCandidate[] {
    const ordered = planSheetReadingOrder(sheet, tables, narrativeBlocks)
        .flat()
        .flatMap((entry) => (entry.kind === "table" ? [entry.table] : []));
    return [...ordered, ...tables.filter((table) => !ordered.includes(table))];
}

function getNarrativeValues(section: LayoutEntry): string[] {
    return (
        section.narrativeBlock?.items
            .flatMap((item) => item.cellValues || [])
            .map((value) => String(value || "").trim())
            .filter(Boolean) || []
    );
}

/** Treats date-and-weekday grids as calendar bodies. */
function isCalendarBodySection(section: LayoutEntry): boolean {
    if (section.kind !== "narrative" || !section.narrativeBlock) return false;
    return section.narrativeBlock.items.some((item) => {
        const values = (item.cellValues || []).map((value) => String(value || "").trim()).filter(Boolean);
        return values.length >= 5 && values.every((value) => isIsoDateToken(value) || isWeekdayToken(value));
    });
}

/** Treats weekday rows and month titles as calendar headers. */
function isCalendarHeaderSection(section: LayoutEntry): boolean {
    if (section.kind !== "narrative" || !section.narrativeBlock) return false;
    const values = getNarrativeValues(section);
    if (values.length === 0 || values.length > 10) return false;
    const allWeekdays = values.length >= 5 && values.every((value) => isWeekdayToken(value));
    const hasMonthTitle = values.some((value) => isMonthTitleToken(value));
    return allWeekdays || hasMonthTitle;
}

/**
 * Moves calendar headers before the main date grid and side calendars after it.
 * createSidebarEntry builds the marker before each side calendar.
 */
function createCalendarAwareSectionEntries<T extends LayoutEntry>(
    entries: T[],
    createSidebarEntry: (sortRow: number, sortCol: number) => T
): T[] {
    const mainCalendarEntries = entries
        .filter((entry) => isCalendarBodySection(entry))
        .sort((left, right) => {
            if (left.sortCol !== right.sortCol) return left.sortCol - right.sortCol;
            return left.sortRow - right.sortRow;
        });
    if (mainCalendarEntries.length === 0) {
        return entries;
    }
    const main = mainCalendarEntries[0];
    const headerEntries = entries.filter(
        (entry) => entry !== main && isCalendarHeaderSection(entry) && entry.sortRow <= main.sortRow + 1
    );
    const sidebarEntries = entries.filter(
        (entry) =>
            entry !== main &&
            !headerEntries.includes(entry) &&
            isCalendarBodySection(entry) &&
            entry.sortCol >= main.sortCol + 10
    );
    const remainingEntries = entries.filter(
        (entry) => entry !== main && !headerEntries.includes(entry) && !sidebarEntries.includes(entry)
    );
    if (headerEntries.length === 0 && sidebarEntries.length === 0) {
        return entries;
    }
    const reordered: T[] = [];
    reordered.push(...headerEntries);
    reordered.push(main);
    if (sidebarEntries.length > 0) {
        const firstSidebar = sidebarEntries[0];
        reordered.push(createSidebarEntry(firstSidebar.sortRow - 0.1, firstSidebar.sortCol));
        reordered.push(...sidebarEntries);
    }
    reordered.push(...remainingEntries);
    return reordered;
}

export {
    createCalendarAwareSectionEntries,
    createGroupedSections,
    extractSectionBlocks,
    groupLayoutEntries,
    orderTablesForReading,
    planSheetReadingOrder,
    sortLayoutEntries,
};
