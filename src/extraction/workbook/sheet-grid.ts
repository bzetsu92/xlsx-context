import type { GridSpan, SheetGrid } from "../../model/worksheet.js";
import { xmlToDocument } from "../xml/xml-utils.js";

const EMU_PER_PX = 9525;
const EMU_PER_PT = 12700;
const MAX_COL = 16384;
const MAX_ROW = 1048576;

/** Font of the Normal style. Column widths count in digits of this font. */
type NormalFont = { name: string; sizePt: number };

// Maximum digit width in pixels at 96 DPI, measured in Excel. Keys are NFKC names in lower case without spaces.
const KNOWN_DIGIT_WIDTH_PX: Record<string, Record<number, number>> = {
    calibri: { 11: 7 },
    arial: { 10: 7 },
    mspゴシック: { 11: 8 },
    msゴシック: { 11: 8 },
    mspgothic: { 11: 8 },
    msgothic: { 11: 8 },
    游ゴシック: { 11: 8 },
    yugothic: { 11: 8 },
};

const DEFAULT_NORMAL_FONT: NormalFont = { name: "Calibri", sizePt: 11 };
const DEFAULT_ROW_HEIGHT_PT = 15;

function fontKey(name: string): string {
    return name.normalize("NFKC").toLowerCase().replace(/\s+/g, "");
}

/**
 * Returns the maximum digit width of the font.
 * Unknown fonts use half the em size. CJK fonts have half-width digits, and most Latin fonts are close to that.
 */
function digitWidthPx(font: NormalFont): { px: number; known: boolean } {
    const known = KNOWN_DIGIT_WIDTH_PX[fontKey(font.name)]?.[font.sizePt];
    if (known) return { px: known, known: true };
    const emPx = Math.round((font.sizePt * 96) / 72);
    return { px: Math.max(1, Math.round(emPx * 0.5)), known: false };
}

/** Converts a stored column width in characters to pixels. This is the formula in ECMA-376 Part 1, 18.3.1.13. */
function columnWidthPx(width: number, digitPx: number): number {
    return Math.trunc(((256 * width + Math.trunc(128 / digitPx)) / 256) * digitPx);
}

/** Excel adds 5 px of padding to the base width and rounds the result up to a multiple of 8 px. */
function baseColumnWidthPx(baseChars: number, digitPx: number): number {
    return Math.ceil((baseChars * digitPx + 5) / 8) * 8;
}

function readNumber(node: Element | null | undefined, name: string): number | null {
    const raw = node?.getAttribute(name);
    if (raw === null || raw === undefined || raw.trim() === "") return null;
    const value = Number(raw);
    return Number.isFinite(value) && value >= 0 ? value : null;
}

/** OOXML writers use both "1" and "true". */
function readFlag(node: Element | null | undefined, name: string): boolean {
    const raw = node?.getAttribute(name);
    return raw === "1" || raw === "true";
}

const normalFontCache = new WeakMap<Map<string, Uint8Array>, NormalFont>();

/** Reads the first font in styles.xml. Excel uses font 0 for the Normal style. */
function readNormalFont(files: Map<string, Uint8Array>, decodeXmlText: (bytes: Uint8Array) => string): NormalFont {
    const cached = normalFontCache.get(files);
    if (cached) return cached;
    let font = DEFAULT_NORMAL_FONT;
    const bytes = files.get("xl/styles.xml");
    // Read font 0 inside fonts. Differential formats (dxfs) also hold font elements.
    const fontsNode = bytes ? xmlToDocument(decodeXmlText(bytes)).getElementsByTagName("fonts")[0] : undefined;
    const first = fontsNode?.getElementsByTagName("font")[0];
    if (first) {
        const sizePt = readNumber(first.getElementsByTagName("sz")[0], "val");
        font = {
            name: first.getElementsByTagName("name")[0]?.getAttribute("val") || DEFAULT_NORMAL_FONT.name,
            sizePt: sizePt && sizePt > 0 ? sizePt : DEFAULT_NORMAL_FONT.sizePt,
        };
    }
    normalFontCache.set(files, font);
    return font;
}

/**
 * Merges raw spans into sorted spans that do not overlap.
 * A later span wins where spans overlap, because Excel applies col elements in document order.
 */
function normalizeSpans(raw: GridSpan[], maxIndex: number): GridSpan[] {
    const sizes = new Map<number, number>();
    const ranges: GridSpan[] = [];
    for (const span of raw) {
        const start = Math.max(1, Math.trunc(span.start));
        const end = Math.min(maxIndex, Math.trunc(span.end));
        if (end < start) continue;
        ranges.push({ start, end, sizeEmu: span.sizeEmu });
    }
    if (ranges.every((span, index) => index === 0 || span.start > (ranges[index - 1] as GridSpan).end)) {
        return mergeEqualNeighbors(ranges.sort((a, b) => a.start - b.start));
    }
    // Overlap is rare. Resolve it per index, which is safe because overlapping writers use short ranges.
    for (const span of ranges) {
        for (let index = span.start; index <= span.end; index += 1) sizes.set(index, span.sizeEmu);
    }
    const sorted = [...sizes.entries()].sort((a, b) => a[0] - b[0]);
    return mergeEqualNeighbors(sorted.map(([index, sizeEmu]) => ({ start: index, end: index, sizeEmu })));
}

function mergeEqualNeighbors(spans: GridSpan[]): GridSpan[] {
    const merged: GridSpan[] = [];
    for (const span of spans) {
        const last = merged[merged.length - 1];
        if (last && last.end + 1 === span.start && last.sizeEmu === span.sizeEmu) last.end = span.end;
        else merged.push({ ...span });
    }
    return merged;
}

/**
 * Builds the grid from worksheet parts.
 *
 * @param formatPr The sheetFormatPr element, or null.
 * @param colNodes The col elements in document order.
 * @param rowNodes The row elements in document order.
 */
function buildSheetGrid(
    formatPr: Element | null,
    colNodes: Element[],
    rowNodes: Element[],
    font: NormalFont = DEFAULT_NORMAL_FONT
): SheetGrid {
    const digit = digitWidthPx(font);
    const defaultColChars = readNumber(formatPr, "defaultColWidth");
    const defaultColPx =
        defaultColChars !== null
            ? columnWidthPx(defaultColChars, digit.px)
            : baseColumnWidthPx(readNumber(formatPr, "baseColWidth") ?? 8, digit.px);
    const defaultColWidthEmu = defaultColPx * EMU_PER_PX;

    const colSpans: GridSpan[] = [];
    for (const node of colNodes) {
        const min = readNumber(node, "min");
        const max = readNumber(node, "max") ?? min;
        if (min === null || max === null) continue;
        const width = readNumber(node, "width");
        const sizeEmu = readFlag(node, "hidden")
            ? 0
            : width !== null
              ? columnWidthPx(width, digit.px) * EMU_PER_PX
              : defaultColWidthEmu;
        colSpans.push({ start: min, end: max, sizeEmu });
    }

    // A listed row without ht keeps the declared default height, even when zeroHeight hides unlisted rows.
    const declaredRowPt = readNumber(formatPr, "defaultRowHeight") ?? DEFAULT_ROW_HEIGHT_PT;
    const defaultRowHeightEmu = readFlag(formatPr, "zeroHeight") ? 0 : Math.round(declaredRowPt * EMU_PER_PT);
    const rowSpans: GridSpan[] = [];
    let previousRow = 0;
    for (const node of rowNodes) {
        // The r attribute is optional. A row without it follows the previous row.
        const row = readNumber(node, "r") ?? previousRow + 1;
        previousRow = row;
        const height = readNumber(node, "ht");
        const sizeEmu = readFlag(node, "hidden") ? 0 : Math.round((height ?? declaredRowPt) * EMU_PER_PT);
        if (sizeEmu !== defaultRowHeightEmu) rowSpans.push({ start: row, end: row, sizeEmu });
    }

    return {
        defaultColWidthEmu,
        defaultRowHeightEmu,
        colSpans: normalizeSpans(colSpans, MAX_COL),
        rowSpans: normalizeSpans(rowSpans, MAX_ROW),
        confidence: digit.known ? "EXACT" : "ESTIMATED",
    };
}

/** Builds a grid with one size for every cell. Used when the worksheet gives no grid. */
function createUniformGrid(colWidthEmu: number, rowHeightEmu: number): SheetGrid {
    return {
        defaultColWidthEmu: colWidthEmu,
        defaultRowHeightEmu: rowHeightEmu,
        colSpans: [],
        rowSpans: [],
        confidence: "ESTIMATED",
    };
}

type SpanIndex = { firsts: number[]; starts: number[] };

// First index and start offset of each span. Computed once per span list.
const spanIndexCache = new WeakMap<GridSpan[], SpanIndex>();

function indexSpans(spans: GridSpan[], defaultSize: number): SpanIndex {
    const cached = spanIndexCache.get(spans);
    if (cached) return cached;
    const firsts: number[] = [];
    const starts: number[] = [];
    let offset = 0;
    let next = 1;
    for (const span of spans) {
        offset += (span.start - next) * defaultSize;
        firsts.push(span.start);
        starts.push(offset);
        offset += (span.end - span.start + 1) * span.sizeEmu;
        next = span.end + 1;
    }
    const index = { firsts, starts };
    spanIndexCache.set(spans, index);
    return index;
}

/** Finds the last index whose value is at most the target. Returns -1 when none is. */
function lastAtMost(values: number[], target: number): number {
    let low = 0;
    let high = values.length - 1;
    let found = -1;
    while (low <= high) {
        const mid = (low + high) >> 1;
        if ((values[mid] as number) <= target) {
            found = mid;
            low = mid + 1;
        } else {
            high = mid - 1;
        }
    }
    return found;
}

/** Returns the start offset of a 1-based column or row. */
function edgeOffset(spans: GridSpan[], defaultSize: number, index: number): number {
    const { firsts, starts } = indexSpans(spans, defaultSize);
    const at = lastAtMost(firsts, index);
    if (at < 0) return (index - 1) * defaultSize;
    const span = spans[at] as GridSpan;
    const start = starts[at] as number;
    if (index <= span.end) return start + (index - span.start) * span.sizeEmu;
    return start + (span.end - span.start + 1) * span.sizeEmu + (index - span.end - 1) * defaultSize;
}

/** Returns the 1-based column or row that holds an offset. Offsets before the grid map to index 1. */
function indexAtOffset(spans: GridSpan[], defaultSize: number, offset: number, maxIndex: number): number {
    if (offset <= 0) return 1;
    const { starts } = indexSpans(spans, defaultSize);
    const at = lastAtMost(starts, offset);
    let base = 0;
    let next = 1;
    if (at >= 0) {
        const span = spans[at] as GridSpan;
        const start = starts[at] as number;
        const spanEnd = start + (span.end - span.start + 1) * span.sizeEmu;
        if (offset < spanEnd) return span.start + Math.floor((offset - start) / span.sizeEmu);
        base = spanEnd;
        next = span.end + 1;
    }
    // Hidden default cells have no size. The offset then sits after all of them.
    if (defaultSize <= 0) return Math.min(maxIndex, next);
    return Math.min(maxIndex, next + Math.floor((offset - base) / defaultSize));
}

/** Returns the left edge of a 1-based column. */
function columnLeftEmu(grid: SheetGrid, col: number): number {
    return edgeOffset(grid.colSpans, grid.defaultColWidthEmu, col);
}

/** Returns the top edge of a 1-based row. */
function rowTopEmu(grid: SheetGrid, row: number): number {
    return edgeOffset(grid.rowSpans, grid.defaultRowHeightEmu, row);
}

/** Returns the 1-based column at a horizontal offset. */
function columnAtEmu(grid: SheetGrid, x: number): number {
    return indexAtOffset(grid.colSpans, grid.defaultColWidthEmu, x, MAX_COL);
}

/** Returns the 1-based row at a vertical offset. */
function rowAtEmu(grid: SheetGrid, y: number): number {
    return indexAtOffset(grid.rowSpans, grid.defaultRowHeightEmu, y, MAX_ROW);
}

/** Returns a 0-based column position with a fraction inside the column. Hidden columns take no room. */
function columnPositionAt(grid: SheetGrid, x: number): number {
    const col = columnAtEmu(grid, x);
    const left = columnLeftEmu(grid, col);
    const width = columnLeftEmu(grid, col + 1) - left;
    return col - 1 + (width > 0 ? Math.min(1, Math.max(0, (x - left) / width)) : 0);
}

/** Returns a 0-based row position with a fraction inside the row. Hidden rows take no room. */
function rowPositionAt(grid: SheetGrid, y: number): number {
    const row = rowAtEmu(grid, y);
    const top = rowTopEmu(grid, row);
    const height = rowTopEmu(grid, row + 1) - top;
    return row - 1 + (height > 0 ? Math.min(1, Math.max(0, (y - top) / height)) : 0);
}

export {
    buildSheetGrid,
    columnPositionAt,
    columnAtEmu,
    columnLeftEmu,
    columnWidthPx,
    createUniformGrid,
    digitWidthPx,
    readNormalFont,
    rowAtEmu,
    rowPositionAt,
    rowTopEmu,
};
export type { NormalFont };
