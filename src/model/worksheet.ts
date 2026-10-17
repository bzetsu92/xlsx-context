import type { MergeRange, ParsedCell } from "./cell.js";
import type { ParsedTable } from "./table.js";
/** Full parsed worksheet with cells, tables, and assets. */
export type ParsedSheet = {
    name: string;
    index: number;
    path: string;
    cells: ParsedCell[];
    merges: MergeRange[];
    tables: ParsedTable[];
    images: ParsedImageAsset[];
    charts: ParsedChartAsset[];
    shapes: ParsedShapeAsset[];
    comments: ParsedCellComment[];
    maxRow: number;
    maxCol: number;
    /** Column widths and row heights. Drawings use the same space. Missing on sheets built by hand. */
    grid?: SheetGrid;
    /** Visibility in Excel. Missing on sheets built by hand, which count as visible. */
    state?: SheetState;
    /** One-based rows that Excel hides. */
    hiddenRows?: number[];
    /** One-based columns that Excel hides. */
    hiddenCols?: number[];
};

/** Sheet visibility. A veryHidden sheet can only be shown again through VBA. */
export type SheetState = "visible" | "hidden" | "veryHidden";

/** Run of columns or rows with one size. Indexes are 1-based and inclusive. */
export type GridSpan = {
    start: number;
    end: number;
    sizeEmu: number;
};

/**
 * Column widths and row heights in EMU.
 * Spans are sorted and do not overlap. Indexes outside all spans use the default size.
 */
export type SheetGrid = {
    defaultColWidthEmu: number;
    defaultRowHeightEmu: number;
    colSpans: GridSpan[];
    rowSpans: GridSpan[];
    /** EXACT when the digit width of the normal font is known. ESTIMATED when it is approximate. */
    confidence: "EXACT" | "ESTIMATED";
};

/** Embedded image extracted from a drawing. */
export type ParsedImageAsset = {
    sheetName: string;
    filename: string;
    path: string;
    anchor: string;
    data: Uint8Array;
    mediaPath: string;
    /** Picture box in the same space as shape boxes. Missing on images built by hand. */
    bbox?: { left: number; top: number; right: number; bottom: number };
};

/** Chart series reference to a source range. */
export type ParsedChartSeries = {
    name: string;
    categoriesRef: string;
    valuesRef: string;
    /** Category labels from the chart cache. Empty when the chart stores no cache. */
    cachedCategories: string[];
    /** Values from the chart cache. Null marks a missing point. Empty when the chart stores no cache. */
    cachedValues: (number | null)[];
    axis: "primary" | "secondary";
};

/** Chart metadata extracted from a drawing. */
export type ParsedChartAsset = {
    sheetName: string;
    anchor: string;
    chartPath: string;
    title: string;
    chartType: string;
    series: ParsedChartSeries[];
};

/**
 * Shapes a connector joins, by drawing id. A null id marks a loose end.
 * Arrow tells which ends carry an arrowhead.
 */
export type ParsedShapeConnection = {
    startId: string | null;
    endId: string | null;
    arrow: "none" | "start" | "end" | "both";
    /** Line start in the same space as `bbox`. Resolves the start shape when `startId` is missing. */
    startPoint?: { x: number; y: number } | null;
    /** Line end in the same space as `bbox`. Resolves the end shape when `endId` is missing. */
    endPoint?: { x: number; y: number } | null;
};

/**
 * Visual style of a shape. Null means the file does not say, so the application default applies.
 * Colors are "#RRGGBB" after theme colors and color transforms resolve.
 */
export type ShapeLook = {
    /** "#RRGGBB", "none", or "picture". */
    fill: string | null;
    /** Line color as "#RRGGBB", or "none". */
    line: string | null;
    /** OOXML preset dash, for example "dash" or "sysDot", or "custom". Null draws a solid line. */
    dash: string | null;
    lineWidthEmu: number | null;
};

/** Raw shape metadata extracted from a drawing. */
export type ParsedShapeAsset = {
    sheetName: string;
    anchor: string;
    /** Drawing-local shape id (cNvPr id). Connectors refer to shapes by this id. */
    drawingId: string;
    /** Stable id scoped to one drawing file. Combines the drawing path, group path, and drawing id. */
    stableId?: string;
    /** Slash-joined names of ancestor group shapes. Null outside groups. */
    groupPath?: string | null;
    /** Anchor flavor that placed the shape. Unknown when the anchor is missing. */
    anchorKind?: string;
    /** Geometry trust level. Estimated values use default cell sizes. */
    geometryConfidence?: "EXACT" | "ESTIMATED" | "UNKNOWN";
    /** Rotation in degrees from the shape transform. Null when absent. */
    rotationDeg?: number | null;
    flipH?: boolean;
    flipV?: boolean;
    name: string;
    kind: string;
    text: string;
    /** Text with Markdown marks, for example ~~struck~~. Missing when it equals `text`. */
    markedText?: string;
    look?: ShapeLook;
    /** Set for connectors and line shapes only. */
    connection: ParsedShapeConnection | null;
    /** Point that a callout points at, in the same space as `bbox`. Null for other shapes. */
    calloutTip?: { x: number; y: number } | null;
    widthEmu: number | null;
    heightEmu: number | null;
    elementName: string;
    anchorElementName: string;
    rawEntries: {
        key: string;
        value: string;
    }[];
    bbox: {
        left: number;
        top: number;
        right: number;
        bottom: number;
    };
    svgFilename: string | null;
    svgPath: string | null;
    svgData: Uint8Array | null;
    /** Human-readable note when the SVG is approximate. Null for exact renders. */
    svgWarning?: string | null;
};

/** Note or threaded comment attached to a cell. */
export type ParsedCellComment = {
    address: string;
    kind: "note" | "threaded";
    author: string;
    text: string;
    dateTime: string;
};

/** Group of shapes that share one anchor region. */
export type ShapeBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    shapeIndexes: number[];
};

/** One prose item inside a narrative block. */
export type NarrativeItem = {
    row: number;
    startCol: number;
    text: string;
    cellValues: string[];
};

/** Prose block extracted outside table candidates. */
export type NarrativeBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    lines: string[];
    items: NarrativeItem[];
};

/** Rectangular section grouping for one sheet region. */
export type SectionBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
};
