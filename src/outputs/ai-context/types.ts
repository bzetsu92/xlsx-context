import type { ShapeLook } from "../../model/index.js";

import type { FormulaResolutionStatus } from "../../model/index.js";

/** Counts tokens in text for one specific model or estimator. */
export type TokenCounter = {
    name: string;
    count: (text: string) => number;
};

/** Anything that turns text into a token id array, for example a model tokenizer's encoder. */
export type TextEncoderLike = {
    name: string;
    encode: (text: string) => ArrayLike<unknown>;
};

/** js-tiktoken encoding names. Declared here so public types do not depend on js-tiktoken. */
export type TiktokenEncodingName = "o200k_base" | "cl100k_base" | "p50k_base" | "p50k_edit" | "r50k_base" | "gpt2";

/** Token budget with an optional model-specific counter. */
export type ContextBudget = {
    maxTokens: number;
    counter?: TokenCounter;
};

/** Provenance pointer back to the source workbook location. */
export type SourceReference = {
    sheet: string;
    /** A1-style range for tables, empty for non-tabular blocks. */
    range: string;
    /** Anchor cell for assets, empty when not applicable. */
    anchor: string;
};

/** Formula metadata attached to a formula cell. */
export type IrFormula = {
    text: string;
    status: FormulaResolutionStatus;
    source: string | null;
};

/** One IR cell with raw and display values plus merge state. */
export type IrCell = {
    /** A1-style address of the cell. */
    address: string;
    /** One-based row number. */
    row: number;
    /** One-based column number. */
    col: number;
    /** Raw internal value of the cell. */
    raw: string;
    /** Display value of the cell. */
    display: string;
    /** Internal value type of the cell. */
    valueType: string;
    /** True when a merge range covers the cell. */
    merged: boolean;
    /** Formula metadata, null for non-formula cells. */
    formula: IrFormula | null;
    /** Cell hyperlink, null when the cell has none. */
    hyperlink: IrHyperlink | null;
};

/** Cell hyperlink. target is a URL for external links and a location such as Sheet2!A1 for internal ones. */
export type IrHyperlink = {
    kind: "external" | "internal";
    target: string;
    tooltip: string;
};

/** One chart series with the ranges it plots. */
export type IrChartSeries = {
    name: string;
    /** Category range, for example Sheet1!$B$4:$B$6. Empty when absent. */
    categories: string;
    /** Value range. Empty when absent. */
    values: string;
    /** Category labels from the chart cache. Empty when the chart stores no cache. */
    cachedCategories: string[];
    /** Values from the chart cache. Null marks a missing point. */
    cachedValues: (number | null)[];
    /** Id of the table on this sheet that holds the values, for example t001. Null when no table holds them. */
    sourceTable: string | null;
    axis: "primary" | "secondary";
};

/** One chart. Series ranges point back to the source cells. */
export type IrChart = {
    /** Chart id within the sheet, for example c001. Matches the Markdown "Chart: 001" heading. */
    id: string;
    /** Top-left anchor cell. */
    anchor: string;
    title: string;
    /** Chart type label, for example Line Chart. */
    type: string;
    series: IrChartSeries[];
};

/** One embedded image. Bytes stay in the parsed workbook. Path matches the Markdown ZIP export. */
export type IrImage = {
    /** Image id within the sheet, for example i001. Matches the Markdown "Image: 001" heading. */
    id: string;
    anchor: string;
    filename: string;
    path: string;
};

/** One drawing shape with its text. */
export type IrShape = {
    /** Shape id within the sheet, for example sh001. Matches the Markdown "Shape: 001" heading. */
    id: string;
    anchor: string;
    /** Shape kind, for example rect or flowChartProcess. */
    kind: string;
    name: string;
    /** Text inside the shape, empty when none. */
    text: string;
    /** Text with Markdown marks, for example ~~struck~~. Missing when it equals `text`. */
    markedText?: string;
    /** Fill, line color, and dash. Missing when the drawing gives no style data. */
    look?: ShapeLook;
    /**
     * Set for connectors only. Shape ids of the joined shapes. A null end is not attached to a shape.
     * With arrow "end", the arrowhead is at `to`. "none" is a plain line. "both" has two arrowheads.
     */
    connection: { from: string | null; to: string | null; arrow: "none" | "end" | "both" } | null;
    /**
     * Places on the sheet that a callout, a loose line end, or a marker refers to.
     * `cell` is under the point or marker. `text` is the nearest visible text, found by `method`.
     * Missing when the shape refers to nothing.
     */
    targets?: IrShapeTarget[];
};

/** One place that a shape refers to. */
export type IrShapeTarget = {
    role: "callout" | "lineStart" | "lineEnd" | "marker";
    cell: string;
    textAddress: string | null;
    text: string | null;
    method: "inside" | "overflow" | "right" | "image" | "nearest" | "none";
    /** Id of the shape that holds the point or carries the labeled text, for example sh012. */
    shape?: string;
    /** Id of the picture that holds the point or the marker, for example i002. */
    image?: string;
};

/** One detected table with its cell grid and provenance. */
export type IrTable = {
    /** Table id within the sheet, for example t001. Matches the Markdown "Table: 001" heading. */
    id: string;
    /** A1-style range of the table. */
    range: string;
    /** Name of the Excel table when the file declares this range as a table. */
    name?: string;
    /** False when the first row reads as data. Then headers is empty and every row is a data row. */
    hasHeader: boolean;
    /** Display values of the header row, empty when hasHeader is false. */
    headers: string[];
    /** Kept cell grid, header row first when hasHeader. Empty rows and columns follow AiIrOptions. */
    rows: IrCell[][];
    /** Provenance pointer for the table. */
    source: SourceReference;
};

/** One prose item inside an IR narrative block. */
export type IrNarrativeItem = {
    /** One-based row number of the item. */
    row: number;
    /** Text of the item. */
    text: string;
};

/** One prose block outside table candidates. */
export type IrNarrative = {
    /** One-based start row of the block. */
    startRow: number;
    /** Rendered lines of the block. */
    lines: string[];
    /** Structured items of the block. */
    items: IrNarrativeItem[];
};

/** One comment attached to a cell. */
export type IrComment = {
    /** A1-style address of the commented cell. */
    address: string;
    /** Comment kind. */
    kind: "note" | "threaded";
    /** Comment author, empty when unknown. */
    author: string;
    /** Comment text. */
    text: string;
};

/** One entry of a reading-order section: a table by id or a narrative block by index. */
export type IrSectionEntry = { kind: "table"; id: string } | { kind: "narrative"; index: number };

/** One IR sheet with tables, narratives, and comments. */
export type IrSheet = {
    /** Sheet name. */
    name: string;
    /** One-based sheet index in workbook order, as in Markdown file names. */
    index: number;
    /** Detected tables in reading order. */
    tables: IrTable[];
    /** Narrative blocks in row order. */
    narratives: IrNarrative[];
    /**
     * Reading order shared with Markdown. Groups related content into sections.
     * Lists tables and narratives in order, so captions stay next to tables.
     */
    sections: IrSectionEntry[][];
    /** Cell comments in address order. */
    comments: IrComment[];
    /** Charts in drawing order. */
    charts: IrChart[];
    /** Images in drawing order. */
    images: IrImage[];
    /** Shapes in drawing order. Empty when parsed with includeShapeDetails: false. */
    shapes: IrShape[];
    /** Visibility in Excel. */
    state: "visible" | "hidden" | "veryHidden";
    /** One-based rows that Excel hides. Their cells stay in the IR unless parsed with includeHidden: false. */
    hiddenRows: number[];
    /** One-based columns that Excel hides. */
    hiddenCols: number[];
};

/** Workbook-level intermediate representation for AI context. */
export type WorkbookIR = {
    /** Source workbook file name. */
    workbook: string;
    /** Sheets in workbook order. */
    sheets: IrSheet[];
};

/** Options for IR construction. */
export type AiIrOptions = {
    /** Table detection strategy. Defaults to balanced. */
    tableDetectionMode?: "balanced" | "border" | "planner-aware";
    /** Drops rows without values, like Markdown. Defaults to true. */
    removeEmptyRows?: boolean;
    /** Drops columns without values, like Markdown. Defaults to true. */
    removeEmptyColumns?: boolean;
};

/**
 * How tables are written for a model.
 * rows: a pipe table, header once (fewest tokens).
 * records: one self-describing line per row, "[B13] key: value | ..." (best for retrieval).
 */
export type TableFormat = "rows" | "records";

/** Rendering options for compact Markdown output. */
export type CompactRenderOptions = {
    /** Includes cell comments. Defaults to true. */
    includeComments?: boolean;
    /** Includes charts, images, and shapes. Defaults to true. */
    includeAssets?: boolean;
    /** Table layout. Defaults to rows. */
    tableFormat?: TableFormat;
};

/** Options for splitting the IR into retrieval chunks. */
export type ChunkOptions = {
    /** Token budget per chunk. Defaults to 512. A single table row larger than this becomes its own chunk. */
    maxTokens?: number;
    /** Counter for the target model. Defaults to the character estimator. */
    counter?: TokenCounter;
    /** Emits comment chunks. Defaults to true. */
    includeComments?: boolean;
    /** Emits chart, image, and shape chunks. Defaults to true. */
    includeAssets?: boolean;
    /** Table layout. Defaults to rows. Record chunks do not repeat a header row. */
    tableFormat?: TableFormat;
};

/** One retrieval chunk with provenance. */
export type ContextChunk = {
    /** Stable id within the workbook, for example s1/t001/2, s1/n/1, s1/c/1, s1/a/1. */
    id: string;
    kind: "table" | "narrative" | "comments" | "assets";
    /** Sheet name. */
    sheet: string;
    /** IR table id for table chunks, empty otherwise. */
    tableId: string;
    /** Source location. For table chunks, range covers the data rows in this chunk. */
    source: SourceReference;
    /** Rendered chunk text. Table chunks repeat the sheet and table, plus the header row in rows format. */
    text: string;
    /** Token count of text under the counter used. */
    tokens: number;
};

/** Options for toAIContext: IR, rendering, and chunking in one object. */
export type AiContextOptions = AiIrOptions & ChunkOptions;

/** One workbook prepared for a model: full text plus retrieval chunks. */
export type AiContext = {
    /** Intermediate representation the text and chunks were built from. */
    ir: WorkbookIR;
    /** Whole workbook as compact text in the chosen table format. */
    text: string;
    /** Token count of text under the counter used. Compare it with your prompt budget. */
    tokens: number;
    /** Retrieval chunks, each within maxTokens except single oversized rows. */
    chunks: ContextChunk[];
};
