/** Cell primitives: values, styles, formulas, and links. */
/** Text decoration flags shared by styles, cells, and rich text. */
export type RichTextStyle = {
    bold: boolean;
    italic: boolean;
    strike: boolean;
    underline: boolean;
};

/** One styled text span inside a cell value. */
export type RichTextRun = RichTextStyle & {
    text: string;
};

/** Border presence on the four sides of a cell. */
export type BorderFlags = {
    top: boolean;
    bottom: boolean;
    left: boolean;
    right: boolean;
};

/**
 * Text placement in a cell. Decides where long text overflows.
 * `horizontal` keeps the OOXML value, for example "general", "left", "center", or "right".
 */
export type CellAlignment = {
    horizontal: string;
    /** True when the text wraps or shrinks to fit. Such text never overflows. */
    contained: boolean;
};

/** Style inputs that control the display value of a cell. */
export type CellStyleInfo = {
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
    textStyle: RichTextStyle;
    /** Missing when the style gives no alignment. */
    alignment?: CellAlignment;
};

/** Outcome of formula resolution for one cell. */
export type FormulaResolutionStatus = "resolved" | "fallback_formula" | "unsupported_external" | null;

/** Resolution path that produced the output value of a formula cell. `reference_lookup` reads a cell or defined name directly. */
export type FormulaResolutionSource =
    "cached_value" | "ast_evaluator" | "reference_lookup" | "formula_text" | "external_unsupported" | null;

/** Cached value state stored in the workbook for a formula cell. */
export type CachedValueState = "present_nonempty" | "present_empty" | "absent" | null;

/** Rectangular cell range with an A1-style reference. */
export type MergeRange = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    ref: string;
};

/** Hyperlink attached to a cell. */
export type CellHyperlink = {
    kind: "external" | "internal";
    target: string;
    location: string;
    tooltip: string;
    display: string;
};

/** Full parsed cell produced by the worksheet parser. */
export type ParsedCell = {
    address: string;
    row: number;
    col: number;
    valueType: string;
    rawValue: string;
    outputValue: string;
    formulaText: string;
    resolutionStatus: FormulaResolutionStatus;
    resolutionSource: FormulaResolutionSource;
    cachedValueState: CachedValueState;
    styleIndex: number;
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
    textStyle: RichTextStyle;
    /** Missing when the cell style gives no alignment. */
    alignment?: CellAlignment;
    richTextRuns: RichTextRun[] | null;
    formulaType: string;
    spillRef: string;
    hyperlink: CellHyperlink | null;
};

/** Parsed sheet-scoped range reference. */
export type ParsedRangeRef = {
    sheetName: string;
    start: string;
    end: string;
};

/** Formula diagnostics entry reported in a file summary. */
export type FormulaDiagnostic = {
    address: string;
    formulaText: string;
    status: FormulaResolutionStatus;
    source: string | null;
    outputValue: string;
};

/** Raw and numeric values resolved for one range. */
export type RangeEntries = {
    rawValues: string[];
    numericValues: number[];
};

/** Resolves one cell address to its output value. */
export type ResolveCellValue = (sheetName: string, address: string) => string;

/** Resolves one range to its raw and numeric values. */
export type ResolveRangeEntries = (sheetName: string, rangeText: string) => RangeEntries;

/** Resolves one defined name to its scalar value. */
export type ResolveDefinedNameScalarValue = ((sheetName: string, name: string) => string | null) | null;
