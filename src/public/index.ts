/*
 * Public package entry. Only names exported here are public API.
 * Runtime environments must provide DOMParser, Blob.stream, and
 * DecompressionStream; the Node entry (`node.ts`) installs DOMParser.
 */

export {
    parseWorkbook,
    convertSheetToMarkdown,
    convertWorkbookToMarkdown,
    createCombinedMarkdown,
    createMarkdownArchive,
    createSummaryText,
    getMarkdownOptionChoices,
} from "./markdown.js";
export type { MarkdownDocument, MarkdownOptionChoices, ParseOptions, WorkbookInput } from "./markdown.js";
export {
    toAIContext,
    buildWorkbookIR,
    renderCompactMarkdown,
    createContextChunks,
    createCharEstimateCounter,
    createEncoderCounter,
    createTiktokenCounter,
    createWhitespaceCounter,
} from "./context.js";

export type {
    CellHyperlink,
    MarkdownEncoding,
    MarkdownExportOptions,
    MarkdownFile,
    MarkdownOptions,
    MergeRange,
    ParsedCell,
    ParsedCellComment,
    ParsedChartAsset,
    ParsedImageAsset,
    ParsedShapeAsset,
    ParsedSheet,
    ParsedWorkbook,
    ParseWorkbookOptions,
    RichTextRun,
} from "../model/index.js";
export type {
    AiContext,
    AiContextOptions,
    AiIrOptions,
    ChunkOptions,
    CompactRenderOptions,
    ContextBudget,
    ContextChunk,
    IrCell,
    IrChart,
    IrChartSeries,
    IrComment,
    IrHyperlink,
    IrImage,
    IrFormula,
    IrNarrative,
    IrNarrativeItem,
    IrShape,
    IrSheet,
    IrTable,
    SourceReference,
    TableFormat,
    TextEncoderLike,
    TiktokenEncodingName,
    TokenCounter,
    WorkbookIR,
} from "../outputs/ai-context/types.js";
