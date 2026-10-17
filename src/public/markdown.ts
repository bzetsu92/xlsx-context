import type {
    MarkdownExportOptions,
    MarkdownFile,
    MarkdownOptions,
    ParsedWorkbook,
    ParseWorkbookOptions,
} from "../model/index.js";
import { convertSheetToMarkdown, convertWorkbookToMarkdownFiles, parseWorkbook as parseCoreWorkbook } from "../core.js";
import {
    FORMATTING_MODES,
    FORMATTING_MODE_ALIASES,
    OUTPUT_MODES,
    TABLE_DETECTION_MODE_ALIASES,
    TABLE_DETECTION_MODES,
} from "../outputs/markdown/markdown-options.js";
import {
    createCombinedMarkdownExportPayload,
    createSummaryText,
    createWorkbookExportArchive,
} from "../outputs/markdown/markdown-export.js";

/** Workbook bytes accepted by the public API. Node Buffers are Uint8Arrays. */
export type WorkbookInput = ArrayBuffer | Uint8Array;

/** Options for {@link parseWorkbook}. */
export type ParseOptions = ParseWorkbookOptions & {
    /** Workbook file name used in output headers and file names. Defaults to workbook.xlsx. */
    fileName?: string;
};

/** Copies a byte view into a standalone ArrayBuffer. Returns ArrayBuffer input unchanged. */
function toArrayBuffer(input: WorkbookInput): ArrayBuffer {
    if (input instanceof ArrayBuffer) {
        return input;
    }
    return input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength) as ArrayBuffer;
}

/** Parses .xlsx bytes into the shared workbook model. */
export function parseWorkbook(input: WorkbookInput, options: ParseOptions = {}): Promise<ParsedWorkbook> {
    const { fileName, ...parseOptions } = options;
    return parseCoreWorkbook(toArrayBuffer(input), fileName || "workbook.xlsx", parseOptions);
}

/** Converts one parsed sheet to a Markdown file. */
export { convertSheetToMarkdown };

/** Converts every sheet of a parsed workbook to Markdown files in workbook order. */
export function convertWorkbookToMarkdown(workbook: ParsedWorkbook, options: MarkdownOptions = {}): MarkdownFile[] {
    return convertWorkbookToMarkdownFiles(workbook, options);
}

/** One combined Markdown document for a workbook, encoded for writing. */
export type MarkdownDocument = {
    /** Suggested file name, for example book.md. */
    fileName: string;
    /** Markdown text. */
    content: string;
    /** Encoded bytes of content, including a BOM when requested. */
    data: Uint8Array;
    /** MIME type with charset. */
    mimeType: string;
};

/** Accepted values for the enum-like Markdown options. */
export type MarkdownOptionChoices = {
    outputModes: readonly string[];
    formattingModes: readonly string[];
    /** Legacy names accepted for formattingMode, mapped to current names. */
    formattingModeAliases: Readonly<Record<string, string>>;
    tableDetectionModes: readonly string[];
    /** Legacy names accepted for tableDetectionMode, mapped to current names. */
    tableDetectionModeAliases: Readonly<Record<string, string>>;
};

/** Combines per-sheet Markdown files into one document with optional front matter. */
export function createCombinedMarkdown(
    workbook: ParsedWorkbook,
    files: MarkdownFile[],
    options: MarkdownExportOptions = {}
): MarkdownDocument {
    return createCombinedMarkdownExportPayload(workbook, files, options);
}

/** Packs the combined Markdown and extracted assets (images) into a ZIP archive. */
export function createMarkdownArchive(
    workbook: ParsedWorkbook,
    files: MarkdownFile[],
    options: MarkdownExportOptions = {}
): Uint8Array {
    return createWorkbookExportArchive(workbook, files, options);
}

/** Returns a human-readable summary of one converted sheet. */
export { createSummaryText };

/** Returns the accepted values for Markdown options, for CLI or UI validation. */
export function getMarkdownOptionChoices(): MarkdownOptionChoices {
    return {
        outputModes: OUTPUT_MODES,
        formattingModes: FORMATTING_MODES,
        formattingModeAliases: FORMATTING_MODE_ALIASES,
        tableDetectionModes: TABLE_DETECTION_MODES,
        tableDetectionModeAliases: TABLE_DETECTION_MODE_ALIASES,
    };
}
