/** Workbook parts: sheets, strings, options, and export files. */
import type { FormulaDiagnostic } from "./cell.js";
import type { TableScoreDetail } from "./table.js";
import type { RichTextRun } from "./cell.js";
import type { ParsedSheet } from "./worksheet.js";
/** One shared string entry from sharedStrings.xml. */
export type SharedStringEntry = {
    text: string;
    runs: RichTextRun[] | null;
};
/** Text encodings supported for Markdown export. shift_jis needs the optional iconv-lite peer. */
export type MarkdownEncoding = "utf-8" | "shift_jis" | "utf-16le" | "utf-16be" | "utf-32le" | "utf-32be";

/** Options for combined Markdown and ZIP export. */
export type MarkdownExportOptions = {
    /** Output text encoding. Defaults to utf-8. */
    encoding?: MarkdownEncoding | string | null;
    /** Byte order mark. Defaults to off. Not allowed with shift_jis. */
    bom?: "off" | "on" | string | null;
    /** Tool version written to front matter. */
    toolVersion?: string | null;
    /** Shape detail label written to front matter. Defaults to exclude. */
    shapeDetails?: "include" | "exclude" | string | null;
    /** Includes YAML front matter. Defaults to include. */
    frontMatter?: "include" | "exclude" | string | null;
};

/** Options for workbook parsing. */
export type ParseWorkbookOptions = {
    /** Parses shape details from drawings. Defaults to true. */
    includeShapeDetails?: boolean;
    /** Keeps hidden sheets, rows, and columns. Defaults to true. False drops them before any output. */
    includeHidden?: boolean;
};

/** Full parsed workbook with sheets, strings, and defined names. */
export type ParsedWorkbook = {
    name: string;
    sheets: ParsedSheet[];
    sharedStrings: SharedStringEntry[];
    definedNames: {
        name: string;
        formulaText: string;
        localSheetName: string | null;
    }[];
    /** Date system of the serials in rawValue. "1904" counts days from 1904-01-01. Missing means "1900". */
    dateSystem?: "1900" | "1904";
    /** Locale that the file shows through format locale tags or an East Asian default font. Null when unknown. */
    locale?: string | null;
};

/** Conversion options for sheet Markdown output. */
export type MarkdownOptions = {
    treatFirstRowAsHeader?: boolean;
    trimText?: boolean;
    removeEmptyRows?: boolean;
    removeEmptyColumns?: boolean;
    includeShapeDetails?: boolean;
    outputMode?: "display" | "raw" | "both";
    /** `github` is a deprecated alias of `gfm`. */
    formattingMode?: "plain" | "gfm" | "github";
    tableDetectionMode?: "balanced" | "border" | "planner-aware";
};

/** One converted sheet with its summary counters. */
export type MarkdownFile = {
    fileName: string;
    sheetName: string;
    markdown: string;
    summary: {
        outputMode: "display" | "raw" | "both";
        formattingMode: "plain" | "gfm";
        tableDetectionMode: "balanced" | "border" | "planner-aware";
        sections: number;
        tables: number;
        narrativeBlocks: number;
        merges: number;
        images: number;
        charts: number;
        comments: number;
        cells: number;
        tableScores: TableScoreDetail[];
        formulaDiagnostics: FormulaDiagnostic[];
    };
};

/** Named file entry for export archives. */
export type ExportEntry = { name: string; data: Uint8Array };
