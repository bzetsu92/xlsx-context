import type { ExportEntry, FormulaDiagnostic, MarkdownExportOptions, MarkdownFile } from "../../model/index.js";
import { createStoredZip } from "../../extraction/zip/zip-io.js";
import { normalizeMarkdownText } from "./markdown-normalize.js";
import { escapeMarkdownTableCell } from "./markdown-table-escape.js";
import { createTextMimeType, encodeText } from "./text-encoding.js";

type MarkdownEncoding = "utf-8" | "shift_jis" | "utf-16le" | "utf-16be" | "utf-32le" | "utf-32be";
type MarkdownBomMode = "off" | "on";
type MarkdownEncodingOptions = {
    encoding?: MarkdownEncoding | string | null;
    bom?: MarkdownBomMode | string | null;
};

type WorkbookLike = {
    name: string;
    sheets: Array<{
        images: Array<{
            path: string;
            data: Uint8Array;
        }>;
        shapes?: Array<{
            svgPath: string | null;
            svgData: Uint8Array | null;
        }>;
    }>;
};

const FORMULA_STATUSES = ["resolved", "fallback_formula", "unsupported_external"] as const;

function normalizeMarkdownLineBreaks(text: string): string {
    return normalizeMarkdownText(text);
}

function escapeMarkdownCell(text: string): string {
    return escapeMarkdownTableCell(text);
}

function createBlankRow(columnCount: number): string[] {
    return new Array(columnCount).fill("");
}

/** Renders a GFM table. Adds one blank row when a lone header would not render. */
function renderMarkdownTable(rows: string[][], treatFirstRowAsHeader: boolean): string {
    if (rows.length === 0) {
        return "";
    }
    const workingRows = rows.map((row) => row.map((cell) => escapeMarkdownCell(cell)));
    if (workingRows.length === 1 && treatFirstRowAsHeader) {
        workingRows.push(createBlankRow(workingRows[0].length));
    }
    const header = treatFirstRowAsHeader ? workingRows[0] : createBlankRow(workingRows[0].length);
    const body = treatFirstRowAsHeader ? workingRows.slice(1) : workingRows;
    const lines = [`| ${header.join(" | ")} |`, `| ${header.map(() => "---").join(" | ")} |`];
    for (const row of body) {
        lines.push(`| ${row.join(" | ")} |`);
    }
    return lines.join("\n");
}

/** Sanitizes file name segments to letters, digits, dots, and dashes. Falls back when empty. */
function sanitizeFileNameSegment(value: string, fallback: string): string {
    const normalized = String(value || "").normalize("NFKC");
    const sanitized = normalized
        .replace(/[\\/:*?"<>|]/g, "_")
        .replace(/\s+/g, "_")
        .replace(/[^\p{L}\p{N}._-]+/gu, "_")
        .replace(/_+/g, "_")
        .replace(/^[_ .-]+|[_ .-]+$/g, "");
    return sanitized || fallback;
}

function stripWorkbookExtension(workbookName: string): string {
    return String(workbookName || "").replace(/\.xlsx$/i, "");
}

function createCombinedMarkdownFileName(workbookName: string): string {
    const baseName = stripWorkbookExtension(String(workbookName || "workbook")) || "workbook";
    return `${baseName}.md`;
}

function createExportEntryName(relativePath: string): string {
    return `output/${relativePath}`;
}

/** Quotes strings for YAML double-quoted scalars. */
function quoteYamlString(value: string): string {
    return `"${String(value).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n")}"`;
}

/**
 * Names per-sheet files as book_001_sheet.md. Ignores output and formatting modes.
 *
 * @param sheetIndex 1-based sheet position.
 */
function createOutputFileName(workbookName: string, sheetIndex: number, sheetName: string): string {
    const bookBase = sanitizeFileNameSegment(stripWorkbookExtension(workbookName), "workbook");
    const safeSheetName = sanitizeFileNameSegment(sheetName, `Sheet${sheetIndex}`);
    return `${bookBase}_${String(sheetIndex).padStart(3, "0")}_${safeSheetName}.md`;
}

function countFormulaStatuses(
    formulaDiagnostics: FormulaDiagnostic[]
): Record<(typeof FORMULA_STATUSES)[number], number> {
    const counts = {
        resolved: 0,
        fallback_formula: 0,
        unsupported_external: 0,
    };
    for (const item of formulaDiagnostics) {
        if (item.status && item.status in counts) {
            counts[item.status] += 1;
        }
    }
    return counts;
}

/** Renders conversion stats as human-readable lines. */
function createSummaryText(markdownFile: MarkdownFile): string {
    const formulaCounts = countFormulaStatuses(markdownFile.summary.formulaDiagnostics);
    return [
        `Output file: ${markdownFile.fileName}`,
        `Output mode: ${markdownFile.summary.outputMode}`,
        `Formatting mode: ${markdownFile.summary.formattingMode}`,
        `Table detection mode: ${markdownFile.summary.tableDetectionMode}`,
        `Sections: ${markdownFile.summary.sections}`,
        `Tables: ${markdownFile.summary.tables}`,
        `Narrative blocks: ${markdownFile.summary.narrativeBlocks}`,
        `Merged ranges: ${markdownFile.summary.merges}`,
        `Images: ${markdownFile.summary.images}`,
        `Charts: ${markdownFile.summary.charts}`,
        `Comments: ${markdownFile.summary.comments || 0}`,
        `Analyzed cells: ${markdownFile.summary.cells}`,
        ...FORMULA_STATUSES.map((status) => `Formula ${status}: ${formulaCounts[status]}`),
        ...markdownFile.summary.tableScores.map(
            (detail) => `Table candidate ${detail.range}: score ${detail.score} / ${detail.reasons.join(", ")}`
        ),
    ].join("\n");
}

/** Removes the shared book heading before combining sheets. */
function stripBookHeading(markdown: string, bookHeading: string): string {
    const lines = String(markdown || "").split("\n");
    if (lines[0] === bookHeading) {
        lines.shift();
        while (lines[0] === "") {
            lines.shift();
        }
    }
    return lines.join("\n");
}

/** Builds the YAML front matter with tool and mode metadata. */
function createFrontMatter(
    workbook: WorkbookLike,
    markdownFiles: MarkdownFile[],
    options: MarkdownExportOptions = {}
): string {
    const firstSummary = markdownFiles[0]?.summary;
    const workbookName = String(workbook.name || "workbook.xlsx");
    const shapeDetails = String(options.shapeDetails || "exclude");
    return [
        "---",
        `title: ${quoteYamlString(workbookName)}`,
        "type: converted",
        "conversion:",
        "  tool: xlsx-context",
        `  version: ${quoteYamlString(String(options.toolVersion || "unknown"))}`,
        `  output_mode: ${firstSummary?.outputMode || "display"}`,
        `  formatting_mode: ${firstSummary?.formattingMode || "gfm"}`,
        `  table_detection_mode: ${firstSummary?.tableDetectionMode || "balanced"}`,
        `  shape_details: ${shapeDetails}`,
        "---",
    ].join("\n");
}

function shouldIncludeFrontMatter(options: MarkdownExportOptions = {}): boolean {
    return String(options.frontMatter || "include") !== "exclude";
}

/** Combines per-sheet files into one document with front matter. Drops the repeated book heading. */
function createCombinedMarkdownExportFile(
    workbook: WorkbookLike,
    markdownFiles: MarkdownFile[],
    options: MarkdownExportOptions = {}
): { fileName: string; content: string } {
    const fileName = createCombinedMarkdownFileName(workbook.name);
    const bookHeading = `# Book: ${String(workbook.name || "workbook.xlsx")}`;
    const content = [
        ...(shouldIncludeFrontMatter(options) ? [createFrontMatter(workbook, markdownFiles, options)] : []),
        bookHeading,
        ...markdownFiles
            .map((markdownFile) => stripBookHeading(markdownFile.markdown, bookHeading))
            .filter((markdown) => markdown.trim().length > 0),
    ].join("\n\n");
    return { fileName, content };
}

function encodeMarkdownText(text: string, options: MarkdownEncodingOptions = {}): Uint8Array {
    return encodeText(text, options);
}

/** Encodes the combined document with BOM and MIME type. */
function createCombinedMarkdownExportPayload(
    workbook: WorkbookLike,
    markdownFiles: MarkdownFile[],
    options: MarkdownExportOptions = {}
): { fileName: string; content: string; data: Uint8Array; mimeType: string } {
    const combined = createCombinedMarkdownExportFile(workbook, markdownFiles, options);
    return {
        ...combined,
        data: encodeMarkdownText(`${combined.content}\n`, options),
        mimeType: createTextMimeType(options),
    };
}

/** Builds the Markdown ZIP entry. Returns null without Markdown files. */
function createMarkdownExportEntry(
    workbook: WorkbookLike,
    markdownFiles: MarkdownFile[],
    options: MarkdownExportOptions = {}
): ExportEntry | null {
    if (markdownFiles.length === 0) {
        return null;
    }
    const combined = createCombinedMarkdownExportPayload(workbook, markdownFiles, options);
    return {
        name: createExportEntryName(combined.fileName),
        data: combined.data,
    };
}

/** Collects image bytes and rendered shape SVGs. Skips shapes without SVG output. */
function createAssetExportEntries(workbook: WorkbookLike): ExportEntry[] {
    const entries: ExportEntry[] = [];
    for (const sheet of workbook.sheets) {
        for (const image of sheet.images) {
            entries.push({
                name: createExportEntryName(image.path),
                data: image.data,
            });
        }
        for (const shape of sheet.shapes || []) {
            if (!shape.svgPath || !shape.svgData) continue;
            entries.push({
                name: createExportEntryName(shape.svgPath),
                data: shape.svgData,
            });
        }
    }
    return entries;
}

/** Lists the Markdown entry first, then asset entries. */
function createExportEntries(
    workbook: WorkbookLike,
    markdownFiles: MarkdownFile[],
    options: MarkdownExportOptions = {}
): ExportEntry[] {
    const entries = createAssetExportEntries(workbook);
    const markdownEntry = createMarkdownExportEntry(workbook, markdownFiles, options);
    if (markdownEntry) {
        entries.unshift(markdownEntry);
    }
    return entries;
}

/** Packs Markdown and assets into one ZIP archive. */
function createWorkbookExportArchive(
    workbook: WorkbookLike,
    markdownFiles: MarkdownFile[],
    options: MarkdownExportOptions = {}
): Uint8Array {
    return createStoredZip(createExportEntries(workbook, markdownFiles, options));
}

/** Renders sheets as Markdown files and packs them with assets into ZIP archives. */
export {
    encodeMarkdownText,
    createCombinedMarkdownExportPayload,
    escapeMarkdownCell,
    renderMarkdownTable,
    sanitizeFileNameSegment,
    stripWorkbookExtension,
    createCombinedMarkdownFileName,
    createExportEntryName,
    createOutputFileName,
    createSummaryText,
    createFrontMatter,
    createCombinedMarkdownExportFile,
    createMarkdownExportEntry,
    createAssetExportEntries,
    createExportEntries,
    createWorkbookExportArchive,
    normalizeMarkdownLineBreaks,
};
