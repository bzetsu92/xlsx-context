import type { MarkdownOptions, ParsedCell, ParsedSheet, ParsedWorkbook } from "../../model/index.js";
import type { NarrativeCellFormatter } from "../../structure/layout/narrative-blocks.js";
import { resolveMarkdownOptions } from "./markdown-options.js";
import { toMarkdownLinkTarget } from "../../structure/formatting/link-target.js";
import { escapeMarkdownLiteralText } from "./markdown-escape.js";
import { createRichTextRenderer } from "./rich-text/rich-text-renderer.js";

type CellMarkdownValues = {
    displayValue: string;
    rawValue: string;
    displayMarkdown: string;
};

type CellMarkdownContext = {
    workbook: ParsedWorkbook | null;
    sheet: ParsedSheet | null;
    options: MarkdownOptions;
};

type CellMarkdownRenderer = (cell: ParsedCell, values: CellMarkdownValues, context: CellMarkdownContext) => string;

/** Cell rendering dependencies. Only text normalization is injected. */
type CellMarkdownDeps = {
    normalizeMarkdownText?: (text: string) => string;
};

/** Renders cells to Markdown text in display, raw, or both modes. */
export function createCellMarkdown(deps: CellMarkdownDeps = {}) {
    const richTextRenderer = createRichTextRenderer({
        normalizeMarkdownText: deps.normalizeMarkdownText,
    });

    function buildCellMap(sheet: ParsedSheet): Map<string, ParsedCell> {
        const map = new Map<string, ParsedCell>();
        for (const cell of sheet.cells) {
            map.set(`${cell.row}:${cell.col}`, cell);
        }
        return map;
    }

    // Internally resolved options stay private and frozen. Per-cell formatting reuses them.
    // Caller objects always resolve fresh.
    const internallyResolvedOptions = new WeakSet<MarkdownOptions>();

    function createHeadingFragment(text: string): string {
        return String(text || "")
            .trim()
            .toLowerCase()
            .replace(/<[^>]+>/g, "")
            .replace(/[^\p{L}\p{N}\s_-]+/gu, "")
            .replace(/\s+/g, "-");
    }

    function parseInternalHyperlinkLocation(
        location: string,
        currentSheetName: string
    ): { sheetName: string; refText: string } {
        const normalized = String(location || "")
            .trim()
            .replace(/^#/, "");
        if (!normalized) {
            return { sheetName: currentSheetName, refText: "" };
        }
        const match = normalized.match(/^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/);
        if (match) {
            return {
                sheetName: (match[1] || match[2] || currentSheetName).replace(/''/g, "'"),
                refText: (match[3] || "").trim(),
            };
        }
        return {
            sheetName: currentSheetName,
            refText: normalized,
        };
    }

    function renderHyperlinkMarkdown(
        cell: ParsedCell,
        text: string,
        workbook: ParsedWorkbook | null,
        sheet: ParsedSheet | null
    ): string {
        const hyperlink = cell.hyperlink;
        const label = String(text || "").trim();
        if (!hyperlink || !label) return text;
        if (hyperlink.kind === "external") {
            const target = String(hyperlink.target || "").trim();
            const href = toMarkdownLinkTarget(target);
            if (href) return `[${label}](${href})`;
            // Unsafe or empty target: keep it visible as plain text, never as a link.
            return target ? `${label} (${escapeMarkdownLiteralText(target)})` : label;
        }
        const currentSheetName = sheet?.name || "";
        const { sheetName, refText } = parseInternalHyperlinkLocation(
            hyperlink.location || hyperlink.target,
            currentSheetName
        );
        const traceText = [sheetName, refText].filter(Boolean).join("!");
        const targetSheet = workbook?.sheets.find((entry) => entry.name === sheetName) || null;
        if (!targetSheet || !workbook) {
            return traceText ? `${label} (${traceText})` : label;
        }
        // Matches the slug of the `## Sheet:` heading.
        const href = `#${createHeadingFragment(`Sheet: ${targetSheet.name}`)}`;
        return traceText && traceText !== targetSheet.name
            ? `[${label}](${href}) (${traceText})`
            : `[${label}](${href})`;
    }

    function createDisplayCellForFormatting(cell: ParsedCell, formattingMode: "plain" | "gfm"): ParsedCell {
        if (formattingMode !== "gfm" || !cell.hyperlink) {
            return cell;
        }
        return {
            ...cell,
            textStyle: {
                ...cell.textStyle,
                underline: false,
            },
            richTextRuns:
                cell.richTextRuns?.map((run) => ({
                    ...run,
                    underline: false,
                })) || null,
        };
    }

    function createCellMarkdownValues(cell: ParsedCell, formattingMode: "plain" | "gfm"): CellMarkdownValues {
        const displayCell = createDisplayCellForFormatting(cell, formattingMode);
        return {
            displayValue: richTextRenderer.compactText(String(cell.outputValue || "")),
            rawValue: richTextRenderer.compactText(String(cell.rawValue || "")),
            displayMarkdown: richTextRenderer.renderCellDisplayText(displayCell, formattingMode),
        };
    }

    function renderCellWithHyperlink(cell: ParsedCell, text: string, context: CellMarkdownContext): string {
        return renderHyperlinkMarkdown(cell, text, context.workbook, context.sheet);
    }

    function renderDisplayModeCell(cell: ParsedCell, values: CellMarkdownValues, context: CellMarkdownContext): string {
        return renderCellWithHyperlink(cell, values.displayMarkdown, context);
    }

    function renderRawModeCell(cell: ParsedCell, values: CellMarkdownValues, context: CellMarkdownContext): string {
        return renderCellWithHyperlink(cell, values.rawValue || values.displayValue, context);
    }

    function renderBothModeCell(cell: ParsedCell, values: CellMarkdownValues, context: CellMarkdownContext): string {
        if (values.rawValue && values.rawValue !== values.displayValue) {
            if (values.displayMarkdown) {
                return `${renderCellWithHyperlink(cell, values.displayMarkdown, context)} [raw=${values.rawValue}]`;
            }
            return `[raw=${values.rawValue}]`;
        }
        return renderCellWithHyperlink(cell, values.displayMarkdown || values.rawValue, context);
    }

    function getCellMarkdownRenderer(outputMode: "display" | "raw" | "both"): CellMarkdownRenderer {
        if (outputMode === "raw") {
            return renderRawModeCell;
        }
        if (outputMode === "both") {
            return renderBothModeCell;
        }
        return renderDisplayModeCell;
    }

    function formatCellForMarkdown(
        cell: ParsedCell | undefined,
        options: MarkdownOptions,
        workbook: ParsedWorkbook | null = null,
        sheet: ParsedSheet | null = null
    ): string {
        if (!cell) return "";
        const resolvedOptions = internallyResolvedOptions.has(options)
            ? (options as ReturnType<typeof resolveMarkdownOptions>)
            : resolveMarkdownOptions(options);
        const formattingMode = resolvedOptions.formattingMode;
        const values = createCellMarkdownValues(cell, formattingMode);
        const renderCell = getCellMarkdownRenderer(resolvedOptions.outputMode);
        return renderCell(cell, values, {
            workbook,
            sheet,
            options,
        });
    }

    // Narrative blocks and sections live in structure/layout, shared with the AI IR.
    // These wrappers keep the Markdown-side signatures and pass the Markdown cell formatter.
    function createMarkdownNarrativeFormatter(
        options: MarkdownOptions,
        workbook: ParsedWorkbook | null,
        sheet: ParsedSheet | null
    ): NarrativeCellFormatter {
        return (cell) => formatCellForMarkdown(cell, options, workbook, sheet);
    }

    /** Resolves options once and caches them for per-cell formatting. */
    function resolveCachedMarkdownOptions(options: MarkdownOptions): ReturnType<typeof resolveMarkdownOptions> {
        const resolved = resolveMarkdownOptions(options);
        internallyResolvedOptions.add(resolved);
        return resolved;
    }

    return {
        buildCellMap,
        formatCellForMarkdown,
        createMarkdownNarrativeFormatter,
        resolveCachedMarkdownOptions,
    };
}
