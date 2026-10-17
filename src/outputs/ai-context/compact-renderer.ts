import type {
    CompactRenderOptions,
    IrCell,
    IrChart,
    IrComment,
    IrImage,
    IrNarrative,
    IrShape,
    IrSheet,
    IrTable,
    TableFormat,
    WorkbookIR,
} from "./types.js";
import { describeHiddenContent } from "../../structure/layout/sheet-visibility.js";
import { toMarkdownLinkTarget } from "../../structure/formatting/link-target.js";

/** Escapes one compact table cell without padding. */
function escapeCompactCell(value: string): string {
    return String(value || "")
        .replace(/\r\n?|\n/g, "<br>")
        .replace(/\|/g, "\\|")
        .trim();
}

/**
 * Renders cell text for a model. Shows the display value plus the link target.
 * External links become [label](url). Internal links become label (Sheet2!A1).
 */
function renderCompactCellText(cell: IrCell): string {
    const label = escapeCompactCell(cell.display);
    const link = cell.hyperlink;
    if (!link || !label) return label;
    const linkTarget = link.kind === "external" ? toMarkdownLinkTarget(link.target) : null;
    if (linkTarget) return `[${label}](${escapeCompactCell(linkTarget)})`;
    return `${label} (${escapeCompactCell(link.target)})`;
}

/** Renders one IR table row as a tight pipe row of cell texts. */
function renderCompactRow(row: IrCell[]): string {
    return `| ${row.map(renderCompactCellText).join(" | ")} |`;
}

/**
 * Builds record keys from the header row. Empty headers use column letters.
 * Repeated headers append column letters, so keys stay unique.
 * Without a header row, pass any row with useLetters.
 */
function recordKeys(header: IrCell[], useLetters = false): string[] {
    const columnOf = (cell: IrCell) => cell.address.replace(/\d+$/, "");
    const names = header.map((cell) => (useLetters ? "" : escapeCompactCell(cell.display)) || columnOf(cell));
    const counts = new Map<string, number>();
    for (const name of names) counts.set(name, (counts.get(name) || 0) + 1);
    return names.map((name, index) =>
        (counts.get(name) as number) > 1 ? `${name} (${columnOf(header[index])})` : name
    );
}

/** Renders one data row as a self-describing record. Empty values are omitted. */
function renderCompactRecord(keys: string[], row: IrCell[]): string {
    const pairs = row
        .map((cell, index) => [keys[index], renderCompactCellText(cell)])
        .filter(([, value]) => value !== "")
        .map(([key, value]) => `${key}: ${value}`);
    return [`[${row[0].address}]`, pairs.join(" | ")].filter(Boolean).join(" ");
}

/**
 * Splits a table into an optional header line and one line per data row.
 * Chunking packs the lines and repeats the header. Full rendering prints both.
 * A records table without data rows keeps its header as a pipe row.
 * A table without a header row has only data lines. Records then key by column letter.
 */
function renderTableLines(table: IrTable, format: TableFormat = "rows"): { header: string | null; lines: string[] } {
    if (table.hasHeader === false) {
        if (format === "records" && table.rows.length > 0) {
            const keys = recordKeys(table.rows[0], true);
            return { header: null, lines: table.rows.map((row) => renderCompactRecord(keys, row)) };
        }
        return { header: null, lines: table.rows.map(renderCompactRow) };
    }
    const [header, ...dataRows] = table.rows;
    if (format === "records" && dataRows.length > 0) {
        const keys = recordKeys(header);
        return { header: null, lines: dataRows.map((row) => renderCompactRecord(keys, row)) };
    }
    return { header: renderCompactRow(header), lines: dataRows.map(renderCompactRow) };
}

/** Renders one IR table with its range in the chosen format. */
function renderCompactTable(table: IrTable, format?: TableFormat): string {
    const { header, lines } = renderTableLines(table, format);
    const heading = table.name ? `### ${table.range} ${table.name}` : `### ${table.range}`;
    return [heading, ...(header === null ? [] : [header]), ...lines].join("\n");
}

/** Renders one IR narrative block as tight lines. */
function renderCompactNarrative(narrative: IrNarrative): string {
    return narrative.lines.map((line) => line.trim()).join("\n");
}

/** Renders one cell comment as a list item with its address. */
function renderCompactComment(comment: IrComment): string {
    const author = comment.author ? `${comment.author}: ` : "";
    return `- [${comment.address}] ${author}${comment.text}`.trimEnd();
}

/** Collapses multi-line asset text onto one line. */
function oneLine(value: string): string {
    return escapeCompactCell(String(value || "").replace(/\s*\r?\n\s*/g, " / "));
}

/** Joins cached points as `label=value` pairs. Uses the point number when the chart has no categories. */
function renderCachedPoints(series: IrChart["series"][number], categories: string[]): string {
    return series.cachedValues
        .map((value, index) => `${oneLine(categories[index] ?? String(index + 1))}=${value === null ? "" : value}`)
        .join(", ");
}

/**
 * Lists each series with its ranges. A series held by a table names the table.
 * Other series carry their cached points, so the reader needs no range lookup.
 */
function renderCompactChart(chart: IrChart): string {
    const title = chart.title ? oneLine(chart.title) : "(untitled)";
    const lines = [`[${chart.anchor}] Chart ${chart.id}: ${title} (${chart.type})`];
    const categories = chart.series.find((series) => series.cachedCategories?.length > 0)?.cachedCategories || [];
    for (const series of chart.series) {
        const axis = series.axis === "secondary" ? " (secondary axis)" : "";
        const ranges = [
            series.categories ? `categories ${series.categories}` : "",
            series.values ? `values ${series.values}` : "",
            series.sourceTable ? `table ${series.sourceTable}` : "",
            !series.sourceTable && series.cachedValues?.length > 0
                ? `data ${renderCachedPoints(series, categories)}`
                : "",
        ].filter(Boolean);
        lines.push(`- ${oneLine(series.name) || "(unnamed)"}${axis}${ranges.length ? `: ${ranges.join("; ")}` : ""}`);
    }
    return lines.join("\n");
}

/** Connector lines show what they join, for example `sh001 -> sh002`. A loose end shows as `?`. */
function renderCompactShape(shape: IrShape): string {
    const label = [shape.kind, oneLine(shape.name)].filter(Boolean).join(", ");
    const text = oneLine(shape.markedText ?? shape.text);
    const arrows = { none: "--", end: "->", both: "<->" } as const;
    const link = shape.connection
        ? `${shape.connection.from ?? "?"} ${arrows[shape.connection.arrow]} ${shape.connection.to ?? "?"}`
        : "";
    const refers = (shape.targets || [])
        .map((target) => {
            const at =
                target.textAddress && target.text ? `${target.textAddress} "${oneLine(target.text)}"` : target.cell;
            return `${target.role === "lineStart" ? "start at" : target.role === "marker" ? "marks" : "points to"} ${at}`;
        })
        .join("; ");
    const body = [link, text, refers ? `(${refers})` : ""].filter(Boolean).join(" ");
    return `[${shape.anchor}] Shape ${shape.id}${label ? ` (${label})` : ""}${body ? `: ${body}` : ""}`;
}

function renderCompactImage(image: IrImage): string {
    return `[${image.anchor}] Image ${image.id}: ${image.path}`;
}

/** One entry per chart, shape, and image, in the Markdown section order. Each starts with its anchor. */
function renderCompactAssets(sheet: IrSheet): string[] {
    return [
        ...(sheet.charts ?? []).map(renderCompactChart),
        ...(sheet.shapes ?? []).map(renderCompactShape),
        ...(sheet.images ?? []).map(renderCompactImage),
    ];
}

/**
 * Lists body parts in reading order. Follows sheet.sections when present, so captions stay next to tables.
 * Falls back to tables then narratives for anything sections miss.
 */
function renderCompactBody(sheet: IrSheet, options: CompactRenderOptions): string[] {
    const parts: string[] = [];
    const renderedTables = new Set<IrTable>();
    const renderedNarratives = new Set<IrNarrative>();
    const pushTable = (table: IrTable | undefined) => {
        if (!table || renderedTables.has(table)) return;
        renderedTables.add(table);
        parts.push(renderCompactTable(table, options.tableFormat));
    };
    const pushNarrative = (narrative: IrNarrative | undefined) => {
        if (!narrative || renderedNarratives.has(narrative)) return;
        renderedNarratives.add(narrative);
        const rendered = renderCompactNarrative(narrative);
        if (rendered) parts.push(rendered);
    };
    for (const section of sheet.sections ?? []) {
        for (const entry of section) {
            if (entry.kind === "table") pushTable(sheet.tables.find((table) => table.id === entry.id));
            else pushNarrative(sheet.narratives[entry.index]);
        }
    }
    sheet.tables.forEach(pushTable);
    sheet.narratives.forEach(pushNarrative);
    return parts;
}

/** Renders one IR sheet with tables, narratives, comments, and assets. */
function renderCompactSheet(sheet: IrSheet, options: CompactRenderOptions): string {
    const hiddenNote = describeHiddenContent(sheet);
    const parts = [`# ${sheet.name}`, ...(hiddenNote ? [`> ${hiddenNote}`] : []), ...renderCompactBody(sheet, options)];
    if (options.includeComments !== false) {
        for (const comment of sheet.comments) {
            parts.push(renderCompactComment(comment));
        }
    }
    if (options.includeAssets !== false) {
        parts.push(...renderCompactAssets(sheet));
    }
    return parts.join("\n\n");
}

/**
 * Renders workbook IR as compact Markdown for AI context.
 * Keeps sheet names, table ranges, and cell values. Omits front
 * matter, padded tables, merge tokens, empty sections, and
 * surrounding whitespace inside cells.
 * Merged non-master cells stay empty. The table range keeps
 * the merge recoverable from the source workbook.
 */
function renderCompactMarkdown(ir: WorkbookIR, options: CompactRenderOptions = {}): string {
    return ir.sheets.map((sheet) => renderCompactSheet(sheet, options)).join("\n\n");
}

/** Renders workbook IR as compact Markdown for AI context. */
export {
    renderCompactMarkdown,
    renderCompactRow,
    renderTableLines,
    renderCompactNarrative,
    renderCompactComment,
    renderCompactAssets,
};
