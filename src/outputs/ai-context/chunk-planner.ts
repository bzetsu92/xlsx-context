import type { ChunkOptions, ContextChunk, IrSheet, IrTable, TableFormat, TokenCounter, WorkbookIR } from "./types.js";
import {
    renderCompactAssets,
    renderCompactComment,
    renderCompactNarrative,
    renderTableLines,
} from "./compact-renderer.js";
import { createCharEstimateCounter } from "./token-counter.js";

const DEFAULT_MAX_TOKENS = 512;

/** A run of consecutive units [start, end) rendered under one prefix. */
type PackedRun = {
    start: number;
    end: number;
    text: string;
    tokens: number;
};

function renderRun(prefix: string, units: string[]): string {
    return [prefix, ...units].join("\n");
}

/**
 * Packs units greedily under the token budget, keeping order and never splitting a unit.
 * Per-unit counts drive the greedy pass, then the real count of the joined text is checked,
 * so non-additive model tokenizers still stay within budget. A unit that exceeds the budget
 * on its own becomes a single-unit run.
 */
function packUnits(prefix: string, units: string[], maxTokens: number, counter: TokenCounter): PackedRun[] {
    const prefixTokens = counter.count(prefix);
    const unitTokens = units.map((unit) => counter.count(unit));
    const runs: PackedRun[] = [];
    let start = 0;
    while (start < units.length) {
        let end = start + 1;
        let estimate = prefixTokens + unitTokens[start];
        while (end < units.length && estimate + unitTokens[end] <= maxTokens) {
            estimate += unitTokens[end];
            end += 1;
        }
        let text = renderRun(prefix, units.slice(start, end));
        let tokens = counter.count(text);
        while (tokens > maxTokens && end - start > 1) {
            end -= 1;
            text = renderRun(prefix, units.slice(start, end));
            tokens = counter.count(text);
        }
        runs.push({ start, end, text, tokens });
        start = end;
    }
    return runs;
}

/** Splits one table into chunks of whole data rows. Each chunk repeats the header row. */
function chunkTable(
    sheet: IrSheet,
    table: IrTable,
    format: TableFormat,
    maxTokens: number,
    counter: TokenCounter
): ContextChunk[] {
    const { header, lines: units } = renderTableLines(table, format);
    const prefix = [
        `# ${sheet.name}`,
        `## Table ${table.id} (${table.range})`,
        ...(header === null ? [] : [header]),
    ].join("\n");
    const toChunk = (index: number, text: string, tokens: number, firstRow: number, lastRow: number): ContextChunk => {
        const first = table.rows[firstRow][0];
        const lastCells = table.rows[lastRow];
        return {
            id: `s${sheet.index}/${table.id}/${index + 1}`,
            kind: "table",
            sheet: sheet.name,
            tableId: table.id,
            source: {
                sheet: sheet.name,
                range: `${first.address}-${lastCells[lastCells.length - 1].address}`,
                anchor: "",
            },
            text,
            tokens,
        };
    };
    if (units.length === 0) {
        return [toChunk(0, prefix, counter.count(prefix), 0, 0)];
    }
    // Run indices cover data rows only. table.rows shifts by one when a header row exists.
    const headerOffset = table.hasHeader === false ? 0 : 1;
    return packUnits(prefix, units, maxTokens, counter).map((run, index) =>
        toChunk(index, run.text, run.tokens, run.start + headerOffset, run.end - 1 + headerOffset)
    );
}

/** Treats narrative blocks as units. A block over budget falls back to its lines. */
function narrativeUnits(sheet: IrSheet, maxTokens: number, counter: TokenCounter): string[] {
    const units: string[] = [];
    for (const narrative of sheet.narratives) {
        const block = renderCompactNarrative(narrative);
        if (!block) continue;
        if (counter.count(block) <= maxTokens) {
            units.push(block);
        } else {
            units.push(...block.split("\n").filter((line) => line.trim() !== ""));
        }
    }
    return units;
}

function chunkSheetText(
    sheet: IrSheet,
    kind: "narrative" | "comments" | "assets",
    units: string[],
    maxTokens: number,
    counter: TokenCounter
): ContextChunk[] {
    const prefix = {
        narrative: `# ${sheet.name}`,
        comments: `# ${sheet.name}\n## Comments`,
        assets: `# ${sheet.name}\n## Charts, shapes, and images`,
    }[kind];
    const idPart = { narrative: "n", comments: "c", assets: "a" }[kind];
    return packUnits(prefix, units, maxTokens, counter).map((run, index) => ({
        id: `s${sheet.index}/${idPart}/${index + 1}`,
        kind,
        sheet: sheet.name,
        tableId: "",
        source: { sheet: sheet.name, range: "", anchor: "" },
        text: run.text,
        tokens: run.tokens,
    }));
}

/**
 * Splits workbook IR into retrieval chunks in reading order.
 * Orders tables, narrative text, comments, charts, shapes, and images per sheet.
 * Stays deterministic for the same IR, options, and counter.
 */
function planContextChunks(ir: WorkbookIR, options: ChunkOptions = {}): ContextChunk[] {
    const maxTokens = options.maxTokens ?? DEFAULT_MAX_TOKENS;
    if (!Number.isFinite(maxTokens) || maxTokens <= 0) {
        throw new Error(`maxTokens must be a positive number: ${maxTokens}`);
    }
    const counter = options.counter || createCharEstimateCounter();
    const chunks: ContextChunk[] = [];
    for (const sheet of ir.sheets) {
        for (const table of sheet.tables) {
            chunks.push(...chunkTable(sheet, table, options.tableFormat || "rows", maxTokens, counter));
        }
        chunks.push(
            ...chunkSheetText(sheet, "narrative", narrativeUnits(sheet, maxTokens, counter), maxTokens, counter)
        );
        if (options.includeComments !== false) {
            chunks.push(
                ...chunkSheetText(sheet, "comments", sheet.comments.map(renderCompactComment), maxTokens, counter)
            );
        }
        if (options.includeAssets !== false) {
            chunks.push(...chunkSheetText(sheet, "assets", renderCompactAssets(sheet), maxTokens, counter));
        }
    }
    return chunks;
}

/** Splits workbook IR into token-limited retrieval chunks. */
export { planContextChunks };
