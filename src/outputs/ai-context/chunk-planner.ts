import type {
    ChunkOptions,
    ContextChunk,
    IrNarrative,
    IrSheet,
    IrTable,
    TableFormat,
    TokenCounter,
    WorkbookIR,
} from "./types.js";
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

/** One packable piece of sheet text with the cells or rows it came from. */
type SourcedUnit = {
    text: string;
    /** Excel reference of the piece: a row span such as 4:8, or a cell such as B3. */
    ref: string;
};

/** Joins the refs of a run. Row spans merge into one span. Cells become a list such as B3,D10. */
function joinRefs(units: SourcedUnit[]): string {
    const spans = units.map((unit) => unit.ref.match(/^(\d+):(\d+)$/));
    if (spans.length > 0 && spans.every(Boolean)) {
        const first = Math.min(...spans.map((span) => Number(span![1])));
        const last = Math.max(...spans.map((span) => Number(span![2])));
        return `${first}:${last}`;
    }
    return [...new Set(units.map((unit) => unit.ref).filter(Boolean))].join(",");
}

/** Treats narrative blocks as units. A block over budget falls back to its lines. */
function narrativeUnits(narratives: IrNarrative[], maxTokens: number, counter: TokenCounter): SourcedUnit[] {
    const units: SourcedUnit[] = [];
    for (const narrative of narratives) {
        const block = renderCompactNarrative(narrative);
        if (!block) continue;
        const lastRow = Math.max(narrative.startRow, ...narrative.items.map((item) => item.row));
        const ref = `${narrative.startRow}:${lastRow}`;
        if (counter.count(block) <= maxTokens) {
            units.push({ text: block, ref });
        } else {
            // Rendered lines do not map one to one to items, so each line keeps the block span.
            for (const line of block.split("\n")) {
                if (line.trim() !== "") units.push({ text: line, ref });
            }
        }
    }
    return units;
}

/** Lists asset texts with their anchor cells, in the order of renderCompactAssets. */
function assetUnits(sheet: IrSheet): SourcedUnit[] {
    const anchors = [
        ...(sheet.charts ?? []).map((chart) => chart.anchor),
        ...(sheet.shapes ?? []).map((shape) => shape.anchor),
        ...(sheet.images ?? []).map((image) => image.anchor),
    ];
    return renderCompactAssets(sheet).map((text, index) => ({ text, ref: anchors[index] || "" }));
}

function chunkSheetText(
    sheet: IrSheet,
    kind: "narrative" | "comments" | "assets",
    units: SourcedUnit[],
    firstNumber: number,
    maxTokens: number,
    counter: TokenCounter
): ContextChunk[] {
    const prefix = {
        narrative: `# ${sheet.name}`,
        comments: `# ${sheet.name}\n## Comments`,
        assets: `# ${sheet.name}\n## Charts, shapes, and images`,
    }[kind];
    const idPart = { narrative: "n", comments: "c", assets: "a" }[kind];
    const texts = units.map((unit) => unit.text);
    return packUnits(prefix, texts, maxTokens, counter).map((run, index) => {
        const runUnits = units.slice(run.start, run.end);
        return {
            id: `s${sheet.index}/${idPart}/${firstNumber + index}`,
            kind,
            sheet: sheet.name,
            tableId: "",
            source: {
                sheet: sheet.name,
                range: joinRefs(runUnits),
                anchor: kind === "assets" ? runUnits[0]?.ref || "" : "",
            },
            text: run.text,
            tokens: run.tokens,
        };
    });
}

type BodyEntry = { kind: "table"; table: IrTable } | { kind: "narrative"; narrative: IrNarrative };

/**
 * Lists tables and narrative blocks in reading order, as the compact text does.
 * Follows sheet.sections, then adds anything the sections miss.
 */
function bodyInReadingOrder(sheet: IrSheet): BodyEntry[] {
    const entries: BodyEntry[] = [];
    const seenTables = new Set<IrTable>();
    const seenNarratives = new Set<IrNarrative>();
    const pushTable = (table: IrTable | undefined) => {
        if (!table || seenTables.has(table)) return;
        seenTables.add(table);
        entries.push({ kind: "table", table });
    };
    const pushNarrative = (narrative: IrNarrative | undefined) => {
        if (!narrative || seenNarratives.has(narrative)) return;
        seenNarratives.add(narrative);
        entries.push({ kind: "narrative", narrative });
    };
    for (const section of sheet.sections ?? []) {
        for (const entry of section) {
            if (entry.kind === "table") pushTable(sheet.tables.find((table) => table.id === entry.id));
            else pushNarrative(sheet.narratives[entry.index]);
        }
    }
    sheet.tables.forEach(pushTable);
    sheet.narratives.forEach(pushNarrative);
    return entries;
}

/**
 * Splits workbook IR into retrieval chunks in reading order.
 * Per sheet: tables and narrative text in the order of the compact text, then comments,
 * then charts, shapes, and images. Every chunk points back to its rows or cells.
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
        let pendingNarratives: IrNarrative[] = [];
        let narrativeNumber = 1;
        // Consecutive narrative blocks pack together. A table between them starts a new run.
        const flushNarratives = () => {
            const narrativeChunks = chunkSheetText(
                sheet,
                "narrative",
                narrativeUnits(pendingNarratives, maxTokens, counter),
                narrativeNumber,
                maxTokens,
                counter
            );
            narrativeNumber += narrativeChunks.length;
            chunks.push(...narrativeChunks);
            pendingNarratives = [];
        };
        for (const entry of bodyInReadingOrder(sheet)) {
            if (entry.kind === "narrative") {
                pendingNarratives.push(entry.narrative);
                continue;
            }
            flushNarratives();
            chunks.push(...chunkTable(sheet, entry.table, options.tableFormat || "rows", maxTokens, counter));
        }
        flushNarratives();
        if (options.includeComments !== false) {
            const commentUnits = sheet.comments.map((comment) => ({
                text: renderCompactComment(comment),
                ref: comment.address,
            }));
            chunks.push(...chunkSheetText(sheet, "comments", commentUnits, 1, maxTokens, counter));
        }
        if (options.includeAssets !== false) {
            chunks.push(...chunkSheetText(sheet, "assets", assetUnits(sheet), 1, maxTokens, counter));
        }
    }
    return chunks;
}

/** Splits workbook IR into token-limited retrieval chunks. */
export { planContextChunks };
