/*
 * Baseline benchmark for xlsx-context (Excel -> AI Context foundation).
 *
 * Measures per-fixture: stage timings (parse / convert / payload / archive),
 * memory deltas, output sizes, and semantic fidelity counters from the public
 * package API. No new dependencies beyond tinybench; token figures are
 * whitespace estimates only (labeled as such, not model tokens).
 *
 * Usage:
 *   node scripts/bench-xlsx-context-baseline.mjs [--out <file.json>] [--iterations <n>] [--filter <substr>]
 */
import fs from "node:fs";
import path from "node:path";
import { performance } from "node:perf_hooks";
import { fileURLToPath } from "node:url";
import { Bench } from "tinybench";

import * as xlsxContext from "xlsx-context";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const ROOT = path.resolve(__dirname, "..");
const FIXTURE_ROOT = path.join(ROOT, "tests", "fixtures");

const DEFAULT_OPTIONS = {
    treatFirstRowAsHeader: true,
    trimText: true,
    removeEmptyRows: true,
    removeEmptyColumns: true,
    includeShapeDetails: false,
    outputMode: "display",
    formattingMode: "gfm",
    tableDetectionMode: "balanced",
};

function parseArgs(argv) {
    const args = { outPath: null, iterations: 10, filter: null };
    for (let index = 0; index < argv.length; index += 1) {
        const arg = argv[index];
        if (arg === "--out") {
            args.outPath = argv[index + 1] || null;
            index += 1;
        } else if (arg === "--iterations") {
            args.iterations = Math.max(1, Number.parseInt(argv[index + 1] || "10", 10) || 10);
            index += 1;
        } else if (arg === "--filter") {
            args.filter = argv[index + 1] || null;
            index += 1;
        } else if (arg === "--help") {
            console.log(
                "Usage: node scripts/bench-xlsx-context-baseline.mjs [--out <file.json>] [--iterations <n>] [--filter <substr>]"
            );
            process.exit(0);
        }
    }
    return args;
}

function listFixtures() {
    const out = [];
    const walk = (dir) => {
        for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
            const full = path.join(dir, entry.name);
            if (entry.isDirectory()) {
                walk(full);
            } else if (entry.name.endsWith(".xlsx")) {
                out.push(full);
            }
        }
    };
    walk(FIXTURE_ROOT);
    return out.sort();
}

function snapshotMemory() {
    const usage = process.memoryUsage();
    return { rss: usage.rss, heapUsed: usage.heapUsed, heapTotal: usage.heapTotal, external: usage.external };
}

function toArrayBuffer(buffer) {
    return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
}

function estimateWhitespaceTokens(text) {
    const matches = String(text || "").match(/\S+/g);
    return matches ? matches.length : 0;
}

function summarizeFormulaDiagnostics(files) {
    const counts = { resolved: 0, fallback_formula: 0, unsupported_external: 0 };
    for (const file of files) {
        for (const diagnostic of file.summary.formulaDiagnostics || []) {
            if (diagnostic.status in counts) {
                counts[diagnostic.status] += 1;
            }
        }
    }
    return counts;
}

async function measureFixture(fixturePath) {
    const inputBytes = fs.readFileSync(fixturePath);
    const workbookName = path.basename(fixturePath);
    const memBefore = snapshotMemory();

    const t0 = performance.now();
    const workbook = await xlsxContext.parseWorkbook(toArrayBuffer(inputBytes), { fileName: workbookName });
    const tParse = performance.now();

    const files = xlsxContext.convertWorkbookToMarkdown(workbook, { ...DEFAULT_OPTIONS });
    const tConvert = performance.now();

    const combined = xlsxContext.createCombinedMarkdown(workbook, files, {
        encoding: "utf-8",
        bom: "off",
        toolVersion: "bench",
        shapeDetails: "exclude",
        frontMatter: "include",
    });
    const tPayload = performance.now();

    const zipBytes = xlsxContext.createMarkdownArchive(workbook, files, {
        encoding: "utf-8",
        bom: "off",
        toolVersion: "bench",
        shapeDetails: "exclude",
        frontMatter: "include",
    });
    const tArchive = performance.now();
    const memAfter = snapshotMemory();

    // AI context is measured after the Markdown stages so existing timings stay comparable.
    const aiContext = {};
    for (const tableFormat of ["rows", "records"]) {
        const tAi = performance.now();
        const context = xlsxContext.toAIContext(workbook, { maxTokens: 512, tableFormat });
        aiContext[tableFormat] = {
            ms: round3(performance.now() - tAi),
            textChars: context.text.length,
            tokens: context.tokens,
            chunks: context.chunks.length,
            maxChunkTokens: context.chunks.reduce((max, chunk) => Math.max(max, chunk.tokens), 0),
        };
    }

    const combinedText = Buffer.from(combined.data).toString("utf8");
    const summary = {
        sections: 0,
        tables: 0,
        narrativeBlocks: 0,
        merges: 0,
        images: 0,
        charts: 0,
        comments: 0,
        cells: 0,
        tableScoreEntries: 0,
    };
    for (const file of files) {
        const s = file.summary;
        summary.sections += s.sections;
        summary.tables += s.tables;
        summary.narrativeBlocks += s.narrativeBlocks;
        summary.merges += s.merges;
        summary.images += s.images;
        summary.charts += s.charts;
        summary.comments += s.comments;
        summary.cells += s.cells;
        summary.tableScoreEntries += (s.tableScores || []).length;
    }

    return {
        fixture: path.relative(ROOT, fixturePath),
        inputBytes: inputBytes.length,
        sheets: workbook.sheets.length,
        timingsMs: {
            parse: round3(tParse - t0),
            convert: round3(tConvert - tParse),
            payload: round3(tPayload - tConvert),
            archive: round3(tArchive - tPayload),
            total: round3(tArchive - t0),
        },
        memory: {
            heapUsedDelta: memAfter.heapUsed - memBefore.heapUsed,
            rssDelta: memAfter.rss - memBefore.rss,
            heapUsedAfter: memAfter.heapUsed,
            rssAfter: memAfter.rss,
        },
        output: {
            markdownChars: combinedText.length,
            markdownLines: combinedText.split("\n").length,
            whitespaceTokenEstimate: estimateWhitespaceTokens(combinedText),
            zipBytes: zipBytes.length,
            files: files.length,
        },
        aiContext,
        semantics: {
            ...summary,
            formulas: summarizeFormulaDiagnostics(files),
        },
    };
}

function round3(value) {
    return Math.round(value * 1000) / 1000;
}

async function microBench(fixturePath, iterations) {
    const inputBytes = fs.readFileSync(fixturePath);
    const workbookName = path.basename(fixturePath);
    const bufferCopy = () => toArrayBuffer(Buffer.from(inputBytes));

    const benchParse = new Bench({ time: 500, iterations });
    benchParse.add("parseWorkbook", async () => {
        await xlsxContext.parseWorkbook(bufferCopy(), { fileName: workbookName });
    });
    await benchParse.run();

    const workbook = await xlsxContext.parseWorkbook(bufferCopy(), { fileName: workbookName });
    const benchConvert = new Bench({ time: 500, iterations });
    benchConvert.add("convertWorkbookToMarkdownFiles", () => {
        xlsxContext.convertWorkbookToMarkdown(workbook, { ...DEFAULT_OPTIONS });
    });
    await benchConvert.run();

    const compact = (bench) =>
        bench.tasks.map((task) => {
            const r = task.result;
            if (!r) return { name: task.name, error: "no-result" };
            return {
                name: task.name,
                latencyMs: { mean: round3(r.latency.mean), p50: round3(r.latency.p50), p99: round3(r.latency.p99) },
                throughputMean: Math.round(r.throughput.mean),
            };
        });
    return [...compact(benchParse), ...compact(benchConvert)];
}

function printHumanTable(results) {
    const rows = results.map((r) => ({
        fixture: r.fixture.replace("tests/fixtures/", ""),
        kb: (r.inputBytes / 1024).toFixed(1),
        parse: r.timingsMs.parse.toFixed(1),
        convert: r.timingsMs.convert.toFixed(1),
        total: r.timingsMs.total.toFixed(1),
        mdChars: r.output.markdownChars,
        tables: r.semantics.tables,
        cells: r.semantics.cells,
    }));
    console.log("fixture | inputKB | parseMs | convertMs | totalMs | mdChars | tables | cells");
    for (const row of rows) {
        console.log(
            `${row.fixture} | ${row.kb} | ${row.parse} | ${row.convert} | ${row.total} | ${row.mdChars} | ${row.tables} | ${row.cells}`
        );
    }
}

async function main() {
    const args = parseArgs(process.argv.slice(2));
    let fixtures = listFixtures();
    if (args.filter) {
        fixtures = fixtures.filter((f) => f.includes(args.filter));
    }
    if (fixtures.length === 0) {
        throw new Error("No fixtures matched.");
    }

    const results = [];
    for (const fixture of fixtures) {
        results.push(await measureFixture(fixture));
    }
    printHumanTable(results);

    const microTargets = [
        "tests/fixtures/table/table-basic-sample01.xlsx",
        "tests/fixtures/table/table-border-priority-sample01.xlsx",
        "tests/fixtures/formula/formula-basic-sample01.xlsx",
    ].filter((rel) => fs.existsSync(path.join(ROOT, rel)));
    const micro = {};
    for (const rel of microTargets) {
        micro[rel] = await microBench(path.join(ROOT, rel), args.iterations);
    }

    const totals = results.reduce(
        (acc, r) => {
            acc.inputBytes += r.inputBytes;
            acc.totalMs += r.timingsMs.total;
            acc.markdownChars += r.output.markdownChars;
            acc.cells += r.semantics.cells;
            for (const format of ["rows", "records"]) {
                acc.aiContext[format].tokens += r.aiContext[format].tokens;
                acc.aiContext[format].chunks += r.aiContext[format].chunks;
                acc.aiContext[format].ms += r.aiContext[format].ms;
            }
            return acc;
        },
        {
            inputBytes: 0,
            totalMs: 0,
            markdownChars: 0,
            cells: 0,
            aiContext: {
                rows: { tokens: 0, chunks: 0, ms: 0 },
                records: { tokens: 0, chunks: 0, ms: 0 },
            },
        }
    );

    const report = {
        meta: {
            tool: "xlsx-context-baseline",
            date: new Date().toISOString(),
            node: process.version,
            options: DEFAULT_OPTIONS,
            iterations: args.iterations,
            fixtureCount: results.length,
            note: "whitespaceTokenEstimate and aiContext.tokens (char-estimate counter, maxTokens 512 per chunk) are NOT model token counts; pass a TokenCounter adapter for exact budgets.",
        },
        totals: {
            ...totals,
            totalMs: round3(totals.totalMs),
            aiContext: {
                rows: { ...totals.aiContext.rows, ms: round3(totals.aiContext.rows.ms) },
                records: { ...totals.aiContext.records, ms: round3(totals.aiContext.records.ms) },
            },
        },
        fixtures: results,
        micro,
    };

    if (args.outPath) {
        const outAbs = path.resolve(args.outPath);
        fs.mkdirSync(path.dirname(outAbs), { recursive: true });
        fs.writeFileSync(outAbs, JSON.stringify(report, null, 2), "utf8");
        console.log(`\nWrote JSON report to ${outAbs}`);
    } else {
        console.log(`\nTotals: ${JSON.stringify(report.totals)}`);
    }
}

await main();
