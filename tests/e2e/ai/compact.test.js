import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import * as aiCompactRenderer from "../../../src/outputs/ai-context/compact-renderer.ts";
import * as aiTokenCounter from "../../../src/outputs/ai-context/token-counter.ts";
import * as aiWorkbookIr from "../../../src/outputs/ai-context/workbook-ir.ts";
import { parseFixture, listFixtures, containsAnswer, normalize, toNaiveCsv } from "./helpers.js";
import * as aiChunkPlanner from "../../../src/outputs/ai-context/chunk-planner.ts";
import { parseCellAddress } from "../../../src/structure/addresses/address-utils.ts";

/*
 * Answerability benchmark: gold questions with known answers derived
 * from fixture contents. It measures information preservation per
 * token, not model behavior. A real LLM judge stays a later step.
 * Compact output must keep every answer the naive dump keeps.
 */
const QUESTIONS = [
    {
        fixture: ["table", "table-basic-sample01.xlsx"],
        asks: [
            { question: "default of altname", answers: ["Hanako"] },
            { question: "register dates of both tables", answers: ["3月13日", "3月15日"] },
            { question: "table count", answers: ["B3-F7", "B9-F13"] },
            { question: "narrative before first table", answers: ["隣接するテーブルのテスト（縦に密接）"] },
        ],
    },
    {
        fixture: ["formula", "formula-basic-sample01.xlsx"],
        // Each answer occurs once in the fixture. arith and sum both give 15, so they cannot prove a row stays.
        asks: [
            { question: "resolved if value", answers: ["OK"] },
            { question: "resolved text value", answers: ["0010"] },
            { question: "resolved value_num value", answers: ["1234.5"] },
            { question: "resolved date value", answers: ["2024/3/17"] },
            { question: "table range", answers: ["A3-B13"] },
        ],
    },
    {
        fixture: ["chart", "chart-basic-sample01.xlsx"],
        asks: [
            { question: "chart title", answers: ["棒グラフのグラフ"] },
            { question: "chart series name", answers: ["値A"] },
            { question: "chart series values ref", answers: ["'chart-basic'!$C$4:$C$7"] },
            { question: "table range", answers: ["B3-D7"] },
            { question: "chart data value", answers: ["28,053"] },
        ],
    },
];

/** True when the chunk source points at the cell. A table chunk also covers its repeated header row. */
function chunkCoversCell(ir, chunk, cell) {
    if (chunk.source.sheet !== cell.sheet) return false;
    const range = chunk.source.range;
    const rowSpan = range.match(/^(\d+):(\d+)$/);
    if (rowSpan) return cell.row >= Number(rowSpan[1]) && cell.row <= Number(rowSpan[2]);
    const rect = range.match(/^([A-Z]+\d+)-([A-Z]+\d+)$/);
    if (rect) {
        const start = parseCellAddress(rect[1]);
        const end = parseCellAddress(rect[2]);
        const table = ir.sheets.find((sheet) => sheet.name === chunk.sheet)?.tables.find((t) => t.id === chunk.tableId);
        const headerRow = table?.hasHeader ? parseCellAddress(table.range.split("-")[0]).row : -1;
        const inCols = cell.col >= start.col && cell.col <= end.col;
        return inCols && ((cell.row >= start.row && cell.row <= end.row) || cell.row === headerRow);
    }
    return range.split(",").includes(cell.address);
}

describe("compact AI context on fixtures", () => {
    it("keeps every gold answer in compact output", async () => {
        const api = core;
        const irApi = aiWorkbookIr;
        const compactApi = aiCompactRenderer;
        const missing = [];
        for (const group of QUESTIONS) {
            const workbook = await parseFixture(api, ...group.fixture);
            const compact = compactApi.renderCompactMarkdown(irApi.buildWorkbookIR(workbook));
            for (const ask of group.asks) {
                for (const answer of ask.answers) {
                    if (!containsAnswer(compact, answer)) {
                        missing.push(`${group.fixture.join("/")} [${ask.question}]: ${answer}`);
                    }
                }
            }
        }
        expect(missing).toEqual([]);
    }, 120000);

    it("reports coverage per token for compact versus naive dump", async () => {
        const api = core;
        const irApi = aiWorkbookIr;
        const compactApi = aiCompactRenderer;
        const counterApi = aiTokenCounter;
        const counter = counterApi.createWhitespaceCounter();
        const rows = [];
        for (const group of QUESTIONS) {
            const workbook = await parseFixture(api, ...group.fixture);
            const compact = compactApi.renderCompactMarkdown(irApi.buildWorkbookIR(workbook));
            const csv = toNaiveCsv(workbook);
            const total = group.asks.reduce((sum, ask) => sum + ask.answers.length, 0);
            const compactHits = group.asks.reduce(
                (sum, ask) => sum + ask.answers.filter((a) => containsAnswer(compact, a)).length,
                0
            );
            const csvHits = group.asks.reduce(
                (sum, ask) => sum + ask.answers.filter((a) => containsAnswer(csv, a)).length,
                0
            );
            rows.push({
                fixture: group.fixture.join("/"),
                compact: `${compactHits}/${total} @ ${counter.count(compact)}tok`,
                csv: `${csvHits}/${total} @ ${counter.count(csv)}tok`,
            });
            expect(compactHits).toBe(total);
        }
        // Opt-in report, so a normal run stays quiet.
        if (process.env.QA_REPORT) console.log(JSON.stringify(rows, null, 1));
    }, 120000);

    it("finds every gold answer in a chunk whose source covers the answer cell", async () => {
        // Small chunks, so retrieval must pick the right piece, as in a real RAG step.
        const problems = [];
        for (const group of QUESTIONS) {
            const workbook = await parseFixture(core, ...group.fixture);
            const ir = aiWorkbookIr.buildWorkbookIR(workbook);
            const chunks = aiChunkPlanner.planContextChunks(ir, { maxTokens: 64 });
            for (const ask of group.asks) {
                for (const answer of ask.answers) {
                    const label = `${group.fixture.join("/")} [${ask.question}]: ${answer}`;
                    const hits = chunks.filter((chunk) => containsAnswer(chunk.text, answer));
                    if (hits.length === 0) {
                        problems.push(`${label}: in no chunk`);
                        continue;
                    }
                    // Answers that are cell values must be cited at their cell. Ranges and chart refs are not cells.
                    const cells = workbook.sheets.flatMap((sheet) =>
                        sheet.cells
                            .filter((cell) => normalize(cell.outputValue) === normalize(answer))
                            .map((cell) => ({ sheet: sheet.name, row: cell.row, col: cell.col, address: cell.address }))
                    );
                    if (cells.length === 0) continue;
                    const cited = hits.some((chunk) => cells.some((cell) => chunkCoversCell(ir, chunk, cell)));
                    if (!cited) {
                        problems.push(
                            `${label}: ${hits.map((chunk) => `${chunk.id}@${chunk.source.range}`).join(" ")}`
                        );
                    }
                }
            }
        }
        expect(problems).toEqual([]);
    }, 120000);

    it("renders compact output smaller than faithful markdown on every fixture", async () => {
        const api = core;
        const irApi = aiWorkbookIr;
        const compactApi = aiCompactRenderer;
        const failures = [];
        for (const fixturePath of listFixtures()) {
            const bytes = readFileSync(fixturePath);
            const arrayBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
            const workbook = await api.parseWorkbook(arrayBuffer, path.basename(fixturePath));
            const files = api.convertWorkbookToMarkdownFiles(workbook, {
                treatFirstRowAsHeader: true,
                trimText: true,
                removeEmptyRows: true,
                removeEmptyColumns: true,
                // Compact always renders assets, so compare fairly.
                includeShapeDetails: true,
                outputMode: "display",
                formattingMode: "gfm",
                tableDetectionMode: "balanced",
            });
            const faithful = files.map((file) => file.markdown).join("\n").length;
            const compact = compactApi.renderCompactMarkdown(irApi.buildWorkbookIR(workbook)).length;
            if (!(compact < faithful)) {
                failures.push(`${path.basename(fixturePath)}: compact=${compact} faithful=${faithful}`);
            }
        }
        expect(failures).toEqual([]);
    }, 180000);
});
