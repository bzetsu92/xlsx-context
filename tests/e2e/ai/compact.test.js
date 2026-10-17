import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import * as aiCompactRenderer from "../../../src/outputs/ai-context/compact-renderer.ts";
import * as aiTokenCounter from "../../../src/outputs/ai-context/token-counter.ts";
import * as aiWorkbookIr from "../../../src/outputs/ai-context/workbook-ir.ts";
import { parseFixture, listFixtures, containsAnswer, toNaiveCsv } from "./helpers.js";

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
