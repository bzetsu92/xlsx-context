import { mkdirSync, writeFileSync } from "node:fs";
import path from "node:path";

import { afterAll, describe, expect, it } from "vitest";
import * as aiWorkbookIr from "../../src/outputs/ai-context/workbook-ir.ts";
import * as tableDetector from "../../src/structure/tables/table-detector.ts";
import { ROOT } from "../helpers/paths.js";

const ROWS = 20000;
const COLS = 10;

function makeBigSheet() {
    const cells = [];
    for (let row = 1; row <= ROWS; row += 1) {
        for (let col = 1; col <= COLS; col += 1) {
            cells.push({
                address: `R${row}C${col}`,
                row,
                col,
                valueType: "s",
                rawValue: `r${row}c${col}`,
                outputValue: `r${row}c${col}`,
                formulaText: "",
                resolutionStatus: null,
                resolutionSource: null,
                cachedValueState: null,
                styleIndex: 0,
                borders: { top: false, bottom: false, left: false, right: false },
                numFmtId: 0,
                formatCode: "",
                textStyle: { bold: false, italic: false, strike: false, underline: false },
                richTextRuns: null,
                formulaType: "",
                spillRef: "",
                hyperlink: null,
            });
        }
    }
    return {
        name: "big",
        index: 0,
        path: "xl/worksheets/sheet1.xml",
        cells,
        merges: [],
        tables: [],
        images: [],
        charts: [],
        shapes: [],
        comments: [],
        maxRow: ROWS,
        maxCol: COLS,
    };
}

// A step fails when it goes over one of these limits.
// About 5x the values measured on a local machine (2026-10), so a slower CI machine still passes.
const BUDGET = {
    detectTableCandidates: { wallMs: 1500, cpuMs: 2000, heapMb: 150 },
    buildWorkbookIR: { wallMs: 2000, cpuMs: 2500, heapMb: 300 },
};
const REPORT = path.join(ROOT, "bench-results", "performance-report.json");
const MB = 1024 * 1024;
const results = [];

function measure(step, run) {
    const heapBefore = process.memoryUsage().heapUsed;
    const cpuBefore = process.cpuUsage();
    const started = performance.now();
    const value = run();
    const wallMs = performance.now() - started;
    const cpu = process.cpuUsage(cpuBefore);
    const heapMb = (process.memoryUsage().heapUsed - heapBefore) / MB;
    const result = {
        step,
        rows: ROWS,
        wallMs: Math.round(wallMs),
        cpuMs: Math.round((cpu.user + cpu.system) / 1000),
        heapMb: Math.round(heapMb),
        peakRssMb: Math.round(process.resourceUsage().maxRSS / 1024),
        rowsPerSecond: Math.round((ROWS / wallMs) * 1000),
        budget: BUDGET[step],
    };
    result.pass =
        result.wallMs <= result.budget.wallMs &&
        result.cpuMs <= result.budget.cpuMs &&
        result.heapMb <= result.budget.heapMb;
    results.push(result);
    return { value, result };
}

describe(`large sheet (${ROWS} rows x ${COLS} columns)`, () => {
    afterAll(() => {
        console.table(results.map(({ budget, ...row }) => row));
        mkdirSync(path.dirname(REPORT), { recursive: true });
        writeFileSync(
            REPORT,
            `${JSON.stringify({ node: process.version, platform: process.platform, results }, null, 2)}\n`
        );
        console.log(`Report: ${path.relative(ROOT, REPORT)}`);
    });

    it("detectTableCandidates() stays within the CPU, RAM, and time budget", () => {
        const sheet = makeBigSheet();
        const { value, result } = measure("detectTableCandidates", () =>
            tableDetector.detectTableCandidates(
                sheet,
                (target) => new Map(target.cells.map((cell) => [`${cell.row}:${cell.col}`, cell]))
            )
        );
        expect(value.length).toBeGreaterThan(0);
        expect(result).toMatchObject({ pass: true });
    }, 60000);

    it("buildWorkbookIR() stays within the CPU, RAM, and time budget", () => {
        const workbook = { name: "big.xlsx", sheets: [makeBigSheet()], sharedStrings: [], definedNames: [] };
        const { value, result } = measure("buildWorkbookIR", () => aiWorkbookIr.buildWorkbookIR(workbook));
        expect(value.sheets[0].tables.length).toBeGreaterThan(0);
        expect(result).toMatchObject({ pass: true });
    }, 90000);
});
