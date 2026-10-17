import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import * as aiChunkPlanner from "../../../src/outputs/ai-context/chunk-planner.ts";
import * as aiCompactRenderer from "../../../src/outputs/ai-context/compact-renderer.ts";
import * as aiTokenCounter from "../../../src/outputs/ai-context/token-counter.ts";
import * as aiWorkbookIr from "../../../src/outputs/ai-context/workbook-ir.ts";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("planContextChunks()", () => {
    const renderRow = (row) => aiCompactRenderer.renderCompactRow(row);
    async function buildIr(fixtureRelPath) {
        const bytes = readFileSync(path.resolve(FIXTURE_ROOT, fixtureRelPath));
        const workbook = await core.parseWorkbook(
            bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
            path.basename(fixtureRelPath)
        );
        return aiWorkbookIr.buildWorkbookIR(workbook);
    }

    function listFixtures() {
        return readdirSync(FIXTURE_ROOT, { recursive: true })
            .map(String)
            .filter((name) => name.endsWith(".xlsx"))
            .sort();
    }

    const planner = () => aiChunkPlanner;

    it("keeps every table row exactly once, in order, under the budget, on every fixture", async () => {
        const counter = aiTokenCounter.createCharEstimateCounter();
        for (const fixture of listFixtures()) {
            const ir = await buildIr(fixture);
            for (const tableFormat of ["rows", "records"]) {
                for (const maxTokens of [24, 512]) {
                    const chunks = planner().planContextChunks(ir, { maxTokens, tableFormat });
                    expect(new Set(chunks.map((chunk) => chunk.id)).size, fixture).toBe(chunks.length);

                    for (const sheet of ir.sheets) {
                        for (const table of sheet.tables) {
                            const label = `${fixture} ${tableFormat} ${table.id}`;
                            const dataRows = table.hasHeader === false ? table.rows : table.rows.slice(1);
                            const tableChunks = chunks.filter(
                                (chunk) => chunk.sheet === sheet.name && chunk.tableId === table.id
                            );
                            // Records repeat no header; a records table without data rows keeps its header row.
                            // A table without a header row has data lines only.
                            const withHeader =
                                table.hasHeader !== false && (tableFormat === "rows" || dataRows.length === 0);
                            const prefix = [
                                `# ${sheet.name}`,
                                `## Table ${table.id} (${table.range})`,
                                ...(withHeader ? [renderRow(table.rows[0])] : []),
                            ].join("\n");
                            const dataLines = [];
                            for (const chunk of tableChunks) {
                                expect(chunk.text.startsWith(prefix), label).toBe(true);
                                expect(chunk.tokens).toBe(counter.count(chunk.text));
                                const lines = chunk.text.slice(prefix.length).split("\n").filter(Boolean);
                                if (lines.length > 1) {
                                    expect(chunk.tokens, label).toBeLessThanOrEqual(maxTokens);
                                }
                                dataLines.push(...lines);
                            }
                            if (tableFormat === "rows") {
                                expect(dataLines, label).toEqual(dataRows.map(renderRow));
                            } else {
                                const anchors = dataLines.map((line) => line.match(/^\[([A-Z]+\d+)\]/)?.[1]);
                                expect(anchors, label).toEqual(dataRows.map((row) => row[0].address));
                            }
                        }
                    }
                }
            }
        }
    });

    it("points table chunk ranges at the data rows they contain", async () => {
        const ir = await buildIr("table/table-basic-sample01.xlsx");
        const table = ir.sheets[0].tables.find((candidate) => candidate.rows.length > 3);
        const chunks = planner()
            .planContextChunks(ir, { maxTokens: 24 })
            .filter((chunk) => chunk.tableId === table.id && chunk.sheet === ir.sheets[0].name);

        expect(chunks.length).toBeGreaterThan(1);
        expect(chunks[0].id).toBe(`s1/${table.id}/1`);
        expect(chunks[0].source.range.startsWith(`${table.rows[1][0].address}-`)).toBe(true);
        const lastRow = table.rows[table.rows.length - 1];
        expect(chunks.at(-1).source.range.endsWith(`-${lastRow[lastRow.length - 1].address}`)).toBe(true);
    });

    it("is deterministic and honors a custom counter and includeComments", async () => {
        const ir = await buildIr("table/table-basic-sample01.xlsx");
        const first = planner().planContextChunks(ir, { maxTokens: 64 });
        expect(JSON.stringify(planner().planContextChunks(ir, { maxTokens: 64 }))).toBe(JSON.stringify(first));

        const lengthCounter = { name: "length", count: (text) => text.length };
        const byLength = planner().planContextChunks(ir, { maxTokens: 200, counter: lengthCounter });
        for (const chunk of byLength) {
            expect(chunk.tokens).toBe(chunk.text.length);
        }

        // Fixtures hold no comments (fixture-hygiene forbids them), so the IR gets one by hand.
        ir.sheets[0].comments = [{ address: "C4", kind: "note", author: "Reviewer", text: "check code" }];
        const withComments = planner()
            .planContextChunks(ir)
            .filter((chunk) => chunk.kind === "comments");
        expect(withComments).toHaveLength(1);
        expect(withComments[0].text).toContain("- [C4] Reviewer: check code");
        const noComments = planner().planContextChunks(ir, { includeComments: false });
        expect(noComments.some((chunk) => chunk.kind === "comments")).toBe(false);
    });

    it("emits chart and image entries as asset chunks with anchors", async () => {
        const ir = await buildIr("image/image-basic-sample02.xlsx");
        const chunks = planner().planContextChunks(ir);
        const assets = chunks.filter((chunk) => chunk.kind === "assets");

        expect(assets.map((chunk) => chunk.id)).toEqual(["s1/a/1"]);
        expect(assets[0].text).toContain("[B9] Chart c001: このグラフのタイトル (Line Chart)");
        expect(assets[0].text).toContain("- 値A: categories image!$B$4:$B$6; values image!$C$4:$C$6");
        expect(assets[0].text).toContain("[H3] Image i001: assets/image/image_001.png");
        expect(
            planner()
                .planContextChunks(ir, { includeAssets: false })
                .some((c) => c.kind === "assets")
        ).toBe(false);
    });

    it("rejects a non-positive budget", async () => {
        const ir = await buildIr("table/table-basic-sample01.xlsx");
        expect(() => planner().planContextChunks(ir, { maxTokens: 0 })).toThrow("maxTokens must be a positive number");
    });
});
