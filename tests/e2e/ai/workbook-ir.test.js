import { readdirSync } from "node:fs";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import * as aiCompactRenderer from "../../../src/outputs/ai-context/compact-renderer.ts";
import * as aiWorkbookIr from "../../../src/outputs/ai-context/workbook-ir.ts";
import { formatRange } from "../../../src/structure/addresses/address-utils.ts";
import { parseFixture } from "./helpers.js";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("buildWorkbookIR()", () => {
    it("keeps raw and display values without dropping cells", async () => {
        const api = core;
        const irApi = aiWorkbookIr;
        const workbook = await parseFixture(api, "table/table-basic-sample01.xlsx");

        const ir = irApi.buildWorkbookIR(workbook);

        const sheet = ir.sheets[0];
        const gridCells = sheet.tables.flatMap((table) => table.rows.flat());
        // Two 5x5 tables, B3:F7 and B9:F13, so every cell of both ranges must be present once.
        const expectedAddresses = [3, 4, 5, 6, 7, 9, 10, 11, 12, 13].flatMap((row) =>
            ["B", "C", "D", "E", "F"].map((col) => `${col}${row}`)
        );
        expect(gridCells.map((cell) => cell.address)).toEqual(expectedAddresses);
        // E6 stores serial 46094 with a date format.
        expect(gridCells.find((cell) => cell.address === "E6")).toMatchObject({ raw: "46094", display: "3月13日" });
    });

    it("keeps narrative text free of Markdown escapes and heading anchors", async () => {
        const api = core;
        const irApi = aiWorkbookIr;
        const spill = irApi.buildWorkbookIR(await parseFixture(api, "formula/formula-spill-sample01.xlsx"));
        const link = irApi.buildWorkbookIR(await parseFixture(api, "hyperlink/hyperlink-basic-sample01.xlsx"));

        const spillLines = spill.sheets[0].narratives.flatMap((narrative) => narrative.lines);
        const linkLines = link.sheets[0].narratives.flatMap((narrative) => narrative.lines);
        expect(spillLines).toContain("spill_sum");
        expect(linkLines).toContain("[Open example](https://example.com/)");
        expect(linkLines).toContain("Jump to Other (Other!A1)");
    });

    it("keeps each caption with its own table in Markdown and in the IR reading order", async () => {
        const api = core;
        const irApi = aiWorkbookIr;
        const workbook = await parseFixture(api, "table/table-basic-sample02.xlsx");

        const markdown = api.convertWorkbookToMarkdownFiles(workbook, { formattingMode: "gfm" })[0].markdown;
        const [left, right] = markdown.split("\n---\n");
        expect(left).toContain("**隣接するテーブルその1**\n\n### Table: 001 (B3-F7)");
        expect(left).not.toContain("その2");
        expect(right).toContain("**隣接するテーブルその2**\n\n### Table: 002 (H3-L7)");

        const sheet = irApi.buildWorkbookIR(workbook).sheets[0];
        const readable = sheet.sections.map((section) =>
            section.map((entry) => (entry.kind === "table" ? entry.id : sheet.narratives[entry.index].lines[0]))
        );
        expect(readable[0].slice(0, 3)).toEqual([
            "隣接するテーブルのテスト（横に密接）",
            "隣接するテーブルその1",
            "t001",
        ]);
        expect(readable[1]).toEqual(["隣接するテーブルその2", "t002"]);
    });

    it("marks a first row of data as no header in Markdown and the IR", async () => {
        const api = core;
        const irApi = aiWorkbookIr;
        const keyValue = await parseFixture(api, "formula/formula-basic-sample01.xlsx");
        const yearChart = await parseFixture(api, "chart/chart-mixed-sample01.xlsx");

        const markdown = api.convertWorkbookToMarkdownFiles(keyValue, {})[0].markdown;
        expect(markdown).toContain("|  |  |\n| --- | --- |\n| base1 | 10 |");
        const table = irApi.buildWorkbookIR(keyValue).sheets[0].tables[0];
        expect(table.hasHeader).toBe(false);
        expect(table.headers).toEqual([]);
        expect(irApi.buildWorkbookIR(yearChart).sheets[0].tables[0].hasHeader).toBe(true);
    });

    it("joins spacer-separated columns only when they cover the same rows", async () => {
        const api = core;
        const spill = await parseFixture(api, "formula/formula-spill-sample01.xlsx");
        const named = await parseFixture(api, "named-range/named-range-sample01.xlsx");
        const shared = await parseFixture(api, "formula/formula-shared-sample01.xlsx");

        const ranges = (workbook) =>
            api
                .detectTableCandidates(workbook.sheets[0])
                .map((table) => formatRange(table.startRow, table.startCol, table.endRow, table.endCol));
        expect(ranges(spill)).toEqual(["A3-C6"]);
        expect(ranges(named)).toEqual(["A3-B5"]);
        expect(ranges(shared)).toEqual(["A1-B11"]);
    });

    it("builds deterministically across runs", async () => {
        const api = core;
        const irApi = aiWorkbookIr;
        const workbook = await parseFixture(api, "table/table-basic-sample01.xlsx");

        const first = irApi.buildWorkbookIR(workbook);
        const second = irApi.buildWorkbookIR(workbook);

        expect(JSON.stringify(second)).toBe(JSON.stringify(first));
    });

    it("matches Markdown table ids and ranges on every fixture", async () => {
        const api = core;
        const irApi = aiWorkbookIr;
        const fixtureRoot = FIXTURE_ROOT;
        const fixtures = readdirSync(fixtureRoot, { recursive: true })
            .filter((name) => String(name).endsWith(".xlsx"))
            .map(String)
            .sort();
        expect(fixtures.length).toBeGreaterThan(0);

        for (const fixture of fixtures) {
            const workbook = await parseFixture(api, fixture);
            const files = api.convertWorkbookToMarkdownFiles(workbook);
            const ir = irApi.buildWorkbookIR(workbook);
            ir.sheets.forEach((sheet, index) => {
                const markdownTables = [...files[index].markdown.matchAll(/^### Table: (\d{3}) \(([^)]+)\)$/gm)].map(
                    (match) => `t${match[1]} ${match[2]}`
                );
                const irTables = sheet.tables.map((table) => `${table.id} ${table.range}`);
                expect(irTables, `${fixture} / ${sheet.name}`).toEqual(markdownTables);
            });
        }
    });

    it("keeps external and internal hyperlinks on cells and in compact text", async () => {
        const workbook = await parseFixture(core, "hyperlink/hyperlink-basic-sample01.xlsx");
        const ir = aiWorkbookIr.buildWorkbookIR(workbook);
        const cells = ir.sheets.flatMap((sheet) => sheet.tables.flatMap((table) => table.rows.flat()));

        expect(cells.find((cell) => cell.address === "B5")?.hyperlink).toEqual({
            kind: "external",
            target: "https://example.com/docs",
            tooltip: "",
        });
        expect(cells.find((cell) => cell.address === "B6")?.hyperlink).toMatchObject({
            kind: "internal",
            target: "Other!A1",
        });
        expect(cells.filter((cell) => !cell.hyperlink).every((cell) => cell.hyperlink === null)).toBe(true);

        const text = aiCompactRenderer.renderCompactMarkdown(ir);
        expect(text).toContain("[Open example](https://example.com/docs)");
        expect(text).toContain("Jump to Other (Other!A1)");
    });

    it("keeps charts with series ranges and images with export paths", async () => {
        const workbook = await parseFixture(core, "image/image-basic-sample02.xlsx");
        const sheet = aiWorkbookIr.buildWorkbookIR(workbook).sheets[0];

        expect(sheet.charts).toEqual([
            {
                id: "c001",
                anchor: "B9",
                title: "このグラフのタイトル",
                type: "Line Chart",
                series: [
                    {
                        name: "値A",
                        categories: "image!$B$4:$B$6",
                        values: "image!$C$4:$C$6",
                        cachedCategories: ["2024年", "2025年", "2026年"],
                        cachedValues: [13568, 19287, 25051],
                        sourceTable: "t001",
                        axis: "primary",
                    },
                    {
                        name: "値B",
                        categories: "image!$B$4:$B$6",
                        values: "image!$D$4:$D$6",
                        cachedCategories: ["2024年", "2025年", "2026年"],
                        cachedValues: [9072, 18058, 32012],
                        sourceTable: "t001",
                        axis: "primary",
                    },
                ],
            },
        ]);
        expect(sheet.images).toEqual([
            { id: "i001", anchor: "H3", filename: "image_001.png", path: "assets/image/image_001.png" },
        ]);
    });

    it("links flowchart connectors to the shapes they join", async () => {
        const workbook = await parseFixture(core, "shape/shape-flowchart-sample01.xlsx");
        const shapes = aiWorkbookIr.buildWorkbookIR(workbook).sheets[0].shapes;

        expect(shapes.map((shape) => [shape.id, shape.connection])).toEqual([
            ["sh001", null],
            ["sh002", null],
            ["sh003", null],
            ["sh004", null],
            ["sh005", { from: "sh001", to: "sh002", arrow: "end" }],
            ["sh006", { from: "sh002", to: "sh003", arrow: "end" }],
            ["sh007", { from: "sh003", to: "sh004", arrow: "end" }],
        ]);
    });

    it("keeps shape text", async () => {
        const workbook = await parseFixture(core, "shape/shape-basic-sample01.xlsx");
        const shapes = aiWorkbookIr.buildWorkbookIR(workbook).sheets[0].shapes;

        // Text box, arrow connector, and rectangle in xl/drawings/drawing1.xml. Only the text box has text.
        expect(shapes.map((shape) => [shape.id, shape.text])).toEqual([
            ["sh001", "テキストボックスの例"],
            ["sh002", ""],
            ["sh003", ""],
        ]);
    });

    it("matches Markdown chart and image ids and anchors on every fixture", async () => {
        const api = core;
        const irApi = aiWorkbookIr;
        const fixtureRoot = FIXTURE_ROOT;
        const fixtures = readdirSync(fixtureRoot, { recursive: true })
            .map(String)
            .filter((name) => name.endsWith(".xlsx"))
            .sort();
        let checked = 0;

        for (const fixture of fixtures) {
            const workbook = await parseFixture(api, fixture);
            const files = api.convertWorkbookToMarkdownFiles(workbook);
            irApi.buildWorkbookIR(workbook).sheets.forEach((sheet, index) => {
                for (const [kind, prefix, assets] of [
                    ["Chart", "c", sheet.charts],
                    ["Image", "i", sheet.images],
                ]) {
                    const headings = [
                        ...files[index].markdown.matchAll(new RegExp(`^### ${kind}: (\\d{3}) \\(([^)]+)\\)$`, "gm")),
                    ].map((match) => `${prefix}${match[1]} ${match[2]}`);
                    expect(
                        assets.map((asset) => `${asset.id} ${asset.anchor}`).sort(),
                        `${fixture} / ${sheet.name} / ${kind}`
                    ).toEqual(headings.sort());
                    checked += headings.length;
                }
            });
        }
        expect(checked).toBeGreaterThan(0);
    });

    it("orders comments by row, then column", async () => {
        const workbook = await parseFixture(core, "table/table-basic-sample01.xlsx");
        const comment = (address) => ({ address, kind: "note", author: "", text: address });
        workbook.sheets[0].comments = [comment("B10"), comment("C2"), comment("B2"), comment("AA1")];

        const ir = aiWorkbookIr.buildWorkbookIR(workbook);

        expect(ir.sheets[0].comments.map((entry) => entry.address)).toEqual(["AA1", "B2", "C2", "B10"]);
    });

    it("explains why each table range was chosen", async () => {
        const workbook = await parseFixture(core, "table/table-basic-sample01.xlsx");

        const [table] = aiWorkbookIr.buildWorkbookIR(workbook, { tableDetectionMode: "border" }).sheets[0].tables;

        expect(table.detection.mode).toBe("border");
        expect(table.detection.score).toBeGreaterThan(0);
        expect(table.detection.reasons).toContain("Has borders (+3)");
    });
});
