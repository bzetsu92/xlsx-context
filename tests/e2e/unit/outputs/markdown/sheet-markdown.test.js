import { describe, expect, it } from "vitest";
import * as sheetMarkdown from "../../../../../src/outputs/markdown/sheet-markdown.ts";

function createDeps(overrides = {}) {
    return {
        renderNarrativeBlock: (block) => `### ${block.lines[0]}`,
        detectTableCandidates: () => [],
        matrixFromCandidate: () => [],
        renderMarkdownTable: (rows) => rows.map((row) => `| ${row.join(" | ")} |`).join("\n"),
        createOutputFileName: (
            _workbookName,
            sheetIndex,
            sheetName,
            _outputMode = "display",
            formattingMode = "plain"
        ) => `${sheetIndex}_${sheetName}${formattingMode === "plain" ? "" : `_${formattingMode}`}.md`,
        extractShapeBlocks: () => [],
        renderHierarchicalRawEntries: () => [],
        parseCellAddress: (address) => {
            const match = String(address || "").match(/^([A-Z]+)(\d+)$/i);
            if (!match) return { row: 0, col: 0 };
            let col = 0;
            for (const ch of match[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
            return { row: Number(match[2]), col };
        },
        formatRange: (startRow, startCol, endRow, endCol) => `${startRow}:${startCol}-${endRow}:${endCol}`,
        colToLetters: (col) => String.fromCharCode(64 + col),
        normalizeMarkdownText: (text) =>
            String(text || "")
                .replace(/\r\n?|\n/g, " ")
                .replace(/\s+/g, " ")
                .trim(),
        defaultCellWidthEmu: 1,
        defaultCellHeightEmu: 1,
        shapeBlockGapXEmu: 1,
        shapeBlockGapYEmu: 1,
        ...overrides,
    };
}

describe("xlsx-context sheet markdown", () => {
    it("builds cell maps and narrative row segments", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const sheet = {
            cells: [
                { row: 1, col: 1, outputValue: "A", rawValue: "A" },
                { row: 1, col: 2, outputValue: "B", rawValue: "B" },
                { row: 1, col: 8, outputValue: "C", rawValue: "C" },
            ],
        };

        expect(api.buildCellMap(sheet).get("1:2")?.outputValue).toBe("B");
        expect(api.splitNarrativeRowSegments(sheet.cells, {})).toEqual([
            { startCol: 1, values: ["A", "B"] },
            { startCol: 8, values: ["C"] },
        ]);
    });

    it("collects narrative cells by row and builds narrative items in row order", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            cells: [
                { row: 2, col: 4, outputValue: "Tail", rawValue: "Tail" },
                { row: 1, col: 1, outputValue: "Head", rawValue: "Head" },
                { row: 2, col: 1, outputValue: "Body", rawValue: "Body" },
                { row: 5, col: 1, outputValue: "TableCell", rawValue: "TableCell" },
            ],
        };
        const tables = [{ startRow: 5, startCol: 1, endRow: 5, endCol: 1 }];

        const rowMap = api.collectNarrativeCellsByRow(sheet, tables);
        const items = api.buildNarrativeItems(workbook, sheet, tables, {});

        expect(Array.from(rowMap.keys())).toEqual([2, 1]);
        expect(items).toEqual([
            { row: 1, startCol: 1, text: "Head", cellValues: ["Head"] },
            { row: 2, startCol: 1, text: "Body Tail", cellValues: ["Body", "Tail"] },
        ]);
    });

    it("extracts narrative blocks outside detected tables", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            cells: [
                { row: 1, col: 1, outputValue: "Heading", rawValue: "Heading" },
                { row: 2, col: 2, outputValue: "Detail", rawValue: "Detail" },
                { row: 5, col: 1, outputValue: "TableCell", rawValue: "TableCell" },
            ],
            charts: [],
            images: [],
        };
        const tables = [{ startRow: 5, startCol: 1, endRow: 5, endCol: 1 }];

        expect(api.extractNarrativeBlocks(workbook, sheet, tables, {})).toHaveLength(1);
        expect(api.extractNarrativeBlocks(workbook, sheet, tables, {})[0].lines).toEqual(["Heading", "Detail"]);
    });

    it("groups nearby section anchors and splits distant ones", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const sheet = {
            cells: [],
            images: [
                { anchor: "B2", filename: "img1.png", path: "assets/img1.png", data: new Uint8Array([1]) },
                { anchor: "D5", filename: "img2.png", path: "assets/img2.png", data: new Uint8Array([2]) },
                { anchor: "A12", filename: "img3.png", path: "assets/img3.png", data: new Uint8Array([3]) },
            ],
            charts: [{ anchor: "C13", title: "Chart", chartType: "bar", series: [] }],
        };
        const tables = [{ startRow: 4, startCol: 1, endRow: 6, endCol: 3, score: 1, reasonSummary: ["test"] }];
        const narrativeBlocks = [
            {
                startRow: 1,
                startCol: 1,
                endRow: 1,
                lines: ["Intro"],
                items: [{ row: 1, startCol: 1, text: "Intro", cellValues: ["Intro"] }],
            },
        ];

        expect(api.extractSectionBlocks(sheet, tables, narrativeBlocks)).toEqual([
            { startRow: 1, startCol: 1, endRow: 6, endCol: 4 },
            { startRow: 12, startCol: 1, endRow: 13, endCol: 3 },
        ]);
    });

    it("splits same-row section anchors when they are far apart horizontally", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const sheet = {
            cells: [],
            images: [],
            charts: [],
        };
        const tables = [];
        const narrativeBlocks = [
            {
                startRow: 1,
                startCol: 1,
                endRow: 1,
                lines: ["イベント チェックリスト"],
                items: [
                    { row: 1, startCol: 1, text: "イベント チェックリスト", cellValues: ["イベント チェックリスト"] },
                ],
            },
            {
                startRow: 1,
                startCol: 8,
                endRow: 1,
                lines: ["イベント カテゴリ"],
                items: [{ row: 1, startCol: 8, text: "イベント カテゴリ", cellValues: ["イベント カテゴリ"] }],
            },
        ];

        expect(api.extractSectionBlocks(sheet, tables, narrativeBlocks)).toEqual([
            { startRow: 1, startCol: 1, endRow: 1, endCol: 1 },
            { startRow: 1, startCol: 8, endRow: 1, endCol: 8 },
        ]);
    });

    it("keeps nearby calendar-like anchors in one section block even when horizontally separated", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const sheet = {
            cells: [],
            images: [],
            charts: [],
        };
        const tables = [];
        const narrativeBlocks = [
            {
                startRow: 11,
                startCol: 3,
                endRow: 17,
                lines: ["2021-01-03 ..."],
                items: [
                    {
                        row: 11,
                        startCol: 3,
                        text: "2021-01-03 ...",
                        cellValues: [
                            "2021-01-03",
                            "2021-01-04",
                            "2021-01-05",
                            "2021-01-06",
                            "2021-01-07",
                            "2021-01-08",
                            "2021-01-09",
                        ],
                    },
                ],
            },
            {
                startRow: 24,
                startCol: 25,
                endRow: 30,
                lines: ["2020-12-01 ..."],
                items: [
                    {
                        row: 24,
                        startCol: 25,
                        text: "2020-12-01 ...",
                        cellValues: ["2020-12-01", "2020-12-02", "2020-12-03", "2020-12-04", "2020-12-05"],
                    },
                ],
            },
        ];

        expect(api.extractSectionBlocks(sheet, tables, narrativeBlocks)).toEqual([
            { startRow: 11, startCol: 3, endRow: 30, endCol: 25 },
        ]);
    });

    it("sorts and renders grouped sections with separators only between blocks", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const sections = [
            { sortRow: 12, sortCol: 1, markdown: "Third\n", kind: "narrative" },
            { sortRow: 2, sortCol: 3, markdown: "Second\n", kind: "table" },
            { sortRow: 2, sortCol: 1, markdown: "First\n", kind: "narrative" },
        ];
        const sectionBlocks = [
            { startRow: 1, startCol: 1, endRow: 5, endCol: 5 },
            { startRow: 10, startCol: 1, endRow: 15, endCol: 5 },
        ];

        api.sortContentSections(sections);
        const grouped = api.createGroupedSections(sectionBlocks, sections);

        expect(grouped).toEqual([
            {
                block: { startRow: 1, startCol: 1, endRow: 5, endCol: 5 },
                entries: [
                    { sortRow: 2, sortCol: 1, markdown: "First\n", kind: "narrative" },
                    { sortRow: 2, sortCol: 3, markdown: "Second\n", kind: "table" },
                ],
            },
            {
                block: { startRow: 10, startCol: 1, endRow: 15, endCol: 5 },
                entries: [{ sortRow: 12, sortCol: 1, markdown: "Third\n", kind: "narrative" }],
            },
        ]);
        expect(api.renderGroupedSectionBody(grouped)).toBe("First\n\nSecond\n\n---\n\nThird");
    });

    it("reorders calendar-like grouped entries to place header and sidebar around the main block", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const header = {
            sortRow: 2,
            sortCol: 3,
            markdown: "2021年1月\n",
            kind: "narrative",
            narrativeBlock: {
                startRow: 2,
                startCol: 3,
                endRow: 2,
                lines: ["2021年1月"],
                items: [{ row: 2, startCol: 3, text: "2021年1月", cellValues: ["2021年1月"] }],
            },
        };
        const main = {
            sortRow: 11,
            sortCol: 3,
            markdown: "2021-01-03 | 2021-01-04 | 2021-01-05 | 2021-01-06 | 2021-01-07 | 2021-01-08 | 2021-01-09\n",
            kind: "narrative",
            narrativeBlock: {
                startRow: 11,
                startCol: 3,
                endRow: 11,
                lines: ["2021-01-03 ..."],
                items: [
                    {
                        row: 11,
                        startCol: 3,
                        text: "2021-01-03 ...",
                        cellValues: [
                            "2021-01-03",
                            "2021-01-04",
                            "2021-01-05",
                            "2021-01-06",
                            "2021-01-07",
                            "2021-01-08",
                            "2021-01-09",
                        ],
                    },
                ],
            },
        };
        const sidebar = {
            sortRow: 24,
            sortCol: 25,
            markdown: "2020-12-01\n",
            kind: "narrative",
            narrativeBlock: {
                startRow: 24,
                startCol: 25,
                endRow: 24,
                lines: ["2020-12-01"],
                items: [
                    {
                        row: 24,
                        startCol: 25,
                        text: "2020-12-01",
                        cellValues: ["2020-12-01", "2020-12-02", "2020-12-03", "2020-12-04", "2020-12-05"],
                    },
                ],
            },
        };

        const reordered = api.createCalendarAwareSectionEntries([sidebar, main, header]);

        expect(reordered.map((entry) => entry.markdown.trim())).toEqual([
            "2021年1月",
            "2021-01-03 | 2021-01-04 | 2021-01-05 | 2021-01-06 | 2021-01-07 | 2021-01-08 | 2021-01-09",
            "### Sidebar",
            "2020-12-01",
        ]);
    });

    it("renders image and chart sections with stable headings and metadata", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const sheet = {
            images: [
                {
                    anchor: "B2",
                    filename: "image_001.png",
                    path: "assets/Sheet1/image_001.png",
                    data: new Uint8Array([1]),
                },
            ],
        };
        const charts = [
            {
                anchor: "C4",
                title: "",
                chartType: "line",
                series: [
                    {
                        name: "Sales",
                        categoriesRef: "Sheet1!A2:A5",
                        valuesRef: "Sheet1!B2:B5",
                        cachedCategories: [],
                        cachedValues: [],
                        axis: "secondary",
                    },
                ],
            },
        ];

        const entries = api.createAssetEntries(sheet, charts, [], [], { includeShapeDetails: false });
        const chartSection = entries[0].markdown;
        const imageSection = entries[1].markdown;

        expect(entries.map((entry) => [entry.kind, entry.sortRow, entry.sortCol])).toEqual([
            ["asset", 4, 3],
            ["asset", 2, 2],
        ]);
        expect(imageSection).toContain("### Image: 001 (B2)");
        expect(imageSection).toContain("- File: assets/Sheet1/image_001.png");
        expect(imageSection).toContain("![image_001.png](assets/Sheet1/image_001.png)");
        expect(chartSection).toContain("### Chart: 001 (C4)");
        expect(chartSection).toContain("- Title: (none)");
        expect(chartSection).toContain("- Type: line");
        expect(chartSection).toContain("- Series:");
        expect(chartSection).toContain("  - Sales");
        expect(chartSection).toContain("    - Axis: secondary");
        expect(chartSection).toContain("    - categories: Sheet1!A2:A5");
        expect(chartSection).toContain("    - values: Sheet1!B2:B5");
        // No source link and no cache: nothing to add.
        expect(chartSection).not.toContain("- Source:");
        expect(chartSection).not.toContain("- Data (chart cache):");
    });

    it("renders cached chart data when no table on the sheet holds the series", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const charts = [
            {
                anchor: "B2",
                title: "Mix",
                chartType: "bar",
                series: [
                    {
                        name: "Sales",
                        categoriesRef: "Data!$A$2:$A$4",
                        valuesRef: "Data!$B$2:$B$4",
                        cachedCategories: ["Q1", "Q2", "Q3"],
                        cachedValues: [10, null, 30.5],
                        axis: "primary",
                    },
                    {
                        name: "Cost",
                        categoriesRef: "",
                        valuesRef: "",
                        cachedCategories: [],
                        cachedValues: [1, 2, 3, 4],
                        axis: "primary",
                    },
                ],
            },
        ];

        const [entry] = api.createAssetEntries({ images: [] }, charts, [], [], {
            includeShapeDetails: false,
            chartSources: [{ tableLabels: [], complete: false }],
        });

        expect(entry.markdown).toContain(
            [
                "- Data (chart cache):",
                "",
                "| Category | Sales | Cost |",
                "| Q1 | 10 | 1 |",
                "| Q2 |  | 2 |",
                "| Q3 | 30.5 | 3 |",
                "| 4 |  | 4 |",
            ].join("\n")
        );
    });

    it("names the source tables and skips cached data when the tables hold every series", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const charts = [
            {
                anchor: "B9",
                title: "",
                chartType: "bar",
                series: [
                    {
                        name: "Sales",
                        categoriesRef: "",
                        valuesRef: "",
                        cachedCategories: ["Q1"],
                        cachedValues: [10],
                        axis: "primary",
                    },
                ],
            },
        ];

        const [entry] = api.createAssetEntries({ images: [] }, charts, [], [], {
            includeShapeDetails: false,
            chartSources: [{ tableLabels: ["Table 001 (2:2-6:3)"], complete: true }],
        });

        expect(entry.markdown).toContain("- Type: bar\n- Source: Table 001 (2:2-6:3)");
        expect(entry.markdown).not.toContain("- Data (chart cache):");
    });

    it("places a chart right after the table it plots and links them in the converted sheet", () => {
        const table = { startRow: 2, startCol: 2, endRow: 4, endCol: 3, score: 9, reasonSummary: [] };
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
                detectTableCandidates: () => [table],
                matrixFromCandidate: () => [
                    ["Year", "Sales"],
                    ["2024", "10"],
                    ["2025", "20"],
                ],
            })
        );
        const cell = (address, row, col, value) => ({
            address,
            row,
            col,
            outputValue: value,
            rawValue: value,
            formulaText: "",
            resolutionStatus: null,
            resolutionSource: null,
        });
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                cell("A1", 1, 1, "Title"),
                cell("B2", 2, 2, "Year"),
                cell("C2", 2, 3, "Sales"),
                cell("B3", 3, 2, "2024"),
                cell("C3", 3, 3, "10"),
                cell("B4", 4, 2, "2025"),
                cell("C4", 4, 3, "20"),
                cell("A20", 20, 1, "Footer"),
            ],
            merges: [],
            images: [],
            shapes: [],
            charts: [
                {
                    anchor: "B6",
                    title: "Sales",
                    chartType: "bar",
                    series: [
                        {
                            name: "Sales",
                            categoriesRef: "Sheet1!$B$3:$B$4",
                            valuesRef: "'sheet1'!$C$3:$C$4",
                            cachedCategories: ["2024", "2025"],
                            cachedValues: [10, 20],
                            axis: "primary",
                        },
                    ],
                },
            ],
        };

        const markdown = api.convertSheetToMarkdown({ name: "book.xlsx", sheets: [] }, sheet, {}).markdown;

        expect(markdown).toContain("- Source: Table 001 (2:2-4:3)");
        expect(markdown).not.toContain("- Data (chart cache):");
        const tableAt = markdown.indexOf("### Table: 001");
        const chartAt = markdown.indexOf("### Chart: 001 (B6)");
        const footerAt = markdown.indexOf("Footer");
        expect(tableAt).toBeGreaterThan(-1);
        expect(chartAt).toBeGreaterThan(tableAt);
        expect(footerAt).toBeGreaterThan(chartAt);
    });

    it("renders comments as a separate sheet section and counts them in summary", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "Cell",
                    rawValue: "Cell",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
            comments: [
                { address: "A1", kind: "note", author: "Alice", text: "Legacy note", dateTime: "" },
                {
                    address: "B2",
                    kind: "threaded",
                    author: "Bob",
                    text: "Threaded reply",
                    dateTime: "2026-07-02T10:00:00Z",
                },
            ],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, {});

        expect(result.markdown).toContain("### Comments");
        expect(result.markdown).toContain("- [comment-1] (A1; note; author=Alice) Legacy note");
        expect(result.markdown).toContain(
            "- [comment-2] (B2; threaded; author=Bob; date=2026-07-02T10:00:00Z) Threaded reply"
        );
        expect(result.summary.comments).toBe(2);
    });

    it("renders ungrouped shapes after grouped shape blocks", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderHierarchicalRawEntries: () => ["- kind: rect"],
            })
        );
        const shapes = [
            {
                anchor: "B3",
                rawEntries: [{ key: "kind", value: "rect" }],
                svgFilename: null,
                svgPath: null,
                svgData: null,
            },
            {
                anchor: "E8",
                rawEntries: [{ key: "kind", value: "rect" }],
                svgFilename: "shape_002.svg",
                svgPath: "assets/Sheet1/shape_002.svg",
                svgData: new Uint8Array([2]),
            },
        ];
        const shapeBlocks = [{ startRow: 3, startCol: 2, endRow: 3, endCol: 2, shapeIndexes: [0] }];

        const entries = api.createAssetEntries({ images: [] }, [], shapes, shapeBlocks, { includeShapeDetails: true });
        const section = entries.map((entry) => entry.markdown).join("\n");

        expect(entries.map((entry) => [entry.sortRow, entry.sortCol])).toEqual([
            [3, 2],
            [8, 5],
        ]);
        expect(section).toContain("### Shape Block: 001 (3:2-3:2)");
        expect(section).toContain("- Shapes: Shape 001");
        expect(section).toContain("### Ungrouped Shapes");
        expect(section).toContain("#### Shape: 002 (E8)");
        expect(section).toContain("![shape_002.svg](assets/Sheet1/shape_002.svg)");
    });

    it("draws joined shapes as a Mermaid flowchart and lists loose connectors", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const box = (left, top) => ({ left, top, right: left + 10, bottom: top + 10 });
        const shape = (fields) => ({ rawEntries: [], svgPath: null, connection: null, ...fields });
        const shapes = [
            shape({
                anchor: "B2",
                drawingId: "2",
                kind: "Shape (flowChartTerminator)",
                text: "Start",
                bbox: box(0, 0),
            }),
            shape({
                anchor: "B5",
                drawingId: "3",
                kind: "Shape (flowChartDecision)",
                text: 'Is "A" > #1?',
                bbox: box(0, 50),
            }),
            shape({ anchor: "B8", drawingId: "4", kind: "Rectangle", text: "Line 1\nLine 2", bbox: box(0, 100) }),
            shape({
                anchor: "B3",
                drawingId: "5",
                kind: "Straight Arrow Connector",
                text: "",
                bbox: box(0, 10),
                connection: { startId: "2", endId: "3", arrow: "end" },
            }),
            shape({
                anchor: "B6",
                drawingId: "6",
                kind: "Connector (bentConnector3)",
                text: "yes",
                bbox: box(0, 60),
                // Arrowhead only at the start: the flow runs from the end shape to the start shape.
                connection: { startId: "4", endId: "3", arrow: "start" },
            }),
            shape({
                anchor: "D6",
                drawingId: "7",
                kind: "Straight Arrow Connector",
                text: "",
                bbox: box(30, 60),
                connection: { startId: "4", endId: null, arrow: "end" },
            }),
        ];
        const shapeBlocks = [{ startRow: 2, startCol: 2, endRow: 9, endCol: 4, shapeIndexes: [0, 1, 2, 3, 4, 5] }];

        const [entry] = api.createAssetEntries({ images: [] }, [], shapes, shapeBlocks, { includeShapeDetails: false });

        expect(entry.markdown).toContain(
            [
                "- Items:",
                "  - Shape 001 (Shape (flowChartTerminator)): Start",
                '  - Shape 002 (Shape (flowChartDecision)): Is "A" > #1?',
                "  - Shape 003 (Rectangle): Line 1 / Line 2",
                "  - Shape 004 (Straight Arrow Connector): Shape 001 → Shape 002",
                "  - Shape 005 (Connector (bentConnector3)): Shape 002 → Shape 003 (yes)",
                "  - Shape 006 (Straight Arrow Connector): not attached",
            ].join("\n")
        );
        expect(entry.markdown).toContain(
            [
                "```mermaid",
                "flowchart TD",
                '    sh001(["Start"])',
                '    sh002{"Is #quot;A#quot; #gt; #35;1?"}',
                '    sh003["Line 1<br>Line 2"]',
                "    sh001 --> sh002",
                '    sh002 -->|"yes"| sh003',
                "```",
            ].join("\n")
        );
        expect(entry.markdown).not.toContain("#### Shape:");
    });

    it("converts a minimal sheet into markdown", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "Hello",
                    rawValue: "Hello",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
                {
                    address: "B2",
                    row: 2,
                    col: 2,
                    outputValue: "World",
                    rawValue: "World",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, {});

        expect(result.fileName).toBe("1_Sheet1.md");
        expect(result.markdown).toContain("# Book: book.xlsx");
        expect(result.markdown).toContain("## Sheet: Sheet1");
        expect(result.summary.narrativeBlocks).toBe(1);
    });

    it("keeps nearby calendar rows in one narrative block even when later rows shift right", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "2021-01-03",
                    rawValue: "2021-01-03",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
                {
                    address: "B1",
                    row: 1,
                    col: 2,
                    outputValue: "2021-01-04",
                    rawValue: "2021-01-04",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
                {
                    address: "C1",
                    row: 1,
                    col: 3,
                    outputValue: "2021-01-05",
                    rawValue: "2021-01-05",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
                {
                    address: "D1",
                    row: 1,
                    col: 4,
                    outputValue: "2021-01-06",
                    rawValue: "2021-01-06",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
                {
                    address: "E1",
                    row: 1,
                    col: 5,
                    outputValue: "2021-01-07",
                    rawValue: "2021-01-07",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
                {
                    address: "F1",
                    row: 1,
                    col: 6,
                    outputValue: "2021-01-08",
                    rawValue: "2021-01-08",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
                {
                    address: "G1",
                    row: 1,
                    col: 7,
                    outputValue: "2021-01-09",
                    rawValue: "2021-01-09",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
                {
                    address: "J2",
                    row: 2,
                    col: 10,
                    outputValue: "仕事",
                    rawValue: "仕事",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
                {
                    address: "K2",
                    row: 2,
                    col: 11,
                    outputValue: "私用",
                    rawValue: "私用",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const blocks = api.extractNarrativeBlocks(workbook, sheet, [], {});

        expect(blocks).toHaveLength(1);
        expect(blocks[0].lines).toEqual([
            "2021-01-03 2021-01-04 2021-01-05 2021-01-06 2021-01-07 2021-01-08 2021-01-09",
            "仕事 私用",
        ]);
    });

    it("collects render state and uses empty-body fallback text", () => {
        const api = sheetMarkdown.createSheetMarkdown(createDeps());
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const state = api.collectSheetRenderState(workbook, sheet, {});
        const markdown = api.createSheetMarkdownText(workbook, sheet, state);
        const summary = api.createSheetSummary(sheet, state);

        expect(state.body).toBe("");
        expect(state.groupedSections).toEqual([]);
        expect(markdown).toContain("_No extractable body content was found._");
        expect(summary.sections).toBe(0);
        expect(summary.tables).toBe(0);
        expect(summary.narrativeBlocks).toBe(0);
    });

    it("normalizes table detection compatibility aliases in core conversion", () => {
        const detectedModes = [];
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                detectTableCandidates: (_sheet, _buildCellMap, tableDetectionMode) => {
                    detectedModes.push(tableDetectionMode);
                    return [];
                },
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, { tableDetectionMode: "border-priority" });

        expect(detectedModes).toEqual(["border"]);
        expect(result.summary.tableDetectionMode).toBe("border");
    });

    it("normalizes cell line breaks into spaces in plain mode", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "Line1\nLine2",
                    rawValue: "Raw1\nRaw2",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, {});

        expect(result.markdown).toContain("Line1 Line2");
        expect(result.markdown).not.toContain("Line1\nLine2");
    });

    it("renders cell line breaks as <br> in gfm mode", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "Line1\nLine2",
                    rawValue: "Raw1\nRaw2",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, { formattingMode: "gfm" });

        expect(result.markdown).toContain("Line1<br>Line2");
        expect(result.markdown).not.toContain("Line1 Line2");

        const legacy = api.convertSheetToMarkdown(workbook, sheet, { formattingMode: "github" });
        expect(legacy.markdown).toBe(result.markdown);
        expect(legacy.summary.formattingMode).toBe("gfm");
    });

    it("keeps the shape summary and drops raw details when includeShapeDetails is disabled", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                extractShapeBlocks: () => [{ startRow: 3, startCol: 2, endRow: 4, endCol: 3, shapeIndexes: [0] }],
                renderHierarchicalRawEntries: () => ["- kind: rect", "- text: dummy"],
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [],
            merges: [],
            images: [],
            charts: [],
            shapes: [
                {
                    anchor: "B3",
                    kind: "Rectangle",
                    text: "dummy",
                    rawEntries: [{ key: "kind", value: "rect" }],
                    svgFilename: null,
                    svgPath: null,
                    svgData: null,
                },
            ],
        };

        const enabled = api.convertSheetToMarkdown(workbook, sheet, {});
        const disabled = api.convertSheetToMarkdown(workbook, sheet, { includeShapeDetails: false });

        expect(enabled.markdown).toContain("### Shape Block: 001 (3:2-4:3)");
        expect(enabled.markdown).toContain("#### Shape: 001 (B3)");
        expect(disabled.markdown).toContain("### Shape Block: 001 (3:2-4:3)");
        expect(disabled.markdown).toContain("  - Shape 001 (Rectangle): dummy");
        expect(disabled.markdown).not.toContain("#### Shape:");
    });

    it("keeps a blank line between shape items when svg output is present", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                extractShapeBlocks: () => [],
                renderHierarchicalRawEntries: () => ["- kind: rect"],
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [],
            merges: [],
            images: [],
            charts: [],
            shapes: [
                {
                    anchor: "H3",
                    rawEntries: [{ key: "kind", value: "rect" }],
                    svgFilename: "shape_001.svg",
                    svgPath: "assets/Sheet1/shape_001.svg",
                    svgData: new Uint8Array([1]),
                },
                {
                    anchor: "K3",
                    rawEntries: [{ key: "kind", value: "rect" }],
                    svgFilename: "shape_002.svg",
                    svgPath: "assets/Sheet1/shape_002.svg",
                    svgData: new Uint8Array([2]),
                },
            ],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, {});

        expect(result.markdown).toContain("![shape_001.svg](assets/Sheet1/shape_001.svg)\n\n#### Shape: 002 (K3)");
    });

    it("keeps line-start markdown markers literal in narrative output", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "# heading",
                    rawValue: "# heading",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
                {
                    address: "A2",
                    row: 2,
                    col: 1,
                    outputValue: "- item",
                    rawValue: "- item",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, {});

        expect(result.markdown).toContain("\\# heading");
        expect(result.markdown).toContain("\\- item");
    });

    it("keeps ordered-list and quote markers literal in narrative output", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "1. item",
                    rawValue: "1. item",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
                {
                    address: "A2",
                    row: 2,
                    col: 1,
                    outputValue: "> quote",
                    rawValue: "> quote",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, {});

        expect(result.markdown).toContain("1\\. item");
        expect(result.markdown).toContain("&gt; quote");
    });

    it("keeps image-like markdown and code spans literal in narrative output", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "![alt](img.png)",
                    rawValue: "![alt](img.png)",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
                {
                    address: "A2",
                    row: 2,
                    col: 1,
                    outputValue: "`code`",
                    rawValue: "`code`",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, {});

        expect(result.markdown).toContain("\\!\\[alt\\]\\(img.png\\)");
        expect(result.markdown).toContain("\\`code\\`");
    });

    it("keeps additional list markers and ampersands literal in narrative output", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "+ plus",
                    rawValue: "+ plus",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
                {
                    address: "A2",
                    row: 2,
                    col: 1,
                    outputValue: "* star",
                    rawValue: "* star",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
                {
                    address: "A3",
                    row: 3,
                    col: 1,
                    outputValue: "a & b",
                    rawValue: "a & b",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, {});

        expect(result.markdown).toContain("\\+ plus");
        expect(result.markdown).toContain("\\* star");
        expect(result.markdown).toContain("a &amp; b");
    });

    it("renders external and workbook hyperlinks as markdown links", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = {
            name: "book.xlsx",
            sheets: [
                { name: "Sheet1", index: 1, cells: [], merges: [], images: [], charts: [], shapes: [] },
                { name: "Other Sheet", index: 2, cells: [], merges: [], images: [], charts: [], shapes: [] },
            ],
        };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "Open",
                    rawValue: "Open",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                    hyperlink: {
                        kind: "external",
                        target: "https://example.com/",
                        location: "",
                        tooltip: "",
                        display: "",
                    },
                },
                {
                    address: "A2",
                    row: 2,
                    col: 1,
                    outputValue: "Jump",
                    rawValue: "Jump",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                    hyperlink: {
                        kind: "internal",
                        target: "'Other Sheet'!C3",
                        location: "'Other Sheet'!C3",
                        tooltip: "",
                        display: "",
                    },
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, { formattingMode: "gfm" });

        expect(result.markdown).toContain("[Open](https://example.com/)");
        expect(result.markdown).toContain("[Jump](#sheet-other-sheet) (Other Sheet!C3)");
    });

    it("suppresses underline markup for hyperlink cells in gfm mode", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = {
            name: "book.xlsx",
            sheets: [{ name: "Sheet1", index: 1, cells: [], merges: [], images: [], charts: [], shapes: [] }],
        };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "Linked",
                    rawValue: "Linked",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: true },
                    richTextRuns: null,
                    hyperlink: {
                        kind: "external",
                        target: "https://example.com/",
                        location: "",
                        tooltip: "",
                        display: "",
                    },
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, { formattingMode: "gfm" });

        expect(result.markdown).toContain("[Linked](https://example.com/)");
        expect(result.markdown).not.toContain("<ins>Linked</ins>");
    });

    it("uses raw values for raw output mode while preserving hyperlinks", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = {
            name: "book.xlsx",
            sheets: [{ name: "Sheet1", index: 1, cells: [], merges: [], images: [], charts: [], shapes: [] }],
        };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "Displayed",
                    rawValue: "https://raw.example/",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                    hyperlink: {
                        kind: "external",
                        target: "https://example.com/",
                        location: "",
                        tooltip: "",
                        display: "",
                    },
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, { outputMode: "raw" });

        expect(result.markdown).toContain("[https://raw.example/](https://example.com/)");
        expect(result.markdown).not.toContain("[Displayed](https://example.com/)");
    });

    it("appends raw values only in both output mode when display and raw differ", () => {
        const api = sheetMarkdown.createSheetMarkdown(
            createDeps({
                renderNarrativeBlock: (block) => block.lines.join("\n"),
            })
        );
        const workbook = { name: "book.xlsx", sheets: [] };
        const sheet = {
            name: "Sheet1",
            index: 1,
            cells: [
                {
                    address: "A1",
                    row: 1,
                    col: 1,
                    outputValue: "Displayed",
                    rawValue: "RawValue",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
                {
                    address: "A2",
                    row: 2,
                    col: 1,
                    outputValue: "SameValue",
                    rawValue: "SameValue",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    textStyle: { bold: false, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
            ],
            merges: [],
            images: [],
            charts: [],
            shapes: [],
        };

        const result = api.convertSheetToMarkdown(workbook, sheet, { outputMode: "both" });

        expect(result.markdown).toContain("Displayed [raw=RawValue]");
        expect(result.markdown).toContain("SameValue");
        expect(result.markdown).not.toContain("SameValue [raw=SameValue]");
    });
});
