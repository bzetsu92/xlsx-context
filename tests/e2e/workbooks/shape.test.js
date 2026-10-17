import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { unzipEntries } from "../../../src/extraction/zip/zip-io.ts";
import { createWorkbookExportArchive } from "../../../src/outputs/markdown/markdown-export.ts";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("shape workbooks", () => {
    const fixtureDir = FIXTURE_ROOT;

    it("parses the shape-basic fixture workbook without misclassifying drawing shapes", async () => {
        const api = core;
        const fixtureName = "shape-basic-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
        const fixtureBytes = readFileSync(fixturePath);
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });
        const sheet = workbook.sheets[0];
        const markdownFile = files[0];

        expect(workbook.sheets).toHaveLength(1);
        expect(sheet.name).toBe("shape-basic");
        expect(sheet.maxRow).toBe(6);
        expect(sheet.maxCol).toBe(5);
        expect(sheet.cells).toHaveLength(17);
        expect(sheet.images).toHaveLength(0);
        expect(sheet.charts).toHaveLength(0);
        expect(sheet.shapes).toHaveLength(3);
        expect(sheet.shapes[0]).toMatchObject({
            sheetName: "shape-basic",
            anchor: "H3",
            name: "テキスト ボックス 1",
            kind: "Text Box",
            text: "テキストボックスの例",
            widthEmu: 1980029,
            heightEmu: 392608,
            elementName: "xdr:sp",
            svgFilename: "shape_001.svg",
            svgPath: "assets/shape-basic/shape_001.svg",
        });
        expect(sheet.shapes[0].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:oneCellAnchor/xdr:from/xdr:col#text", value: "7" },
                { key: "xdr:oneCellAnchor/xdr:from/xdr:row#text", value: "2" },
                { key: "xdr:oneCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvPr@name", value: "テキスト ボックス 1" },
                { key: "xdr:oneCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvSpPr@txBox", value: "1" },
                { key: "xdr:oneCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "テキストボックスの例" },
            ])
        );
        expect(sheet.shapes[1]).toMatchObject({
            sheetName: "shape-basic",
            anchor: "H8",
            name: "直線矢印コネクタ 3",
            kind: "Straight Arrow Connector",
            text: "",
            widthEmu: 1308100,
            heightEmu: 0,
            elementName: "xdr:cxnSp",
            svgFilename: "shape_002.svg",
            svgPath: "assets/shape-basic/shape_002.svg",
        });
        expect(sheet.shapes[1].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:cxnSp/xdr:nvCxnSpPr/xdr:cNvPr@name", value: "直線矢印コネクタ 3" },
                { key: "xdr:twoCellAnchor/xdr:cxnSp/xdr:spPr/a:prstGeom@prst", value: "straightConnector1" },
            ])
        );
        expect(sheet.shapes[2]).toMatchObject({
            sheetName: "shape-basic",
            anchor: "K3",
            name: "正方形/長方形 4",
            kind: "Rectangle",
            text: "",
            widthEmu: 1511300,
            heightEmu: 1155700,
            elementName: "xdr:sp",
            svgFilename: "shape_003.svg",
            svgPath: "assets/shape-basic/shape_003.svg",
        });
        expect(sheet.shapes[2].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvPr@name", value: "正方形/長方形 4" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "rect" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:xfrm/a:ext@cx", value: "1511300" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:xfrm/a:ext@cy", value: "1155700" },
            ])
        );

        expect(markdownFile.fileName).toBe("shape-basic-sample01_001_shape-basic.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.charts).toBe(0);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-E6"]);
        expect(markdownFile.markdown).toContain("## Sheet: shape-basic");
        expect(markdownFile.markdown).toContain("# Book: shape-basic-sample01.xlsx");
        expect(markdownFile.markdown).toContain("図形サンプル");
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-E6)");
        expect(markdownFile.markdown).toContain("| 項目 | 値A | 値B | 値C |");
        expect(markdownFile.markdown).toContain("| 2026年 | 25,051 | 32,012 | 14,850 |");
        expect(markdownFile.markdown).toContain("### Shape Block: 001 (");
        expect(markdownFile.markdown).toContain("- Shapes: Shape 001, Shape 002, Shape 003");
        expect(markdownFile.markdown).not.toContain("### Shapes");
        expect(markdownFile.markdown).toContain("#### Shape: 001 (H3)");
        expect(markdownFile.markdown).toContain("- `xdr:oneCellAnchor`");
        expect(markdownFile.markdown).toContain("    - `xdr:from`");
        expect(markdownFile.markdown).toContain("        - `xdr:col#text`: `7`");
        expect(markdownFile.markdown).toContain("    - `xdr:sp`");
        expect(markdownFile.markdown).toContain("- `xdr:cNvPr@name`: `テキスト ボックス 1`");
        expect(markdownFile.markdown).toContain("- `a:t#text`: `テキストボックスの例`");
        expect(markdownFile.markdown).toContain("- SVG: assets/shape-basic/shape_001.svg");
        expect(markdownFile.markdown).toContain("![shape_001.svg](assets/shape-basic/shape_001.svg)");
        expect(markdownFile.markdown).toContain("#### Shape: 002 (H8)");
        expect(markdownFile.markdown).toContain("- `xdr:twoCellAnchor`");
        expect(markdownFile.markdown).toContain("    - `xdr:cxnSp`");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `straightConnector1`");
        expect(markdownFile.markdown).toContain("- `a:ext@cx`: `1308100`");
        expect(markdownFile.markdown).toContain("#### Shape: 003 (K3)");
        expect(markdownFile.markdown).toContain("- `xdr:cNvPr@name`: `正方形/長方形 4`");
        expect(markdownFile.markdown).toContain("![shape_003.svg](assets/shape-basic/shape_003.svg)");
        expect(markdownFile.markdown).not.toContain("### Image:");
        expect(markdownFile.markdown).not.toContain("### Chart:");
    });

    it("skips shape parsing when includeShapeDetails is disabled at parse time", async () => {
        const api = core;
        const fixtureName = "shape-basic-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
        const fixtureBytes = readFileSync(fixturePath);
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await api.parseWorkbook(arrayBuffer, fixtureName, { includeShapeDetails: false });
        const sheet = workbook.sheets[0];
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            includeShapeDetails: false,
        });

        expect(sheet.shapes).toEqual([]);
        expect(files[0].markdown).not.toContain("### Shape Block:");
        expect(files[0].markdown).not.toContain("#### Shape:");
    });

    it("exports shape SVG assets into the markdown+assets archive", async () => {
        const api = core;
        const fixtureName = "shape-basic-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
        const fixtureBytes = readFileSync(fixturePath);
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });
        const archive = createWorkbookExportArchive(workbook, files);
        const extracted = await unzipEntries(
            archive.buffer.slice(archive.byteOffset, archive.byteOffset + archive.byteLength)
        );

        expect(extracted.has("output/shape-basic-sample01.md")).toBe(true);
        expect(extracted.has("output/assets/shape-basic/shape_001.svg")).toBe(true);
        expect(extracted.has("output/assets/shape-basic/shape_002.svg")).toBe(true);
        expect(extracted.has("output/assets/shape-basic/shape_003.svg")).toBe(true);

        const markdownText = new TextDecoder().decode(extracted.get("output/shape-basic-sample01.md"));
        expect(markdownText).toContain("#### Shape: 001 (H3)");
        expect(markdownText).toContain("![shape_001.svg](assets/shape-basic/shape_001.svg)");

        const shape1Svg = new TextDecoder().decode(extracted.get("output/assets/shape-basic/shape_001.svg"));
        const shape2Svg = new TextDecoder().decode(extracted.get("output/assets/shape-basic/shape_002.svg"));
        const shape3Svg = new TextDecoder().decode(extracted.get("output/assets/shape-basic/shape_003.svg"));
        expect(shape1Svg).toContain("<svg");
        expect(shape1Svg).toContain("<text");
        expect(shape2Svg).toContain("<svg");
        expect(shape2Svg).toContain("<line");
        expect(shape3Svg).toContain("<svg");
        expect(shape3Svg).toContain("<rect");
    });

    it("parses the shape-flowchart fixture workbook with flowchart raw metadata and connector SVG export", async () => {
        const api = core;
        const fixtureName = "shape-flowchart-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
        const fixtureBytes = readFileSync(fixturePath);
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });
        const sheet = workbook.sheets[0];
        const markdownFile = files[0];

        expect(workbook.sheets).toHaveLength(1);
        expect(sheet.name).toBe("shape-flowchart");
        expect(sheet.maxRow).toBe(6);
        expect(sheet.maxCol).toBe(7);
        expect(sheet.cells).toHaveLength(18);
        expect(sheet.images).toHaveLength(0);
        expect(sheet.charts).toHaveLength(0);
        expect(sheet.shapes).toHaveLength(7);

        expect(sheet.shapes[0]).toMatchObject({
            sheetName: "shape-flowchart",
            anchor: "H3",
            name: "フローチャート: 端子 2",
            kind: "Shape (flowChartTerminator)",
            text: "開始",
            widthEmu: 1689100,
            heightEmu: 584200,
            svgPath: "assets/shape-flowchart/shape_001.svg",
        });
        expect(sheet.shapes[0].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvPr@name", value: "フローチャート: 端子 2" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "flowChartTerminator" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "開始" },
            ])
        );
        expect(sheet.shapes[1]).toMatchObject({
            anchor: "K3",
            name: "フローチャート: 処理 5",
            kind: "Shape (flowChartProcess)",
            text: "処理",
            svgPath: "assets/shape-flowchart/shape_002.svg",
        });
        expect(sheet.shapes[1].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "flowChartProcess" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "処理" },
            ])
        );
        expect(sheet.shapes[2]).toMatchObject({
            anchor: "N3",
            name: "フローチャート: 判断 6",
            kind: "Shape (flowChartDecision)",
            text: "条件判断",
            svgPath: "assets/shape-flowchart/shape_003.svg",
        });
        expect(sheet.shapes[2].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "flowChartDecision" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "条件判断" },
            ])
        );
        expect(sheet.shapes[3]).toMatchObject({
            anchor: "Q3",
            name: "フローチャート: データ 7",
            kind: "Shape (flowChartInputOutput)",
            text: "データ",
            svgPath: "assets/shape-flowchart/shape_004.svg",
        });
        expect(sheet.shapes[3].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "flowChartInputOutput" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "データ" },
            ])
        );
        expect(sheet.shapes[4]).toMatchObject({
            anchor: "I4",
            name: "直線矢印コネクタ 9",
            kind: "Straight Arrow Connector",
            text: "",
            svgFilename: "shape_005.svg",
            svgPath: "assets/shape-flowchart/shape_005.svg",
            drawingId: "10",
            connection: { startId: "3", endId: "6", arrow: "end" },
        });
        expect(sheet.shapes[0]).toMatchObject({ drawingId: "3", connection: null });
        expect(sheet.shapes[4].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:cxnSp/xdr:spPr/a:prstGeom@prst", value: "straightConnector1" },
                { key: "xdr:twoCellAnchor/xdr:cxnSp/xdr:nvCxnSpPr/xdr:cNvCxnSpPr/a:endCxn@id", value: "6" },
            ])
        );

        expect(markdownFile.fileName).toBe("shape-flowchart-sample01_001_shape-flowchart.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.charts).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(0);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-E6"]);
        expect(markdownFile.markdown).toContain("## Sheet: shape-flowchart");
        expect(markdownFile.markdown).toContain("# Book: shape-flowchart-sample01.xlsx");
        expect(markdownFile.markdown).toContain("フローチャート図形サンプル");
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-E6)");
        expect(markdownFile.markdown).toContain("### Shape Block: 001 (H3-S7)");
        expect(markdownFile.markdown).toContain(
            "- Shapes: Shape 001, Shape 002, Shape 003, Shape 004, Shape 005, Shape 006, Shape 007"
        );
        expect(markdownFile.markdown).not.toContain("### Shapes");
        expect(markdownFile.markdown).toContain("#### Shape: 001 (H3)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `flowChartTerminator`");
        expect(markdownFile.markdown).toContain("- `a:t#text`: `開始`");
        expect(markdownFile.markdown).toContain("#### Shape: 003 (N3)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `flowChartDecision`");
        expect(markdownFile.markdown).toContain("- `a:t#text`: `条件判断`");
        expect(markdownFile.markdown).toContain("#### Shape: 005 (I4)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `straightConnector1`");
        expect(markdownFile.markdown).toContain("- `a:endCxn@id`: `6`");
        expect(markdownFile.markdown).toContain("![shape_005.svg](assets/shape-flowchart/shape_005.svg)");
        expect(markdownFile.markdown).toContain("![shape_006.svg](assets/shape-flowchart/shape_006.svg)");
        expect(markdownFile.markdown).toContain("![shape_007.svg](assets/shape-flowchart/shape_007.svg)");
        expect(markdownFile.markdown).toContain(
            [
                "```mermaid",
                "flowchart LR",
                '    sh001(["開始"])',
                '    sh002["処理"]',
                '    sh003{"条件判断"}',
                '    sh004[/"データ"/]',
                "    sh001 --> sh002",
                "    sh002 --> sh003",
                "    sh003 --> sh004",
                "```",
            ].join("\n")
        );
        expect(markdownFile.markdown).toContain("  - Shape 005 (Straight Arrow Connector): Shape 001 → Shape 002");
        expect(markdownFile.markdown).not.toContain("### Image:");
        expect(markdownFile.markdown).not.toContain("### Chart:");
    });

    it("exports flowchart connector SVG assets into the markdown+assets archive", async () => {
        const api = core;
        const fixtureName = "shape-flowchart-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
        const fixtureBytes = readFileSync(fixturePath);
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });
        const archive = createWorkbookExportArchive(workbook, files);
        const extracted = await unzipEntries(
            archive.buffer.slice(archive.byteOffset, archive.byteOffset + archive.byteLength)
        );

        expect(extracted.has("output/shape-flowchart-sample01.md")).toBe(true);
        expect(extracted.has("output/assets/shape-flowchart/shape_005.svg")).toBe(true);
        expect(extracted.has("output/assets/shape-flowchart/shape_006.svg")).toBe(true);
        expect(extracted.has("output/assets/shape-flowchart/shape_007.svg")).toBe(true);
        // Flowchart presets have no exact outline. They get a dashed fallback box with the preset name.
        expect(extracted.has("output/assets/shape-flowchart/shape_001.svg")).toBe(true);
        const shape1Svg = new TextDecoder().decode(extracted.get("output/assets/shape-flowchart/shape_001.svg"));
        expect(shape1Svg).toContain('data-prst="flowChartTerminator"');
        expect(shape1Svg).toContain('stroke-dasharray="6 3"');
        expect(shape1Svg).toContain("開始");

        const markdownText = new TextDecoder().decode(extracted.get("output/shape-flowchart-sample01.md"));
        expect(markdownText).toContain("#### Shape: 001 (H3)");
        expect(markdownText).toContain("![shape_005.svg](assets/shape-flowchart/shape_005.svg)");

        const shape5Svg = new TextDecoder().decode(extracted.get("output/assets/shape-flowchart/shape_005.svg"));
        const shape6Svg = new TextDecoder().decode(extracted.get("output/assets/shape-flowchart/shape_006.svg"));
        const shape7Svg = new TextDecoder().decode(extracted.get("output/assets/shape-flowchart/shape_007.svg"));
        expect(shape5Svg).toContain("<svg");
        expect(shape5Svg).toContain("<line");
        expect(shape6Svg).toContain("<svg");
        expect(shape6Svg).toContain("<line");
        expect(shape7Svg).toContain("<svg");
        expect(shape7Svg).toContain("<line");
    });

    it("parses the shape-block-arrow fixture workbook with block-arrow raw metadata", async () => {
        const api = core;
        const fixtureName = "shape-block-arrow-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
        const fixtureBytes = readFileSync(fixturePath);
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });
        const sheet = workbook.sheets[0];
        const markdownFile = files[0];

        expect(workbook.sheets).toHaveLength(1);
        expect(sheet.name).toBe("shape-block-arrow");
        expect(sheet.maxRow).toBe(6);
        expect(sheet.maxCol).toBe(7);
        expect(sheet.cells).toHaveLength(18);
        expect(sheet.images).toHaveLength(0);
        expect(sheet.charts).toHaveLength(0);
        expect(sheet.shapes).toHaveLength(5);

        expect(sheet.shapes[0]).toMatchObject({
            sheetName: "shape-block-arrow",
            anchor: "H3",
            name: "右矢印 22",
            kind: "Shape (rightArrow)",
            text: "右矢印",
            widthEmu: 2108200,
            heightEmu: 1066800,
            svgPath: "assets/shape-block-arrow/shape_001.svg",
        });
        expect(sheet.shapes[0].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvPr@name", value: "右矢印 22" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "rightArrow" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "右矢印" },
            ])
        );
        expect(sheet.shapes[1]).toMatchObject({
            anchor: "K3",
            name: "左右矢印 24",
            kind: "Shape (leftRightArrow)",
            text: "",
            svgPath: "assets/shape-block-arrow/shape_002.svg",
        });
        expect(sheet.shapes[1].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "leftRightArrow" },
            ])
        );
        expect(sheet.shapes[2]).toMatchObject({
            anchor: "N3",
            name: "上矢印 25",
            kind: "Shape (upArrow)",
            text: "",
            svgPath: "assets/shape-block-arrow/shape_003.svg",
        });
        expect(sheet.shapes[2].rawEntries).toEqual(
            expect.arrayContaining([{ key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "upArrow" }])
        );
        expect(sheet.shapes[3]).toMatchObject({
            anchor: "Q3",
            name: "U ターン矢印 26",
            kind: "Shape (uturnArrow)",
            text: "",
            svgPath: "assets/shape-block-arrow/shape_004.svg",
        });
        expect(sheet.shapes[3].rawEntries).toEqual(
            expect.arrayContaining([{ key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "uturnArrow" }])
        );
        expect(sheet.shapes[4]).toMatchObject({
            anchor: "H8",
            name: "四方向矢印 27",
            kind: "Shape (quadArrow)",
            text: "",
            svgPath: "assets/shape-block-arrow/shape_005.svg",
        });
        expect(sheet.shapes[4].rawEntries).toEqual(
            expect.arrayContaining([{ key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "quadArrow" }])
        );

        expect(markdownFile.fileName).toBe("shape-block-arrow-sample01_001_shape-block-arrow.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.charts).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(0);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-E6"]);
        expect(markdownFile.markdown).toContain("## Sheet: shape-block-arrow");
        expect(markdownFile.markdown).toContain("# Book: shape-block-arrow-sample01.xlsx");
        expect(markdownFile.markdown).toContain("ブロック矢印サンプル");
        expect(markdownFile.markdown).toContain("### Shape Block: 001 (H3-R14)");
        expect(markdownFile.markdown).toContain("- Shapes: Shape 001, Shape 002, Shape 003, Shape 004, Shape 005");
        expect(markdownFile.markdown).not.toContain("### Shapes");
        expect(markdownFile.markdown).toContain("#### Shape: 001 (H3)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `rightArrow`");
        expect(markdownFile.markdown).toContain("- `a:t#text`: `右矢印`");
        expect(markdownFile.markdown).toContain("#### Shape: 002 (K3)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `leftRightArrow`");
        expect(markdownFile.markdown).toContain("#### Shape: 003 (N3)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `upArrow`");
        expect(markdownFile.markdown).toContain("#### Shape: 004 (Q3)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `uturnArrow`");
        expect(markdownFile.markdown).toContain("#### Shape: 005 (H8)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `quadArrow`");
        expect(markdownFile.markdown).toContain("![shape_001.svg](assets/shape-block-arrow/shape_001.svg)");
        expect(markdownFile.markdown).not.toContain("### Image:");
        expect(markdownFile.markdown).not.toContain("### Chart:");
    });

    it("exports fallback SVG assets for the block-arrow fixture archive", async () => {
        const api = core;
        const fixtureName = "shape-block-arrow-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
        const fixtureBytes = readFileSync(fixturePath);
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });
        const archive = createWorkbookExportArchive(workbook, files);
        const extracted = await unzipEntries(
            archive.buffer.slice(archive.byteOffset, archive.byteOffset + archive.byteLength)
        );

        expect(extracted.has("output/shape-block-arrow-sample01.md")).toBe(true);
        expect(Array.from(extracted.keys()).sort()).toEqual([
            "output/assets/shape-block-arrow/shape_001.svg",
            "output/assets/shape-block-arrow/shape_002.svg",
            "output/assets/shape-block-arrow/shape_003.svg",
            "output/assets/shape-block-arrow/shape_004.svg",
            "output/assets/shape-block-arrow/shape_005.svg",
            "output/shape-block-arrow-sample01.md",
        ]);
        const shape4Svg = new TextDecoder().decode(extracted.get("output/assets/shape-block-arrow/shape_004.svg"));
        expect(shape4Svg).toContain('data-prst="uturnArrow"');

        const markdownText = new TextDecoder().decode(extracted.get("output/shape-block-arrow-sample01.md"));
        expect(markdownText).toContain("#### Shape: 001 (H3)");
        expect(markdownText).toContain("- `a:prstGeom@prst`: `uturnArrow`");
        expect(markdownText).toContain("![shape_001.svg](assets/shape-block-arrow/shape_001.svg)");
    });

    it("parses the shape-callout fixture workbook with callout raw metadata and text extraction", async () => {
        const api = core;
        const fixtureName = "shape-callout-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
        const fixtureBytes = readFileSync(fixturePath);
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });
        const sheet = workbook.sheets[0];
        const markdownFile = files[0];

        expect(workbook.sheets).toHaveLength(1);
        expect(sheet.name).toBe("shape-callout");
        expect(sheet.maxRow).toBe(6);
        expect(sheet.maxCol).toBe(7);
        expect(sheet.cells).toHaveLength(18);
        expect(sheet.images).toHaveLength(0);
        expect(sheet.charts).toHaveLength(0);
        expect(sheet.shapes).toHaveLength(4);

        expect(sheet.shapes[0]).toMatchObject({
            sheetName: "shape-callout",
            anchor: "H3",
            name: "角丸四角形吹き出し 2",
            kind: "Shape (wedgeRoundRectCallout)",
            text: "角四角",
            widthEmu: 2374900,
            heightEmu: 901700,
            svgPath: "assets/shape-callout/shape_001.svg",
        });
        expect(sheet.shapes[0].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:nvSpPr/xdr:cNvPr@name", value: "角丸四角形吹き出し 2" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "wedgeRoundRectCallout" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "角四角" },
            ])
        );
        expect(sheet.shapes[1]).toMatchObject({
            anchor: "K3",
            name: "円形吹き出し 3",
            kind: "Shape (wedgeEllipseCallout)",
            text: "楕円",
            svgPath: "assets/shape-callout/shape_002.svg",
        });
        expect(sheet.shapes[1].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "wedgeEllipseCallout" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "楕円" },
            ])
        );
        expect(sheet.shapes[2]).toMatchObject({
            anchor: "N3",
            name: "雲形吹き出し 4",
            kind: "Shape (cloudCallout)",
            text: "雲",
            svgPath: "assets/shape-callout/shape_003.svg",
        });
        expect(sheet.shapes[2].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "cloudCallout" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "雲" },
            ])
        );
        expect(sheet.shapes[3]).toMatchObject({
            anchor: "H8",
            name: "強調線吹き出し 1 (枠付き) 1",
            kind: "Shape (accentBorderCallout1)",
            text: "注釈",
            svgPath: "assets/shape-callout/shape_004.svg",
        });
        expect(sheet.shapes[3].rawEntries).toEqual(
            expect.arrayContaining([
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:spPr/a:prstGeom@prst", value: "accentBorderCallout1" },
                { key: "xdr:twoCellAnchor/xdr:sp/xdr:txBody/a:p/a:r/a:t#text", value: "注釈" },
            ])
        );

        expect(markdownFile.fileName).toBe("shape-callout-sample01_001_shape-callout.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.charts).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(0);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-E6"]);
        expect(markdownFile.markdown).toContain("## Sheet: shape-callout");
        expect(markdownFile.markdown).toContain("# Book: shape-callout-sample01.xlsx");
        expect(markdownFile.markdown).toContain("吹き出しサンプル");
        expect(markdownFile.markdown).toContain("### Shape Block: 001 (H3-P11)");
        expect(markdownFile.markdown).toContain("- Shapes: Shape 001, Shape 002, Shape 003, Shape 004");
        expect(markdownFile.markdown).not.toContain("### Shapes");
        expect(markdownFile.markdown).toContain("#### Shape: 001 (H3)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `wedgeRoundRectCallout`");
        expect(markdownFile.markdown).toContain("- `a:t#text`: `角四角`");
        expect(markdownFile.markdown).toContain("#### Shape: 002 (K3)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `wedgeEllipseCallout`");
        expect(markdownFile.markdown).toContain("- `a:t#text`: `楕円`");
        expect(markdownFile.markdown).toContain("#### Shape: 003 (N3)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `cloudCallout`");
        expect(markdownFile.markdown).toContain("- `a:t#text`: `雲`");
        expect(markdownFile.markdown).toContain("#### Shape: 004 (H8)");
        expect(markdownFile.markdown).toContain("- `a:prstGeom@prst`: `accentBorderCallout1`");
        expect(markdownFile.markdown).toContain("- `a:t#text`: `注釈`");
        expect(markdownFile.markdown).toContain("![shape_001.svg](assets/shape-callout/shape_001.svg)");
        expect(markdownFile.markdown).toContain("  - Shape 003 (Shape (cloudCallout)): 雲");
        expect(markdownFile.markdown).not.toContain("### Image:");
        expect(markdownFile.markdown).not.toContain("### Chart:");
    });

    it("exports fallback SVG assets for the callout fixture archive", async () => {
        const api = core;
        const fixtureName = "shape-callout-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "shape", fixtureName);
        const fixtureBytes = readFileSync(fixturePath);
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });
        const archive = createWorkbookExportArchive(workbook, files);
        const extracted = await unzipEntries(
            archive.buffer.slice(archive.byteOffset, archive.byteOffset + archive.byteLength)
        );

        expect(extracted.has("output/shape-callout-sample01.md")).toBe(true);
        expect(Array.from(extracted.keys()).sort()).toEqual([
            "output/assets/shape-callout/shape_001.svg",
            "output/assets/shape-callout/shape_002.svg",
            "output/assets/shape-callout/shape_003.svg",
            "output/assets/shape-callout/shape_004.svg",
            "output/shape-callout-sample01.md",
        ]);
        const shape3Svg = new TextDecoder().decode(extracted.get("output/assets/shape-callout/shape_003.svg"));
        expect(shape3Svg).toContain('data-prst="cloudCallout"');
        expect(shape3Svg).toContain("雲");

        const markdownText = new TextDecoder().decode(extracted.get("output/shape-callout-sample01.md"));
        expect(markdownText).toContain("#### Shape: 001 (H3)");
        expect(markdownText).toContain("- `a:prstGeom@prst`: `cloudCallout`");
        expect(markdownText).toContain("![shape_001.svg](assets/shape-callout/shape_001.svg)");
    });
});
