import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("table workbooks", () => {
    const fixtureDir = FIXTURE_ROOT;

    it("parses the table-basic-sample01 fixture workbook as two vertically adjacent independent tables", async () => {
        const api = core;
        const fixtureName = "table-basic-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
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
        expect(sheet.name).toBe("table-basic");
        expect(sheet.maxRow).toBe(13);
        expect(sheet.maxCol).toBe(6);
        expect(sheet.cells).toHaveLength(53);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe(
            "隣接するテーブルのテスト（縦に密接）"
        );
        expect(sheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("隣接するテーブルその1");
        expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe("隣接するテーブルその2");
        expect(sheet.cells.find((cell) => cell.address === "B5")?.formulaText).toBe("=B4+1");
        expect(sheet.cells.find((cell) => cell.address === "B13")?.formulaText).toBe("=B12+1");
        expect(sheet.cells.find((cell) => cell.address === "E12")?.outputValue).toBe("3月15日");
        expect(sheet.cells.find((cell) => cell.address === "F13")?.outputValue).toBe("更新した日");

        expect(markdownFile.fileName).toBe("table-basic-sample01_001_table-basic.md");
        expect(markdownFile.summary.tables).toBe(2);
        expect(markdownFile.summary.merges).toBe(0);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(6);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-F7", "B9-F13"]);
        expect(markdownFile.markdown).toContain("## Sheet: table-basic");
        expect(markdownFile.markdown).toContain("# Book: table-basic-sample01.xlsx");
        expect(markdownFile.markdown).toContain("隣接するテーブルのテスト（縦に密接）");
        expect(markdownFile.markdown).toContain("隣接するテーブルその1");
        expect(markdownFile.markdown).toContain("隣接するテーブルその2");
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-F7)");
        expect(markdownFile.markdown).toContain("### Table: 002 (B9-F13)");
        expect(markdownFile.markdown).toContain("| 1 | コード | code | 101 | 何かのコード |");
        expect(markdownFile.markdown).toContain("| 3 | 登録日 | createdAt | 3月15日 | 登録した日 |");
    });

    it("parses the table-basic-sample02 fixture workbook as two horizontally adjacent independent tables", async () => {
        const api = core;
        const fixtureName = "table-basic-sample02.xlsx";
        const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
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
        expect(sheet.name).toBe("table-basic");
        expect(sheet.maxRow).toBe(7);
        expect(sheet.maxCol).toBe(12);
        expect(sheet.cells).toHaveLength(56);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe(
            "隣接するテーブルのテスト（横に密接）"
        );
        expect(sheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("隣接するテーブルその1");
        expect(sheet.cells.find((cell) => cell.address === "H2")?.outputValue).toBe("隣接するテーブルその2");
        expect(sheet.cells.find((cell) => cell.address === "G4")?.outputValue).toBe("確認");
        expect(sheet.cells.find((cell) => cell.address === "G6")?.outputValue).toBe("日付");
        expect(sheet.cells.find((cell) => cell.address === "H7")?.formulaText).toBe("=H6+1");
        expect(sheet.cells.find((cell) => cell.address === "K6")?.outputValue).toBe("3月15日");

        expect(markdownFile.fileName).toBe("table-basic-sample02_001_table-basic.md");
        expect(markdownFile.summary.tables).toBe(2);
        expect(markdownFile.summary.merges).toBe(0);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(6);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-F7", "H3-L7"]);
        expect(markdownFile.markdown).toContain("## Sheet: table-basic");
        expect(markdownFile.markdown).toContain("# Book: table-basic-sample02.xlsx");
        expect(markdownFile.markdown).toContain("隣接するテーブルのテスト（横に密接）");
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-F7)");
        expect(markdownFile.markdown).toContain("### Table: 002 (H3-L7)");
        expect(markdownFile.markdown).toContain("| 1 | コード | code | 101 | 何かのコード |");
        expect(markdownFile.markdown).toContain("| 2 | 別名 | altname | Hanako | 何かの別名 |");
        expect(markdownFile.markdown).toContain("確認");
        expect(markdownFile.markdown).toContain("日付");
    });

    it("parses the table-basic-sample03 fixture workbook as four tightly adjacent independent tables", async () => {
        const api = core;
        const fixtureName = "table-basic-sample03.xlsx";
        const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
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
        expect(sheet.name).toBe("table-basic");
        expect(sheet.maxRow).toBe(13);
        expect(sheet.maxCol).toBe(12);
        expect(sheet.cells).toHaveLength(111);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe(
            "隣接するテーブルのテスト（縦横に密接）"
        );
        expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe("隣接するテーブルその3");
        expect(sheet.cells.find((cell) => cell.address === "H8")?.outputValue).toBe("隣接するテーブルその4");
        expect(sheet.cells.find((cell) => cell.address === "G10")?.outputValue).toBe("確認");
        expect(sheet.cells.find((cell) => cell.address === "H13")?.formulaText).toBe("=H12+1");
        expect(sheet.cells.find((cell) => cell.address === "K13")?.outputValue).toBe("3月19日");
        expect(sheet.cells.find((cell) => cell.address === "L11")?.outputValue).toBe("何かの別名");

        expect(markdownFile.fileName).toBe("table-basic-sample03_001_table-basic.md");
        expect(markdownFile.summary.tables).toBe(4);
        expect(markdownFile.summary.merges).toBe(0);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(12);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual([
            "B3-F7",
            "H3-L7",
            "B9-F13",
            "H9-L13",
        ]);
        expect(markdownFile.markdown).toContain("## Sheet: table-basic");
        expect(markdownFile.markdown).toContain("# Book: table-basic-sample03.xlsx");
        // Numbered in reading order: the left column of tables first, then the right column.
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-F7)");
        expect(markdownFile.markdown).toContain("### Table: 002 (B9-F13)");
        expect(markdownFile.markdown).toContain("### Table: 003 (H3-L7)");
        expect(markdownFile.markdown).toContain("### Table: 004 (H9-L13)");
        expect(markdownFile.markdown).not.toContain("### Table: 005");
        expect(markdownFile.markdown).not.toContain("### Table: 002 (B3-L13)");
        expect(markdownFile.markdown).toContain("| 1 | コード | code | 301 | 何かのコード |");
        expect(markdownFile.markdown).toContain("| 2 | 別名 | altname | Sawada | 何かの別名 |");
    });

    it("parses the table-basic-sample11 fixture workbook as one merge-heavy graph-paper table", async () => {
        const api = core;
        const fixtureName = "table-basic-sample11.xlsx";
        const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
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
        expect(sheet.name).toBe("table-basic");
        expect(sheet.maxRow).toBe(7);
        expect(sheet.maxCol).toBe(20);
        expect(sheet.cells).toHaveLength(97);
        expect(sheet.merges).toHaveLength(20);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("方眼紙的様式のテスト");
        expect(sheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("テーブルその1");
        expect(sheet.cells.find((cell) => cell.address === "B5")?.formulaText).toBe("=B4+1");
        expect(sheet.cells.find((cell) => cell.address === "P4")?.outputValue).toBe("何かのコード");
        expect(sheet.cells.find((cell) => cell.address === "L6")?.outputValue).toBe("3月13日");

        expect(markdownFile.fileName).toBe("table-basic-sample11_001_table-basic.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.merges).toBe(20);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(3);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-T7"]);
        expect(markdownFile.markdown).toContain("## Sheet: table-basic");
        expect(markdownFile.markdown).toContain("# Book: table-basic-sample11.xlsx");
        expect(markdownFile.markdown).toContain("方眼紙的様式のテスト");
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-T7)");
        expect(markdownFile.markdown).toContain("| 1 | コード | code | 101 | 何かのコード |");
        expect(markdownFile.markdown).toContain("| 4 | 更新日 | updatedate | 3月14日 | 何かの更新日 |");
    });

    it("parses the table-basic-sample12 fixture workbook as two vertically separated merge-heavy graph-paper tables", async () => {
        const api = core;
        const fixtureName = "table-basic-sample12.xlsx";
        const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
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
        expect(sheet.name).toBe("table-basic");
        expect(sheet.maxRow).toBe(14);
        expect(sheet.maxCol).toBe(20);
        expect(sheet.cells).toHaveLength(194);
        expect(sheet.merges).toHaveLength(40);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("方眼紙的様式のテスト");
        expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe(
            "方眼紙風のためにセル結合が多用されます"
        );
        expect(sheet.cells.find((cell) => cell.address === "B9")?.outputValue).toBe("テーブルその2");
        expect(sheet.cells.find((cell) => cell.address === "B14")?.formulaText).toBe("=B13+1");
        expect(sheet.cells.find((cell) => cell.address === "L13")?.outputValue).toBe("3月13日");

        expect(markdownFile.fileName).toBe("table-basic-sample12_001_table-basic.md");
        expect(markdownFile.summary.tables).toBe(2);
        expect(markdownFile.summary.merges).toBe(40);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(6);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-T7", "B10-T14"]);
        expect(markdownFile.markdown).toContain("# Book: table-basic-sample12.xlsx");
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-T7)");
        expect(markdownFile.markdown).toContain("### Table: 002 (B10-T14)");
        expect(markdownFile.markdown).toContain("方眼紙風のためにセル結合が多用されます");
        expect(markdownFile.markdown).toContain("テーブルその2");
    });

    it("parses the table-basic-sample13 fixture workbook as four merge-heavy graph-paper tables", async () => {
        const api = core;
        const fixtureName = "table-basic-sample13.xlsx";
        const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
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
        expect(sheet.name).toBe("table-basic");
        expect(sheet.maxRow).toBe(14);
        expect(sheet.maxCol).toBe(40);
        expect(sheet.cells).toHaveLength(389);
        expect(sheet.merges).toHaveLength(80);
        expect(sheet.cells.find((cell) => cell.address === "B9")?.outputValue).toBe("テーブルその3");
        expect(sheet.cells.find((cell) => cell.address === "V9")?.outputValue).toBe("テーブルその4");
        expect(sheet.cells.find((cell) => cell.address === "V14")?.formulaText).toBe("=V13+1");
        expect(sheet.cells.find((cell) => cell.address === "AJ11")?.outputValue).toBe("何かのコード");
        expect(sheet.cells.find((cell) => cell.address === "AF12")?.outputValue).toBe("Sabro");

        expect(markdownFile.fileName).toBe("table-basic-sample13_001_table-basic.md");
        expect(markdownFile.summary.tables).toBe(4);
        expect(markdownFile.summary.merges).toBe(80);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(12);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual([
            "B3-T7",
            "V3-AN7",
            "B10-T14",
            "V10-AN14",
        ]);
        expect(markdownFile.markdown).toContain("# Book: table-basic-sample13.xlsx");
        // Numbered in reading order: the left column of tables first, then the right column.
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-T7)");
        expect(markdownFile.markdown).toContain("### Table: 002 (B10-T14)");
        expect(markdownFile.markdown).toContain("### Table: 003 (V3-AN7)");
        expect(markdownFile.markdown).toContain("### Table: 004 (V10-AN14)");
        expect(markdownFile.markdown).toContain("| 1 | コード | code | 401 | 何かのコード |");
        expect(markdownFile.markdown).toContain("| 2 | 名前 | name | Sabro | 何かの名前 |");
        // U4 and U11 hold an unbordered "○" next to the tables. It stays in the output as text.
        expect(markdownFile.markdown.match(/^○$/gm)).toHaveLength(2);
        expect(markdownFile.markdown).toContain("| 2 | 名前 | name | Jiro | 何かの名前 |");
    });

    it("parses the table-basic-sample14 fixture workbook as one graph-paper table even with a few unmerged cells", async () => {
        const api = core;
        const fixtureName = "table-basic-sample14.xlsx";
        const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
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
        expect(sheet.name).toBe("table-basic");
        expect(sheet.maxRow).toBe(8);
        expect(sheet.maxCol).toBe(20);
        expect(sheet.cells).toHaveLength(98);
        expect(sheet.merges).toHaveLength(18);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("方眼紙的様式のテスト");
        expect(sheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("テーブルその1");
        expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toContain(
            "たまに結合漏れのセルがある場合"
        );
        expect(sheet.cells.find((cell) => cell.address === "L5")?.outputValue).toBe("Taro");
        expect(sheet.cells.find((cell) => cell.address === "L6")?.outputValue).toBe("3月13日");

        expect(markdownFile.fileName).toBe("table-basic-sample14_001_table-basic.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.merges).toBe(18);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(3);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-T7"]);
        expect(markdownFile.markdown).toContain("# Book: table-basic-sample14.xlsx");
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-T7)");
        expect(markdownFile.markdown).toContain("| 2 | 名前 | name | Taro | 何かの名前 |");
        // G6:K6 is not merged, so row 6 also tests the merge gap.
        expect(markdownFile.markdown).toContain("| 3 | 登録日 | entrydate | 3月13日 | 何かの登録日 |");
        expect(markdownFile.markdown).toContain("たまに結合漏れのセルがある場合");
    });

    it("parses the table-basic-sample15 fixture workbook as one graph-paper table with a vertical merged note", async () => {
        const api = core;
        const fixtureName = "table-basic-sample15.xlsx";
        const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
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
        expect(sheet.name).toBe("table-basic");
        expect(sheet.maxRow).toBe(8);
        expect(sheet.maxCol).toBe(20);
        expect(sheet.cells).toHaveLength(98);
        expect(sheet.merges).toHaveLength(19);
        expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe("※方眼紙＋結合＋さらに縦結合");
        expect(sheet.cells.find((cell) => cell.address === "P5")?.outputValue).toBe("何かの名前");
        expect(sheet.cells.find((cell) => cell.address === "P6")?.outputValue).toBe("登録および更新日");
        expect(sheet.cells.find((cell) => cell.address === "P7")?.outputValue).toBe("");

        expect(markdownFile.fileName).toBe("table-basic-sample15_001_table-basic.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.merges).toBe(19);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(3);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-T7"]);
        expect(markdownFile.markdown).toContain("# Book: table-basic-sample15.xlsx");
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-T7)");
        expect(markdownFile.markdown).toContain("| 3 | 登録日 | entrydate | 3月13日 | 登録および更新日 |");
        expect(markdownFile.markdown).toContain("| 4 | 更新日 | updatedate | 3月14日 | [↑M↑] |");
        expect(markdownFile.markdown).toContain("※方眼紙＋結合＋さらに縦結合");
    });

    it("parses the table-basic-sample16 fixture workbook and keeps the extra value column caused by a merge gap", async () => {
        const api = core;
        const fixtureName = "table-basic-sample16.xlsx";
        const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
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
        expect(sheet.name).toBe("table-basic");
        expect(sheet.maxRow).toBe(8);
        expect(sheet.maxCol).toBe(20);
        expect(sheet.cells).toHaveLength(98);
        expect(sheet.merges).toHaveLength(18);
        expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toContain(
            "たまに結合漏れのセルがあって"
        );
        expect(sheet.cells.find((cell) => cell.address === "L5")?.outputValue).toBe("Taro");
        expect(sheet.cells.find((cell) => cell.address === "N5")?.outputValue).toBe("Ito");

        expect(markdownFile.fileName).toBe("table-basic-sample16_001_table-basic.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.merges).toBe(18);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(3);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-T7"]);
        expect(markdownFile.markdown).toContain("# Book: table-basic-sample16.xlsx");
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-T7)");
        expect(markdownFile.markdown).toContain("| 項番 | 項目名称 | 物理名 | デフォルト値 | [←M←] | 備考 |");
        expect(markdownFile.markdown).toContain("| 2 | 名前 | name | Taro | Ito | 何かの名前 |");
        expect(markdownFile.markdown).toContain("たまに結合漏れのセルがあって、さらに複数文字が登場");
    });

    it("parses the table-grid-layout-sample01 fixture workbook as two graph-paper tables", async () => {
        const api = core;
        const fixtureName = "table-grid-layout-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
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
        expect(sheet.name).toBe("grid-layout");
        expect(sheet.maxRow).toBe(16);
        expect(sheet.maxCol).toBe(22);
        expect(sheet.cells).toHaveLength(281);
        expect(sheet.merges).toHaveLength(70);
        expect(sheet.cells.find((cell) => cell.address === "C8")?.outputValue).toBe("項番");
        expect(sheet.cells.find((cell) => cell.address === "R9")?.outputValue).toBe("担当コードの値");
        expect(sheet.cells.find((cell) => cell.address === "C15")?.formulaText).toBe("=C14+1");
        expect(sheet.cells.find((cell) => cell.address === "C16")?.formulaText).toBe("=C15+1");
        expect(sheet.cells.find((cell) => cell.address === "R16")?.outputValue).toBe("システムへの更新日");

        expect(markdownFile.fileName).toBe("table-grid-layout-sample01_001_grid-layout.md");
        expect(markdownFile.summary.tables).toBe(2);
        expect(markdownFile.summary.merges).toBe(70);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(10);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B2-U6", "C8-V16"]);
        expect(markdownFile.markdown).toContain("# Book: table-grid-layout-sample01.xlsx");
        expect(markdownFile.markdown).toContain("方眼紙的様式の動作確認");
        expect(markdownFile.markdown).toContain("### Table: 001 (B2-U6)");
        expect(markdownFile.markdown).toContain("### Table: 002 (C8-V16)");
        expect(markdownFile.markdown).toContain("| 1 | 担当コード | tantocode | 101 | 担当コードの値 |");
        expect(markdownFile.markdown).toContain("| 2 | 担当者名 | tantoname | Taro | 担当の名前 |");
        expect(markdownFile.markdown).toContain("| 6 | 備考 | remarks | なし | 補足事項 |");
        expect(markdownFile.markdown).toContain("| 8 | 更新日 | updatedate |  | システムへの更新日 |");
    });

    it("parses the table-border-priority-sample01 fixture workbook differently between balanced and border modes", async () => {
        const api = core;
        const fixtureName = "table-border-priority-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "table", fixtureName);
        const fixtureBytes = readFileSync(fixturePath);
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await api.parseWorkbook(arrayBuffer, fixtureName);
        const balancedFiles = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            tableDetectionMode: "balanced",
        });
        const borderFiles = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            tableDetectionMode: "border",
        });
        const sheet = workbook.sheets[0];
        const balancedFile = balancedFiles[0];
        const borderFile = borderFiles[0];

        expect(workbook.sheets).toHaveLength(1);
        expect(sheet.name).toBe("border-priority");
        expect(sheet.maxRow).toBe(6);
        expect(sheet.maxCol).toBe(2);
        expect(sheet.cells).toHaveLength(6);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("border-priority fixture");
        expect(sheet.cells.find((cell) => cell.address === "A3")?.outputValue).toBe("項目");
        expect(sheet.cells.find((cell) => cell.address === "B4")?.outputValue).toBe("100");

        expect(balancedFile.fileName).toBe("table-border-priority-sample01_001_border-priority.md");
        expect(balancedFile.summary.tables).toBe(1);
        expect(balancedFile.summary.tableDetectionMode).toBe("balanced");
        expect(balancedFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A3-B4"]);
        expect(balancedFile.markdown).toContain("# Book: table-border-priority-sample01.xlsx");
        expect(balancedFile.markdown).toContain("### Table: 001 (A3-B4)");
        expect(balancedFile.markdown).toContain("| 項目 | 値 |");
        expect(balancedFile.markdown).toContain("| A | 100 |");

        expect(borderFile.fileName).toBe("table-border-priority-sample01_001_border-priority.md");
        expect(borderFile.summary.tables).toBe(0);
        expect(borderFile.summary.tableDetectionMode).toBe("border");
        expect(borderFile.summary.tableScores).toHaveLength(0);
        expect(borderFile.markdown).toContain("# Book: table-border-priority-sample01.xlsx");
        expect(borderFile.markdown).not.toContain("### Table: 001");
        expect(borderFile.markdown).toContain("項目");
        expect(borderFile.markdown).toContain("100");
        expect(borderFile.markdown).toContain("※罫線優先モード確認用");
    });
});
