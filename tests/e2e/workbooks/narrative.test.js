import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("narrative workbooks", () => {
    const fixtureDir = FIXTURE_ROOT;

    it("parses the narrative fixture workbook with concrete narrative and table expectations", async () => {
        const api = core;
        const fixtureName = "narrative-vs-table-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "narrative", fixtureName);
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
        expect(sheet.name).toBe("narrative-vs-table");
        expect(sheet.maxRow).toBe(13);
        expect(sheet.maxCol).toBe(6);
        expect(sheet.cells).toHaveLength(26);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("地の文と表の判定");
        expect(sheet.cells.find((cell) => cell.address === "A3")?.outputValue).toBe(
            "この設計書は受注入力画面を説明する。"
        );
        expect(sheet.cells.find((cell) => cell.address === "A4")?.outputValue).toBe(
            "外部システムとの連携条件を以下に示す。"
        );
        expect(sheet.cells.find((cell) => cell.address === "A5")?.outputValue).toBe("本文は罫線なしのままにする。");
        expect(sheet.cells.find((cell) => cell.address === "A7")?.outputValue).toBe("項目一覧");
        expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe("項番");
        expect(sheet.cells.find((cell) => cell.address === "F8")?.outputValue).toBe("備考");
        expect(sheet.cells.find((cell) => cell.address === "B10")?.formulaText).toBe("=B9+1");
        expect(sheet.cells.find((cell) => cell.address === "B10")?.outputValue).toBe("2");
        expect(sheet.cells.find((cell) => cell.address === "B11")?.formulaText).toBe("=B10+1");
        expect(sheet.cells.find((cell) => cell.address === "B11")?.outputValue).toBe("3");
        expect(sheet.cells.find((cell) => cell.address === "E11")?.outputValue).toBe("3月13日");
        expect(sheet.cells.find((cell) => cell.address === "B13")?.outputValue).toBe("※注記: この表はサンプルです。");

        expect(markdownFile.fileName).toBe("narrative-vs-table-sample01_001_narrative-vs-table.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.merges).toBe(0);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(2);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B8-F11"]);
        expect(markdownFile.markdown).toContain("## Sheet: narrative-vs-table");
        expect(markdownFile.markdown).toContain("# Book: narrative-vs-table-sample01.xlsx");
        expect(markdownFile.markdown).toContain("地の文と表の判定");
        expect(markdownFile.markdown).toContain("この設計書は受注入力画面を説明する。");
        expect(markdownFile.markdown).toContain("外部システムとの連携条件を以下に示す。");
        expect(markdownFile.markdown).toContain("本文は罫線なしのままにする。");
        expect(markdownFile.markdown).toContain("項目一覧");
        expect(markdownFile.markdown).toContain("### Table: 001 (B8-F11)");
        expect(markdownFile.markdown).toContain("| 項番 | 項目名称 | 物理名 | 初期値 | 備考 |");
        expect(markdownFile.markdown).toContain("| 1 | コード | code | 101 | 何かのコード |");
        expect(markdownFile.markdown).toContain("| 3 | 登録日 | entrydate | 3月13日 | 何かの登録日 |");
        expect(markdownFile.markdown).toContain("※注記: この表はサンプルです。");
    });

    it("renders indented narrative rows as heading plus markdown bullets", () => {
        const api = core;
        const workbook = { name: "todo.xlsx" };
        const sheet = {
            name: "To Do リスト",
            index: 1,
            path: "xl/worksheets/sheet1.xml",
            merges: [],
            tables: [],
            images: [],
            maxRow: 8,
            maxCol: 2,
            cells: [
                {
                    row: 1,
                    col: 1,
                    address: "A1",
                    valueType: "str",
                    rawValue: "To Do リスト",
                    outputValue: "To Do リスト",
                    formulaText: "",
                    resolutionStatus: null,
                    styleIndex: 0,
                    borders: { top: false, bottom: false, left: false, right: false },
                    numFmtId: 0,
                    formatCode: "General",
                },
                {
                    row: 2,
                    col: 1,
                    address: "A2",
                    valueType: "str",
                    rawValue: "学校が始まる前に確認します",
                    outputValue: "学校が始まる前に確認します",
                    formulaText: "",
                    resolutionStatus: null,
                    styleIndex: 0,
                    borders: { top: false, bottom: false, left: false, right: false },
                    numFmtId: 0,
                    formatCode: "General",
                },
                {
                    row: 3,
                    col: 1,
                    address: "A3",
                    valueType: "str",
                    rawValue: "完了 タスク",
                    outputValue: "完了 タスク",
                    formulaText: "",
                    resolutionStatus: null,
                    styleIndex: 0,
                    borders: { top: false, bottom: false, left: false, right: false },
                    numFmtId: 0,
                    formatCode: "General",
                },
                {
                    row: 4,
                    col: 2,
                    address: "B4",
                    valueType: "str",
                    rawValue: "登録フォームに記入する",
                    outputValue: "登録フォームに記入する",
                    formulaText: "",
                    resolutionStatus: null,
                    styleIndex: 0,
                    borders: { top: false, bottom: false, left: false, right: false },
                    numFmtId: 0,
                    formatCode: "General",
                },
                {
                    row: 5,
                    col: 2,
                    address: "B5",
                    valueType: "str",
                    rawValue: "健康診断をスケジュールする",
                    outputValue: "健康診断をスケジュールする",
                    formulaText: "",
                    resolutionStatus: null,
                    styleIndex: 0,
                    borders: { top: false, bottom: false, left: false, right: false },
                    numFmtId: 0,
                    formatCode: "General",
                },
                {
                    row: 6,
                    col: 2,
                    address: "B6",
                    valueType: "str",
                    rawValue: "予防接種を確認する",
                    outputValue: "予防接種を確認する",
                    formulaText: "",
                    resolutionStatus: null,
                    styleIndex: 0,
                    borders: { top: false, bottom: false, left: false, right: false },
                    numFmtId: 0,
                    formatCode: "General",
                },
                {
                    row: 7,
                    col: 2,
                    address: "B7",
                    valueType: "str",
                    rawValue: "教師に会う",
                    outputValue: "教師に会う",
                    formulaText: "",
                    resolutionStatus: null,
                    styleIndex: 0,
                    borders: { top: false, bottom: false, left: false, right: false },
                    numFmtId: 0,
                    formatCode: "General",
                },
            ],
        };

        const markdownFile = api.convertSheetToMarkdown(workbook, sheet, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });

        expect(markdownFile.markdown).toContain("To Do リスト");
        expect(markdownFile.markdown).toContain("学校が始まる前に確認します");
        expect(markdownFile.markdown).toContain("**完了 タスク**");
        expect(markdownFile.markdown).toContain("- 登録フォームに記入する");
        expect(markdownFile.markdown).toContain("- 健康診断をスケジュールする");
        expect(markdownFile.markdown).toContain("- 予防接種を確認する");
        expect(markdownFile.markdown).toContain("- 教師に会う");
    });
});
