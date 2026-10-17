import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { applyMergeTokens } from "../../../src/structure/tables/table-detector.ts";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("merge workbooks", () => {
    const fixtureDir = FIXTURE_ROOT;

    it("parses the merge-pattern fixture workbook with concrete merge expectations", async () => {
        const api = core;
        const fixtureName = "merge-pattern-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "merge", fixtureName);
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
        expect(sheet.name).toBe("merge");
        expect(sheet.maxRow).toBe(19);
        expect(sheet.maxCol).toBe(5);
        expect(sheet.cells).toHaveLength(68);
        expect(sheet.merges.map((merge) => merge.ref)).toEqual([
            "B15:C16",
            "B17:C18",
            "D15:E16",
            "D17:E18",
            "B2:C2",
            "B3:D3",
            "B4:E4",
            "D2:E2",
            "B8:B9",
            "B10:B11",
            "C8:C10",
            "D8:D11",
        ]);

        expect(sheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("横結合");
        expect(sheet.cells.find((cell) => cell.address === "C2")?.outputValue).toBe("");
        expect(sheet.cells.find((cell) => cell.address === "D2")?.outputValue).toBe("横結合");
        expect(sheet.cells.find((cell) => cell.address === "E2")?.outputValue).toBe("");
        expect(sheet.cells.find((cell) => cell.address === "A9")?.formulaText).toBe("=A8+1");
        expect(sheet.cells.find((cell) => cell.address === "A9")?.outputValue).toBe("2");
        expect(sheet.cells.find((cell) => cell.address === "B9")?.outputValue).toBe("");
        expect(sheet.cells.find((cell) => cell.address === "C9")?.outputValue).toBe("");
        expect(sheet.cells.find((cell) => cell.address === "D11")?.outputValue).toBe("");
        expect(sheet.cells.find((cell) => cell.address === "B15")?.outputValue).toBe("2x2結合");
        expect(sheet.cells.find((cell) => cell.address === "C15")?.outputValue).toBe("");
        expect(sheet.cells.find((cell) => cell.address === "E16")?.outputValue).toBe("");
        expect(sheet.cells.find((cell) => cell.address === "C18")?.outputValue).toBe("");

        expect(markdownFile.fileName).toBe("merge-pattern-sample01_001_merge.md");
        expect(markdownFile.summary.tables).toBe(3);
        expect(markdownFile.summary.merges).toBe(12);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A1-E4", "A7-D11", "A14-E18"]);
        expect(markdownFile.markdown).toContain("## Sheet: merge");
        expect(markdownFile.markdown).toContain("# Book: merge-pattern-sample01.xlsx");
        expect(markdownFile.markdown).toContain("### Table: 001 (A1-E4)");
        expect(markdownFile.markdown).toContain("### Table: 002 (A7-D11)");
        expect(markdownFile.markdown).toContain("### Table: 003 (A14-E18)");
        expect(markdownFile.markdown).toContain("| 1 | 横結合 | [←M←] | 横結合 | [←M←] |");
        expect(markdownFile.markdown).toContain("| 2 | [↑M↑] | [↑M↑] | [↑M↑] |");
        expect(markdownFile.markdown).toContain("| 1 | 2x2結合 | [←M←] | 2x2結合 | [←M←] |");
        expect(markdownFile.markdown).toContain("※横結合のサンプルです");
        expect(markdownFile.markdown).toContain("※縦結合のサンプルです");
        expect(markdownFile.markdown).toContain("※2x2結合のサンプルです");
        expect(markdownFile.markdown).toContain("[←M←]");
        expect(markdownFile.markdown).toContain("[↑M↑]");
    });

    it("parses the merge-multiline fixture workbook with concrete multiline-merge expectations", async () => {
        const api = core;
        const fixtureName = "merge-multiline-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "merge", fixtureName);
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
        expect(sheet.name).toBe("merge-multiline");
        expect(sheet.maxRow).toBe(6);
        expect(sheet.maxCol).toBe(3);
        expect(sheet.cells).toHaveLength(11);
        expect(sheet.merges.map((merge) => merge.ref)).toEqual(["B3:C4"]);

        expect(sheet.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("1行目\n2行目");
        expect(sheet.cells.find((cell) => cell.address === "C3")?.outputValue).toBe("");
        expect(sheet.cells.find((cell) => cell.address === "B4")?.outputValue).toBe("");
        expect(sheet.cells.find((cell) => cell.address === "A6")?.outputValue).toBe("※結合セル内の改行確認用");

        expect(markdownFile.fileName).toBe("merge-multiline-sample01_001_merge-multiline.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.merges).toBe(1);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(0);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A1-C4"]);
        expect(markdownFile.markdown).toContain("## Sheet: merge-multiline");
        expect(markdownFile.markdown).toContain("# Book: merge-multiline-sample01.xlsx");
        expect(markdownFile.markdown).toContain("### Table: 001 (A1-C4)");
        expect(markdownFile.markdown).toContain("| 1 | 1行目 2行目 | [←M←] |");
        expect(markdownFile.markdown).toContain("| 2 | [↑M↑] | [↑M↑] |");
        expect(markdownFile.markdown).toContain("※結合セル内の改行確認用");
    });

    it("expands merged cells with structural tokens", () => {
        const api = core;
        const matrix = [
            ["見出し", "", ""],
            ["", "", ""],
            ["値", "2", "3"],
        ];
        applyMergeTokens(matrix, [{ startRow: 1, startCol: 1, endRow: 2, endCol: 3, ref: "A1:C2" }], 1, 1, 3, 3);

        expect(matrix[0][1]).toBe("[←M←]");
        expect(matrix[0][2]).toBe("[←M←]");
        expect(matrix[1][0]).toBe("[↑M↑]");
        expect(matrix[1][2]).toBe("[↑M↑]");
    });
});
