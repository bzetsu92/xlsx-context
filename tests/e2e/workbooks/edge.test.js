import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("edge workbooks", () => {
    const fixtureDir = FIXTURE_ROOT;

    it("parses the edge-empty fixture workbook with concrete sparse-sheet expectations", async () => {
        const api = core;
        const fixtureName = "edge-empty-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "edge", fixtureName);
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
        expect(sheet.name).toBe("edge-empty");
        expect(sheet.maxRow).toBe(7);
        expect(sheet.maxCol).toBe(3);
        expect(sheet.cells).toHaveLength(2);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("空系境界サンプル");
        expect(sheet.cells.find((cell) => cell.address === "C7")?.outputValue).toBe("only-value");

        expect(markdownFile.fileName).toBe("edge-empty-sample01_001_edge-empty.md");
        expect(markdownFile.summary.tables).toBe(0);
        expect(markdownFile.summary.merges).toBe(0);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.tableScores).toHaveLength(0);
        expect(markdownFile.markdown).toContain("## Sheet: edge-empty");
        expect(markdownFile.markdown).toContain("# Book: edge-empty-sample01.xlsx");
        expect(markdownFile.markdown).toContain("空系境界サンプル");
        expect(markdownFile.markdown).toContain("only-value");
        expect(markdownFile.markdown).not.toContain("### Table:");
    });

    it("parses the edge-weird-sheetname fixture workbook with concrete sheet-name expectations", async () => {
        const api = core;
        const fixtureName = "edge-weird-sheetname-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "edge", fixtureName);
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
        expect(sheet.name).toBe("A B-東京&大阪.01");
        expect(sheet.maxRow).toBe(4);
        expect(sheet.maxCol).toBe(4);
        expect(sheet.cells).toHaveLength(16);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("項番");
        expect(sheet.cells.find((cell) => cell.address === "B1")?.outputValue).toBe("名称");
        expect(sheet.cells.find((cell) => cell.address === "C4")?.outputValue).toBe("3月13日");
        expect(sheet.cells.find((cell) => cell.address === "D4")?.outputValue).toBe("何かの登録日");

        expect(markdownFile.fileName).toBe("edge-weird-sheetname-sample01_001_A_B-東京_大阪.01.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.merges).toBe(0);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A1-D4"]);
        expect(markdownFile.markdown).toContain("## Sheet: A B-東京&大阪.01");
        expect(markdownFile.markdown).toContain("# Book: edge-weird-sheetname-sample01.xlsx");
        expect(markdownFile.markdown).not.toContain("- Sheet: A B-東京&大阪.01");
        expect(markdownFile.markdown).toContain("### Table: 001 (A1-D4)");
        expect(markdownFile.markdown).toContain("| 項番 | 名称 | 値 | 備考 |");
        expect(markdownFile.markdown).toContain("| 1 | コード | 101 | 何かのコード |");
        expect(markdownFile.markdown).toContain("| 3 | 登録日 | 3月13日 | 何かの登録日 |");
    });
});
