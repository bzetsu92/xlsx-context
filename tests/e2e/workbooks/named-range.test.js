import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("named-range workbooks", () => {
    const fixtureDir = FIXTURE_ROOT;

    it("parses the named-range fixture workbook with concrete defined-name expectations", async () => {
        const api = core;
        const fixtureName = "named-range-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "named-range", fixtureName);
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
        const summarySheet = workbook.sheets[0];
        const otherSheet = workbook.sheets[1];
        const summaryFile = files[0];
        const otherFile = files[1];

        expect(workbook.definedNames).toEqual([
            {
                name: "BaseName",
                formulaText: "=Summary!$B$3",
                localSheetName: null,
            },
            {
                name: "BaseRange",
                formulaText: "=Summary!$B$4:$B$5",
                localSheetName: null,
            },
            {
                name: "LocalCross",
                formulaText: "=Other!$B$2",
                localSheetName: "Other",
            },
        ]);

        expect(workbook.sheets).toHaveLength(2);

        expect(summarySheet.name).toBe("Summary");
        expect(summarySheet.maxRow).toBe(13);
        expect(summarySheet.maxCol).toBe(4);
        expect(summarySheet.cells).toHaveLength(25);
        expect(summarySheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("definedNames サンプル");
        expect(summarySheet.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("Base");
        expect(summarySheet.cells.find((cell) => cell.address === "D3")?.formulaText).toBe("=BaseName");
        expect(summarySheet.cells.find((cell) => cell.address === "D3")?.outputValue).toBe("Base");
        expect(summarySheet.cells.find((cell) => cell.address === "D3")?.resolutionStatus).toBe("resolved");
        expect(summarySheet.cells.find((cell) => cell.address === "D4")?.formulaText).toBe("=SUM(BaseRange)");
        expect(summarySheet.cells.find((cell) => cell.address === "D4")?.outputValue).toBe("30");
        expect(summarySheet.cells.find((cell) => cell.address === "D4")?.resolutionStatus).toBe("resolved");

        expect(otherSheet.name).toBe("Other");
        expect(otherSheet.maxRow).toBe(2);
        expect(otherSheet.maxCol).toBe(4);
        expect(otherSheet.cells).toHaveLength(3);
        expect(otherSheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("LocalCross元");
        expect(otherSheet.cells.find((cell) => cell.address === "B2")?.outputValue).toBe("CrossRef");
        expect(otherSheet.cells.find((cell) => cell.address === "D2")?.formulaText).toBe("=LocalCross");
        expect(otherSheet.cells.find((cell) => cell.address === "D2")?.outputValue).toBe("CrossRef");
        expect(otherSheet.cells.find((cell) => cell.address === "D2")?.resolutionStatus).toBe("resolved");

        expect(summaryFile.fileName).toBe("named-range-sample01_001_Summary.md");
        expect(summaryFile.summary.tables).toBe(1);
        expect(summaryFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A3-B5"]);
        expect(summaryFile.summary.formulaDiagnostics).toHaveLength(2);
        expect(summaryFile.summary.formulaDiagnostics.every((diagnostic) => diagnostic.source === "cached_value")).toBe(
            true
        );
        expect(summaryFile.markdown).toContain("## Sheet: Summary");
        expect(summaryFile.markdown).toContain("# Book: named-range-sample01.xlsx");
        expect(summaryFile.markdown).toContain("| BaseName元 | Base |");
        expect(summaryFile.markdown).toContain("| BaseRange1 | 10 |");
        expect(summaryFile.markdown).toContain("| BaseRange2 | 20 |");

        expect(otherFile.fileName).toBe("named-range-sample01_002_Other.md");
        expect(otherFile.summary.tables).toBe(0);
        expect(otherFile.summary.tableScores).toHaveLength(0);
        expect(otherFile.summary.formulaDiagnostics).toHaveLength(1);
        expect(otherFile.markdown).toContain("## Sheet: Other");
        expect(otherFile.markdown).toContain("LocalCross元");
        expect(otherFile.markdown).toContain("CrossRef");
    });
});
