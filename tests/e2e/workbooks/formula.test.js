import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { createSummaryText } from "../../../src/outputs/markdown/markdown-export.ts";
import { createStoredZip } from "./stored-zip.js";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("formula workbooks", () => {
    const fixtureDir = FIXTURE_ROOT;

    it("parses the formula-basic fixture workbook with concrete formula expectations", async () => {
        const api = core;
        const fixtureName = "formula-basic-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "formula", fixtureName);
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
        expect(sheet.name).toBe("formula");
        expect(sheet.maxRow).toBe(13);
        expect(sheet.maxCol).toBe(2);
        expect(sheet.cells).toHaveLength(23);

        expect(sheet.cells.find((cell) => cell.address === "B5")?.formulaText).toBe("=B3");
        expect(sheet.cells.find((cell) => cell.address === "B5")?.outputValue).toBe("10");
        expect(sheet.cells.find((cell) => cell.address === "B5")?.resolutionStatus).toBe("resolved");
        expect(sheet.cells.find((cell) => cell.address === "B6")?.formulaText).toBe("=B3+B4");
        expect(sheet.cells.find((cell) => cell.address === "B6")?.outputValue).toBe("15");
        expect(sheet.cells.find((cell) => cell.address === "B7")?.formulaText).toBe('=IF(B3>B4,"OK","NG")');
        expect(sheet.cells.find((cell) => cell.address === "B7")?.outputValue).toBe("OK");
        expect(sheet.cells.find((cell) => cell.address === "B8")?.formulaText).toBe("=SUM(B3:B4)");
        expect(sheet.cells.find((cell) => cell.address === "B8")?.outputValue).toBe("15");
        expect(sheet.cells.find((cell) => cell.address === "B9")?.formulaText).toBe('=COUNTIF(B3:B4,">7")');
        expect(sheet.cells.find((cell) => cell.address === "B9")?.outputValue).toBe("1");
        expect(sheet.cells.find((cell) => cell.address === "B10")?.formulaText).toBe('=TEXT(B3,"0000")');
        expect(sheet.cells.find((cell) => cell.address === "B10")?.outputValue).toBe("0010");
        expect(sheet.cells.find((cell) => cell.address === "B11")?.formulaText).toBe("=DATE(2024,3,17)");
        expect(sheet.cells.find((cell) => cell.address === "B11")?.outputValue).toBe("2024/3/17");
        expect(sheet.cells.find((cell) => cell.address === "B12")?.formulaText).toBe('=VALUE("1,234.5")');
        expect(sheet.cells.find((cell) => cell.address === "B12")?.outputValue).toBe("1234.5");
        expect(sheet.cells.find((cell) => cell.address === "B13")?.formulaText).toBe('=VALUE("2024/03/17")');
        expect(sheet.cells.find((cell) => cell.address === "B13")?.outputValue).toBe("45368");

        expect(markdownFile.fileName).toBe("formula-basic-sample01_001_formula.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.merges).toBe(0);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(9);
        expect(
            markdownFile.summary.formulaDiagnostics.every((diagnostic) => diagnostic.source === "cached_value")
        ).toBe(true);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A3-B13"]);
        expect(markdownFile.markdown).toContain("## Sheet: formula");
        expect(markdownFile.markdown).toContain("# Book: formula-basic-sample01.xlsx");
        expect(markdownFile.markdown).toContain("| ref | 10 |");
        expect(markdownFile.markdown).toContain("| arith | 15 |");
        expect(markdownFile.markdown).toContain("| if | OK |");
        expect(markdownFile.markdown).toContain("| sum | 15 |");
        expect(markdownFile.markdown).toContain("| countif | 1 |");
        expect(markdownFile.markdown).toContain("| text | 0010 |");
        expect(markdownFile.markdown).toContain("| date | 2024/3/17 |");
        expect(markdownFile.markdown).toContain("| value\\_num | 1234.5 |");
        expect(markdownFile.markdown).toContain("| value\\_date | 45368 |");
    });

    it("parses the formula-spill fixture workbook with concrete spill-like expectations", async () => {
        const api = core;
        const fixtureName = "formula-spill-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "formula", fixtureName);
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
        expect(sheet.name).toBe("spill-sample");
        expect(sheet.maxRow).toBe(6);
        expect(sheet.maxCol).toBe(5);
        expect(sheet.cells).toHaveLength(11);

        expect(sheet.cells.find((cell) => cell.address === "A4")?.outputValue).toBe("1");
        expect(sheet.cells.find((cell) => cell.address === "A6")?.outputValue).toBe("3");
        expect(sheet.cells.find((cell) => cell.address === "C4")?.formulaText).toBe("=_xlfn.SEQUENCE(3)");
        expect(sheet.cells.find((cell) => cell.address === "C4")?.outputValue).toBe("1");
        expect(sheet.cells.find((cell) => cell.address === "C4")?.resolutionStatus).toBe("resolved");
        expect(sheet.cells.find((cell) => cell.address === "C5")?.outputValue).toBe("2");
        expect(sheet.cells.find((cell) => cell.address === "C6")?.outputValue).toBe("3");
        expect(sheet.cells.find((cell) => cell.address === "E4")?.formulaText).toBe("=SUM(_xlfn.ANCHORARRAY(C4))");
        expect(sheet.cells.find((cell) => cell.address === "E4")?.outputValue).toBe("6");
        expect(sheet.cells.find((cell) => cell.address === "E4")?.resolutionStatus).toBe("resolved");

        expect(markdownFile.fileName).toBe("formula-spill-sample01_001_spill-sample.md");
        // Columns A and C share rows 3-6 across the empty spacer column B; E3:E4 stays separate.
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.markdown).toContain("### Table: 001 (A3-C6)");
        expect(markdownFile.summary.merges).toBe(0);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(2);
        expect(
            markdownFile.summary.formulaDiagnostics.every((diagnostic) => diagnostic.source === "cached_value")
        ).toBe(true);
        expect(markdownFile.summary.tableScores).toHaveLength(1);
        expect(markdownFile.markdown).toContain("## Sheet: spill-sample");
        expect(markdownFile.markdown).toContain("# Book: formula-spill-sample01.xlsx");
        expect(markdownFile.markdown).toContain("spill サンプル");
        expect(markdownFile.markdown).toContain("| src1 | spill\\_ref |");
        expect(markdownFile.markdown).toContain("| 1 | 1 |");
        expect(markdownFile.markdown).toContain("| 3 | 3 |");
        expect(markdownFile.markdown).toContain("spill\\_sum");
    });

    it("parses the formula-crosssheet fixture workbook with concrete cross-sheet expectations", async () => {
        const api = core;
        const fixtureName = "formula-crosssheet-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "formula", fixtureName);
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
        const sheet1 = workbook.sheets[0];
        const sheet2 = workbook.sheets[1];
        const sheet3 = workbook.sheets[2];
        const sheet1File = files[0];
        const sheet2File = files[1];
        const sheet3File = files[2];

        expect(workbook.sheets).toHaveLength(3);

        expect(sheet1.name).toBe("Sheet1");
        expect(sheet1.maxRow).toBe(13);
        expect(sheet1.maxCol).toBe(2);
        expect(sheet1.cells).toHaveLength(23);
        expect(sheet1.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("複数シート参照サンプル");
        expect(sheet1.cells.find((cell) => cell.address === "B3")?.formulaText).toBe("=Sheet2!B3");
        expect(sheet1.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("CrossValue");
        expect(sheet1.cells.find((cell) => cell.address === "B3")?.resolutionStatus).toBe("resolved");
        expect(sheet1.cells.find((cell) => cell.address === "B4")?.formulaText).toBe("=日本語シート!C4");
        expect(sheet1.cells.find((cell) => cell.address === "B4")?.outputValue).toBe("日本語参照値");
        expect(sheet1.cells.find((cell) => cell.address === "B4")?.resolutionStatus).toBe("resolved");
        expect(sheet1.cells.find((cell) => cell.address === "B5")?.formulaText).toBe("=SUM(Sheet2!A1:B2)");
        expect(sheet1.cells.find((cell) => cell.address === "B5")?.outputValue).toBe("10");
        expect(sheet1.cells.find((cell) => cell.address === "B5")?.resolutionStatus).toBe("resolved");

        expect(sheet2.name).toBe("Sheet2");
        expect(sheet2.maxRow).toBe(3);
        expect(sheet2.maxCol).toBe(2);
        expect(sheet2.cells).toHaveLength(5);
        expect(sheet2.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("CrossValue");

        expect(sheet3.name).toBe("日本語シート");
        expect(sheet3.maxRow).toBe(4);
        expect(sheet3.maxCol).toBe(3);
        expect(sheet3.cells).toHaveLength(1);
        expect(sheet3.cells.find((cell) => cell.address === "C4")?.outputValue).toBe("日本語参照値");

        expect(sheet1File.fileName).toBe("formula-crosssheet-sample01_001_Sheet1.md");
        expect(sheet1File.summary.tables).toBe(1);
        expect(sheet1File.summary.tableScores.map((detail) => detail.range)).toEqual(["A3-B5"]);
        expect(sheet1File.summary.formulaDiagnostics).toHaveLength(3);
        expect(sheet1File.summary.formulaDiagnostics.every((diagnostic) => diagnostic.source === "cached_value")).toBe(
            true
        );
        expect(sheet1File.markdown).toContain("## Sheet: Sheet1");
        expect(sheet1File.markdown).toContain("# Book: formula-crosssheet-sample01.xlsx");
        expect(sheet1File.markdown).toContain("| sheet2\\_ref | CrossValue |");
        expect(sheet1File.markdown).toContain("| jp\\_sheet\\_ref | 日本語参照値 |");
        expect(sheet1File.markdown).toContain("| sum\\_range | 10 |");

        expect(sheet2File.fileName).toBe("formula-crosssheet-sample01_002_Sheet2.md");
        expect(sheet2File.summary.tables).toBe(1);
        expect(sheet2File.summary.tableScores.map((detail) => detail.range)).toEqual(["A1-B3"]);
        expect(sheet2File.markdown).toContain("| 1 | 2 |");
        expect(sheet2File.markdown).toContain("| 3 | 4 |");
        expect(sheet2File.markdown).toContain("|  | CrossValue |");

        expect(sheet3File.fileName).toBe("formula-crosssheet-sample01_003_日本語シート.md");
        expect(sheet3File.summary.tables).toBe(0);
        expect(sheet3File.summary.tableScores).toHaveLength(0);
        expect(sheet3File.markdown).toContain("## Sheet: 日本語シート");
        expect(sheet3File.markdown).toContain("日本語参照値");
    });

    it("parses the formula-shared fixture workbook with concrete shared-formula expectations", async () => {
        const api = core;
        const fixtureName = "formula-shared-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "formula", fixtureName);
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
        expect(sheet.name).toBe("formula");
        expect(sheet.maxRow).toBe(13);
        expect(sheet.maxCol).toBe(4);
        expect(sheet.cells).toHaveLength(27);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("No");
        expect(sheet.cells.find((cell) => cell.address === "B1")?.outputValue).toBe("連番");
        expect(sheet.cells.find((cell) => cell.address === "D1")?.outputValue).toBe("shared formula サンプル");
        expect(sheet.cells.find((cell) => cell.address === "B3")?.formulaText).toBe("=B2+1");
        expect(sheet.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("2");
        expect(sheet.cells.find((cell) => cell.address === "B3")?.resolutionStatus).toBe("resolved");
        expect(sheet.cells.find((cell) => cell.address === "B4")?.formulaText).toBe("=B3+1");
        expect(sheet.cells.find((cell) => cell.address === "B4")?.outputValue).toBe("3");
        expect(sheet.cells.find((cell) => cell.address === "B5")?.formulaText).toBe("=B4+1");
        expect(sheet.cells.find((cell) => cell.address === "B5")?.outputValue).toBe("4");
        expect(sheet.cells.find((cell) => cell.address === "B6")?.formulaText).toBe("=B5+1");
        expect(sheet.cells.find((cell) => cell.address === "B6")?.outputValue).toBe("5");
        expect(sheet.cells.find((cell) => cell.address === "B11")?.formulaText).toBe("=B10+1");
        expect(sheet.cells.find((cell) => cell.address === "B11")?.outputValue).toBe("10");
        expect(sheet.cells.find((cell) => cell.address === "B11")?.resolutionStatus).toBe("resolved");

        expect(markdownFile.fileName).toBe("formula-shared-sample01_001_formula.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.merges).toBe(0);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.formulaDiagnostics).toHaveLength(9);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A1-B11"]);
        expect(markdownFile.markdown).toContain("## Sheet: formula");
        expect(markdownFile.markdown).toContain("# Book: formula-shared-sample01.xlsx");
        expect(markdownFile.markdown).toContain("| No | 連番 |");
        expect(markdownFile.markdown).toContain("| 1 | 1 |");
        expect(markdownFile.markdown).toContain("| 2 | 2 |");
        expect(markdownFile.markdown).toContain("| 5 | 5 |");
        expect(markdownFile.markdown).toContain("| 10 | 10 |");
    });

    it("resolves simple same-workbook formulas when cached value is missing", async () => {
        const api = core;
        const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <definedNames>
        <definedName name="BaseName">Summary!$A$1</definedName>
        <definedName name="BaseRange">Summary!$G$1:$H$1</definedName>
        <definedName name="LocalCross" localSheetId="0">'Other Sheet'!$B$2</definedName>
    </definedNames>
    <sheets>
        <sheet name="Summary" sheetId="1" r:id="rId1"/>
        <sheet name="Other Sheet" sheetId="2" r:id="rId2"/>
    </sheets>
</workbook>`;
        const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`;
        const sheet1Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr"><is><t>Base</t></is></c>
            <c r="B1"><f>A1</f></c>
            <c r="C1"><f>'Other Sheet'!B2</f></c>
            <c r="D1"><f>BaseName</f></c>
            <c r="E1"><f>SUM(BaseRange)</f></c>
            <c r="F1"><f>LocalCross</f></c>
            <c r="G1"><v>10</v></c>
            <c r="H1"><v>20</v></c>
        </row>
    </sheetData>
</worksheet>`;
        const sheet2Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="2">
            <c r="B2" t="inlineStr"><is><t>CrossRef</t></is></c>
        </row>
    </sheetData>
</worksheet>`;
        const zip = createStoredZip([
            {
                name: "[Content_Types].xml",
                data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`,
            },
            { name: "xl/workbook.xml", data: workbookXml },
            { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
            { name: "xl/worksheets/sheet1.xml", data: sheet1Xml },
            { name: "xl/worksheets/sheet2.xml", data: sheet2Xml },
        ]);

        const workbook = await api.parseWorkbook(zip, "formula.xlsx");
        const summarySheet = workbook.sheets[0];
        const b1 = summarySheet.cells.find((cell) => cell.address === "B1");
        const c1 = summarySheet.cells.find((cell) => cell.address === "C1");
        const d1 = summarySheet.cells.find((cell) => cell.address === "D1");
        const e1 = summarySheet.cells.find((cell) => cell.address === "E1");
        const f1 = summarySheet.cells.find((cell) => cell.address === "F1");
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });

        expect(b1?.outputValue).toBe("Base");
        expect(b1?.resolutionStatus).toBe("resolved");
        expect(c1?.outputValue).toBe("CrossRef");
        expect(c1?.resolutionStatus).toBe("resolved");
        expect(d1?.outputValue).toBe("Base");
        expect(e1?.outputValue).toBe("30");
        expect(f1?.outputValue).toBe("CrossRef");
        expect(files[0].summary.formulaDiagnostics).toHaveLength(5);
        expect(files[0].summary.formulaDiagnostics[0].status).toBe("resolved");
        expect(files[0].summary.formulaDiagnostics[4].status).toBe("resolved");
        expect(createSummaryText(files[0])).toContain("Formula resolved: 5");
    });

    it("evaluates simple arithmetic formulas over same-workbook references", async () => {
        const api = core;
        const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Summary" sheetId="1" r:id="rId1"/>
        <sheet name="Other" sheetId="2" r:id="rId2"/>
    </sheets>
</workbook>`;
        const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`;
        const sheet1Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1"><v>10</v></c>
            <c r="B1"><v>5</v></c>
            <c r="C1"><f>A1+B1</f></c>
            <c r="D1"><f>(A1-B1)*2</f></c>
            <c r="E1"><f>Other!B2/2</f></c>
            <c r="F1"><f>$A$1+$B1</f></c>
            <c r="G1"><f>Other!$B$2/2</f></c>
            <c r="H1"><f>ROUND(Other!B2/3,2)</f></c>
            <c r="I1"><f>ROUNDUP(Other!B2/3,2)</f></c>
            <c r="J1"><f>ROUNDDOWN(Other!B2/3,2)</f></c>
            <c r="K1"><f>INT(Other!B2/3)</f></c>
        </row>
    </sheetData>
</worksheet>`;
        const sheet2Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="2">
            <c r="B2"><v>8</v></c>
        </row>
    </sheetData>
</worksheet>`;
        const zip = createStoredZip([
            {
                name: "[Content_Types].xml",
                data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`,
            },
            { name: "xl/workbook.xml", data: workbookXml },
            { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
            { name: "xl/worksheets/sheet1.xml", data: sheet1Xml },
            { name: "xl/worksheets/sheet2.xml", data: sheet2Xml },
        ]);

        const workbook = await api.parseWorkbook(zip, "arith.xlsx");
        const summarySheet = workbook.sheets[0];
        const c1 = summarySheet.cells.find((cell) => cell.address === "C1");
        const d1 = summarySheet.cells.find((cell) => cell.address === "D1");
        const e1 = summarySheet.cells.find((cell) => cell.address === "E1");
        const f1 = summarySheet.cells.find((cell) => cell.address === "F1");
        const g1 = summarySheet.cells.find((cell) => cell.address === "G1");
        const h1 = summarySheet.cells.find((cell) => cell.address === "H1");
        const i1 = summarySheet.cells.find((cell) => cell.address === "I1");
        const j1 = summarySheet.cells.find((cell) => cell.address === "J1");
        const k1 = summarySheet.cells.find((cell) => cell.address === "K1");

        expect(c1?.outputValue).toBe("15");
        expect(d1?.outputValue).toBe("10");
        expect(e1?.outputValue).toBe("4");
        expect(f1?.outputValue).toBe("15");
        expect(g1?.outputValue).toBe("4");
        expect(h1?.outputValue).toBe("2.67");
        expect(i1?.outputValue).toBe("2.67");
        expect(j1?.outputValue).toBe("2.66");
        expect(k1?.outputValue).toBe("2");
        expect(c1?.resolutionStatus).toBe("resolved");
        expect(d1?.resolutionStatus).toBe("resolved");
        expect(e1?.resolutionStatus).toBe("resolved");
        expect(f1?.resolutionStatus).toBe("resolved");
        expect(g1?.resolutionStatus).toBe("resolved");
        expect(h1?.resolutionStatus).toBe("resolved");
        expect(i1?.resolutionStatus).toBe("resolved");
        expect(j1?.resolutionStatus).toBe("resolved");
        expect(k1?.resolutionStatus).toBe("resolved");
    });

    it("expands shared formulas and resolves follower cells", async () => {
        const api = core;
        const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Summary" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>`;
        const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
        const sheetXml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1"><v>10</v></c>
            <c r="B1"><f t="shared" ref="B1:B3" si="0">A1+1</f><v>11</v></c>
        </row>
        <row r="2">
            <c r="A2"><v>20</v></c>
            <c r="B2"><f t="shared" si="0"/><v>21</v></c>
        </row>
        <row r="3">
            <c r="A3"><v>30</v></c>
            <c r="B3"><f t="shared" si="0"/></c>
        </row>
    </sheetData>
</worksheet>`;
        const zip = createStoredZip([
            {
                name: "[Content_Types].xml",
                data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`,
            },
            { name: "xl/workbook.xml", data: workbookXml },
            { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
            { name: "xl/worksheets/sheet1.xml", data: sheetXml },
        ]);

        const workbook = await api.parseWorkbook(zip, "shared-formula.xlsx");
        const summarySheet = workbook.sheets[0];

        expect(summarySheet.cells.find((cell) => cell.address === "B1")?.formulaText).toBe("=A1+1");
        expect(summarySheet.cells.find((cell) => cell.address === "B2")?.formulaText).toBe("=A2+1");
        expect(summarySheet.cells.find((cell) => cell.address === "B3")?.formulaText).toBe("=A3+1");
        expect(summarySheet.cells.find((cell) => cell.address === "B3")?.outputValue).toBe("31");
        expect(summarySheet.cells.find((cell) => cell.address === "B3")?.resolutionStatus).toBe("resolved");
    });

    it("evaluates SUM AVERAGE MIN MAX over ranges", async () => {
        const api = core;
        const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Summary" sheetId="1" r:id="rId1"/>
        <sheet name="Other" sheetId="2" r:id="rId2"/>
    </sheets>
</workbook>`;
        const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`;
        const sheet1Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1"><v>1</v></c>
            <c r="B1"><v>2</v></c>
            <c r="C1" t="inlineStr"><is><t>text</t></is></c>
        </row>
        <row r="2">
            <c r="A2"><v>3</v></c>
            <c r="B2"><v>4</v></c>
        </row>
        <row r="4">
            <c r="A4"><f>SUM(A1:B2)</f></c>
            <c r="B4"><f>AVERAGE(A1:B2)</f></c>
            <c r="C4"><f>MIN(A1:B2)</f></c>
            <c r="D4"><f>MAX(Other!A1:B2)</f></c>
            <c r="E4"><f>COUNT(A1:C2)</f></c>
            <c r="F4"><f>COUNTA(A1:C2)</f></c>
            <c r="G4"><f>SUM(A1,B1,10)</f></c>
            <c r="H4"><f>SUM($A$1:$B$2)</f></c>
            <c r="I4"><f>VLOOKUP(2,Other!A1:B2,2,FALSE)</f></c>
            <c r="J4"><f>VLOOKUP(3,$A$1:$B$2,2,0)</f></c>
            <c r="K4"><f>COUNTIF(A1:B2,">2")</f></c>
            <c r="L4"><f>SUMIF(A1:B2,">2")</f></c>
            <c r="M4"><f>COUNTIF(A1:C2,"text")</f></c>
            <c r="N4"><f>AVERAGEIF(A1:B2,">2")</f></c>
            <c r="O4"><f>COUNTIFS(A1:B2,"&gt;1",A1:B2,"&lt;4")</f></c>
            <c r="P4"><f>SUMIFS(A1:B2,A1:B2,"&gt;1",A1:B2,"&lt;4")</f></c>
            <c r="Q4"><f>AVERAGEIFS(A1:B2,A1:B2,"&gt;1",A1:B2,"&lt;4")</f></c>
            <c r="R4"><f>HLOOKUP("K2",A6:C7,2,FALSE)</f></c>
            <c r="S4"><f>MATCH("K2",A6:C6,0)</f></c>
            <c r="T4"><f>INDEX(A7:C7,1,2)</f></c>
            <c r="U4"><f>INDEX(A7:C7,1,MATCH("K2",A6:C6,0))</f></c>
            <c r="V4"><f>XLOOKUP("K2",A6:C6,A7:C7)</f></c>
            <c r="W4"><f>XLOOKUP("ZZ",A6:C6,A7:C7,"NF")</f></c>
            <c r="X4"><f>CHOOSE(2,"A","B","C")</f></c>
            <c r="Y4"><f>CHOOSE(MATCH("K2",A6:C6,0),"A","B","C")</f></c>
        </row>
        <row r="6">
            <c r="A6" t="inlineStr"><is><t>K1</t></is></c>
            <c r="B6" t="inlineStr"><is><t>K2</t></is></c>
            <c r="C6" t="inlineStr"><is><t>K3</t></is></c>
        </row>
        <row r="7">
            <c r="A7"><v>11</v></c>
            <c r="B7"><v>22</v></c>
            <c r="C7"><v>33</v></c>
        </row>
    </sheetData>
</worksheet>`;
        const sheet2Xml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1"><v>1</v></c>
            <c r="B1"><v>7</v></c>
        </row>
        <row r="2">
            <c r="A2"><v>2</v></c>
            <c r="B2"><v>8</v></c>
        </row>
    </sheetData>
</worksheet>`;
        const zip = createStoredZip([
            {
                name: "[Content_Types].xml",
                data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`,
            },
            { name: "xl/workbook.xml", data: workbookXml },
            { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
            { name: "xl/worksheets/sheet1.xml", data: sheet1Xml },
            { name: "xl/worksheets/sheet2.xml", data: sheet2Xml },
        ]);

        const workbook = await api.parseWorkbook(zip, "func.xlsx");
        const summarySheet = workbook.sheets[0];

        expect(summarySheet.cells.find((cell) => cell.address === "A4")?.outputValue).toBe("10");
        expect(summarySheet.cells.find((cell) => cell.address === "B4")?.outputValue).toBe("2.5");
        expect(summarySheet.cells.find((cell) => cell.address === "C4")?.outputValue).toBe("1");
        expect(summarySheet.cells.find((cell) => cell.address === "D4")?.outputValue).toBe("8");
        expect(summarySheet.cells.find((cell) => cell.address === "E4")?.outputValue).toBe("4");
        expect(summarySheet.cells.find((cell) => cell.address === "F4")?.outputValue).toBe("5");
        expect(summarySheet.cells.find((cell) => cell.address === "G4")?.outputValue).toBe("13");
        expect(summarySheet.cells.find((cell) => cell.address === "H4")?.outputValue).toBe("10");
        expect(summarySheet.cells.find((cell) => cell.address === "I4")?.outputValue).toBe("8");
        expect(summarySheet.cells.find((cell) => cell.address === "J4")?.outputValue).toBe("4");
        expect(summarySheet.cells.find((cell) => cell.address === "K4")?.outputValue).toBe("2");
        expect(summarySheet.cells.find((cell) => cell.address === "L4")?.outputValue).toBe("7");
        expect(summarySheet.cells.find((cell) => cell.address === "M4")?.outputValue).toBe("1");
        expect(summarySheet.cells.find((cell) => cell.address === "N4")?.outputValue).toBe("3.5");
        expect(summarySheet.cells.find((cell) => cell.address === "O4")?.outputValue).toBe("2");
        expect(summarySheet.cells.find((cell) => cell.address === "P4")?.outputValue).toBe("5");
        expect(summarySheet.cells.find((cell) => cell.address === "Q4")?.outputValue).toBe("2.5");
        expect(summarySheet.cells.find((cell) => cell.address === "R4")?.outputValue).toBe("22");
        expect(summarySheet.cells.find((cell) => cell.address === "S4")?.outputValue).toBe("2");
        expect(summarySheet.cells.find((cell) => cell.address === "T4")?.outputValue).toBe("22");
        expect(summarySheet.cells.find((cell) => cell.address === "U4")?.outputValue).toBe("22");
        expect(summarySheet.cells.find((cell) => cell.address === "V4")?.outputValue).toBe("22");
        expect(summarySheet.cells.find((cell) => cell.address === "W4")?.outputValue).toBe("NF");
        expect(summarySheet.cells.find((cell) => cell.address === "X4")?.outputValue).toBe("B");
        expect(summarySheet.cells.find((cell) => cell.address === "Y4")?.outputValue).toBe("B");
    });

    it("evaluates comparison expressions and IF formulas", async () => {
        const api = core;
        const workbookXml = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Summary" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>`;
        const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
        const sheetXml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1"><v>10</v></c>
            <c r="B1"><v>5</v></c>
            <c r="C1"><f>A1&gt;B1</f></c>
            <c r="D1"><f>A1=B1</f></c>
            <c r="E1"><f>IF(A1&gt;B1,1,0)</f></c>
            <c r="F1"><f>IF(B1&gt;A1,"OK","NG")</f></c>
            <c r="G1"><f>IF((A1-B1)=5,"MATCH","MISS")</f></c>
            <c r="H1"><f>AND(A1&gt;B1,B1&gt;0)</f></c>
            <c r="I1"><f>OR(A1&lt;B1,B1&gt;0)</f></c>
            <c r="J1"><f>NOT(A1=B1)</f></c>
            <c r="K1"><f>IF(AND(A1&gt;B1,B1&gt;0),"OK","NG")</f></c>
            <c r="L1"><f>"ID-"&amp;A1</f></c>
            <c r="M1"><f>A1&amp;" / "&amp;B1</f></c>
            <c r="N1"><f>"Result:"&amp;IF(A1&gt;B1,"OK","NG")</f></c>
            <c r="O1"><f>"ABS-"&amp;$A$1&amp;"-"&amp;$B1</f></c>
            <c r="P1"><f>LEFT("ABCDE",2)</f></c>
            <c r="Q1"><f>RIGHT("ABCDE",3)</f></c>
            <c r="R1"><f>MID("ABCDE",2,2)</f></c>
            <c r="S1"><f>LEN("AB CD")</f></c>
            <c r="T1"><f>TRIM("  A   B  ")</f></c>
            <c r="U1"><f>LEFT("ID-"&amp;A1,4)</f></c>
            <c r="V1"><f>SUBSTITUTE("A-B-B","B","X")</f></c>
            <c r="W1"><f>SUBSTITUTE("A-B-B","B","X",2)</f></c>
            <c r="X1"><f>REPLACE("ABCDE",2,2,"ZZ")</f></c>
            <c r="Y1"><f>TEXT(1234.5,"#,##0.00")</f></c>
            <c r="Z1"><f>TEXT(45292,"yyyy-mm-dd")</f></c>
            <c r="AA1"><f>TEXT(0.5,"hh:mm:ss")</f></c>
            <c r="AB1"><f>YEAR(45292)</f></c>
            <c r="AC1"><f>MONTH(45292)</f></c>
            <c r="AD1"><f>DAY(45292)</f></c>
            <c r="AE1"><f>YEAR(TEXT(45292,"yyyy-mm-dd"))</f></c>
            <c r="AF1"><f>DATE(2024,1,1)</f></c>
            <c r="AG1"><f>TEXT(DATE(2024,1,1),"yyyy-mm-dd")</f></c>
            <c r="AH1"><f>VALUE("1,234.50")</f></c>
            <c r="AI1"><f>VALUE("2024-01-01")</f></c>
            <c r="AJ1"><f>IFERROR(Unknown!A1,"ALT")</f></c>
            <c r="AK1"><f>IFERROR(A1+B1,"ALT")</f></c>
            <c r="AL1"><f>ISBLANK(C2)</f></c>
            <c r="AM1"><f>ISNUMBER(A1)</f></c>
            <c r="AN1"><f>ISTEXT("HELLO")</f></c>
            <c r="AO1"><f>IF(ISBLANK(C2),"EMPTY","FILLED")</f></c>
            <c r="AP1"><f>ISERROR(Unknown!A1)</f></c>
            <c r="AQ1"><f>ISNA(VLOOKUP(99,A1:B1,2,FALSE))</f></c>
        </row>
    </sheetData>
</worksheet>`;
        const zip = createStoredZip([
            {
                name: "[Content_Types].xml",
                data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`,
            },
            { name: "xl/workbook.xml", data: workbookXml },
            { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
            { name: "xl/worksheets/sheet1.xml", data: sheetXml },
        ]);

        const workbook = await api.parseWorkbook(zip, "if.xlsx");
        const summarySheet = workbook.sheets[0];

        expect(summarySheet.cells.find((cell) => cell.address === "C1")?.outputValue).toBe("TRUE");
        expect(summarySheet.cells.find((cell) => cell.address === "D1")?.outputValue).toBe("FALSE");
        expect(summarySheet.cells.find((cell) => cell.address === "E1")?.outputValue).toBe("1");
        expect(summarySheet.cells.find((cell) => cell.address === "F1")?.outputValue).toBe("NG");
        expect(summarySheet.cells.find((cell) => cell.address === "G1")?.outputValue).toBe("MATCH");
        expect(summarySheet.cells.find((cell) => cell.address === "H1")?.outputValue).toBe("TRUE");
        expect(summarySheet.cells.find((cell) => cell.address === "I1")?.outputValue).toBe("TRUE");
        expect(summarySheet.cells.find((cell) => cell.address === "J1")?.outputValue).toBe("TRUE");
        expect(summarySheet.cells.find((cell) => cell.address === "K1")?.outputValue).toBe("OK");
        expect(summarySheet.cells.find((cell) => cell.address === "L1")?.outputValue).toBe("ID-10");
        expect(summarySheet.cells.find((cell) => cell.address === "M1")?.outputValue).toBe("10 / 5");
        expect(summarySheet.cells.find((cell) => cell.address === "N1")?.outputValue).toBe("Result:OK");
        expect(summarySheet.cells.find((cell) => cell.address === "O1")?.outputValue).toBe("ABS-10-5");
        expect(summarySheet.cells.find((cell) => cell.address === "P1")?.outputValue).toBe("AB");
        expect(summarySheet.cells.find((cell) => cell.address === "Q1")?.outputValue).toBe("CDE");
        expect(summarySheet.cells.find((cell) => cell.address === "R1")?.outputValue).toBe("BC");
        expect(summarySheet.cells.find((cell) => cell.address === "S1")?.outputValue).toBe("5");
        expect(summarySheet.cells.find((cell) => cell.address === "T1")?.outputValue).toBe("A B");
        expect(summarySheet.cells.find((cell) => cell.address === "U1")?.outputValue).toBe("ID-1");
        expect(summarySheet.cells.find((cell) => cell.address === "V1")?.outputValue).toBe("A-X-X");
        expect(summarySheet.cells.find((cell) => cell.address === "W1")?.outputValue).toBe("A-B-X");
        expect(summarySheet.cells.find((cell) => cell.address === "X1")?.outputValue).toBe("AZZDE");
        expect(summarySheet.cells.find((cell) => cell.address === "Y1")?.outputValue).toBe("1,234.50");
        expect(summarySheet.cells.find((cell) => cell.address === "Z1")?.outputValue).toBe("2024-01-01");
        expect(summarySheet.cells.find((cell) => cell.address === "AA1")?.outputValue).toBe("12:00:00");
        expect(summarySheet.cells.find((cell) => cell.address === "AB1")?.outputValue).toBe("2024");
        expect(summarySheet.cells.find((cell) => cell.address === "AC1")?.outputValue).toBe("1");
        expect(summarySheet.cells.find((cell) => cell.address === "AD1")?.outputValue).toBe("1");
        expect(summarySheet.cells.find((cell) => cell.address === "AE1")?.outputValue).toBe("2024");
        expect(summarySheet.cells.find((cell) => cell.address === "AF1")?.outputValue).toBe("45292");
        expect(summarySheet.cells.find((cell) => cell.address === "AG1")?.outputValue).toBe("2024-01-01");
        expect(summarySheet.cells.find((cell) => cell.address === "AH1")?.outputValue).toBe("1234.5");
        expect(summarySheet.cells.find((cell) => cell.address === "AI1")?.outputValue).toBe("45292");
        expect(summarySheet.cells.find((cell) => cell.address === "AJ1")?.outputValue).toBe("ALT");
        expect(summarySheet.cells.find((cell) => cell.address === "AK1")?.outputValue).toBe("15");
        expect(summarySheet.cells.find((cell) => cell.address === "AL1")?.outputValue).toBe("TRUE");
        expect(summarySheet.cells.find((cell) => cell.address === "AM1")?.outputValue).toBe("TRUE");
        expect(summarySheet.cells.find((cell) => cell.address === "AN1")?.outputValue).toBe("TRUE");
        expect(summarySheet.cells.find((cell) => cell.address === "AO1")?.outputValue).toBe("EMPTY");
        expect(summarySheet.cells.find((cell) => cell.address === "AP1")?.outputValue).toBe("TRUE");
        expect(summarySheet.cells.find((cell) => cell.address === "AQ1")?.outputValue).toBe("TRUE");
    });

    it("includes formula resolution source in diagnostics", () => {
        const api = core;
        const workbook = { name: "source.xlsx" };
        const sheet = {
            name: "Formula Source",
            index: 1,
            path: "xl/worksheets/sheet1.xml",
            merges: [],
            tables: [],
            images: [],
            maxRow: 2,
            maxCol: 2,
            cells: [
                {
                    row: 1,
                    col: 1,
                    address: "A1",
                    valueType: "str",
                    rawValue: "label",
                    outputValue: "label",
                    formulaText: "",
                    resolutionStatus: null,
                    resolutionSource: null,
                    cachedValueState: null,
                    styleIndex: 0,
                    borders: { top: true, bottom: true, left: true, right: true },
                    numFmtId: 0,
                    formatCode: "General",
                    spillRef: "",
                },
                {
                    row: 1,
                    col: 2,
                    address: "B1",
                    valueType: "formula",
                    rawValue: "2",
                    outputValue: "2",
                    formulaText: "=1+1",
                    resolutionStatus: "resolved",
                    resolutionSource: "ast_evaluator",
                    cachedValueState: "absent",
                    styleIndex: 0,
                    borders: { top: true, bottom: true, left: true, right: true },
                    numFmtId: 0,
                    formatCode: "General",
                    spillRef: "",
                },
                {
                    row: 2,
                    col: 1,
                    address: "A2",
                    valueType: "formula",
                    rawValue: "=UNKNOWN(A1)",
                    outputValue: "=UNKNOWN(A1)",
                    formulaText: "=UNKNOWN(A1)",
                    resolutionStatus: "fallback_formula",
                    resolutionSource: "formula_text",
                    cachedValueState: "absent",
                    styleIndex: 0,
                    borders: { top: true, bottom: true, left: true, right: true },
                    numFmtId: 0,
                    formatCode: "General",
                    spillRef: "",
                },
                {
                    row: 2,
                    col: 2,
                    address: "B2",
                    valueType: "formula",
                    rawValue: "#REF!",
                    outputValue: "=[other.xlsx]Sheet1!A1",
                    formulaText: "=[other.xlsx]Sheet1!A1",
                    resolutionStatus: "unsupported_external",
                    resolutionSource: "external_unsupported",
                    cachedValueState: "absent",
                    styleIndex: 0,
                    borders: { top: true, bottom: true, left: true, right: true },
                    numFmtId: 0,
                    formatCode: "General",
                    spillRef: "",
                },
            ],
        };

        const markdownFile = api.convertSheetToMarkdown(workbook, sheet, {
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            treatFirstRowAsHeader: true,
        });

        expect(markdownFile.summary.formulaDiagnostics).toEqual([
            { address: "B1", formulaText: "=1+1", status: "resolved", source: "ast_evaluator", outputValue: "2" },
            {
                address: "A2",
                formulaText: "=UNKNOWN(A1)",
                status: "fallback_formula",
                source: "formula_text",
                outputValue: "=UNKNOWN(A1)",
            },
            {
                address: "B2",
                formulaText: "=[other.xlsx]Sheet1!A1",
                status: "unsupported_external",
                source: "external_unsupported",
                outputValue: "=[other.xlsx]Sheet1!A1",
            },
        ]);
    });

    it("treats empty-string cached formula results as cached", async () => {
        const api = core;
        const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>`;
        const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
        const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1"><f>IF(1=1,"","X")</f><v></v></c>
        </row>
    </sheetData>
</worksheet>`;

        const zip = createStoredZip([
            {
                name: "[Content_Types].xml",
                data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
            },
            { name: "xl/workbook.xml", data: workbookXml },
            { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
            { name: "xl/worksheets/sheet1.xml", data: sheetXml },
        ]);

        const workbook = await api.parseWorkbook(zip, "empty-cache.xlsx");
        const cell = workbook.sheets[0].cells.find((entry) => entry.address === "A1");
        const markdownFile = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        })[0];

        expect(cell?.rawValue).toBe("");
        expect(cell?.outputValue).toBe("");
        expect(cell?.resolutionStatus).toBe("resolved");
        expect(cell?.resolutionSource).toBe("cached_value");
        expect(cell?.cachedValueState).toBe("present_empty");
        expect(markdownFile.summary.formulaDiagnostics).toEqual([
            {
                address: "A1",
                formulaText: '=IF(1=1,"","X")',
                status: "resolved",
                source: "cached_value",
                outputValue: "",
            },
        ]);
    });
});
