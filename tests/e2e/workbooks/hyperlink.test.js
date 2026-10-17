import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { createStoredZip } from "./stored-zip.js";
import { fixture } from "../../helpers/paths.js";

describe("hyperlink workbooks", () => {
    it("renders external and workbook hyperlinks into markdown", async () => {
        const api = core;
        const zip = createStoredZip([
            {
                name: "[Content_Types].xml",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
            },
            {
                name: "_rels/.rels",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
            },
            {
                name: "xl/workbook.xml",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheets>
        <sheet name="Summary" sheetId="1" r:id="rId1"/>
        <sheet name="Other" sheetId="2" r:id="rId2"/>
    </sheets>
</workbook>`,
            },
            {
                name: "xl/_rels/workbook.xml.rels",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`,
            },
            {
                name: "xl/worksheets/sheet1.xml",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr"><is><t>Open</t></is></c>
        </row>
        <row r="2">
            <c r="A2" t="inlineStr"><is><t>Jump</t></is></c>
        </row>
    </sheetData>
    <hyperlinks>
        <hyperlink ref="A1" r:id="rId1"/>
        <hyperlink ref="A2" location="Other!A1"/>
    </hyperlinks>
</worksheet>`,
            },
            {
                name: "xl/worksheets/_rels/sheet1.xml.rels",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
</Relationships>`,
            },
            {
                name: "xl/worksheets/sheet2.xml",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="inlineStr"><is><t>Target</t></is></c>
        </row>
    </sheetData>
</worksheet>`,
            },
        ]);

        const workbook = await api.parseWorkbook(zip, "link-book.xlsx");
        const markdownFiles = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            formattingMode: "gfm",
        });

        expect(markdownFiles[0].markdown).toContain("[Open](https://example.com/)");
        expect(markdownFiles[0].markdown).toContain("[Jump](#sheet-other) (Other!A1)");
        expect(markdownFiles[0].markdown).not.toContain("<ins>Open</ins>");
        expect(markdownFiles[0].markdown).not.toContain("<ins>Jump</ins>");
        expect(markdownFiles[1].markdown).toContain("## Sheet: Other");
    });

    it("parses the hyperlink-basic fixture workbook with external and internal links", async () => {
        const fixtureName = "hyperlink-basic-sample01.xlsx";
        const fixtureBytes = readFileSync(fixture("hyperlink", fixtureName));
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await core.parseWorkbook(arrayBuffer, fixtureName);
        const files = core.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            formattingMode: "gfm",
        });
        const summarySheet = workbook.sheets[0];
        const summaryFile = files[0];
        const otherFile = files[1];
        const hyperlinkOf = (address) => summarySheet.cells.find((cell) => cell.address === address)?.hyperlink;

        expect(workbook.sheets.map((sheet) => sheet.name)).toEqual(["Summary", "Other"]);
        expect(summarySheet.cells).toHaveLength(8);
        expect(hyperlinkOf("A1")).toMatchObject({ kind: "external", target: "https://example.com/" });
        expect(hyperlinkOf("A2")).toMatchObject({ kind: "internal", target: "Other!A1", location: "Other!A1" });
        expect(hyperlinkOf("B5")).toMatchObject({ kind: "external", target: "https://example.com/docs" });
        expect(hyperlinkOf("B6")).toMatchObject({ kind: "internal", target: "Other!A1", location: "Other!A1" });
        expect(hyperlinkOf("A5")).toBeNull();

        expect(summaryFile.fileName).toBe("hyperlink-basic-sample01_001_Summary.md");
        expect(summaryFile.summary.tables).toBe(1);
        expect(summaryFile.summary.tableScores.map((detail) => detail.range)).toEqual(["A4-B6"]);
        expect(summaryFile.markdown).toContain("[Open example](https://example.com/)");
        expect(summaryFile.markdown).toContain("[Jump to Other](#sheet-other) (Other!A1)");
        expect(summaryFile.markdown).toContain("### Table: 001 (A4-B6)");
        expect(summaryFile.markdown).toContain("| External | [Open example](https://example.com/docs) |");
        expect(summaryFile.markdown).toContain("| Internal | [Jump to Other](#sheet-other) (Other!A1) |");
        // Excel underlines link cells. The link syntax already shows that, so gfm mode adds no <ins>.
        expect(summaryFile.markdown).not.toContain("<ins>");

        expect(otherFile.fileName).toBe("hyperlink-basic-sample01_002_Other.md");
        expect(otherFile.markdown).toContain("## Sheet: Other");
        expect(otherFile.markdown).toContain("Target cell");
    });
});
