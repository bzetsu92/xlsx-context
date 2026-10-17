import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { createSummaryText } from "../../../src/outputs/markdown/markdown-export.ts";
import { createStoredZip } from "./stored-zip.js";
import { buildWorkbookIR } from "../../../src/outputs/ai-context/workbook-ir.ts";
import { renderCompactMarkdown } from "../../../src/outputs/ai-context/compact-renderer.ts";

describe("core convert", () => {
    it("parses a minimal workbook and converts a sheet to markdown", async () => {
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
        const sharedStringsXml = `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="5" uniqueCount="5">
    <si><t>Region</t></si>
    <si><t>Sales</t></si>
    <si><t>East</t></si>
    <si><t>West</t></si>
    <si><t>Overview text</t></si>
</sst>`;
        const stylesXml = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <borders count="3">
        <border><left/><right/><top/><bottom/></border>
        <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border>
        <border><left/><right/><top/><bottom/></border>
    </borders>
    <cellXfs count="3">
        <xf borderId="0"/>
        <xf borderId="1"/>
        <xf borderId="2" numFmtId="14" applyNumberFormat="1"/>
    </cellXfs>
</styleSheet>`;
        const sheetXml = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="s"><v>4</v></c>
        </row>
        <row r="3">
            <c r="A3" t="s" s="1"><v>0</v></c>
            <c r="B3" t="s" s="1"><v>1</v></c>
            <c r="C3" s="1"><v>100</v></c>
        </row>
        <row r="4">
            <c r="A4" t="s" s="1"><v>2</v></c>
            <c r="B4" t="s" s="1"><v>3</v></c>
            <c r="C4" s="1"><v>90</v></c>
        </row>
        <row r="6">
            <c r="A6" t="inlineStr"><is><t>Date</t></is></c>
            <c r="B6" s="2"><v>45292</v></c>
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
            { name: "xl/sharedStrings.xml", data: sharedStringsXml },
            { name: "xl/styles.xml", data: stylesXml },
            { name: "xl/worksheets/sheet1.xml", data: sheetXml },
        ]);

        const workbook = await api.parseWorkbook(zip, "sales.xlsx");
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });
        const rawFiles = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            outputMode: "raw",
        });
        const bothFiles = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            outputMode: "both",
        });

        expect(workbook.sheets).toHaveLength(1);
        expect(workbook.sheets[0].name).toBe("Summary");
        expect(files).toHaveLength(1);
        expect(files[0].fileName).toBe("sales_001_Summary.md");
        expect(rawFiles[0].fileName).toBe("sales_001_Summary.md");
        expect(bothFiles[0].fileName).toBe("sales_001_Summary.md");
        expect(files[0].markdown).toContain("## Sheet: Summary");
        expect(files[0].markdown).toContain("# Book: sales.xlsx");
        expect(files[0].markdown).toContain("Overview text");
        expect(files[0].markdown).toContain("| Region | Sales | 100 |");
        expect(files[0].markdown).toContain("| East | West | 90 |");
        // Built-in date 14 with no locale in the file renders as ISO.
        expect(files[0].markdown).toContain("2024-01-01");
        expect(rawFiles[0].markdown).toContain("Date 45292");
        expect(bothFiles[0].markdown).toContain("Date 2024-01-01 [raw=45292]");
        expect(createSummaryText(bothFiles[0])).toContain("Output mode: both");
        expect(files[0].summary.images).toBe(0);
        expect(files[0].summary.tableScores).toHaveLength(1);
        expect(files[0].summary.tableScores[0].range).toBe("A3-C4");
        expect(files[0].summary.tableScores[0].score).toBeGreaterThanOrEqual(4);
        expect(files[0].summary.tableScores[0].reasons.join(" ")).toContain("Has borders");
    });

    it("renders accounting zero format as dash instead of currency zero", () => {
        const api = core;
        const workbook = { name: "accounting.xlsx" };
        const sheet = {
            name: "Summary",
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
                    rawValue: "費用",
                    outputValue: "費用",
                    formulaText: "",
                    resolutionStatus: null,
                    styleIndex: 0,
                    borders: { top: true, bottom: true, left: true, right: true },
                    numFmtId: 0,
                    formatCode: "General",
                },
                {
                    row: 1,
                    col: 2,
                    address: "B1",
                    valueType: "str",
                    rawValue: "合計費用",
                    outputValue: "合計費用",
                    formulaText: "",
                    resolutionStatus: null,
                    styleIndex: 0,
                    borders: { top: true, bottom: true, left: true, right: true },
                    numFmtId: 0,
                    formatCode: "General",
                },
                {
                    row: 2,
                    col: 1,
                    address: "A2",
                    valueType: "n",
                    rawValue: "0",
                    outputValue: "¥ -",
                    formulaText: "",
                    resolutionStatus: null,
                    styleIndex: 0,
                    borders: { top: true, bottom: true, left: true, right: true },
                    numFmtId: 44,
                    formatCode: '_ "¥"* #,##0.00_ ;_ "¥"* \\-#,##0.00_ ;_ "¥"* "-"??_ ;_ @_ ',
                },
                {
                    row: 2,
                    col: 2,
                    address: "B2",
                    valueType: "n",
                    rawValue: "0",
                    outputValue: "¥ -",
                    formulaText: "",
                    resolutionStatus: null,
                    styleIndex: 0,
                    borders: { top: true, bottom: true, left: true, right: true },
                    numFmtId: 44,
                    formatCode: '_ "¥"* #,##0.00_ ;_ "¥"* \\-#,##0.00_ ;_ "¥"* "-"??_ ;_ @_ ',
                },
            ],
        };

        const markdownFile = api.convertSheetToMarkdown(workbook, sheet, {
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            treatFirstRowAsHeader: true,
        });

        expect(markdownFile.markdown).toContain("| ¥ - | ¥ - |");
    });
});

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

/** Builds a workbook package. Each sheet is { name, state?, xml, rels?, parts? }. */
function buildWorkbook({ sheets, styles = "", workbookPr = "" }) {
    const sheetEntries = sheets.map(
        (sheet, index) =>
            `<sheet name="${sheet.name}" sheetId="${index + 1}" r:id="rId${index + 1}"${sheet.state ? ` state="${sheet.state}"` : ""}/>`
    );
    const relEntries = sheets.map(
        (_, index) =>
            `<Relationship Id="rId${index + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
    );
    return createStoredZip([
        {
            name: "[Content_Types].xml",
            data: `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
        },
        {
            name: "xl/workbook.xml",
            data: `<workbook xmlns="${MAIN}" xmlns:r="${REL}">${workbookPr}<sheets>${sheetEntries.join("")}</sheets></workbook>`,
        },
        {
            name: "xl/_rels/workbook.xml.rels",
            data: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relEntries.join("")}</Relationships>`,
        },
        ...(styles ? [{ name: "xl/styles.xml", data: `<styleSheet xmlns="${MAIN}">${styles}</styleSheet>` }] : []),
        ...sheets.flatMap((sheet, index) => [
            {
                name: `xl/worksheets/sheet${index + 1}.xml`,
                data: `<worksheet xmlns="${MAIN}" xmlns:r="${REL}">${sheet.xml}</worksheet>`,
            },
            ...(sheet.rels
                ? [
                      {
                          name: `xl/worksheets/_rels/sheet${index + 1}.xml.rels`,
                          data: `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheet.rels}</Relationships>`,
                      },
                  ]
                : []),
            ...(sheet.parts || []),
        ]),
    ]);
}

const inline = (ref, text) => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;

describe("workbook settings that the file declares", () => {
    it("reads dates, currency, and names by the locale and date system of the file", async () => {
        const zip = buildWorkbook({
            workbookPr: `<workbookPr date1904="1"/>`,
            styles: `<numFmts count="1"><numFmt numFmtId="164" formatCode="[$-42A]dddd d mmmm yyyy"/></numFmts>
                <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
                <borders count="1"><border/></borders>
                <cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="44"/><xf numFmtId="164"/></cellXfs>`,
            sheets: [
                {
                    name: "Sales",
                    xml: `<sheetData><row r="1"><c r="A1" s="1"><v>0</v></c></row>
                        <row r="2"><c r="A2" s="2"><v>1234.5</v></c></row>
                        <row r="3"><c r="A3" s="3"><v>43904</v></c></row></sheetData>`,
                },
            ],
        });

        const workbook = await core.parseWorkbook(zip, "vi.xlsx");
        const value = (address) => workbook.sheets[0].cells.find((cell) => cell.address === address)?.outputValue;

        expect(workbook.dateSystem).toBe("1904");
        expect(workbook.locale).toBe("vi-VN");
        // Serial 0 is 1904-01-01 in the 1904 date system. Vietnamese short dates put the day first.
        expect(value("A1")).toBe("1/1/1904");
        expect(value("A2")).toBe("₫ 1,234.50");
        expect(value("A3")).toBe("Thứ Sáu 15 Tháng 3 2024");
    });

    it("marks hidden sheets, rows, and columns, and drops them on request", async () => {
        const zip = buildWorkbook({
            sheets: [
                {
                    name: "Visible",
                    xml: `<cols><col min="3" max="3" hidden="1"/></cols>
                        <sheetData><row r="1">${inline("A1", "keep")}${inline("C1", "secret column")}</row>
                        <row r="2" hidden="1">${inline("A2", "secret row")}</row></sheetData>`,
                },
                {
                    name: "Lookup",
                    state: "hidden",
                    xml: `<sheetData><row r="1">${inline("A1", "hidden sheet text")}</row></sheetData>`,
                },
            ],
        });

        const workbook = await core.parseWorkbook(zip, "hidden.xlsx");
        const files = core.convertWorkbookToMarkdownFiles(workbook);
        const ir = buildWorkbookIR(workbook);

        expect(workbook.sheets.map((sheet) => [sheet.name, sheet.state])).toEqual([
            ["Visible", "visible"],
            ["Lookup", "hidden"],
        ]);
        expect(files[0].markdown).toContain("> Hidden rows: 2. Hidden columns: C.");
        expect(files[0].markdown).toContain("secret row");
        expect(files[1].markdown).toContain("> Hidden sheet.");
        expect(ir.sheets[0]).toMatchObject({ state: "visible", hiddenRows: [2], hiddenCols: [3] });
        expect(renderCompactMarkdown(ir)).toContain("# Lookup\n\n> Hidden sheet.");

        const visibleOnly = await core.parseWorkbook(zip, "hidden.xlsx", { includeHidden: false });
        const text = core
            .convertWorkbookToMarkdownFiles(visibleOnly)
            .map((file) => file.markdown)
            .join("\n");
        expect(visibleOnly.sheets.map((sheet) => sheet.name)).toEqual(["Visible"]);
        expect(text).toContain("keep");
        expect(text).not.toContain("secret");
    });

    it("uses declared Excel tables with their names and header setting", async () => {
        const zip = buildWorkbook({
            sheets: [
                {
                    name: "Stock",
                    xml: `<sheetData>
                        <row r="1">${inline("A1", "Pen")}<c r="B1"><v>3</v></c></row>
                        <row r="2">${inline("A2", "Ink")}<c r="B2"><v>5</v></c></row>
                        <row r="3">${inline("A3", "Pad")}<c r="B3"><v>7</v></c></row></sheetData>
                        <tableParts count="1"><tablePart r:id="rId1"/></tableParts>`,
                    rels: `<Relationship Id="rId1" Type="${REL}/table" Target="../tables/table1.xml"/>`,
                    parts: [
                        {
                            name: "xl/tables/table1.xml",
                            data: `<table xmlns="${MAIN}" id="1" name="Stock" displayName="StockList" ref="A1:B3" headerRowCount="0">
                                <tableColumns count="2"><tableColumn id="1" name="Column1"/><tableColumn id="2" name="Column2"/></tableColumns></table>`,
                        },
                    ],
                },
            ],
        });

        const workbook = await core.parseWorkbook(zip, "stock.xlsx");
        const markdown = core.convertWorkbookToMarkdownFiles(workbook)[0].markdown;
        const ir = buildWorkbookIR(workbook);

        expect(workbook.sheets[0].tables[0]).toMatchObject({ displayName: "StockList", headerRowCount: 0 });
        expect(markdown).toContain("### Table: 001 (A1-B3)\n\n- Excel table: StockList");
        expect(markdown).toContain("|  |  |\n| --- | --- |\n| Pen | 3 |");
        expect(ir.sheets[0].tables[0]).toMatchObject({ name: "StockList", range: "A1-B3", hasHeader: false });
        expect(renderCompactMarkdown(ir)).toContain("### A1-B3 StockList");
    });
});
