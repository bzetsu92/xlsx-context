import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { unzipEntries } from "../../../src/extraction/zip/zip-io.ts";
import { buildWorkbookIR } from "../../../src/outputs/ai-context/workbook-ir.ts";
import { renderCompactMarkdown } from "../../../src/outputs/ai-context/compact-renderer.ts";
import { createStoredZip } from "./stored-zip.js";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("chart workbooks", () => {
    const fixtureDir = FIXTURE_ROOT;

    it("extracts chart metadata blocks from drawing charts", async () => {
        const api = core;
        const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets>
                    <sheet name="ChartSheet" sheetId="1" r:id="rId1"/>
                </sheets>
            </workbook>`;
        const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
            </Relationships>`;
        const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheetData>
                    <row r="1"><c r="A1" t="inlineStr"><is><t>Chart Sample</t></is></c></row>
                </sheetData>
                <drawing r:id="rId1"/>
            </worksheet>`;
        const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
            </Relationships>`;
        const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <xdr:twoCellAnchor>
                    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>17</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
                    <xdr:to><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>34</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
                    <xdr:graphicFrame>
                        <xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
                        <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                                <c:chart r:id="rId1"/>
                            </a:graphicData>
                        </a:graphic>
                    </xdr:graphicFrame>
                    <xdr:clientData/>
                </xdr:twoCellAnchor>
            </xdr:wsDr>`;
        const drawingRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
            </Relationships>`;
        const chartXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <c:chart>
                    <c:title><c:tx><c:rich><a:p><a:r><a:t>Sample Sales Chart</a:t></a:r></a:p></c:rich></c:tx></c:title>
                    <c:plotArea>
                        <c:barChart>
                            <c:ser>
                                <c:tx><c:v>Series A</c:v></c:tx>
                                <c:cat><c:strRef><c:f>ChartSheet!$B$3:$B$6</c:f></c:strRef></c:cat>
                                <c:val><c:numRef><c:f>ChartSheet!$E$3:$E$6</c:f></c:numRef></c:val>
                            </c:ser>
                            <c:ser>
                                <c:tx><c:v>Series B</c:v></c:tx>
                                <c:cat><c:strRef><c:f>ChartSheet!$B$3:$B$6</c:f></c:strRef></c:cat>
                                <c:val><c:numRef><c:f>ChartSheet!$D$3:$D$6</c:f></c:numRef></c:val>
                            </c:ser>
                        </c:barChart>
                    </c:plotArea>
                </c:chart>
            </c:chartSpace>`;
        const arrayBuffer = createStoredZip([
            {
                name: "[Content_Types].xml",
                data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                    <Default Extension="xml" ContentType="application/xml"/>
                </Types>`,
            },
            {
                name: "_rels/.rels",
                data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>`,
            },
            { name: "xl/workbook.xml", data: workbookXml },
            { name: "xl/_rels/workbook.xml.rels", data: workbookRels },
            { name: "xl/worksheets/sheet1.xml", data: sheetXml },
            { name: "xl/worksheets/_rels/sheet1.xml.rels", data: sheetRels },
            { name: "xl/drawings/drawing1.xml", data: drawingXml },
            { name: "xl/drawings/_rels/drawing1.xml.rels", data: drawingRels },
            { name: "xl/charts/chart1.xml", data: chartXml },
        ]);

        const workbook = await api.parseWorkbook(arrayBuffer, "chart-sample.xlsx");
        const markdownFile = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        })[0];

        expect(workbook.sheets[0].charts).toHaveLength(1);
        expect(workbook.sheets[0].charts[0].anchor).toBe("B18");
        expect(workbook.sheets[0].charts[0].title).toBe("Sample Sales Chart");
        expect(workbook.sheets[0].charts[0].chartType).toContain("Bar Chart");
        expect(workbook.sheets[0].charts[0].series).toHaveLength(2);
        expect(workbook.sheets[0].charts[0].series[0].axis).toBe("primary");
        expect(markdownFile.summary.charts).toBe(1);
        expect(markdownFile.markdown).toContain("### Chart: 001 (B18)");
        expect(markdownFile.markdown).toContain("- Title: Sample Sales Chart");
        expect(markdownFile.markdown).toContain("- Type: Bar Chart");
        expect(markdownFile.markdown).toContain("  - Series A");
        expect(markdownFile.markdown).toContain("    - categories: ChartSheet!$B$3:$B$6");
        expect(markdownFile.markdown).toContain("    - values: ChartSheet!$E$3:$E$6");
    });

    it("reads chart titles, 3D types, and stacked layouts from edited chart XML", async () => {
        const api = core;
        const fixtureBytes = readFileSync(path.resolve(fixtureDir, "chart", "chart-basic-sample01.xlsx"));
        const entries = await unzipEntries(
            fixtureBytes.buffer.slice(fixtureBytes.byteOffset, fixtureBytes.byteOffset + fixtureBytes.byteLength)
        );
        const decoder = new TextDecoder();
        const original = decoder.decode(entries.get("xl/charts/chart1.xml"));
        const convert = async (edit) => {
            const zipEntries = Array.from(entries, ([name, data]) => ({
                name,
                data: name === "xl/charts/chart1.xml" ? edit(original) : data,
            }));
            const workbook = await api.parseWorkbook(createStoredZip(zipEntries), "edited.xlsx");
            return {
                chart: workbook.sheets[0].charts[0],
                markdown: api.convertWorkbookToMarkdownFiles(workbook)[0].markdown,
            };
        };
        const axisTitle =
            "<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>売上（円）</a:t></a:r></a:p></c:rich></c:tx></c:title>";

        const withAxisTitle = await convert((xml) => xml.replace("</c:valAx>", `${axisTitle}</c:valAx>`));
        // The axis title is not part of the chart title.
        expect(withAxisTitle.chart.title).toBe("棒グラフのグラフ");
        expect(withAxisTitle.markdown).toContain("- Title: 棒グラフのグラフ\n");

        const threeD = await convert((xml) => xml.replaceAll("c:barChart>", "c:bar3DChart>"));
        expect(threeD.chart.chartType).toBe("3D Bar Chart");
        expect(threeD.chart.series.map((series) => [series.name, series.valuesRef, series.cachedValues])).toEqual([
            ["値A", "'chart-basic'!$C$4:$C$7", [13568, 19287, 25051, 28053]],
            ["値B", "'chart-basic'!$D$4:$D$7", [9072, 18058, 32012, 32012]],
        ]);
        expect(threeD.markdown).toContain("- Source: Table 001 (B3-D7)");

        const stacked = await convert((xml) =>
            xml
                .replace('<c:barDir val="col"/>', '<c:barDir val="bar"/>')
                .replace('<c:grouping val="clustered"/>', '<c:grouping val="percentStacked"/>')
        );
        expect(stacked.chart.chartType).toBe("Bar Chart (horizontal, 100% stacked)");

        const linkedTitle = await convert((xml) =>
            xml.replace(
                /<c:title>[\s\S]*?<\/c:title>/,
                '<c:title><c:tx><c:strRef><c:f>\'chart-basic\'!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>グラフ基本サンプル</c:v></c:pt></c:strCache></c:strRef></c:tx></c:title>'
            )
        );
        expect(linkedTitle.chart.title).toBe("グラフ基本サンプル");

        const twoLineTitle = await convert((xml) =>
            xml.replace(
                "<a:t>棒グラフのグラフ</a:t></a:r></a:p>",
                "<a:t>棒グラフ</a:t></a:r></a:p><a:p><a:r><a:t>2024-2027</a:t></a:r></a:p>"
            )
        );
        expect(twoLineTitle.chart.title).toBe("棒グラフ 2024-2027");

        const noTitle = await convert((xml) => xml.replace(/<c:title>[\s\S]*?<\/c:title>/, ""));
        expect(noTitle.chart.title).toBe("");
        expect(noTitle.markdown).toContain("- Title: (none)");
    });

    it("renders chart caches and connector diagrams from synthetic drawings", async () => {
        const api = core;
        const ns =
            'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
        const twoCell = (fromCol, fromRow, toCol, toRow, body) => `
                <xdr:twoCellAnchor>
                    <xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
                    <xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
                    ${body}
                    <xdr:clientData/>
                </xdr:twoCellAnchor>`;
        const chartFrame = (id, relId) => `
                    <xdr:graphicFrame>
                        <xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
                        <xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
                        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="${relId}"/></a:graphicData></a:graphic>
                    </xdr:graphicFrame>`;
        const shape = (id, prst, text) => `
                    <xdr:sp>
                        <xdr:nvSpPr><xdr:cNvPr id="${id}" name="Shape ${id}"/><xdr:cNvSpPr/></xdr:nvSpPr>
                        <xdr:spPr><a:prstGeom prst="${prst}"/></xdr:spPr>
                        <xdr:txBody><a:p><a:r><a:t>${text}</a:t></a:r></a:p></xdr:txBody>
                    </xdr:sp>`;
        const connector = (id, links, line) => `
                    <xdr:cxnSp>
                        <xdr:nvCxnSpPr><xdr:cNvPr id="${id}" name="Connector ${id}"/><xdr:cNvCxnSpPr>${links}</xdr:cNvCxnSpPr></xdr:nvCxnSpPr>
                        <xdr:spPr><a:prstGeom prst="straightConnector1"/><a:ln>${line}</a:ln></xdr:spPr>
                    </xdr:cxnSp>`;
        const drawingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <xdr:wsDr ${ns}>
        ${twoCell(1, 2, 6, 12, chartFrame(2, "rId1"))}
        ${twoCell(1, 19, 6, 28, chartFrame(3, "rId2"))}
        ${twoCell(1, 39, 6, 48, chartFrame(4, "rId3"))}
        ${twoCell(10, 2, 12, 4, shape(10, "rect", "A"))}
        ${twoCell(10, 7, 12, 9, shape(11, "ellipse", "B"))}
        ${twoCell(11, 4, 11, 7, connector(12, '<a:stCxn id="11" idx="0"/><a:endCxn id="10" idx="2"/>', '<a:headEnd type="triangle"/>'))}
        ${twoCell(12, 4, 13, 6, connector(13, '<a:stCxn id="10" idx="3"/>', '<a:tailEnd type="triangle"/>'))}
            </xdr:wsDr>`;
        const chartSpace = (plot) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                <c:chart><c:plotArea>${plot}</c:plotArea></c:chart>
            </c:chartSpace>`;
        // Values live on another sheet. One cache point is missing and one idx lies past ptCount.
        // The second series has no categories and a literal value list.
        const barChartXml = chartSpace(`
                    <c:barChart>
                        <c:ser>
                            <c:tx><c:v>Sales</c:v></c:tx>
                            <c:cat><c:strRef><c:f>'Other Sheet'!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/>
                                <c:pt idx="0"><c:v>Q1</c:v></c:pt><c:pt idx="1"><c:v>Q2</c:v></c:pt><c:pt idx="2"><c:v>Q3</c:v></c:pt>
                            </c:strCache></c:strRef></c:cat>
                            <c:val><c:numRef><c:f>'Other Sheet'!$B$2:$B$4</c:f><c:numCache><c:ptCount val="3"/>
                                <c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="2"><c:v>30</c:v></c:pt><c:pt idx="999999"><c:v>99</c:v></c:pt>
                            </c:numCache></c:numRef></c:val>
                        </c:ser>
                        <c:ser>
                            <c:tx><c:strRef><c:f>'Other Sheet'!$C$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Cost</c:v></c:pt></c:strCache></c:strRef></c:tx>
                            <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>5</c:v></c:pt><c:pt idx="1"><c:v>n/a</c:v></c:pt></c:numLit></c:val>
                        </c:ser>
                    </c:barChart>`);
        const scatterChartXml = chartSpace(`
                    <c:scatterChart>
                        <c:ser>
                            <c:tx><c:v>XY</c:v></c:tx>
                            <c:xVal><c:numRef><c:f>Report!$A$30:$A$31</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:xVal>
                            <c:yVal><c:numRef><c:f>Report!$B$30:$B$31</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>4.5</c:v></c:pt></c:numCache></c:numRef></c:yVal>
                        </c:ser>
                    </c:scatterChart>`);
        const multiLevelChartXml = chartSpace(`
                    <c:lineChart>
                        <c:ser>
                            <c:tx><c:v>Visits</c:v></c:tx>
                            <c:cat><c:multiLvlStrRef><c:f>'Other Sheet'!$A$2:$B$3</c:f><c:multiLvlStrCache><c:ptCount val="2"/>
                                <c:lvl><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt></c:lvl>
                                <c:lvl><c:pt idx="0"><c:v>2024</c:v></c:pt></c:lvl>
                            </c:multiLvlStrCache></c:multiLvlStrRef></c:cat>
                            <c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>7</c:v></c:pt><c:pt idx="1"><c:v>8</c:v></c:pt></c:numLit></c:val>
                        </c:ser>
                    </c:lineChart>`);
        const arrayBuffer = createStoredZip([
            {
                name: "[Content_Types].xml",
                data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                    <Default Extension="xml" ContentType="application/xml"/>
                </Types>`,
            },
            {
                name: "_rels/.rels",
                data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>`,
            },
            {
                name: "xl/workbook.xml",
                data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheets><sheet name="Report" sheetId="1" r:id="rId1"/></sheets>
            </workbook>`,
            },
            {
                name: "xl/_rels/workbook.xml.rels",
                data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
            </Relationships>`,
            },
            {
                name: "xl/worksheets/sheet1.xml",
                data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Report</t></is></c></row></sheetData>
                <drawing r:id="rId1"/>
            </worksheet>`,
            },
            {
                name: "xl/worksheets/_rels/sheet1.xml.rels",
                data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
            </Relationships>`,
            },
            { name: "xl/drawings/drawing1.xml", data: drawingXml },
            {
                name: "xl/drawings/_rels/drawing1.xml.rels",
                data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
                <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/>
                <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>
            </Relationships>`,
            },
            { name: "xl/charts/chart1.xml", data: barChartXml },
            { name: "xl/charts/chart2.xml", data: scatterChartXml },
            { name: "xl/charts/chart3.xml", data: multiLevelChartXml },
        ]);

        const workbook = await api.parseWorkbook(arrayBuffer, "synthetic.xlsx");
        const sheet = workbook.sheets[0];
        const markdown = api.convertWorkbookToMarkdownFiles(workbook, { includeShapeDetails: false })[0].markdown;

        expect(sheet.charts.map((chart) => chart.series)).toEqual([
            [
                {
                    name: "Sales",
                    categoriesRef: "'Other Sheet'!$A$2:$A$4",
                    valuesRef: "'Other Sheet'!$B$2:$B$4",
                    cachedCategories: ["Q1", "Q2", "Q3"],
                    cachedValues: [10, null, 30],
                    axis: "primary",
                },
                {
                    name: "Cost",
                    // A series without categories must not borrow the name reference.
                    categoriesRef: "",
                    valuesRef: "",
                    cachedCategories: [],
                    cachedValues: [5, null],
                    axis: "primary",
                },
            ],
            [
                {
                    name: "XY",
                    categoriesRef: "Report!$A$30:$A$31",
                    valuesRef: "Report!$B$30:$B$31",
                    cachedCategories: ["1", "2"],
                    cachedValues: [3, 4.5],
                    axis: "primary",
                },
            ],
            [
                {
                    name: "Visits",
                    categoriesRef: "'Other Sheet'!$A$2:$B$3",
                    valuesRef: "",
                    cachedCategories: ["Jan", "Feb"],
                    cachedValues: [7, 8],
                    axis: "primary",
                },
            ],
        ]);
        expect(sheet.shapes.map((entry) => [entry.drawingId, entry.text, entry.connection])).toEqual([
            ["10", "A", null],
            ["11", "B", null],
            // Line ends also carry their points, so loose ends can resolve by geometry.
            ["12", "", expect.objectContaining({ startId: "11", endId: "10", arrow: "start" })],
            ["13", "", expect.objectContaining({ startId: "10", endId: null, arrow: "end" })],
        ]);

        expect(markdown).toContain(
            [
                "- Data (chart cache):",
                "",
                "| Category | Sales | Cost |",
                "| --- | --- | --- |",
                "| Q1 | 10 | 5 |",
                "| Q2 |  |  |",
                "| Q3 | 30 |  |",
            ].join("\n")
        );
        expect(markdown).toContain("| Category | XY |\n| --- | --- |\n| 1 | 3 |\n| 2 | 4.5 |");
        expect(markdown).toContain("| Category | Visits |\n| --- | --- |\n| Jan | 7 |\n| Feb | 8 |");
        expect(markdown).not.toContain("- Source:");
        // The arrowhead sits at the start shape (B), so the flow runs A to B.
        expect(markdown).toContain("  - Shape 003 (Straight Arrow Connector): Shape 001 → Shape 002");
        // The loose end has no shape, so the line names the cell it reaches.
        expect(markdown).toContain("  - Shape 004 (Straight Arrow Connector): Shape 001 → N7 (no text nearby)");
        expect(markdown).toContain(
            ["```mermaid", "flowchart TD", '    sh001["A"]', '    sh002(("B"))', "    sh001 --> sh002", "```"].join(
                "\n"
            )
        );
        // Assets follow their anchors: chart 1 (B3), shapes (K3), chart 2 (B20), chart 3 (B40).
        const order = [
            "### Chart: 001 (B3)",
            "### Shape Block: 001",
            "### Chart: 002 (B20)",
            "### Chart: 003 (B40)",
        ].map((heading) => markdown.indexOf(heading));
        expect(order.every((position) => position > 0)).toBe(true);
        expect([...order].sort((left, right) => left - right)).toEqual(order);

        const irSheet = buildWorkbookIR(workbook).sheets[0];
        expect(irSheet.charts[0].series.map((series) => series.sourceTable)).toEqual([null, null]);
        expect(irSheet.shapes.map((entry) => entry.connection)).toEqual([
            null,
            null,
            { from: "sh001", to: "sh002", arrow: "end" },
            { from: "sh001", to: null, arrow: "end" },
        ]);
        const compact = renderCompactMarkdown(buildWorkbookIR(workbook));
        expect(compact).toContain(
            "- Sales: categories 'Other Sheet'!$A$2:$A$4; values 'Other Sheet'!$B$2:$B$4; data Q1=10, Q2=, Q3=30"
        );
        expect(compact).toContain("Shape sh003 (Straight Arrow Connector, Connector 12): sh001 -> sh002");
        expect(compact).toContain("Shape sh004 (Straight Arrow Connector, Connector 13): sh001 -> ?");
    });

    it("parses the chart-basic fixture workbook with concrete chart expectations", async () => {
        const api = core;
        const fixtureName = "chart-basic-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "chart", fixtureName);
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
        expect(sheet.name).toBe("chart-basic");
        expect(sheet.maxRow).toBe(7);
        expect(sheet.maxCol).toBe(4);
        expect(sheet.cells).toHaveLength(16);
        expect(sheet.images).toHaveLength(0);
        expect(sheet.charts).toHaveLength(1);
        expect(sheet.charts[0]).toEqual({
            sheetName: "chart-basic",
            anchor: "B10",
            chartPath: "xl/charts/chart1.xml",
            title: "棒グラフのグラフ",
            chartType: "Bar Chart",
            series: [
                {
                    name: "値A",
                    categoriesRef: "'chart-basic'!$B$4:$B$7",
                    valuesRef: "'chart-basic'!$C$4:$C$7",
                    cachedCategories: ["2024年", "2025年", "2026年", "2027年"],
                    cachedValues: [13568, 19287, 25051, 28053],
                    axis: "primary",
                },
                {
                    name: "値B",
                    categoriesRef: "'chart-basic'!$B$4:$B$7",
                    valuesRef: "'chart-basic'!$D$4:$D$7",
                    cachedCategories: ["2024年", "2025年", "2026年", "2027年"],
                    cachedValues: [9072, 18058, 32012, 32012],
                    axis: "primary",
                },
            ],
        });

        expect(markdownFile.fileName).toBe("chart-basic-sample01_001_chart-basic.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.charts).toBe(1);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-D7"]);
        expect(markdownFile.markdown).toContain("## Sheet: chart-basic");
        expect(markdownFile.markdown).toContain("# Book: chart-basic-sample01.xlsx");
        expect(markdownFile.markdown).toContain("グラフ基本サンプル");
        expect(markdownFile.markdown).toContain("### Table: 001 (B3-D7)");
        expect(markdownFile.markdown).toContain("| 項目 | 値A | 値B |");
        expect(markdownFile.markdown).toContain("| 2027年 | 28,053 | 32,012 |");
        expect(markdownFile.markdown).toContain("### Chart: 001 (B10)");
        expect(markdownFile.markdown).toContain("- Title: 棒グラフのグラフ");
        expect(markdownFile.markdown).toContain("- Type: Bar Chart");
        expect(markdownFile.markdown).toContain("    - categories: 'chart-basic'!$B$4:$B$7");
        expect(markdownFile.markdown).toContain("    - values: 'chart-basic'!$D$4:$D$7");
        expect(markdownFile.markdown).toContain("- Type: Bar Chart\n- Source: Table 001 (B3-D7)\n- Series:");
        // The chart sits below its table, so it renders right after the table, not at the end of the sheet.
        expect(markdownFile.markdown.indexOf("### Chart: 001 (B10)")).toBeGreaterThan(
            markdownFile.markdown.indexOf("| 2027年 | 28,053 | 32,012 |")
        );
    });

    it("parses the chart-mixed fixture workbook with concrete mixed-chart expectations", async () => {
        const api = core;
        const fixtureName = "chart-mixed-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "chart", fixtureName);
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
        expect(sheet.name).toBe("chart-mixed");
        expect(sheet.maxRow).toBe(8);
        expect(sheet.maxCol).toBe(5);
        expect(sheet.cells).toHaveLength(25);
        expect(sheet.images).toHaveLength(0);
        expect(sheet.charts).toHaveLength(1);
        expect(sheet.charts[0]).toEqual({
            sheetName: "chart-mixed",
            anchor: "B10",
            chartPath: "xl/charts/chart1.xml",
            title: "棒と折れ線",
            chartType: "Bar Chart + Line Chart (Combined)",
            series: [
                {
                    name: "売上",
                    categoriesRef: "'chart-mixed'!$B$4:$B$8",
                    valuesRef: "'chart-mixed'!$C$4:$C$8",
                    cachedCategories: ["2024年", "2025年", "2026年", "2027年", "2028年"],
                    cachedValues: [13568, 19287, 25051, 28053, 31027],
                    axis: "primary",
                },
                {
                    name: "割引額",
                    categoriesRef: "'chart-mixed'!$B$4:$B$8",
                    valuesRef: "'chart-mixed'!$D$4:$D$8",
                    cachedCategories: ["2024年", "2025年", "2026年", "2027年", "2028年"],
                    cachedValues: [1000, 1200, 1500, 2300, 2500],
                    axis: "primary",
                },
                {
                    name: "利益率",
                    categoriesRef: "'chart-mixed'!$B$4:$B$8",
                    valuesRef: "'chart-mixed'!$E$4:$E$8",
                    cachedCategories: ["2024年", "2025年", "2026年", "2027年", "2028年"],
                    cachedValues: [0.25, 0.27, 0.32, 0.25, 0.18],
                    axis: "secondary",
                },
            ],
        });

        expect(markdownFile.fileName).toBe("chart-mixed-sample01_001_chart-mixed.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.images).toBe(0);
        expect(markdownFile.summary.charts).toBe(1);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-E8"]);
        expect(markdownFile.markdown).toContain("## Sheet: chart-mixed");
        expect(markdownFile.markdown).toContain("# Book: chart-mixed-sample01.xlsx");
        expect(markdownFile.markdown).toContain("| 項目 | 売上 | 割引額 | 利益率 |");
        expect(markdownFile.markdown).toContain("| 2028年 | 31,027 | 2,500 | 18% |");
        expect(markdownFile.markdown).toContain("### Chart: 001 (B10)");
        expect(markdownFile.markdown).toContain("- Title: 棒と折れ線");
        expect(markdownFile.markdown).toContain("- Type: Bar Chart + Line Chart (Combined)");
        expect(markdownFile.markdown).toContain("  - 利益率");
        expect(markdownFile.markdown).toContain("    - Axis: secondary");
        expect(markdownFile.markdown).toContain("    - values: 'chart-mixed'!$E$4:$E$8");
        expect(markdownFile.markdown).toContain("- Source: Table 001 (B3-E8)");
    });
});
