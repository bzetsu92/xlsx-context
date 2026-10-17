import { readFileSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { unzipEntries } from "../../../src/extraction/zip/zip-io.ts";
import { createWorkbookExportArchive } from "../../../src/outputs/markdown/markdown-export.ts";
import { createStoredZip } from "./stored-zip.js";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("image workbooks", () => {
    const fixtureDir = FIXTURE_ROOT;

    it("parses the image-basic fixture workbook with concrete image expectations", async () => {
        const api = core;
        const fixtureName = "image-basic-sample01.xlsx";
        const fixturePath = path.resolve(fixtureDir, "image", fixtureName);
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
        expect(sheet.name).toBe("image");
        expect(sheet.maxRow).toBe(13);
        expect(sheet.maxCol).toBe(6);
        expect(sheet.cells).toHaveLength(39);
        expect(sheet.images).toHaveLength(2);
        expect(
            sheet.images.map((image) => ({
                filename: image.filename,
                path: image.path,
                anchor: image.anchor,
                mediaPath: image.mediaPath,
            }))
        ).toEqual([
            {
                filename: "image_001.png",
                path: "assets/image/image_001.png",
                anchor: "C8",
                mediaPath: "xl/media/image1.png",
            },
            {
                filename: "image_002.png",
                path: "assets/image/image_002.png",
                anchor: "F8",
                mediaPath: "xl/media/image2.png",
            },
        ]);
        expect(sheet.cells.find((cell) => cell.address === "A1")?.outputValue).toBe("画像抽出サンプル");
        expect(sheet.cells.find((cell) => cell.address === "D4")?.outputValue).toBe("123,456");
        expect(sheet.cells.find((cell) => cell.address === "E4")?.outputValue).toBe("1,234.56");
        expect(sheet.cells.find((cell) => cell.address === "F4")?.outputValue).toBe("値1");
        expect(sheet.cells.find((cell) => cell.address === "D6")?.outputValue).toBe("345,678");
        expect(sheet.cells.find((cell) => cell.address === "E6")?.outputValue).toBe("3,456.89");

        expect(markdownFile.fileName).toBe("image-basic-sample01_001_image.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.merges).toBe(0);
        expect(markdownFile.summary.images).toBe(2);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-F6"]);
        expect(markdownFile.markdown).toContain("## Sheet: image");
        expect(markdownFile.markdown).toContain("# Book: image-basic-sample01.xlsx");
        expect(markdownFile.markdown).toContain("| 項目1 | 項目2 | 項目3 | 項目4 | 項目5 |");
        expect(markdownFile.markdown).toContain("| B4 | C4 | 123,456 | 1,234.56 | 値1 |");
        expect(markdownFile.markdown).toContain("| B6 | C6 | 345,678 | 3,456.89 | 値3 |");
        expect(markdownFile.markdown).toContain("### Image: 001 (C8)");
        expect(markdownFile.markdown).toContain("- File: assets/image/image_001.png");
        expect(markdownFile.markdown).toContain("![image_001.png](assets/image/image_001.png)");
        expect(markdownFile.markdown).toContain("### Image: 002 (F8)");
        expect(markdownFile.markdown).toContain("- File: assets/image/image_002.png");
        expect(markdownFile.markdown).toContain("![image_002.png](assets/image/image_002.png)");
    });

    it("parses the image-basic-sample02 fixture workbook with concrete image and chart expectations", async () => {
        const api = core;
        const fixtureName = "image-basic-sample02.xlsx";
        const fixturePath = path.resolve(fixtureDir, "image", fixtureName);
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
        expect(sheet.name).toBe("image");
        expect(sheet.maxRow).toBe(6);
        expect(sheet.maxCol).toBe(4);
        expect(sheet.cells).toHaveLength(13);
        expect(sheet.images).toHaveLength(1);
        expect(sheet.charts).toHaveLength(1);
        expect(sheet.images[0]).toMatchObject({
            filename: "image_001.png",
            path: "assets/image/image_001.png",
            anchor: "H3",
            mediaPath: "xl/media/image1.png",
        });
        expect(sheet.charts[0]).toEqual({
            sheetName: "image",
            anchor: "B9",
            chartPath: "xl/charts/chart1.xml",
            title: "このグラフのタイトル",
            chartType: "Line Chart",
            series: [
                {
                    name: "値A",
                    categoriesRef: "image!$B$4:$B$6",
                    valuesRef: "image!$C$4:$C$6",
                    cachedCategories: ["2024年", "2025年", "2026年"],
                    cachedValues: [13568, 19287, 25051],
                    axis: "primary",
                },
                {
                    name: "値B",
                    categoriesRef: "image!$B$4:$B$6",
                    valuesRef: "image!$D$4:$D$6",
                    cachedCategories: ["2024年", "2025年", "2026年"],
                    cachedValues: [9072, 18058, 32012],
                    axis: "primary",
                },
            ],
        });

        expect(markdownFile.fileName).toBe("image-basic-sample02_001_image.md");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.summary.images).toBe(1);
        expect(markdownFile.summary.charts).toBe(1);
        expect(markdownFile.summary.tableScores.map((detail) => detail.range)).toEqual(["B3-D6"]);
        expect(markdownFile.markdown).toContain("## Sheet: image");
        expect(markdownFile.markdown).toContain("# Book: image-basic-sample02.xlsx");
        expect(markdownFile.markdown).toContain("| 項目 | 値A | 値B |");
        expect(markdownFile.markdown).toContain("| 2024年 | 13,568 | 9,072 |");
        expect(markdownFile.markdown).toContain("### Chart: 001 (B9)");
        expect(markdownFile.markdown).toContain("- Title: このグラフのタイトル");
        expect(markdownFile.markdown).toContain("- Type: Line Chart");
        expect(markdownFile.markdown).toContain("    - values: image!$D$4:$D$6");
        expect(markdownFile.markdown).toContain("### Image: 001 (H3)");
        expect(markdownFile.markdown).toContain("- File: assets/image/image_001.png");
        expect(markdownFile.markdown).toContain("- Source: Table 001 (B3-D6)");
        // The table holds every series, so the chart repeats no cached data.
        expect(markdownFile.markdown).not.toContain("- Data (chart cache):");
    });

    it("extracts drawing image metadata and renders markdown image section", async () => {
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
            <c r="A1" t="inlineStr"><is><t>Sheet with image</t></is></c>
        </row>
    </sheetData>
</worksheet>`;
        const sheetRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`;
        const drawingXml = `<?xml version="1.0" encoding="UTF-8"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <xdr:oneCellAnchor>
        <xdr:from>
            <xdr:col>2</xdr:col>
            <xdr:row>1</xdr:row>
        </xdr:from>
        <xdr:pic>
            <xdr:blipFill>
                <a:blip r:embed="rIdImage1"/>
            </xdr:blipFill>
        </xdr:pic>
    </xdr:oneCellAnchor>
</xdr:wsDr>`;
        const drawingRelsXml = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/sample.png"/>
</Relationships>`;
        const pngBytes = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]);
        const zip = createStoredZip([
            {
                name: "[Content_Types].xml",
                data: `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>`,
            },
            { name: "xl/workbook.xml", data: workbookXml },
            { name: "xl/_rels/workbook.xml.rels", data: workbookRelsXml },
            { name: "xl/worksheets/sheet1.xml", data: sheetXml },
            { name: "xl/worksheets/_rels/sheet1.xml.rels", data: sheetRelsXml },
            { name: "xl/drawings/drawing1.xml", data: drawingXml },
            { name: "xl/drawings/_rels/drawing1.xml.rels", data: drawingRelsXml },
            { name: "xl/media/sample.png", data: pngBytes },
        ]);

        const workbook = await api.parseWorkbook(zip, "diagram.xlsx");
        const files = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
        });

        expect(workbook.sheets[0].images).toHaveLength(1);
        expect(workbook.sheets[0].images[0].anchor).toBe("C2");
        expect(workbook.sheets[0].images[0].path).toBe("assets/Summary/image_001.png");
        expect(files[0].summary.images).toBe(1);
        expect(files[0].markdown).toContain("### Image: 001 (C2)");
        expect(files[0].markdown).toContain("![image_001.png](assets/Summary/image_001.png)");

        const archive = createWorkbookExportArchive(workbook, files);
        const extracted = await unzipEntries(
            archive.buffer.slice(archive.byteOffset, archive.byteOffset + archive.byteLength)
        );
        expect(extracted.has("output/diagram.md")).toBe(true);
        expect(extracted.has("output/assets/Summary/image_001.png")).toBe(true);
        const markdownText = new TextDecoder().decode(extracted.get("output/diagram.md"));
        expect(markdownText).toContain("# Book: diagram.xlsx");
        expect(markdownText).toContain("### Image: 001 (C2)");
        expect(extracted.get("output/assets/Summary/image_001.png")).toBeDefined();
    });
});
