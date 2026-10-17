import { readFileSync } from "node:fs";

import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import { createStoredZip } from "./stored-zip.js";
import { fixture } from "../../helpers/paths.js";

function toArrayBuffer(bytes) {
    return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function loadFixtureMarkdown(fixtureName, formattingMode) {
    const api = core;
    const fixturePath = fixture("rich", fixtureName);
    const fixtureBytes = readFileSync(fixturePath);
    const workbook = await api.parseWorkbook(toArrayBuffer(fixtureBytes), fixtureName);
    const markdownFile = api.convertWorkbookToMarkdownFiles(workbook, {
        treatFirstRowAsHeader: true,
        trimText: true,
        removeEmptyRows: true,
        removeEmptyColumns: true,
        formattingMode,
    })[0];
    return { workbook, markdownFile };
}

describe("rich workbooks", () => {
    it("preserves supported rich text as gfm-compatible markdown", async () => {
        const api = core;
        const workbookBuffer = createStoredZip([
            {
                name: "[Content_Types].xml",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
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
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>`,
            },
            {
                name: "xl/_rels/workbook.xml.rels",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
            },
            {
                name: "xl/worksheets/sheet1.xml",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData>
        <row r="1">
            <c r="A1" t="s"><v>0</v></c>
            <c r="B1" t="inlineStr" s="1"><is><t>WholeCell</t></is></c>
        </row>
    </sheetData>
</worksheet>`,
            },
            {
                name: "xl/sharedStrings.xml",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
    <si>
        <r><rPr><b/></rPr><t>Bold</t></r>
        <r><rPr><i/></rPr><t>Italic</t></r>
        <r><rPr><strike/></rPr><t>Strike</t></r>
        <r><rPr><u/></rPr><t>Under</t></r>
    </si>
</sst>`,
            },
            {
                name: "xl/styles.xml",
                data: `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="2">
        <font/>
        <font><u/></font>
    </fonts>
    <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
    <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="2">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    </cellXfs>
</styleSheet>`,
            },
        ]);

        const workbook = await api.parseWorkbook(workbookBuffer, "rich-text.xlsx");
        const gfmFiles = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            formattingMode: "gfm",
        });
        const plainFiles = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            formattingMode: "plain",
        });

        expect(gfmFiles[0].markdown).toContain("**Bold***Italic*~~Strike~~<ins>Under</ins>");
        expect(gfmFiles[0].markdown).toContain("<ins>WholeCell</ins>");
        expect(gfmFiles[0].summary.formattingMode).toBe("gfm");
        expect(plainFiles[0].markdown).toContain("BoldItalicStrikeUnder");
        expect(plainFiles[0].markdown).not.toContain("<ins>WholeCell</ins>");
    });

    it("converts the rich text fixture into gfm-compatible markdown", async () => {
        const api = core;
        const fixturePath = fixture("rich", "rich-text-gfm-sample01.xlsx");
        const fixtureBytes = readFileSync(fixturePath);
        const arrayBuffer = fixtureBytes.buffer.slice(
            fixtureBytes.byteOffset,
            fixtureBytes.byteOffset + fixtureBytes.byteLength
        );

        const workbook = await api.parseWorkbook(arrayBuffer, "rich-text-gfm-sample01.xlsx");
        const markdownFile = api.convertWorkbookToMarkdownFiles(workbook, {
            treatFirstRowAsHeader: true,
            trimText: true,
            removeEmptyRows: true,
            removeEmptyColumns: true,
            formattingMode: "gfm",
        })[0];

        expect(markdownFile.summary.formattingMode).toBe("gfm");
        expect(markdownFile.markdown).toContain("**bold whole cell**");
        expect(markdownFile.markdown).toContain("*italic whole cell*");
        expect(markdownFile.markdown).toContain("~~strike whole cell~~");
        expect(markdownFile.markdown).toContain("<ins>underline whole cell</ins>");
        expect(markdownFile.markdown).toContain("plain **bold** *italic* strike <ins>underline</ins>");
        expect(markdownFile.markdown).toContain("***bold+italic***");
        expect(markdownFile.markdown).toContain("**<ins>bold+underline</ins>**");
        // A9: only "italic" has no rPr, so only it takes the italic cell font. B9 is italic+strike as a whole cell.
        expect(markdownFile.markdown).toContain("| *italic*+~~strike~~ | *~~italic+strike~~* |");
        expect(markdownFile.markdown).toContain("改行入り文字列で<br>**一部だけ太**字");
        expect(markdownFile.markdown).toContain("重要, <ins>下線</ins>,~~取消線~~,**強調**");
        expect(markdownFile.markdown).toContain("**12345**");
        expect(markdownFile.markdown).toContain("<ins>24690</ins>");
        expect(markdownFile.markdown).toContain("| abc ***<ins>def</ins>*** | def |");
    });

    it("renders rich-markdown-escape-sample01.xlsx in gfm mode and keeps fixture-specific cases stable", async () => {
        const { workbook, markdownFile } = await loadFixtureMarkdown("rich-markdown-escape-sample01.xlsx", "gfm");

        expect(workbook.sheets).toHaveLength(1);
        expect(workbook.sheets[0].cells.length).toBe(36);
        expect(markdownFile.summary.formattingMode).toBe("gfm");
        expect(markdownFile.summary.tables).toBe(2);
        expect(markdownFile.markdown).toContain("line1 \\* x<br>**line2 \\[y\\]\\(z\\)**");
        expect(markdownFile.markdown).toContain("| Header \\| One | Header \\*Two\\* | Header \\[Three\\]\\(x\\) |");
        expect(markdownFile.markdown).toContain("| a<strong>\\*b</strong> | a\\_**b** | a\\~\\~b\\~\\~c |");
        expect(markdownFile.markdown).toContain("| \\# not **heading** | \\- not list | 1\\. ***not*** list |");
        expect(markdownFile.markdown).toContain("| a\\*b | **a\\*b** |");
        expect(markdownFile.markdown).toContain("| a\\_b | *a\\_b* |");
        expect(markdownFile.markdown).toContain("| a\\~\\~b\\~\\~c | ~~a\\~\\~b\\~\\~c~~ |");
        expect(markdownFile.markdown).toContain("| \\# not heading | <ins>\\# not heading</ins> |");
        expect(markdownFile.markdown).toContain("| &lt;tag&gt; | &lt;tag&gt; |");
        expect(markdownFile.markdown).toContain("| \\!\\[alt\\]\\(image.png\\) | \\!\\[alt\\]\\(image.png\\) |");
        expect(markdownFile.markdown).toContain("| code \\`sample\\` | code \\`sample\\` |");
        expect(markdownFile.markdown).toContain("| path\\\\to\\\\file | path\\\\to\\\\file |");
        expect(markdownFile.markdown).toContain(
            "| \\- not list | plain \\* **star** \\_ <ins>under</ins> \\~ ~~tilde~~ |"
        );
        expect(markdownFile.markdown).toContain(
            "| \\[label\\]\\(https://example.com\\) | \\[label\\]\\(https://example.com\\) |"
        );
        expect(markdownFile.markdown).toContain("| a \\| b | a \\| b |");
        expect(markdownFile.markdown).toContain(
            "| \\# a\\*b \\[x\\]\\(y\\) \\~ z \\| q | \\# a\\*b \\[x\\]\\(y\\) \\~ z \\| q |"
        );
    });

    it("renders rich-markdown-escape-sample01.xlsx in plain mode as plain text without <br>", async () => {
        const { markdownFile } = await loadFixtureMarkdown("rich-markdown-escape-sample01.xlsx", "plain");

        expect(markdownFile.summary.formattingMode).toBe("plain");
        expect(markdownFile.summary.tables).toBe(2);
        expect(markdownFile.markdown).toContain("line1 \\* x line2 \\[y\\]\\(z\\)");
        expect(markdownFile.markdown).not.toContain("<br>");
        expect(markdownFile.markdown).toContain("| a\\*b | a\\*b |");
        expect(markdownFile.markdown).toContain("| a\\_b | a\\_b |");
        expect(markdownFile.markdown).toContain("| a\\~\\~b\\~\\~c | a\\~\\~b\\~\\~c |");
        expect(markdownFile.markdown).toContain("| \\# not heading | \\# not heading |");
        expect(markdownFile.markdown).toContain("| a\\*b | a\\_b | a\\~\\~b\\~\\~c |");
        expect(markdownFile.markdown).toContain("| \\# not heading | \\- not list | 1\\. not list |");
        expect(markdownFile.markdown).toContain("| \\!\\[alt\\]\\(image.png\\) | \\!\\[alt\\]\\(image.png\\) |");
        expect(markdownFile.markdown).toContain("| code \\`sample\\` | code \\`sample\\` |");
        expect(markdownFile.markdown).toContain("| path\\\\to\\\\file | path\\\\to\\\\file |");
        expect(markdownFile.markdown).toContain("| \\- not list | plain \\* star \\_ under \\~ tilde |");
        expect(markdownFile.markdown).toContain(
            "| \\[label\\]\\(https://example.com\\) | \\[label\\]\\(https://example.com\\) |"
        );
        expect(markdownFile.markdown).toContain("| a \\| b | a \\| b |");
    });

    it("renders rich-usecase-sample01.xlsx in gfm mode with hyperlinks, inline styling, and <br>", async () => {
        const { workbook, markdownFile } = await loadFixtureMarkdown("rich-usecase-sample01.xlsx", "gfm");

        expect(workbook.sheets).toHaveLength(1);
        expect(markdownFile.summary.formattingMode).toBe("gfm");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.markdown).toContain(
            "| [Apple](https://www.apple.com/) | ***Apple*** の製品が<ins>購入できます</ins>。 |"
        );
        expect(markdownFile.markdown).toContain(
            "| [Google](https://www.google.com/) | とても<ins>有名</ins>な**検索サイト**です。 |"
        );
        expect(markdownFile.markdown).toContain(
            "| [Amazon](https://www.amazon.co.jp/) | <strong><ins>お買い物</ins></strong>でお世話になっています。 |"
        );
        expect(markdownFile.markdown).toContain(
            "| [Yodobashi](https://www.yodobashi.com/) | 実店舗とともに<br>**ネットショップ**でもお世話になっています。 |"
        );
        expect(markdownFile.markdown).toContain(
            "~~池袋の激戦区で、生き残るのはどの店舗か。~~<br>→トルツメ: この部分は文面から外すことを提案。"
        );
        expect(markdownFile.markdown).toContain("| ***Gemini*** さんにもお世話になっています。 |");
        expect(markdownFile.markdown).toContain("| お買い物で**かなり**お世話になっています。 |");
    });

    it("renders rich-usecase-sample01.xlsx in plain mode with links preserved and styling removed", async () => {
        const { markdownFile } = await loadFixtureMarkdown("rich-usecase-sample01.xlsx", "plain");

        expect(markdownFile.summary.formattingMode).toBe("plain");
        expect(markdownFile.summary.tables).toBe(1);
        expect(markdownFile.markdown).toContain("| [Apple](https://www.apple.com/) | Apple の製品が購入できます。 |");
        expect(markdownFile.markdown).toContain("| [Google](https://www.google.com/) | とても有名な検索サイトです。 |");
        expect(markdownFile.markdown).toContain(
            "| [Amazon](https://www.amazon.co.jp/) | お買い物でお世話になっています。 |"
        );
        expect(markdownFile.markdown).toContain(
            "| [Yodobashi](https://www.yodobashi.com/) | 実店舗とともに ネットショップでもお世話になっています。 |"
        );
        expect(markdownFile.markdown).toContain(
            "池袋の激戦区で、生き残るのはどの店舗か。 →トルツメ: この部分は文面から外すことを提案。"
        );
        expect(markdownFile.markdown).not.toContain("<br>");
        expect(markdownFile.markdown).not.toContain("<ins>");
        expect(markdownFile.markdown).not.toContain("~~池袋の激戦区で、生き残るのはどの店舗か。~~");
    });
});
