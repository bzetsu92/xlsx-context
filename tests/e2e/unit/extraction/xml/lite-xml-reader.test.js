import { afterEach, describe, expect, it, vi } from "vitest";
import * as liteXmlReader from "../../../../../src/extraction/xml/lite-xml-reader.ts";

function domCells(xml) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    return Array.from(doc.getElementsByTagName("c"));
}

function describeCell(cell) {
    return {
        r: cell.getAttribute("r"),
        t: cell.getAttribute("t"),
        v: cell.getElementsByTagName("v")[0]?.textContent ?? null,
        f: cell.getElementsByTagName("f")[0]?.textContent ?? null,
        ts: Array.from(cell.getElementsByTagName("t")).map((node) => node.textContent),
        text: cell.textContent,
    };
}

const read = (xml) => {
    const result = liteXmlReader.readContainer(xml, "sheetData", "c");
    return result && { cells: result.items, xmlWithoutSheetData: result.xmlWithoutContainer };
};

describe("lite xml reader", () => {
    afterEach(() => {
        vi.restoreAllMocks();
    });

    it("matches the DOM for entities, CDATA, comments, inline strings and self-closing cells", () => {
        const xml = `<?xml version="1.0"?><worksheet xmlns="main"><dimension ref="A1"/><sheetData>
<row r="1"><c r="A1" t="str"><v>a &lt;b&gt; &amp; &quot;c&quot; &apos;d&apos; &#65;&#x42;</v></c>
<c r="B1" t="inlineStr"><is><r><rPr><b/></rPr><t xml:space="preserve"> x </t></r><r><t><![CDATA[<y>]]></t></r></is></c>
<!-- skipped --><c r="C1" s="2"/><c r="D1"><f>SUM(A1:B1)</f><v>3</v></c>
<c r = 'E1' t="s"
  s="1"><v>0</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>`;
        const result = read(xml);
        expect(result).not.toBeNull();
        expect(result.cells.map(describeCell)).toEqual(domCells(xml).map(describeCell));
        expect(result.xmlWithoutSheetData).toContain("<sheetData/><mergeCells");
        expect(domCells(result.xmlWithoutSheetData)).toHaveLength(0);
    });

    it("normalizes line endings like the DOM parser", () => {
        const xml = `<worksheet><sheetData><row><c r="A1" t="str"><v>a\r\nb\rc</v></c></row></sheetData></worksheet>`;
        expect(read(xml).cells[0].textContent).toBe("a\nb\nc");
    });

    it("falls back to the DOM path for markup outside its envelope", () => {
        expect(read(`<!DOCTYPE x><worksheet><sheetData/></worksheet>`)).toBeNull();
        expect(read(`<worksheet><sheetData><c r="A1"><v>&nbsp;</v></c></sheetData></worksheet>`)).toBeNull();
        expect(read(`<worksheet><sheetData><c r="A1"><v>1</c></sheetData></worksheet>`)).toBeNull();
        expect(read(`<worksheet><sheetData><c r="A1"><v>1</v></c></worksheet>`)).toBeNull();
        expect(read(`<x:worksheet xmlns:x="m"><x:sheetData><x:c/></x:sheetData></x:worksheet>`)).toBeNull();
        expect(read(`<worksheet><sheetData/><c r="A1"/></worksheet>`)).toBeNull();
    });

    it("reads shared strings with rich runs and phonetic hints like the DOM", () => {
        const xml = `<?xml version="1.0"?><sst xmlns="main" count="3" uniqueCount="3">
<si><t>plain &amp; simple</t></si>
<si><r><rPr><b val="1"/></rPr><t>bold</t></r><r><t xml:space="preserve"> tail</t></r></si>
<si><t>kanji</t><rPh sb="0" eb="1"><t>kana</t></rPh><phoneticPr fontId="1"/></si></sst>`;
        const fast = liteXmlReader.readContainer(xml, "sst", "si");
        const dom = new DOMParser().parseFromString(xml, "application/xml");
        const shape = (item) => ({
            text: item.textContent,
            ts: Array.from(item.getElementsByTagName("t")).map((node) => node.textContent),
            children: Array.from(item.childNodes).map((node) => node.nodeType + ":" + (node.localName || "")),
        });
        expect(fast.items.map(shape)).toEqual(Array.from(dom.getElementsByTagName("si")).map(shape));
    });

    it("keeps an empty sheet unchanged", () => {
        const xml = `<worksheet><sheetData/></worksheet>`;
        expect(read(xml)).toEqual({ cells: [], xmlWithoutSheetData: xml });
    });
});
