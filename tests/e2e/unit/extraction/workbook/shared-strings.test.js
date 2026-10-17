import { describe, expect, it } from "vitest";
import * as sharedStrings from "../../../../../src/extraction/workbook/shared-strings.ts";

describe("xlsx-context shared strings", () => {
    it("collects plain and rich text runs", () => {
        const api = sharedStrings;
        const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
        <si><t>通常のテキスト</t></si>
        <si><r><t>分割</t></r><r><t>テキスト</t></r></si>
      </sst>`;
        const files = new Map([["xl/sharedStrings.xml", new TextEncoder().encode(xml)]]);

        expect(api.parseSharedStrings(files)).toEqual([
            { text: "通常のテキスト", runs: null },
            { text: "分割テキスト", runs: null },
        ]);
    });

    it("skips phonetic text nodes and normalizes CRLF to LF", () => {
        const api = sharedStrings;
        const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
        <si>
          <t>line1&#13;&#10;line2</t>
          <rPh sb="0" eb="1"><t>phonetic</t></rPh>
          <phoneticPr fontId="1"/>
        </si>
      </sst>`;
        const files = new Map([["xl/sharedStrings.xml", new TextEncoder().encode(xml)]]);

        expect(api.parseSharedStrings(files)).toEqual([{ text: "line1\nline2", runs: null }]);
    });

    it("preserves supported rich text emphasis flags", () => {
        const api = sharedStrings;
        const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
        <si>
          <r><rPr><b/></rPr><t>Bold</t></r>
          <r><rPr><i/></rPr><t>Italic</t></r>
          <r><rPr><strike/></rPr><t>Strike</t></r>
          <r><rPr><u/></rPr><t>Under</t></r>
        </si>
      </sst>`;
        const files = new Map([["xl/sharedStrings.xml", new TextEncoder().encode(xml)]]);

        expect(api.parseSharedStrings(files)).toEqual([
            {
                text: "BoldItalicStrikeUnder",
                runs: [
                    { text: "Bold", ownStyle: true, bold: true, italic: false, strike: false, underline: false },
                    { text: "Italic", ownStyle: true, bold: false, italic: true, strike: false, underline: false },
                    { text: "Strike", ownStyle: true, bold: false, italic: false, strike: true, underline: false },
                    { text: "Under", ownStyle: true, bold: false, italic: false, strike: false, underline: true },
                ],
            },
        ]);
    });

    it("marks runs without rPr and keeps runs whose rPr sets no emphasis", () => {
        const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
        <si>
          <r><t>cell</t></r>
          <r><rPr><sz val="11"/></rPr><t>plain</t></r>
        </si>
      </sst>`;
        const files = new Map([["xl/sharedStrings.xml", new TextEncoder().encode(xml)]]);

        expect(sharedStrings.parseSharedStrings(files)).toEqual([
            {
                text: "cellplain",
                runs: [
                    { text: "cell", ownStyle: false, bold: false, italic: false, strike: false, underline: false },
                    { text: "plain", ownStyle: true, bold: false, italic: false, strike: false, underline: false },
                ],
            },
        ]);
    });

    it("converts to the public entry shape without font sources", () => {
        const plain = { text: "x", bold: false, italic: false, strike: false, underline: false };

        expect(
            sharedStrings.toSharedStringEntry({
                text: "ab",
                runs: [
                    { ...plain, text: "a", ownStyle: false },
                    { ...plain, text: "b", ownStyle: true },
                ],
            })
        ).toEqual({ text: "ab", runs: null });
        expect(
            sharedStrings.toSharedStringEntry({
                text: "abc",
                runs: [
                    { ...plain, text: "a", ownStyle: false },
                    { ...plain, text: "b", ownStyle: true },
                    { ...plain, text: "c", ownStyle: true, bold: true },
                ],
            })
        ).toEqual({
            text: "abc",
            runs: [
                { text: "ab", bold: false, italic: false, strike: false, underline: false },
                { text: "c", bold: true, italic: false, strike: false, underline: false },
            ],
        });
    });
});
