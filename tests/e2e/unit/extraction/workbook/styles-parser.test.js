import { describe, expect, it } from "vitest";
import * as stylesParser from "../../../../../src/extraction/workbook/styles-parser.ts";

describe("xlsx-context styles parser", () => {
    it("detects border sides from style attributes or child nodes", () => {
        const api = stylesParser;
        const doc = new DOMParser().parseFromString(
            '<root><top style="thin"/><bottom><color rgb="000000"/></bottom><left/><right/></root>',
            "application/xml"
        );

        expect(api.hasBorderSide(doc.getElementsByTagName("top")[0])).toBe(true);
        expect(api.hasBorderSide(doc.getElementsByTagName("bottom")[0])).toBe(true);
        expect(api.hasBorderSide(doc.getElementsByTagName("left")[0])).toBe(false);
        expect(api.hasBorderSide(doc.getElementsByTagName("right")[0])).toBe(false);
    });

    it("parses borders and custom number formats from styles.xml", () => {
        const api = stylesParser;
        const xml = `<?xml version="1.0" encoding="UTF-8"?>
      <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <numFmts count="1">
          <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
        </numFmts>
        <borders count="2">
          <border><left/><right/><top/><bottom/></border>
          <border><left style="thin"/><right/><top style="thin"/><bottom/></border>
        </borders>
        <cellXfs count="2">
          <xf numFmtId="0" borderId="0"/>
          <xf numFmtId="164" borderId="1"/>
        </cellXfs>
      </styleSheet>`;
        const files = new Map([["xl/styles.xml", new TextEncoder().encode(xml)]]);

        const styles = api.parseCellStyles(files);

        expect(styles).toHaveLength(2);
        expect(styles[0]).toEqual({
            borders: { top: false, bottom: false, left: false, right: false },
            numFmtId: 0,
            formatCode: "General",
            textStyle: { bold: false, italic: false, strike: false, underline: false },
        });
        expect(styles[1]).toEqual({
            borders: { top: true, bottom: false, left: true, right: false },
            numFmtId: 164,
            formatCode: "yyyy-mm-dd",
            textStyle: { bold: false, italic: false, strike: false, underline: false },
        });
    });

    it("resolves locale-dependent built-in formats by the locale that the file shows", () => {
        const stylesFor = ({ font = "Calibri", numFmts = "" } = {}) => {
            const xml = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              ${numFmts}
              <fonts count="1"><font><sz val="11"/><name val="${font}"/></font></fonts>
              <borders count="1"><border/></borders>
              <cellXfs count="4">
                <xf numFmtId="14" fontId="0" borderId="0"/>
                <xf numFmtId="44" fontId="0" borderId="0"/>
                <xf numFmtId="31" fontId="0" borderId="0"/>
                <xf numFmtId="3" fontId="0" borderId="0"/>
              </cellXfs>
            </styleSheet>`;
            const files = new Map([["xl/styles.xml", new TextEncoder().encode(xml)]]);
            return {
                codes: stylesParser.parseCellStyles(files).map((style) => style.formatCode),
                locale: stylesParser.readWorkbookLocale(files),
            };
        };
        const tagged = (lcid) =>
            stylesFor({
                numFmts: `<numFmts count="1"><numFmt numFmtId="164" formatCode="[$-${lcid}]d mmmm yyyy"/></numFmts>`,
            });

        // No locale tag and a Western font: the file does not say, so dates use ISO and currency has no symbol.
        expect(stylesFor()).toEqual({
            locale: null,
            codes: ["yyyy-mm-dd", '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)', "yyyy-mm-dd", "#,##0"],
        });
        expect(stylesFor({ font: "游ゴシック" })).toMatchObject({ locale: "ja-JP" });
        expect(stylesFor({ font: "游ゴシック" }).codes.slice(0, 3)).toEqual([
            "yyyy/m/d",
            '_("¥"* #,##0.00_);_("¥"* \\(#,##0.00\\);_("¥"* "-"??_);_(@_)',
            'yyyy"年"m"月"d"日"',
        ]);
        expect(tagged("42A")).toMatchObject({ locale: "vi-VN" });
        expect(tagged("42A").codes[0]).toBe("d/m/yyyy");
        expect(tagged("42A").codes[1]).toContain('"₫"');
        expect(tagged("409").codes[0]).toBe("m/d/yyyy");
        expect(tagged("40C").codes[1]).toContain('"€"');
        expect(tagged("410")).toMatchObject({ locale: "it-IT" });
    });
});

describe("xlsx-context styles parser alignment", () => {
    it("reads horizontal alignment and contained text per cell format", () => {
        const styles = `<styleSheet><fonts><font><sz val="11"/></font></fonts><borders><border/></borders>
            <cellXfs>
                <xf numFmtId="0" fontId="0" borderId="0"/>
                <xf numFmtId="0" fontId="0" borderId="0"><alignment horizontal="right"/></xf>
                <xf numFmtId="0" fontId="0" borderId="0"><alignment wrapText="1"/></xf>
                <xf numFmtId="0" fontId="0" borderId="0"><alignment horizontal="center" shrinkToFit="true"/></xf>
            </cellXfs></styleSheet>`;
        const files = new Map([["xl/styles.xml", new TextEncoder().encode(styles)]]);
        expect(stylesParser.parseCellStyles(files).map((style) => style.alignment)).toEqual([
            undefined,
            { horizontal: "right", contained: false },
            { horizontal: "general", contained: true },
            { horizontal: "center", contained: true },
        ]);
    });
});
