import { describe, expect, it } from "vitest";
import {
    buildSheetGrid,
    columnAtEmu,
    columnLeftEmu,
    columnWidthPx,
    createUniformGrid,
    digitWidthPx,
    readNormalFont,
    rowAtEmu,
    rowTopEmu,
} from "../../../../../src/extraction/workbook/sheet-grid.ts";
import { xmlToDocument } from "../../../../../src/extraction/xml/xml-utils.ts";

const PX = 9525;
const PT = 12700;
const CALIBRI = { name: "Calibri", sizePt: 11 };
const PGOTHIC = { name: "ＭＳ Ｐゴシック", sizePt: 11 };

function sheetParts(xml) {
    const doc = xmlToDocument(`<worksheet>${xml}</worksheet>`);
    return {
        formatPr: doc.getElementsByTagName("sheetFormatPr")[0] || null,
        cols: Array.from(doc.getElementsByTagName("col")),
        rows: Array.from(doc.getElementsByTagName("row")),
    };
}

function gridOf(xml, font = CALIBRI) {
    const parts = sheetParts(xml);
    return buildSheetGrid(parts.formatPr, parts.cols, parts.rows, font);
}

describe("sheet grid", () => {
    it("uses the Excel default of 64 px by 15 pt without sheet data", () => {
        const grid = gridOf("");
        expect(grid.defaultColWidthEmu).toBe(64 * PX);
        expect(grid.defaultRowHeightEmu).toBe(15 * PT);
        expect(grid.confidence).toBe("EXACT");
    });

    it("measures digits of full-width CJK font names", () => {
        expect(digitWidthPx(PGOTHIC)).toEqual({ px: 8, known: true });
        // Japanese Excel shows 8.38 characters for 72 px.
        expect(gridOf("", PGOTHIC).defaultColWidthEmu).toBe(72 * PX);
    });

    it("marks grids of unknown fonts as estimated", () => {
        const grid = gridOf("", { name: "Some Font", sizePt: 11 });
        expect(grid.confidence).toBe("ESTIMATED");
        expect(grid.defaultColWidthEmu).toBeGreaterThan(0);
    });

    it("converts stored widths with the ECMA-376 formula", () => {
        expect(columnWidthPx(9.140625, 7)).toBe(64);
        expect(columnWidthPx(2.5, 8)).toBe(20);
        expect(columnWidthPx(0, 7)).toBe(0);
    });

    it("reads defaultColWidth before baseColWidth", () => {
        const grid = gridOf(
            '<sheetFormatPr defaultColWidth="2.50390625" baseColWidth="20" defaultRowHeight="11.25"/>',
            PGOTHIC
        );
        expect(grid.defaultColWidthEmu).toBe(20 * PX);
        expect(grid.defaultRowHeightEmu).toBe(11.25 * PT);
    });

    it("places columns after custom, hidden, and open-ended spans", () => {
        const grid = gridOf(
            '<cols><col min="2" max="3" width="20" customWidth="1"/><col min="4" max="4" width="9" hidden="true"/>' +
                '<col min="6" max="16384" width="2.5"/></cols>'
        );
        const wide = columnWidthPx(20, 7) * PX;
        expect(columnLeftEmu(grid, 1)).toBe(0);
        expect(columnLeftEmu(grid, 2)).toBe(64 * PX);
        expect(columnLeftEmu(grid, 4)).toBe(64 * PX + 2 * wide);
        // Column D is hidden, so E starts where D starts.
        expect(columnLeftEmu(grid, 5)).toBe(columnLeftEmu(grid, 4));
        expect(columnLeftEmu(grid, 7) - columnLeftEmu(grid, 6)).toBe(columnWidthPx(2.5, 7) * PX);
        expect(columnAtEmu(grid, columnLeftEmu(grid, 4))).toBe(5);
    });

    it("lets a later col element win where spans overlap", () => {
        const grid = gridOf('<cols><col min="1" max="5" width="20"/><col min="3" max="3" width="5"/></cols>');
        expect(columnLeftEmu(grid, 4) - columnLeftEmu(grid, 3)).toBe(columnWidthPx(5, 7) * PX);
        expect(columnLeftEmu(grid, 6) - columnLeftEmu(grid, 5)).toBe(columnWidthPx(20, 7) * PX);
    });

    it("skips col elements without a valid range", () => {
        const grid = gridOf('<cols><col width="20"/><col min="5" max="2" width="20"/><col min="x" width="1"/></cols>');
        expect(grid.colSpans).toEqual([]);
    });

    it("reads row heights, hidden rows, and rows without r", () => {
        const grid = gridOf(
            '<sheetFormatPr defaultRowHeight="15"/><sheetData><row r="2" ht="30"/><row ht="20"/>' +
                '<row r="6" hidden="1" ht="40"/></sheetData>'
        );
        expect(rowTopEmu(grid, 2)).toBe(15 * PT);
        expect(rowTopEmu(grid, 3)).toBe(45 * PT);
        expect(rowTopEmu(grid, 4)).toBe(65 * PT);
        expect(rowTopEmu(grid, 7) - rowTopEmu(grid, 6)).toBe(0);
        expect(rowAtEmu(grid, 46 * PT)).toBe(3);
    });

    it("hides unlisted rows under zeroHeight and keeps listed rows", () => {
        const grid = gridOf('<sheetFormatPr defaultRowHeight="15" zeroHeight="1"/><sheetData><row r="3"/></sheetData>');
        expect(grid.defaultRowHeightEmu).toBe(0);
        expect(rowTopEmu(grid, 3)).toBe(0);
        expect(rowTopEmu(grid, 4)).toBe(15 * PT);
        expect(rowAtEmu(grid, 1)).toBe(3);
        expect(rowAtEmu(grid, 20 * PT)).toBe(4);
    });

    it("maps every cell edge back to its own index", () => {
        const grid = gridOf(
            '<sheetFormatPr defaultRowHeight="13.5"/><cols><col min="2" max="2" width="30"/><col min="5" max="9" width="1.5"/></cols>' +
                '<sheetData><row r="1" ht="47.25"/><row r="4" ht="3"/></sheetData>',
            PGOTHIC
        );
        for (let index = 1; index <= 40; index += 1) {
            expect(columnAtEmu(grid, columnLeftEmu(grid, index))).toBe(index);
            expect(columnAtEmu(grid, columnLeftEmu(grid, index + 1) - 1)).toBe(index);
            expect(rowAtEmu(grid, rowTopEmu(grid, index))).toBe(index);
            expect(rowAtEmu(grid, rowTopEmu(grid, index + 1) - 1)).toBe(index);
        }
    });

    it("clamps offsets outside the sheet", () => {
        const grid = createUniformGrid(64 * PX, 15 * PT);
        expect(columnAtEmu(grid, -5)).toBe(1);
        expect(rowAtEmu(grid, 0)).toBe(1);
        expect(columnAtEmu(grid, Number.MAX_SAFE_INTEGER)).toBe(16384);
    });

    it("reads the normal font from fonts and ignores dxfs fonts", () => {
        const styles =
            '<styleSheet><fonts count="1"><font><sz val="11"/><name val="ＭＳ Ｐゴシック"/></font></fonts>' +
            '<dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>';
        const files = new Map([["xl/styles.xml", new TextEncoder().encode(styles)]]);
        expect(readNormalFont(files, (bytes) => new TextDecoder().decode(bytes))).toEqual(PGOTHIC);
        expect(readNormalFont(new Map(), (bytes) => new TextDecoder().decode(bytes))).toEqual(CALIBRI);
    });
});
