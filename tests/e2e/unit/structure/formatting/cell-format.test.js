import { describe, expect, it } from "vitest";
import * as cellFormat from "../../../../../src/structure/formatting/cell-format.ts";

describe("xlsx-context cell format", () => {
    it("formats date, percent, fraction, and currency values", () => {
        const api = cellFormat;

        // Excel shows serial 13 as 1900-01-13. Serial 60 is the leap day that Excel counts in 1900.
        expect(api.formatCellDisplayValue("13", { borders: {}, numFmtId: 14, formatCode: "yyyy/m/d" })).toBe(
            "1900/1/13"
        );
        expect(api.formatCellDisplayValue("60", { borders: {}, numFmtId: 14, formatCode: "yyyy-mm-dd" })).toBe(
            "1900-02-29"
        );
        expect(api.formatCellDisplayValue("61", { borders: {}, numFmtId: 14, formatCode: "yyyy-mm-dd" })).toBe(
            "1900-03-01"
        );
        expect(api.formatCellDisplayValue("0.987", { borders: {}, numFmtId: 10, formatCode: "0.0%" })).toBe("98.7%");
        expect(api.formatCellDisplayValue("0.75", { borders: {}, numFmtId: 12, formatCode: "# ?/?" })).toBe("3/4");
        expect(api.formatCellDisplayValue("1024768", { borders: {}, numFmtId: 42, formatCode: "¥ * #,##0" })).toBe(
            "¥ 1,024,768"
        );
    });

    it("formats special text and scientific notation patterns", () => {
        const api = cellFormat;

        expect(
            api.formatCellDisplayValue("1023456", { borders: {}, numFmtId: 186, formatCode: "[DBNum3]General" })
        ).toBe("1 0 2 3 4 5 6");
        expect(api.formatCellDisplayValue("1023456", { borders: {}, numFmtId: 11, formatCode: "0.000000E+00" })).toBe(
            "1.023456E+06"
        );
        expect(api.formatCellDisplayValue("0", { borders: {}, numFmtId: 0, formatCode: '0;0;"-"' })).toBe("-");
    });

    it("parses date-like text into Excel serial-compatible numbers", () => {
        const api = cellFormat;

        const serialOf = (text) => {
            const parts = api.parseDateLikeParts(text);
            return (
                parts && api.datePartsToExcelSerial(+parts.yyyy, +parts.mm, +parts.dd, +parts.hh, +parts.mi, +parts.ss)
            );
        };

        expect(serialOf("2000-03-13")).toBe(36598);
        expect(serialOf("3月14日")).toBe(36599);
        expect(serialOf("2000-03-13 12:00")).toBe(36598.5);
        expect(serialOf("not a date")).toBeNull();
    });

    it("applies resolved formula formatting back onto a cell", () => {
        const api = cellFormat;
        const cell = {
            borders: { top: false, bottom: false, left: false, right: false },
            numFmtId: 10,
            formatCode: "0.0%",
            rawValue: "",
            outputValue: "",
            resolutionStatus: null,
            resolutionSource: null,
        };

        api.applyResolvedFormulaValue(cell, "0.125", "reference_lookup");

        expect(cell.rawValue).toBe("0.125");
        expect(cell.outputValue).toBe("12.5%");
        expect(cell.resolutionStatus).toBe("resolved");
        expect(cell.resolutionSource).toBe("reference_lookup");
    });
});
