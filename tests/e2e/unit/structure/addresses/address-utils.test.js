import { describe, expect, it } from "vitest";
import * as addressUtils from "../../../../../src/structure/addresses/address-utils.ts";

describe("xlsx-context address utils", () => {
    it("parses sheet-qualified ranges from chart formulas", () => {
        const api = addressUtils;

        expect(api.parseSheetRangeRef("'chart basic'!$B$4:$B$7")).toEqual({
            sheetName: "chart basic",
            startRow: 4,
            startCol: 2,
            endRow: 7,
            endCol: 2,
        });
        expect(api.parseSheetRangeRef("'It''s'!C3")).toEqual({
            sheetName: "It's",
            startRow: 3,
            startCol: 3,
            endRow: 3,
            endCol: 3,
        });
        expect(api.parseSheetRangeRef("Sheet1!$D$9:$B$2")).toEqual({
            sheetName: "Sheet1",
            startRow: 2,
            startCol: 2,
            endRow: 9,
            endCol: 4,
        });
        expect(api.parseSheetRangeRef("$A$1:$A$3")).toMatchObject({ sheetName: null, startRow: 1, endRow: 3 });
        expect(api.parseSheetRangeRef("(Sheet1!$A$1:$A$3,Sheet1!$C$1:$C$3)")).toBeNull();
        expect(api.parseSheetRangeRef("SalesRange")).toBeNull();
        expect(api.parseSheetRangeRef("Sheet1!$A$0")).toBeNull();
        expect(api.parseSheetRangeRef("")).toBeNull();
    });
});
