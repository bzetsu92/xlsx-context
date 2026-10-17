import { describe, expect, it } from "vitest";
import {
    findTopLevelOperatorIndex,
    parseWholeFunctionCall,
    splitFormulaArguments,
} from "../../../../../src/formula/references/formula-text.ts";

describe("xlsx-context formula text", () => {
    it("splits function arguments with nesting and quoted commas", () => {
        expect(splitFormulaArguments('A1,"x,y",SUM(B1:B2)')).toEqual(["A1", '"x,y"', "SUM(B1:B2)"]);
    });

    it("finds top-level operators outside quotes and parentheses", () => {
        expect(findTopLevelOperatorIndex("'a:b'!A1:INDEX(B1:B3,2)", ":")).toBe(8);
        expect(findTopLevelOperatorIndex('"x:y"', ":")).toBe(-1);
    });

    it("parses a whole call only for allowed names", () => {
        expect(parseWholeFunctionCall("INDEX(A1:A3,2)", ["INDEX"])).toEqual({ name: "INDEX", argsText: "A1:A3,2" });
        expect(parseWholeFunctionCall("INDEX(A1:A3,2)+1", ["INDEX"])).toBeNull();
        expect(parseWholeFunctionCall("SUM(A1:A3)", ["INDEX"])).toBeNull();
    });
});
