import { describe, expect, it } from "vitest";
import * as formulaReferenceUtils from "../../../../../src/formula/references/reference-utils.ts";

describe("xlsx-context formula reference utils", () => {
    it("parses local and qualified cell references", () => {
        const api = formulaReferenceUtils.createFormulaReferenceUtils({
            normalizeFormulaAddress: (address) =>
                String(address || "")
                    .replace(/\$/g, "")
                    .toUpperCase(),
        });

        expect(api.parseSimpleFormulaReference("=A1", "Sheet1")).toEqual({
            sheetName: "Sheet1",
            address: "A1",
        });
        expect(api.parseSimpleFormulaReference("='Sheet 2'!$B$3", "Sheet1")).toEqual({
            sheetName: "Sheet 2",
            address: "B3",
        });
    });

    it("parses sheet-scoped defined-name references but excludes cell refs", () => {
        const api = formulaReferenceUtils.createFormulaReferenceUtils({
            normalizeFormulaAddress: (address) => String(address || ""),
        });

        expect(api.parseSheetScopedDefinedNameReference("'Sheet 2'!Total_Value", "Sheet1")).toEqual({
            sheetName: "Sheet 2",
            name: "Total_Value",
        });
        expect(api.parseSheetScopedDefinedNameReference("Sheet2!B3", "Sheet1")).toBeNull();
    });

    it("resolves qualified ranges, structured references and defined names", () => {
        const api = formulaReferenceUtils.createFormulaReferenceUtils({
            normalizeFormulaAddress: (address) => String(address || ""),
            parseRangeAddress: (rawRange) => {
                const match = String(rawRange || "").match(/^\$?([A-Z]+\$?\d+):\$?([A-Z]+\$?\d+)$/i);
                return match ? { start: match[1].replace(/\$/g, ""), end: match[2].replace(/\$/g, "") } : null;
            },
            getDefinedNameRangeRef: () => (sheetName, name) =>
                name === "Prices" ? { sheetName: "Data", start: "A1", end: "A9" } : null,
            getStructuredRangeRef: () => (sheetName, text) =>
                sheetName === "Sheet1" && text === "Sales[Amount]"
                    ? { sheetName: "Sheet1", start: "D3", end: "D4" }
                    : null,
        });

        expect(api.parseQualifiedRangeReference("'My Sheet'!A1:B2", "Sheet1")).toEqual({
            sheetName: "My Sheet",
            start: "A1",
            end: "B2",
        });
        expect(api.parseQualifiedRangeReference("C1:C3", "Sheet1")).toEqual({
            sheetName: "Sheet1",
            start: "C1",
            end: "C3",
        });
        expect(api.parseQualifiedRangeReference("Sales[Amount]", "Sheet1")).toEqual({
            sheetName: "Sheet1",
            start: "D3",
            end: "D4",
        });
        expect(api.parseQualifiedRangeReference("Prices", "Sheet1")).toEqual({
            sheetName: "Data",
            start: "A1",
            end: "A9",
        });
        expect(api.parseQualifiedRangeReference("Unknown", "Sheet1")).toBeNull();
    });
});
