import { describe, expect, it } from "vitest";
import * as formulaEngine from "../../../../../src/formula/evaluation/engine.ts";

describe("xlsx-context formula engine", () => {
    it("prefers direct defined-name values before other resolvers", () => {
        const api = formulaEngine.createFormulaEngine({
            getDefinedNameScalarValue: () => (sheetName, name) =>
                sheetName === "Sheet1" && name === "Total" ? "42" : null,
            tryResolveFormulaExpressionWithAst: () => "999",
            parseSimpleFormulaReference: () => null,
            parseSheetScopedDefinedNameReference: () => null,
        });

        expect(api.tryResolveFormulaExpressionDetailed("=Total", "Sheet1", () => "")).toEqual({
            value: "42",
            source: "reference_lookup",
        });
    });

    it("resolves through the ast evaluator", () => {
        const api = formulaEngine.createFormulaEngine({
            getDefinedNameScalarValue: () => null,
            tryResolveFormulaExpressionWithAst: () => "12",
            parseSimpleFormulaReference: () => null,
            parseSheetScopedDefinedNameReference: () => null,
        });

        expect(api.tryResolveFormulaExpressionDetailed("=A1+B1", "Sheet1", () => "")).toEqual({
            value: "12",
            source: "ast_evaluator",
        });
    });

    it("resolves scalar operands: literals, references, defined names, then expressions", () => {
        const api = formulaEngine.createFormulaEngine({
            getDefinedNameScalarValue: () => (sheetName, name) => (name === "Rate" ? "0.1" : null),
            tryResolveFormulaExpressionWithAst: (expression) => (expression === "1+2" ? "3" : null),
            parseSimpleFormulaReference: (formulaText, sheetName) =>
                formulaText === "=B2" ? { sheetName, address: "B2" } : null,
            parseSheetScopedDefinedNameReference: () => null,
        });
        const resolveCell = (sheetName, address) => `${sheetName}!${address}`;

        expect(api.resolveScalarFormulaValue('"a""b"', "Sheet1", resolveCell)).toBe('a"b');
        expect(api.resolveScalarFormulaValue("2.50", "Sheet1", resolveCell)).toBe("2.5");
        expect(api.resolveScalarFormulaValue("B2", "Sheet1", resolveCell)).toBe("Sheet1!B2");
        expect(api.resolveScalarFormulaValue("Rate", "Sheet1", resolveCell)).toBe("0.1");
        expect(api.resolveScalarFormulaValue("true", "Sheet1", resolveCell)).toBe("TRUE");
        expect(api.resolveScalarFormulaValue("1+2", "Sheet1", resolveCell)).toBe("3");
    });
});
