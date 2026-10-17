import { describe, expect, it } from "vitest";
import * as formulaAstBridge from "../../../../../src/formula/evaluation/ast-bridge.ts";

function createDeps() {
    return {
        normalizeFormulaAddress: (address) =>
            String(address || "")
                .replace(/\$/g, "")
                .trim()
                .toUpperCase(),
        parseSheetScopedDefinedNameReference: () => null,
        parseRangeAddress: (rawRange) => {
            const match = String(rawRange || "").match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);
            return match ? { start: match[1].toUpperCase(), end: match[2].toUpperCase() } : null;
        },
        parseCellAddress: (address) => {
            const match = String(address || "").match(/^([A-Z]+)(\d+)$/i);
            if (!match) return { row: 0, col: 0 };
            let col = 0;
            for (const ch of match[1].toUpperCase()) {
                col = col * 26 + (ch.charCodeAt(0) - 64);
            }
            return { row: Number(match[2]), col };
        },
    };
}

describe("xlsx-context formula ast", () => {
    it("coerces booleans, numbers, and text scalars", () => {
        const api = formulaAstBridge.createFormulaAstBridge(createDeps());

        expect(api.coerceFormulaAstScalar("TRUE")).toBe(true);
        expect(api.coerceFormulaAstScalar("1,024")).toBe(1024);
        expect(api.coerceFormulaAstScalar("hello")).toBe("hello");
    });

    it("builds a scalar matrix from range entries", () => {
        const api = formulaAstBridge.createFormulaAstBridge(createDeps());

        const matrix = api.createFormulaAstRangeMatrix("Sheet1", "A1", "B2", () => ({
            rawValues: ["1", "TRUE", "x", "4"],
            numericValues: [1, 4],
        }));

        expect(matrix).toEqual([
            [1, true],
            ["x", 4],
        ]);
    });
});
