import { describe, expect, it } from "vitest";
import * as formulaResolver from "../../../../../src/formula/references/resolver.ts";

function createDeps(overrides = {}) {
    return {
        normalizeStructuredTableKey: (value) =>
            String(value || "")
                .normalize("NFKC")
                .trim()
                .toUpperCase(),
        normalizeFormulaSheetName: (rawName) =>
            String(rawName || "")
                .replace(/^'/, "")
                .replace(/'$/, "")
                .replace(/''/g, "'"),
        normalizeDefinedNameKey: (name) =>
            String(name || "")
                .trim()
                .toUpperCase(),
        normalizeFormulaAddress: (address) =>
            String(address || "")
                .trim()
                .replace(/\$/g, "")
                .toUpperCase(),
        parseSimpleFormulaReference: (formulaText, currentSheetName) => {
            const normalized = String(formulaText || "")
                .trim()
                .replace(/^=/, "");
            const local = normalized.match(/^([A-Z]+\d+)$/i);
            return local ? { sheetName: currentSheetName, address: local[1].toUpperCase() } : null;
        },
        resolveScalarFormulaValue: () => null,
        parseQualifiedRangeReference: () => null,
        findTopLevelOperatorIndex: () => -1,
        parseWholeFunctionCall: () => null,
        splitFormulaArguments: () => [],
        parseCellAddress: (address) => {
            const match = String(address || "").match(/^([A-Z]+)(\d+)$/i);
            if (!match) return { row: 0, col: 0 };
            let col = 0;
            for (const ch of match[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
            return { row: Number(match[2]), col };
        },
        colToLetters: (col) => {
            let current = col;
            let result = "";
            while (current > 0) {
                const remainder = (current - 1) % 26;
                result = String.fromCharCode(65 + remainder) + result;
                current = Math.floor((current - 1) / 26);
            }
            return result;
        },
        parseRangeAddress: () => null,
        tryResolveFormulaExpressionDetailed: () => null,
        applyResolvedFormulaValue: (cell, value, source = "reference_lookup") => {
            cell.rawValue = value;
            cell.outputValue = value;
            cell.resolutionStatus = "resolved";
            cell.resolutionSource = source;
        },
        setDefinedNameResolvers: () => {},
        ...overrides,
    };
}

describe("xlsx-context formula resolver", () => {
    it("resolves direct same-sheet references", () => {
        const api = formulaResolver;
        const workbook = {
            sheets: [
                {
                    name: "Sheet1",
                    tables: [],
                    cells: [
                        {
                            address: "A1",
                            rawValue: "1",
                            outputValue: "1",
                            formulaText: "",
                            resolutionStatus: null,
                            resolutionSource: null,
                            valueType: "n",
                            formulaType: "",
                            spillRef: "",
                        },
                        {
                            address: "B1",
                            rawValue: "",
                            outputValue: "",
                            formulaText: "=A1",
                            resolutionStatus: null,
                            resolutionSource: null,
                            valueType: "formula",
                            formulaType: "",
                            spillRef: "",
                        },
                    ],
                },
            ],
            definedNames: [],
        };

        api.resolveSimpleFormulaReferences(workbook, createDeps());

        expect(workbook.sheets[0].cells[1].outputValue).toBe("1");
        expect(workbook.sheets[0].cells[1].resolutionStatus).toBe("resolved");
    });

    it("builds structured table ranges from table metadata", () => {
        const api = formulaResolver;
        const workbook = {
            sheets: [
                {
                    name: "Sheet1",
                    tables: [
                        {
                            sheetName: "Sheet1",
                            name: "Sales",
                            displayName: "Sales",
                            start: "B2",
                            end: "D5",
                            columns: ["Code", "Name", "Amount"],
                            headerRowCount: 1,
                            totalsRowCount: 1,
                        },
                    ],
                    cells: [],
                },
            ],
            definedNames: [],
        };

        const resolver = api.buildFormulaResolver(workbook, createDeps());

        expect(resolver.resolveStructuredRange("Sheet1", "Sales[Amount]")).toEqual({
            sheetName: "Sheet1",
            start: "D3",
            end: "D4",
        });
    });

    it("exposes defined-name resolvers while resolving and clears them afterward", () => {
        const api = formulaResolver;
        const seen = [];
        const workbook = {
            sheets: [{ name: "Sheet1", tables: [], cells: [] }],
            definedNames: [{ name: "Answer", formulaText: "=42", localSheetName: null }],
        };

        api.resolveSimpleFormulaReferences(
            workbook,
            createDeps({
                setDefinedNameResolvers: (scalar, range, structured) => {
                    seen.push([!!scalar, !!range, !!structured]);
                },
            })
        );

        expect(seen[0]).toEqual([true, true, true]);
        expect(seen[seen.length - 1]).toEqual([false, false, false]);
    });
});
