import type { FormulaResolutionSource } from "../../model/index.js";
import type { FormulaEvaluationOptions } from "./evaluator.js";

type FormulaEngineDeps = {
    getDefinedNameScalarValue: () => ((sheetName: string, name: string) => string | null) | null;
    tryResolveFormulaExpressionWithAst: (
        expression: string,
        currentSheetName: string,
        resolveCellValue: (sheetName: string, address: string) => string,
        resolveRangeEntries?: (
            sheetName: string,
            rangeText: string
        ) => { rawValues: string[]; numericValues: number[] },
        currentAddress?: string,
        options?: FormulaEvaluationOptions
    ) => string | null;
    parseSimpleFormulaReference: (
        formulaText: string,
        currentSheetName: string
    ) => { sheetName: string; address: string } | null;
    parseSheetScopedDefinedNameReference: (
        formulaText: string,
        currentSheetName: string
    ) => { sheetName: string; name: string } | null;
};

function createFormulaEngine(deps: FormulaEngineDeps) {
    /** Resolves one formula and reports which resolver produced the value. */
    function tryResolveFormulaExpressionDetailed(
        formulaText: string,
        currentSheetName: string,
        resolveCellValue: (sheetName: string, address: string) => string,
        // Unused. The AST evaluator reads ranges through resolveRangeEntries. The parameter stays for positional callers.
        _resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
        resolveRangeEntries?: (
            sheetName: string,
            rangeText: string
        ) => { rawValues: string[]; numericValues: number[] },
        currentAddress?: string,
        options?: FormulaEvaluationOptions
    ): { value: string; source: FormulaResolutionSource } | null {
        const normalized = String(formulaText || "")
            .trim()
            .replace(/^=/, "");
        if (!normalized) return null;
        const directDefinedNameValue = deps.getDefinedNameScalarValue()?.(currentSheetName, normalized) || null;
        if (directDefinedNameValue != null) {
            return {
                value: directDefinedNameValue,
                source: "reference_lookup",
            };
        }
        const astResolved = deps.tryResolveFormulaExpressionWithAst(
            normalized,
            currentSheetName,
            resolveCellValue,
            resolveRangeEntries,
            currentAddress,
            options
        );
        if (astResolved != null) {
            return {
                value: astResolved,
                source: "ast_evaluator",
            };
        }
        return null;
    }

    function tryResolveFormulaExpression(
        formulaText: string,
        currentSheetName: string,
        resolveCellValue: (sheetName: string, address: string) => string,
        resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
        resolveRangeEntries?: (
            sheetName: string,
            rangeText: string
        ) => { rawValues: string[]; numericValues: number[] },
        currentAddress?: string,
        options?: FormulaEvaluationOptions
    ): string | null {
        return (
            tryResolveFormulaExpressionDetailed(
                formulaText,
                currentSheetName,
                resolveCellValue,
                resolveRangeValues,
                resolveRangeEntries,
                currentAddress,
                options
            )?.value ?? null
        );
    }

    /** Resolves one formula operand: literal, cell reference, defined name, or a nested expression. */
    /** Resolves one scalar operand: quoted strings first, then the AST evaluator. */
    function resolveScalarFormulaValue(
        expression: string,
        currentSheetName: string,
        resolveCellValue: (sheetName: string, address: string) => string,
        resolveRangeValues?: (sheetName: string, rangeText: string) => number[],
        resolveRangeEntries?: (
            sheetName: string,
            rangeText: string
        ) => { rawValues: string[]; numericValues: number[] },
        options?: FormulaEvaluationOptions
    ): string | null {
        const trimmed = String(expression || "").trim();
        if (!trimmed) return null;
        const quotedString = trimmed.match(/^"(.*)"$/);
        if (quotedString) {
            return quotedString[1].replace(/""/g, '"');
        }
        const numeric = Number(trimmed);
        if (!Number.isNaN(numeric)) {
            return String(numeric);
        }
        const simpleRef = deps.parseSimpleFormulaReference(`=${trimmed}`, currentSheetName);
        if (simpleRef) {
            return resolveCellValue(simpleRef.sheetName, simpleRef.address);
        }
        const scopedDefinedNameRef = deps.parseSheetScopedDefinedNameReference(trimmed, currentSheetName);
        if (scopedDefinedNameRef) {
            const scopedValue =
                deps.getDefinedNameScalarValue()?.(scopedDefinedNameRef.sheetName, scopedDefinedNameRef.name) || null;
            if (scopedValue != null) return scopedValue;
        }
        const definedNameValue = deps.getDefinedNameScalarValue()?.(currentSheetName, trimmed) || null;
        if (definedNameValue != null) return definedNameValue;
        if (/^(TRUE|FALSE)$/i.test(trimmed)) {
            return trimmed.toUpperCase();
        }
        return tryResolveFormulaExpression(
            `=${trimmed}`,
            currentSheetName,
            resolveCellValue,
            resolveRangeValues,
            resolveRangeEntries,
            undefined,
            options
        );
    }

    return {
        tryResolveFormulaExpressionDetailed,
        tryResolveFormulaExpression,
        resolveScalarFormulaValue,
    };
}

/** Resolves formula text. Tries direct defined-name values before AST evaluation. */
export { createFormulaEngine };
