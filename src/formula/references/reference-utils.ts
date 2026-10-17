type RangeRef = { sheetName: string; start: string; end: string };

type FormulaReferenceUtilsDeps = {
    normalizeFormulaAddress: (address: string) => string;
    parseRangeAddress: (rawRange: string) => { start: string; end: string } | null;
    // Reads workbook resolvers through getters. Core installs the resolvers after creation.
    getDefinedNameRangeRef: () => ((sheetName: string, name: string) => RangeRef | null) | null;
    getStructuredRangeRef: () => ((sheetName: string, text: string) => RangeRef | null) | null;
};

function createFormulaReferenceUtils(deps: FormulaReferenceUtilsDeps) {
    /** Parses `=Sheet!A1` into sheet and address. Returns null for anything else. */
    function parseSimpleFormulaReference(
        formulaText: string,
        currentSheetName: string
    ): { sheetName: string; address: string } | null {
        const normalizedFormula = String(formulaText || "")
            .trim()
            .replace(/^=/, "");
        const quotedSheetMatch = normalizedFormula.match(/^'((?:[^']|'')+)'!(\$?[A-Z]+\$?\d+)$/i);
        if (quotedSheetMatch) {
            return {
                sheetName: quotedSheetMatch[1].replace(/''/g, "'"),
                address: deps.normalizeFormulaAddress(quotedSheetMatch[2]),
            };
        }
        const sheetMatch = normalizedFormula.match(/^([^'=][^!]*)!(\$?[A-Z]+\$?\d+)$/i);
        if (sheetMatch) {
            return {
                sheetName: sheetMatch[1],
                address: deps.normalizeFormulaAddress(sheetMatch[2]),
            };
        }
        const localMatch = normalizedFormula.match(/^(\$?[A-Z]+\$?\d+)$/i);
        if (localMatch) {
            return {
                sheetName: currentSheetName,
                address: deps.normalizeFormulaAddress(localMatch[1]),
            };
        }
        return null;
    }

    function normalizeFormulaSheetName(rawName: string): string {
        return String(rawName || "")
            .replace(/^'/, "")
            .replace(/'$/, "")
            .replace(/''/g, "'");
    }

    function normalizeDefinedNameKey(name: string): string {
        return String(name || "")
            .trim()
            .toUpperCase();
    }

    /** Parses `Sheet!Name` references. Rejects cell addresses in the name position. */
    function parseSheetScopedDefinedNameReference(
        expression: string,
        currentSheetName: string
    ): { sheetName: string; name: string } | null {
        const normalizedExpression = String(expression || "").trim();
        const quotedSheetMatch = normalizedExpression.match(/^'((?:[^']|'')+)'!([A-Za-z_][A-Za-z0-9_.]*)$/);
        if (quotedSheetMatch) {
            return {
                sheetName: normalizeFormulaSheetName(quotedSheetMatch[1].replace(/''/g, "'")),
                name: quotedSheetMatch[2],
            };
        }
        const sheetMatch = normalizedExpression.match(/^([^'=][^!]*)!([A-Za-z_][A-Za-z0-9_.]*)$/);
        if (!sheetMatch) {
            return null;
        }
        if (/^\$?[A-Z]+\$?\d+$/i.test(sheetMatch[2])) {
            return null;
        }
        return {
            sheetName: normalizeFormulaSheetName(sheetMatch[1] || currentSheetName),
            name: sheetMatch[2],
        };
    }

    /** Resolves `A1:B2`, `Sheet!A1:B2`, a defined name or a structured reference to a range. */
    /** Parses `Sheet!A1:B2`, local ranges, and defined names into range refs. Returns null for anything else. */
    function parseQualifiedRangeReference(argText: string, currentSheetName: string): RangeRef | null {
        const qualifiedRangeMatch = argText.match(
            /^(?:'((?:[^']|'')+)'|([^'=][^!]*))!(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)$/i
        );
        const localRangeMatch = argText.match(/^(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?\d+)$/i);
        if (!qualifiedRangeMatch && !localRangeMatch) {
            const trimmed = String(argText || "").trim();
            const scopedDefinedName = parseSheetScopedDefinedNameReference(trimmed, currentSheetName);
            if (scopedDefinedName) {
                const scopedRange =
                    deps.getDefinedNameRangeRef()?.(scopedDefinedName.sheetName, scopedDefinedName.name) || null;
                if (scopedRange) return scopedRange;
            }
            const structuredRange = deps.getStructuredRangeRef()?.(currentSheetName, trimmed) || null;
            if (structuredRange) return structuredRange;
            const definedRange = deps.getDefinedNameRangeRef()?.(currentSheetName, trimmed) || null;
            if (definedRange) return definedRange;
            return null;
        }
        const sheetName = qualifiedRangeMatch
            ? normalizeFormulaSheetName(qualifiedRangeMatch[1] || qualifiedRangeMatch[2] || currentSheetName)
            : currentSheetName;
        const rangeText = String(qualifiedRangeMatch ? qualifiedRangeMatch[3] : localRangeMatch?.[1] || "");
        const range = deps.parseRangeAddress(rangeText);
        if (!range) return null;
        return { sheetName, start: range.start, end: range.end };
    }

    return {
        parseSimpleFormulaReference,
        parseSheetScopedDefinedNameReference,
        parseQualifiedRangeReference,
        normalizeFormulaSheetName,
        normalizeDefinedNameKey,
    };
}

/** Parses cell, range, and defined-name references in formula text. */
export { createFormulaReferenceUtils };
