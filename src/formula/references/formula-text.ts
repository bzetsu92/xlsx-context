/*
 * Scans raw formula text with quote and parenthesis awareness.
 * Serves places that inspect strings without an AST, such as `A1:INDEX(...)` ranges.
 */

/** Index of `operator` outside quotes and parentheses, or -1. */
function findTopLevelOperatorIndex(expression: string, operator: string): number {
    const target = String(operator || "");
    if (!target) return -1;
    let depth = 0;
    let inSingleQuote = false;
    let inDoubleQuote = false;
    for (let i = 0; i <= expression.length - target.length; i += 1) {
        const ch = expression[i];
        if (ch === "'" && !inDoubleQuote) {
            inSingleQuote = !inSingleQuote;
            continue;
        }
        if (ch === '"' && !inSingleQuote) {
            inDoubleQuote = !inDoubleQuote;
            continue;
        }
        if (inSingleQuote || inDoubleQuote) continue;
        if (ch === "(") {
            depth += 1;
            continue;
        }
        if (ch === ")") {
            depth = Math.max(0, depth - 1);
            continue;
        }
        if (depth === 0 && expression.slice(i, i + target.length) === target) {
            return i;
        }
    }
    return -1;
}

/** Splits `NAME(args)` when the whole expression is one call to an allowed function. */
function parseWholeFunctionCall(expression: string, allowedNames: string[]): { name: string; argsText: string } | null {
    const trimmed = String(expression || "").trim();
    const nameMatch = trimmed.match(/^([A-Z][A-Z0-9]*)\(/i);
    if (!nameMatch) return null;
    const name = nameMatch[1].toUpperCase();
    if (!allowedNames.includes(name)) return null;
    let depth = 0;
    let inSingleQuote = false;
    let inDoubleQuote = false;
    for (let i = name.length; i < trimmed.length; i += 1) {
        const ch = trimmed[i];
        if (ch === "'" && !inDoubleQuote) {
            inSingleQuote = !inSingleQuote;
            continue;
        }
        if (ch === '"' && !inSingleQuote) {
            inDoubleQuote = !inDoubleQuote;
            continue;
        }
        if (inSingleQuote || inDoubleQuote) continue;
        if (ch === "(") {
            depth += 1;
            continue;
        }
        if (ch !== ")") continue;
        depth -= 1;
        if (depth > 0) continue;
        if (depth < 0 || i !== trimmed.length - 1) return null;
        return {
            name,
            argsText: trimmed.slice(name.length + 1, i),
        };
    }
    return null;
}

/** Splits a function argument list on top-level commas, keeping quoted and nested commas. */
function splitFormulaArguments(argText: string): string[] {
    const args: string[] = [];
    let current = "";
    let depth = 0;
    let inSingleQuote = false;
    let inDoubleQuote = false;
    for (let i = 0; i < argText.length; i += 1) {
        const ch = argText[i];
        if (ch === "'" && !inDoubleQuote) {
            inSingleQuote = !inSingleQuote;
            current += ch;
            continue;
        }
        if (ch === '"' && !inSingleQuote) {
            inDoubleQuote = !inDoubleQuote;
            current += ch;
            continue;
        }
        if (!inSingleQuote && !inDoubleQuote) {
            if (ch === "(") {
                depth += 1;
            } else if (ch === ")") {
                depth = Math.max(0, depth - 1);
            } else if (ch === "," && depth === 0) {
                args.push(current.trim());
                current = "";
                continue;
            }
        }
        current += ch;
    }
    if (current.trim()) args.push(current.trim());
    return args;
}

export { findTopLevelOperatorIndex, parseWholeFunctionCall, splitFormulaArguments };
