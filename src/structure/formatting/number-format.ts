/** Renders a number with an Excel number format code, as the file stores it. */
import { splitSections } from "./date-format.js";

type Token =
    | { kind: "literal"; text: string }
    | { kind: "fill" }
    | { kind: "digit"; char: "0" | "#" | "?" }
    | { kind: "comma" }
    | { kind: "point" }
    | { kind: "general" };

type Condition = { op: string; value: number };

const COLOR = /^(black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)$/i;
const CONDITION = /^(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d+)?)$/;

function readCondition(section: string): Condition | null {
    for (const match of section.matchAll(/\[([^\]]*)\]/g)) {
        const condition = match[1].trim().match(CONDITION);
        if (condition) return { op: condition[1], value: Number(condition[2]) };
    }
    return null;
}

function matches(condition: Condition, value: number): boolean {
    switch (condition.op) {
        case "<":
            return value < condition.value;
        case "<=":
            return value <= condition.value;
        case ">":
            return value > condition.value;
        case ">=":
            return value >= condition.value;
        case "=":
            return value === condition.value;
        default:
            return value !== condition.value;
    }
}

/** Splits one section into tokens. Returns null for parts this renderer does not support, such as [DBNum1]. */
function tokenize(section: string): { tokens: Token[]; percent: number } | null {
    const tokens: Token[] = [];
    let percent = 0;
    for (let index = 0; index < section.length; index += 1) {
        const char = section[index];
        if (char === '"') {
            const end = section.indexOf('"', index + 1);
            if (end < 0) return null;
            tokens.push({ kind: "literal", text: section.slice(index + 1, end) });
            index = end;
        } else if (char === "\\") {
            if (index + 1 < section.length) tokens.push({ kind: "literal", text: section[index + 1] });
            index += 1;
        } else if (char === "[") {
            const end = section.indexOf("]", index + 1);
            if (end < 0) return null;
            const content = section.slice(index + 1, end).trim();
            const tag = content.match(/^\$([^-]*)(?:-.*)?$/);
            if (tag) {
                if (tag[1]) tokens.push({ kind: "literal", text: tag[1] });
            } else if (!COLOR.test(content) && !CONDITION.test(content)) {
                return null;
            }
            index = end;
        } else if (char === "_") {
            index += 1;
        } else if (char === "*") {
            tokens.push({ kind: "fill" });
            index += 1;
        } else if (char === "0" || char === "#" || char === "?") {
            tokens.push({ kind: "digit", char });
        } else if (char === ",") {
            tokens.push({ kind: "comma" });
        } else if (char === ".") {
            tokens.push({ kind: "point" });
        } else if (char === "%") {
            percent += 1;
            tokens.push({ kind: "literal", text: "%" });
        } else if (char === "@") {
            continue;
        } else if (section.slice(index, index + 7).toLowerCase() === "general") {
            tokens.push({ kind: "general" });
            index += 6;
        } else {
            tokens.push({ kind: "literal", text: char });
        }
    }
    return { tokens, percent };
}

function generalText(value: number): string {
    if (Number.isInteger(value) && Math.abs(value) < 1e11) return String(value);
    return String(Number(value.toPrecision(11)));
}

function groupThousands(digits: string): string {
    return digits.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

function renderSection(section: string, value: number, negative: boolean): string | null {
    const parsed = tokenize(section);
    if (!parsed) return null;
    const { tokens } = parsed;
    const digitIndexes = tokens.flatMap((token, index) => (token.kind === "digit" ? [index] : []));
    const firstDigit = digitIndexes[0] ?? -1;
    const lastDigit = digitIndexes[digitIndexes.length - 1] ?? -1;
    let pointIndex = -1;
    if (firstDigit >= 0) {
        for (let index = Math.max(0, firstDigit - 1); index <= lastDigit; index += 1) {
            if (tokens[index].kind === "point") {
                pointIndex = index;
                break;
            }
        }
    }
    const intEnd = pointIndex >= 0 ? pointIndex : lastDigit + 1;
    const intDigits = digitIndexes.filter((index) => index < intEnd);
    const fracDigits = digitIndexes.filter((index) => index > pointIndex && pointIndex >= 0);
    const lastIntDigit = intDigits[intDigits.length - 1] ?? -1;

    let grouping = false;
    let scaleCommas = 0;
    const structural = new Set<number>();
    for (let index = firstDigit; index >= 0 && index < intEnd; index += 1) {
        if (tokens[index].kind !== "comma") continue;
        structural.add(index);
        if (index < lastIntDigit) grouping = true;
    }
    for (let index = lastIntDigit + 1; lastIntDigit >= 0 && tokens[index]?.kind === "comma"; index += 1) {
        structural.add(index);
        scaleCommas += 1;
    }
    if (pointIndex >= 0) structural.add(pointIndex);

    const scaled = (value * 100 ** parsed.percent) / 1000 ** scaleCommas;
    const fracCount = fracDigits.length;
    const fixed = scaled.toFixed(fracCount);
    const [intPartText, fracPartText = ""] = fixed.split(".");

    let fracText = fracPartText;
    for (let position = fracCount - 1; position >= 0; position -= 1) {
        const placeholder = (tokens[fracDigits[position]] as { char: string }).char;
        if (fracText[position] !== "0" || placeholder === "0") break;
        fracText = fracText.slice(0, position) + (placeholder === "?" ? " " : "") + fracText.slice(position + 1);
    }

    const embedded = tokens.some(
        (token, index) =>
            index > firstDigit && index < lastIntDigit && token.kind === "literal" && !structural.has(index)
    );
    const digits = intPartText === "0" ? "" : intPartText;
    const intText = new Map<number, string>();
    if (embedded) {
        let remaining = digits;
        for (let position = intDigits.length - 1; position >= 0; position -= 1) {
            const placeholder = (tokens[intDigits[position]] as { char: string }).char;
            const take = position === 0 ? remaining : remaining.slice(-1);
            remaining = position === 0 ? "" : remaining.slice(0, -1);
            intText.set(intDigits[position], take || (placeholder === "0" ? "0" : placeholder === "?" ? " " : ""));
        }
    } else if (intDigits.length > 0) {
        let text = digits;
        const zeros = intDigits.filter((index) => (tokens[index] as { char: string }).char === "0").length;
        text = text.padStart(zeros, "0");
        if (grouping) text = groupThousands(text);
        intText.set(intDigits[0], text);
    }

    let output = "";
    tokens.forEach((token, index) => {
        if (token.kind === "literal") {
            output += token.text;
        } else if (token.kind === "fill") {
            if (!/\s$/u.test(output)) output += " ";
        } else if (token.kind === "general") {
            output += generalText(scaled);
        } else if (token.kind === "digit") {
            if (intText.has(index)) {
                output += intText.get(index);
            } else if (index > pointIndex && pointIndex >= 0 && index === fracDigits[0]) {
                output += fracText;
            }
        } else if (token.kind === "point") {
            output += ".";
        } else if (token.kind === "comma" && !structural.has(index)) {
            output += ",";
        }
    });
    const text = output.trim();
    return negative && text !== "" ? `-${text}` : text;
}

/**
 * Formats a number by an Excel format code: sections, conditions, literals, currency tags, padding, grouping,
 * scaling by trailing commas, and percent. Returns null for codes it cannot render.
 */
function formatNumberByFormatCode(formatCode: string, value: number): string | null {
    if (!Number.isFinite(value)) return null;
    const sections = splitSections(String(formatCode || ""));
    const numeric = sections.slice(0, 3);
    if (numeric.some((section) => readCondition(section))) {
        for (const section of numeric) {
            const condition = readCondition(section);
            if (!condition || matches(condition, value)) {
                return renderSection(section, Math.abs(value), value < 0 && !condition);
            }
        }
        return null;
    }
    if (value < 0 && numeric.length >= 2) return renderSection(numeric[1], -value, false);
    if (value === 0 && numeric.length >= 3) return renderSection(numeric[2], 0, false);
    return renderSection(numeric[0], Math.abs(value), value < 0);
}

export { formatNumberByFormatCode };
