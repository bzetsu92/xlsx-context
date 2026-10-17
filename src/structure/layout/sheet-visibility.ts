import type { ParsedSheet } from "../../model/index.js";
import { colToLetters } from "../addresses/address-utils.js";

function compactRuns(values: number[], label: (value: number) => string): string {
    const sorted = [...new Set(values)].sort((left, right) => left - right);
    const runs: string[] = [];
    for (let index = 0; index < sorted.length; index += 1) {
        const start = sorted[index];
        while (sorted[index + 1] === sorted[index] + 1) index += 1;
        runs.push(start === sorted[index] ? label(start) : `${label(start)}-${label(sorted[index])}`);
    }
    return runs.join(", ");
}

/** One line that says which parts of the sheet Excel hides. Empty when nothing is hidden. */
function describeHiddenContent(sheet: Pick<ParsedSheet, "state" | "hiddenRows" | "hiddenCols">): string {
    const parts: string[] = [];
    if (sheet.state === "hidden") parts.push("Hidden sheet.");
    if (sheet.state === "veryHidden") parts.push("Very hidden sheet.");
    if (sheet.hiddenRows?.length) parts.push(`Hidden rows: ${compactRuns(sheet.hiddenRows, String)}.`);
    if (sheet.hiddenCols?.length) parts.push(`Hidden columns: ${compactRuns(sheet.hiddenCols, colToLetters)}.`);
    return parts.join(" ");
}

export { describeHiddenContent };
