import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";

import { FIXTURE_ROOT, fixture } from "../../helpers/paths.js";

export function normalize(text) {
    return String(text || "")
        .replace(/\s+/g, " ")
        .trim();
}
export function containsAnswer(contextText, answer) {
    return normalize(contextText).includes(normalize(answer));
}
export function toNaiveCsv(workbook) {
    const parts = [];
    for (const sheet of workbook.sheets) {
        parts.push(`# ${sheet.name}`);
        const byRow = new Map();
        for (const cell of sheet.cells) {
            if (!byRow.has(cell.row)) byRow.set(cell.row, []);
            byRow.get(cell.row).push(cell);
        }
        for (const row of [...byRow.keys()].sort((a, b) => a - b)) {
            const cells = byRow.get(row).sort((a, b) => a.col - b.col);
            parts.push(cells.map((cell) => cell.outputValue).join(","));
        }
    }
    return parts.join("\n");
}
export function parseFixture(api, ...parts) {
    const fixturePath = fixture(...parts);
    const bytes = readFileSync(fixturePath);
    const arrayBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    return api.parseWorkbook(arrayBuffer, path.basename(fixturePath));
}
export function listFixtures(dir = FIXTURE_ROOT) {
    const out = [];
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            out.push(...listFixtures(full));
        } else if (entry.name.endsWith(".xlsx")) {
            out.push(full);
        }
    }
    return out.sort();
}
