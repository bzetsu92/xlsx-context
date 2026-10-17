import { describe, expect, it } from "vitest";
import { isIsoDateToken, isWeekdayToken } from "../../../../../src/structure/formatting/date-tokens.ts";
import * as narrativeStructure from "../../../../../src/outputs/markdown/narrative-structure.ts";

describe("date tokens", () => {
    it("recognizes weekday names in Japanese and English", () => {
        for (const name of ["日", "土", "月曜日", "Sun", "sat", " Monday ", "FRIDAY"]) {
            expect(isWeekdayToken(name), name).toBe(true);
        }
        for (const name of ["", "Mo", "Sunday!", "月曜", "Summer", "2024-03-15"]) {
            expect(isWeekdayToken(name), name).toBe(false);
        }
    });

    it("recognizes ISO dates only", () => {
        expect(isIsoDateToken("2024-03-15")).toBe(true);
        expect(isIsoDateToken(" 2024-03-15 ")).toBe(true);
        expect(isIsoDateToken("2024/03/15")).toBe(false);
        expect(isIsoDateToken("2024-3-15")).toBe(false);
    });
});

describe("calendar detection uses the shared weekday names", () => {
    const item = (row, values) => ({ row, startCol: 2, endCol: 8, text: values.join(" "), cellValues: values });
    const week = (names) => [names, names.map((_, index) => String(index + 1))];
    const block = (rows) => ({
        startRow: 1,
        startCol: 2,
        endRow: rows.length,
        endCol: 8,
        lines: rows.map((row) => row.join(" ")),
        items: rows.map((row, index) => item(index + 1, row)),
    });

    it("renders a weekday header row as a heading in Japanese and in English", () => {
        const api = narrativeStructure;
        const japanese = block(week(["日", "月", "火", "水", "木", "金", "土"]));
        const english = block(week(["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]));

        expect(api.renderNarrativeBlock(japanese).split("\n")[0]).toBe("**日 月 火 水 木 金 土**");
        expect(api.renderNarrativeBlock(english).split("\n")[0]).toBe("**Sun Mon Tue Wed Thu Fri Sat**");
    });
});
