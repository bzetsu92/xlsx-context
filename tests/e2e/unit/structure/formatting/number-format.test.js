import { describe, expect, it } from "vitest";
import { formatNumberByFormatCode } from "../../../../../src/structure/formatting/number-format.ts";

const fmt = formatNumberByFormatCode;
const ACCOUNTING_USD = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';

describe("formatNumberByFormatCode()", () => {
    it("keeps currency symbols and units of any locale", () => {
        expect(fmt('"$"#,##0.00', 1234.5)).toBe("$1,234.50");
        expect(fmt('"$"#,##0.00', -1234.5)).toBe("-$1,234.50");
        expect(fmt("[$€-407] #,##0.00", 1234.5)).toBe("€ 1,234.50");
        expect(fmt("#,##0.00 [$₫-42A]", 1234.5)).toBe("1,234.50 ₫");
        expect(fmt('#,##0 "VND"', 1234.5)).toBe("1,235 VND");
        expect(fmt('0.0" kg"', 3.14)).toBe("3.1 kg");
        expect(fmt("¥#,##0", 1500)).toBe("¥1,500");
    });

    it("picks the negative, zero, and conditional sections", () => {
        expect(fmt("#,##0;(#,##0)", -5)).toBe("(5)");
        expect(fmt('#,##0;-#,##0;"-"', 0)).toBe("-");
        expect(fmt("0;-0;;@", 0)).toBe("");
        expect(fmt(ACCOUNTING_USD, 1234.5)).toBe("$ 1,234.50");
        expect(fmt(ACCOUNTING_USD, -5)).toBe("$ (5.00)");
        expect(fmt(ACCOUNTING_USD, 0)).toBe("$ -");
        expect(fmt('[>=1000]#,##0,"K";0', 1500)).toBe("2K");
        expect(fmt('[>=1000]#,##0,"K";0', 999)).toBe("999");
        expect(fmt("[Red]0.00;[Blue]-0.00", -2)).toBe("-2.00");
    });

    it("applies placeholders, scaling, and percent", () => {
        expect(fmt('#,##0,,"M"', 2500000)).toBe("3M");
        expect(fmt("0%", 0.256)).toBe("26%");
        expect(fmt("0.0#", 1.5)).toBe("1.5");
        expect(fmt("0.0#", 1.234)).toBe("1.23");
        expect(fmt("#.##", 0.5)).toBe(".5");
        expect(fmt("00000", 42)).toBe("00042");
        expect(fmt("000-0000", 1234567)).toBe("123-4567");
        expect(fmt('"No. "General', 7)).toBe("No. 7");
    });

    it("returns null for codes it cannot render", () => {
        expect(fmt("[DBNum1]0", 5)).toBe(null);
        expect(fmt("0", Number.NaN)).toBe(null);
    });
});
