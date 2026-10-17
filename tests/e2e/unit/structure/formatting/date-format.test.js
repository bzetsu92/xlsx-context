import { describe, expect, it } from "vitest";
import { formatDateByFormatCode } from "../../../../../src/structure/formatting/date-format.ts";
import {
    datePartsToExcelSerial,
    excelSerialToDateParts,
    formatCellDisplayValue,
    formatTextFunctionValue,
} from "../../../../../src/structure/formatting/cell-format.ts";

const render = (code, serial) => formatDateByFormatCode(code, serial, excelSerialToDateParts);

// 45366 is Friday, 2024-03-15. 0.6 is 14:24:00.
const DATE = 45366;
const NOON_ISH = 0.6;

describe("date format renderer", () => {
    it("prints the literals that the format code contains", () => {
        expect(render('yyyy"年"m"月"d"日"', DATE)).toBe("2024年3月15日");
        expect(render('yyyy"年"mm"月"', DATE)).toBe("2024年03月");
        expect(render('m"月"d"日"', DATE)).toBe("3月15日");
        expect(render("m月d日", DATE)).toBe("3月15日");
        expect(render('d "of" mmmm', DATE)).toBe("15 of March");
        expect(render("yyyy\\-mm", DATE)).toBe("2024-03");
    });

    it("renders common formats of other countries", () => {
        expect(render("yyyy-mm-dd", DATE)).toBe("2024-03-15");
        expect(render("yyyy/m/d", DATE)).toBe("2024/3/15");
        expect(render("yyyy/mm/dd", DATE)).toBe("2024/03/15");
        expect(render("dd/mm/yyyy", DATE)).toBe("15/03/2024");
        expect(render("mm/dd/yyyy", DATE)).toBe("03/15/2024");
        expect(render("d-mmm-yy", DATE)).toBe("15-Mar-24");
        expect(render("mmm d, yyyy", DATE)).toBe("Mar 15, 2024");
        expect(render("mmmm yyyy", DATE)).toBe("March 2024");
        expect(render("dddd, mmmm dd, yyyy", DATE)).toBe("Friday, March 15, 2024");
        expect(render("ddd", DATE)).toBe("Fri");
        expect(render("[$-409]mmmm d", DATE)).toBe("March 15");
        expect(render("[$-407]dd.mm.yyyy", DATE)).toBe("15.03.2024");
    });

    it("uses month and weekday names of the locale tag or of the workbook", () => {
        expect(render("[$-407]mmmm", DATE)).toBe("März");
        expect(render("[$-42A]dddd", DATE)).toBe("Thứ Sáu");
        expect(render("[$-40C]d mmmm yyyy", DATE)).toBe("15 mars 2024");
        expect(render("[$-410]dddd d mmmm", DATE)).toBe("venerdì 15 marzo");
        expect(formatDateByFormatCode("d mmmm yyyy", DATE, excelSerialToDateParts, "vi-VN")).toBe("15 Tháng 3 2024");
        expect(formatDateByFormatCode("[$-409]mmmm", DATE, excelSerialToDateParts, "vi-VN")).toBe("March");
        expect(render('上午/下午h"时"mm"分"', NOON_ISH)).toBe("下午2时24分");
        expect(render("오전/오후 h:mm", 0.25)).toBe("오전 6:00");
    });

    it("uses Japanese names for a Japanese locale tag and for aaa", () => {
        expect(render("[$-411]dddd", DATE)).toBe("金曜日");
        expect(render("[$-411]ddd", DATE)).toBe("金");
        expect(render("[$-ja-JP]mmmm", DATE)).toBe("3月");
        expect(render("aaa", DATE)).toBe("金");
        expect(render("aaaa", DATE)).toBe("金曜日");
    });

    it("renders Japanese eras by the exact start date", () => {
        const serial = (year, month, day) => datePartsToExcelSerial(year, month, day);
        const era = 'ggge"年"m"月"d"日"';

        expect(render(era, DATE)).toBe("令和6年3月15日");
        expect(render(`[$-411]${era}`, DATE)).toBe("令和6年3月15日");
        expect(render(`[$]${era};@`, DATE)).toBe("令和6年3月15日");
        expect(render(era, serial(2019, 4, 30))).toBe("平成31年4月30日");
        expect(render(era, serial(2019, 5, 1))).toBe("令和1年5月1日");
        expect(render(era, serial(1989, 1, 7))).toBe("昭和64年1月7日");
        expect(render(era, serial(1989, 1, 8))).toBe("平成1年1月8日");
        expect(render("gg", DATE)).toBe("令");
        expect(render("g", DATE)).toBe("R");
        expect(render("ggg ee", DATE)).toBe("令和 06");
    });

    it("renders time and tells month from minute", () => {
        expect(render("h:mm:ss", NOON_ISH)).toBe("14:24:00");
        expect(render("h:mm AM/PM", NOON_ISH)).toBe("2:24 PM");
        expect(render("hh:mm:ss am/pm", NOON_ISH)).toBe("02:24:00 pm");
        expect(render("h:mm A/P", NOON_ISH)).toBe("2:24 P");
        expect(render("mm:ss", NOON_ISH)).toBe("24:00");
        expect(render("m/d h:mm", DATE + NOON_ISH)).toBe("3/15 14:24");
        expect(render("yyyy-mm-dd hh:mm:ss", DATE + NOON_ISH)).toBe("2024-03-15 14:24:00");
        expect(render("hh:mm:ss.00", 0.5)).toBe("12:00:00.00");
    });

    it("renders elapsed time", () => {
        expect(render("[h]:mm:ss", 1.5)).toBe("36:00:00");
        expect(render("[mm]:ss", 0.5)).toBe("720:00");
        expect(render("[s]", 0.001)).toBe("86");
    });

    it("uses the first section, and the third section for zero", () => {
        expect(render("yyyy/m/d;@", DATE)).toBe("2024/3/15");
        expect(render('yyyy;"-";mm', DATE)).toBe("2024");
        expect(render('yyyy;"-";mm', 0)).toBe("12");
    });

    it("returns null for formats it cannot render exactly", () => {
        expect(render("General", DATE)).toBe(null);
        expect(render("0.00", DATE)).toBe(null);
        expect(render("[>100]yyyy", DATE)).toBe(null);
        expect(render("yyyy E", DATE)).toBe(null);
        expect(render("e", DATE)).toBe(null);
        expect(render("b", DATE)).toBe(null);
        expect(render("[DBNum1]yyyy", DATE)).toBe(null);
        expect(render("yyyy-mm-dd", -1)).toBe(null);
        expect(render('ggge"年"', datePartsToExcelSerial(1800, 1, 1))).toBe(null);
    });

    it("does not guess when the display depends on the computer of the reader", () => {
        expect(render("[$-F800]dddd, mmmm dd, yyyy", DATE)).toBe(null);
        expect(render("[$-F400]h:mm:ss\\ AM/PM", NOON_ISH)).toBe(null);
        // A workbook locale stands in for the reader computer.
        expect(formatDateByFormatCode("[$-F800]dddd, mmmm dd, yyyy", DATE, excelSerialToDateParts, "en-US")).toBe(
            "Friday, March 15, 2024"
        );
    });
});

describe("date formats in cells and in TEXT", () => {
    it("formats cells by the format code of the file", () => {
        const cell = (formatCode, raw = String(DATE)) => formatCellDisplayValue(raw, { numFmtId: 164, formatCode });

        expect(cell('yyyy"年"m"月"d"日"')).toBe("2024年3月15日");
        expect(cell("mmmm d, yyyy")).toBe("March 15, 2024");
        // The system time tag keeps the neutral 24-hour output.
        expect(cell("[$-F400]h:mm:ss\\ AM/PM", "0.52425925925925931")).toBe("12:34:56");
        // Formats the renderer rejects keep the previous fallback.
        expect(cell("[>100]yyyy")).toBe("2024-03-15");
    });

    it("formats TEXT() with the same renderer", () => {
        expect(formatTextFunctionValue(String(DATE), "mmm d, yyyy")).toBe("Mar 15, 2024");
        expect(formatTextFunctionValue(String(DATE), 'yyyy"年"m"月"')).toBe("2024年3月");
        expect(formatTextFunctionValue(String(DATE), "yyyy-mm-dd")).toBe("2024-03-15");
        expect(formatTextFunctionValue("0.5", "hh:mm:ss")).toBe("12:00:00");
        expect(formatTextFunctionValue("1234.5", "#,##0.00")).toBe("1,234.50");
        expect(formatTextFunctionValue("10", "0000")).toBe(null);
    });
});
