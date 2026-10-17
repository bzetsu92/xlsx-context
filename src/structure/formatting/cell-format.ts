import type { BorderFlags, CellStyleInfo, FormulaResolutionSource } from "../../model/index.js";
import { formatDateByFormatCode } from "./date-format.js";
import { formatNumberByFormatCode } from "./number-format.js";

/** Workbook settings that change how a stored number reads. */
type FormatContext = {
    /** Locale that the file shows. Null when the file does not say. */
    locale?: string | null;
    /** True when the workbook counts dates from 1904-01-01. */
    date1904?: boolean;
};

/** Days between the 1900 and 1904 date systems. */
const DATE_1904_OFFSET = 1462;

type ResolvedCellLike = {
    borders: BorderFlags;
    numFmtId: number;
    formatCode: string;
    rawValue: string;
    outputValue: string;
    resolutionStatus: "resolved" | "fallback_formula" | "unsupported_external" | null;
    resolutionSource: FormulaResolutionSource;
};

/** Detects date codes after removing literals. Excludes General. */
function isDateFormatCode(formatCode: string): boolean {
    const normalized = String(formatCode || "")
        .toLowerCase()
        .replace(/\[[^\]]*]/g, "")
        .replace(/"[^"]*"/g, "")
        .replace(/\\./g, "");
    if (!normalized) return false;
    if (normalized.includes("general")) return false;
    return /[ymdhs]/.test(normalized);
}

function normalizeNumericFormatCode(formatCode: string): string {
    return String(formatCode || "")
        .trim()
        .replace(/\[[^\]]*]/g, "")
        .replace(/"([^"]*)"/g, "$1")
        .replace(/\\(.)/g, "$1")
        .replace(/_.?/g, "")
        .replace(/\*/g, "");
}

/** Converts serials to ISO text. Omits midnight time. */
function excelSerialToIsoText(serial: number): string {
    if (!Number.isFinite(serial)) return String(serial);
    const wholeDays = Math.floor(serial);
    const fractional = serial - wholeDays;
    const utcDays = wholeDays > 59 ? wholeDays - 1 : wholeDays;
    const baseUtcMs = Date.UTC(1899, 11, 31);
    const msPerDay = 24 * 60 * 60 * 1000;
    const date = new Date(baseUtcMs + utcDays * msPerDay + Math.round(fractional * msPerDay));
    const yyyy = String(date.getUTCFullYear()).padStart(4, "0");
    const mm = String(date.getUTCMonth() + 1).padStart(2, "0");
    const dd = String(date.getUTCDate()).padStart(2, "0");
    const hh = String(date.getUTCHours()).padStart(2, "0");
    const mi = String(date.getUTCMinutes()).padStart(2, "0");
    const ss = String(date.getUTCSeconds()).padStart(2, "0");
    if (hh === "00" && mi === "00" && ss === "00") {
        return `${yyyy}-${mm}-${dd}`;
    }
    return `${yyyy}-${mm}-${dd} ${hh}:${mi}:${ss}`;
}

/** Splits serials into padded date parts. Returns null for non-finite input. */
function excelSerialToDateParts(serial: number): {
    year: number;
    month: number;
    day: number;
    hour: number;
    minute: number;
    second: number;
    yyyy: string;
    mm: string;
    dd: string;
    hh: string;
    mi: string;
    ss: string;
} | null {
    if (!Number.isFinite(serial)) return null;
    const wholeDays = Math.floor(serial);
    const fractional = serial - wholeDays;
    const excelEpochOffsetDays = 25569;
    const msPerDay = 24 * 60 * 60 * 1000;
    // Excel counts 1900-02-29, which never existed, as serial 60. Serials before it are one day later.
    const utcDays = (wholeDays < 60 ? wholeDays + 1 : wholeDays) - excelEpochOffsetDays;
    const baseUtcMs = Date.UTC(1970, 0, 1);
    const date = new Date(baseUtcMs + utcDays * msPerDay + Math.round(fractional * msPerDay));
    const leapDay = wholeDays === 60;
    return {
        year: date.getUTCFullYear(),
        month: leapDay ? 2 : date.getUTCMonth() + 1,
        day: leapDay ? 29 : date.getUTCDate(),
        hour: date.getUTCHours(),
        minute: date.getUTCMinutes(),
        second: date.getUTCSeconds(),
        yyyy: String(date.getUTCFullYear()).padStart(4, "0"),
        mm: leapDay ? "02" : String(date.getUTCMonth() + 1).padStart(2, "0"),
        dd: leapDay ? "29" : String(date.getUTCDate()).padStart(2, "0"),
        hh: String(date.getUTCHours()).padStart(2, "0"),
        mi: String(date.getUTCMinutes()).padStart(2, "0"),
        ss: String(date.getUTCSeconds()).padStart(2, "0"),
    };
}

/** Implements TEXT() for date and simple number patterns. Returns null for unsupported patterns. */
function formatTextFunctionValue(value: string, formatText: string): string | null {
    const format = String(formatText || "").trim();
    if (!format) return null;
    const numericValue = Number(value);

    if (!Number.isNaN(numericValue)) {
        if (isDateFormatCode(format)) {
            const rendered = formatDateByFormatCode(format, numericValue, excelSerialToDateParts);
            if (rendered !== null) return rendered;
        }

        if (/^0(?:\.0+)?$/.test(format)) {
            const decimalPlaces = (format.split(".")[1] || "").length;
            return numericValue.toFixed(decimalPlaces);
        }
        if (/^#,##0(?:\.0+)?$/.test(format)) {
            const decimalPlaces = (format.split(".")[1] || "").length;
            return numericValue.toLocaleString("en-US", {
                minimumFractionDigits: decimalPlaces,
                maximumFractionDigits: decimalPlaces,
                useGrouping: true,
            });
        }
    }
    return null;
}

function formatNumberByPattern(value: number, pattern: string): string {
    const normalizedPattern = pattern.trim();
    const decimalPlaces = (normalizedPattern.split(".")[1] || "").replace(/[^0#]/g, "").length;
    const useGrouping = normalizedPattern.includes(",");
    return value.toLocaleString("en-US", {
        minimumFractionDigits: decimalPlaces,
        maximumFractionDigits: decimalPlaces,
        useGrouping,
    });
}

/** Approximates decimals as fractions with denominators up to 100. Returns null without a close match. */
function formatFractionPattern(value: number): string | null {
    if (!Number.isFinite(value)) return null;
    const tolerance = 1e-9;
    for (let denominator = 1; denominator <= 100; denominator += 1) {
        const numerator = Math.round(value * denominator);
        if (Math.abs(value - numerator / denominator) < tolerance) {
            return `${numerator}/${denominator}`;
        }
    }
    return null;
}

/** Spreads characters with spaces for DBNum3 wide display. */
function formatDbNum3Pattern(rawValue: string): string {
    return rawValue.split("").join(" ");
}

/** Splits format codes on semicolons outside quotes. */
function splitFormatSections(formatCode: string): string[] {
    const sections: string[] = [];
    let current = "";
    let inQuotes = false;
    for (let index = 0; index < formatCode.length; index += 1) {
        const char = formatCode[index];
        if (char === '"') {
            inQuotes = !inQuotes;
            current += char;
            continue;
        }
        if (char === ";" && !inQuotes) {
            sections.push(current);
            current = "";
            continue;
        }
        current += char;
    }
    sections.push(current);
    return sections;
}

/** Renders zero sections with dash or currency placeholders. Returns null without a dash literal. */
function formatZeroSection(section: string): string | null {
    const normalizedSection = String(section || "");
    if (!normalizedSection) return null;
    const compact = normalizedSection.replace(/_.|\\.|[*?]/g, "").trim();
    const hasDashLiteral = /"-"|(^|[^a-z0-9])-($|[^a-z0-9])/i.test(compact);
    if (!hasDashLiteral) return null;
    if (compact.includes("¥")) return "¥ -";
    if (compact.includes("$")) return "$ -";
    return "-";
}

/** Formats raw values by the number format code of the file. Returns null when no rule applies. */
function formatCellDisplayValue(
    rawValue: string,
    cellStyle: Pick<CellStyleInfo, "numFmtId" | "formatCode">,
    context: FormatContext = {}
): string | null {
    if (rawValue === "") return null;
    const numericValue = Number(rawValue);
    if (Number.isNaN(numericValue)) return null;
    const formatCode = normalizeNumericFormatCode(cellStyle.formatCode);
    const normalized = formatCode.toLowerCase();

    if (isDateFormatCode(formatCode)) {
        const serial = context.date1904 ? numericValue + DATE_1904_OFFSET : numericValue;
        const parts = excelSerialToDateParts(serial);
        if (!parts) return null;
        const rendered = formatDateByFormatCode(
            cellStyle.formatCode,
            serial,
            excelSerialToDateParts,
            context.locale ?? null
        );
        if (rendered !== null) {
            return rendered;
        }
        const hasDate =
            /y/.test(normalized) ||
            /d/.test(normalized) ||
            /(^|[^a-z])m(?:\/|-)/.test(normalized) ||
            /(?:\/|-)m(?:[^a-z]|$)/.test(normalized);
        const hasTime =
            /h/.test(normalized) || /s/.test(normalized) || normalized.includes(":") || normalized.includes("am/pm");
        if (hasDate && hasTime) {
            return `${parts.yyyy}-${parts.mm}-${parts.dd} ${parts.hh}:${parts.mi}:${parts.ss}`;
        }
        if (hasTime && !hasDate) {
            return `${parts.hh}:${parts.mi}:${parts.ss}`;
        }
        return `${parts.yyyy}-${parts.mm}-${parts.dd}`;
    }

    if (/dbnum3/i.test(cellStyle.formatCode)) {
        return formatDbNum3Pattern(rawValue);
    }
    if (!formatCode || normalized === "general") {
        return null;
    }

    if (/[#0][^;]*e[+-]0+/i.test(formatCode)) {
        const scientificPattern = formatCode.split(";")[0] || formatCode;
        const decimalPartMatch = scientificPattern.match(/\.([0#]+)e[+-]/i);
        const decimalPlaces = (decimalPartMatch?.[1] || "").length;
        const exponentDigits = (scientificPattern.match(/e[+-]([0#]+)/i)?.[1] || "").length;
        const [mantissa, exponentPart] = numericValue.toExponential(decimalPlaces).split("e");
        const exponent = Number(exponentPart || 0);
        const sign = exponent >= 0 ? "+" : "-";
        const paddedExponent = String(Math.abs(exponent)).padStart(exponentDigits, "0");
        return `${mantissa}E${sign}${paddedExponent}`;
    }

    if (/[?#0]\/[?#0]/.test(normalized)) {
        return formatFractionPattern(numericValue);
    }

    return formatNumberByFormatCode(cellStyle.formatCode, numericValue);
}

/** Writes resolved values back into the cell and marks it resolved. */
function applyResolvedFormulaValue(
    cell: ResolvedCellLike,
    resolvedValue: string,
    resolutionSource: FormulaResolutionSource = "reference_lookup",
    context: FormatContext = {}
): void {
    const rawValue = String(resolvedValue || "");
    const formattedValue = formatCellDisplayValue(
        rawValue,
        {
            numFmtId: cell.numFmtId,
            formatCode: cell.formatCode,
        },
        context
    );
    cell.rawValue = rawValue;
    cell.outputValue = formattedValue ?? rawValue;
    cell.resolutionStatus = "resolved";
    cell.resolutionSource = resolutionSource;
}

/** Reads Excel serials, ISO dates, and Japanese date text. Fills missing parts with defaults. */
function parseDateLikeParts(value: string): {
    yyyy: string;
    mm: string;
    dd: string;
    hh: string;
    mi: string;
    ss: string;
} | null {
    const trimmed = String(value || "").trim();
    const numericValue = Number(trimmed);
    if (!Number.isNaN(numericValue)) {
        return excelSerialToDateParts(numericValue);
    }
    const isoMatch = trimmed.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
    if (isoMatch) {
        return {
            yyyy: isoMatch[1],
            mm: isoMatch[2].padStart(2, "0"),
            dd: isoMatch[3].padStart(2, "0"),
            hh: (isoMatch[4] || "00").padStart(2, "0"),
            mi: (isoMatch[5] || "00").padStart(2, "0"),
            ss: (isoMatch[6] || "00").padStart(2, "0"),
        };
    }
    const japaneseMatch = trimmed.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日(?:\s*(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
    if (japaneseMatch) {
        return {
            yyyy: japaneseMatch[1],
            mm: japaneseMatch[2].padStart(2, "0"),
            dd: japaneseMatch[3].padStart(2, "0"),
            hh: (japaneseMatch[4] || "00").padStart(2, "0"),
            mi: (japaneseMatch[5] || "00").padStart(2, "0"),
            ss: (japaneseMatch[6] || "00").padStart(2, "0"),
        };
    }
    const japaneseYearMonthMatch = trimmed.match(/^(\d{4})年(\d{1,2})月$/);
    if (japaneseYearMonthMatch) {
        return {
            yyyy: japaneseYearMonthMatch[1],
            mm: japaneseYearMonthMatch[2].padStart(2, "0"),
            dd: "01",
            hh: "00",
            mi: "00",
            ss: "00",
        };
    }
    const japaneseMonthDayMatch = trimmed.match(/^(\d{1,2})月(\d{1,2})日$/);
    if (japaneseMonthDayMatch) {
        return {
            yyyy: "2000",
            mm: japaneseMonthDayMatch[1].padStart(2, "0"),
            dd: japaneseMonthDayMatch[2].padStart(2, "0"),
            hh: "00",
            mi: "00",
            ss: "00",
        };
    }
    const isoYearMonthMatch = trimmed.match(/^(\d{4})[-/](\d{1,2})$/);
    if (isoYearMonthMatch) {
        return {
            yyyy: isoYearMonthMatch[1],
            mm: isoYearMonthMatch[2].padStart(2, "0"),
            dd: "01",
            hh: "00",
            mi: "00",
            ss: "00",
        };
    }
    return null;
}

/** Converts date parts back to serials. Reapplies the 1900 leap-year offset. */
function datePartsToExcelSerial(
    year: number,
    month: number,
    day: number,
    hour = 0,
    minute = 0,
    second = 0
): number | null {
    if (![year, month, day, hour, minute, second].every(Number.isFinite)) return null;
    const baseUtcMs = Date.UTC(1899, 11, 31);
    const targetUtcMs = Date.UTC(year, month - 1, day, hour, minute, second);
    const msPerDay = 24 * 60 * 60 * 1000;
    let serial = (targetUtcMs - baseUtcMs) / msPerDay;
    if (serial >= 60) {
        serial += 1;
    }
    return serial;
}

/** Formats raw values by Excel number format code. Compensates the 1900 leap-year bug in both directions. */
export type { FormatContext };
export {
    isDateFormatCode,
    normalizeNumericFormatCode,
    excelSerialToIsoText,
    excelSerialToDateParts,
    formatTextFunctionValue,
    formatNumberByPattern,
    formatFractionPattern,
    formatDbNum3Pattern,
    splitFormatSections,
    formatZeroSection,
    formatCellDisplayValue,
    applyResolvedFormulaValue,
    parseDateLikeParts,
    datePartsToExcelSerial,
};
