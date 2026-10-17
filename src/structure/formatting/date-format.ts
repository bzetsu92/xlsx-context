/**
 * Renders an Excel serial date with a date/time number format code.
 */
import { localeFromTagId } from "../../extraction/workbook/workbook-locale.js";

export { splitSections };

export type DateParts = {
    year: number;
    month: number;
    day: number;
    hour: number;
    minute: number;
    second: number;
};

type Token =
    | { kind: "literal"; text: string }
    | { kind: "year"; digits: number }
    | { kind: "month"; count: number }
    | { kind: "day"; count: number }
    | { kind: "weekday"; count: number; japanese: boolean }
    | { kind: "era"; count: number }
    | { kind: "eraYear"; count: number }
    | { kind: "hour"; count: number; elapsed: boolean }
    | { kind: "minute"; count: number; elapsed: boolean }
    | { kind: "second"; count: number; elapsed: boolean }
    | { kind: "ampm"; text: string; am?: string; pm?: string }
    | { kind: "fraction"; digits: number };

type ParsedSection = {
    tokens: Token[];
    /** Locale from a [$-xxx] tag. Null when the code has no known tag. */
    locale: string | null;
    /** A system tag: the reader computer shows its own long date or time. */
    system: "date" | "time" | null;
};

type LocaleTag = { locale: string | null; system: "date" | "time" | null; symbol: string };

const DEFAULT_LOCALE = "en-US";
const LOCAL_AM_PM = [
    { text: "上午/下午", am: "上午", pm: "下午" },
    { text: "오전/오후", am: "오전", pm: "오후" },
];
const nameCache = new Map<string, string[]>();

/** Japanese eras, newest first. The start is the first day of the era. */
const JA_ERAS = [
    { name: "令和", initial: "R", start: [2019, 5, 1] },
    { name: "平成", initial: "H", start: [1989, 1, 8] },
    { name: "昭和", initial: "S", start: [1926, 12, 25] },
    { name: "大正", initial: "T", start: [1912, 7, 30] },
    { name: "明治", initial: "M", start: [1873, 1, 1] },
] as const;

const IGNORED_BRACKETS = /^(black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)$/i;
/** Characters that Excel prints as they are, without quotes. */
const BARE_LITERALS = new Set([
    " ",
    "-",
    "+",
    "/",
    ":",
    ",",
    ".",
    "(",
    ")",
    "$",
    "!",
    "^",
    "&",
    "'",
    "~",
    "{",
    "}",
    "<",
    ">",
    "=",
]);

/** Splits a format code at ";" outside quotes, brackets, and escapes. */
function splitSections(code: string): string[] {
    const sections: string[] = [];
    let current = "";
    let inQuote = false;
    let inBracket = false;
    for (let index = 0; index < code.length; index += 1) {
        const char = code[index];
        if (char === "\\" && !inQuote && index + 1 < code.length) {
            current += char + code[index + 1];
            index += 1;
            continue;
        }
        if (char === '"' && !inBracket) inQuote = !inQuote;
        else if (char === "[" && !inQuote) inBracket = true;
        else if (char === "]" && !inQuote) inBracket = false;
        if (char === ";" && !inQuote && !inBracket) {
            sections.push(current);
            current = "";
            continue;
        }
        current += char;
    }
    sections.push(current);
    return sections;
}

/** Reads a [$symbol-locale] tag. Returns null when the bracket is not a tag of this kind. */
function readLocaleTag(content: string): LocaleTag | null {
    const match = content.match(/^\$([^-]*)(?:-(.+))?$/);
    if (!match) return null;
    const symbol = match[1];
    const id = (match[2] || "").toLowerCase();
    const lcid = /^[0-9a-f]+$/.test(id) ? parseInt(id, 16) & 0xffff : -1;
    if (lcid === 0xf800 || id === "x-sysdate") return { locale: null, system: "date", symbol };
    if (lcid === 0xf400 || id === "x-systime") return { locale: null, system: "time", symbol };
    return { locale: id ? localeFromTagId(id) : null, system: null, symbol };
}

/** Reads one section into tokens. Returns null when any part is not understood. */
function parseSection(section: string): ParsedSection | null {
    const tokens: Token[] = [];
    let locale: string | null = null;
    const lower = section.toLowerCase();
    let index = 0;
    while (index < section.length) {
        const char = section[index];
        const lowerChar = lower[index];

        if (char === '"') {
            const end = section.indexOf('"', index + 1);
            if (end < 0) return null;
            tokens.push({ kind: "literal", text: section.slice(index + 1, end) });
            index = end + 1;
            continue;
        }
        if (char === "\\") {
            if (index + 1 >= section.length) return null;
            tokens.push({ kind: "literal", text: section[index + 1] });
            index += 2;
            continue;
        }
        if (char === "[") {
            const end = section.indexOf("]", index + 1);
            if (end < 0) return null;
            const content = section.slice(index + 1, end);
            const tag = readLocaleTag(content);
            if (tag !== null) {
                // The text after a system tag is only a snapshot. The reader's computer decides the display.
                if (tag.system) return { tokens: [], locale: null, system: tag.system };
                if (tag.locale) locale = tag.locale;
                if (tag.symbol) tokens.push({ kind: "literal", text: tag.symbol });
            } else if (/^(h+|m+|s+)$/i.test(content) && content.length <= 2) {
                const unit = content[0].toLowerCase();
                if (unit === "h") tokens.push({ kind: "hour", count: content.length, elapsed: true });
                else if (unit === "m") tokens.push({ kind: "minute", count: content.length, elapsed: true });
                else tokens.push({ kind: "second", count: content.length, elapsed: true });
            } else if (!IGNORED_BRACKETS.test(content)) {
                // Rejects conditions, digit styles such as [DBNum1], and unknown tags.
                return null;
            }
            index = end + 1;
            continue;
        }
        const localAmPm = LOCAL_AM_PM.find((entry) => section.startsWith(entry.text, index));
        if (localAmPm) {
            tokens.push({ kind: "ampm", text: localAmPm.text, am: localAmPm.am, pm: localAmPm.pm });
            index += localAmPm.text.length;
            continue;
        }
        if (char === "_" || char === "*") {
            // Width padding and fill characters have no meaning in plain text.
            index += 2;
            continue;
        }
        if (lowerChar === "a") {
            if (lower.startsWith("am/pm", index)) {
                tokens.push({ kind: "ampm", text: section.slice(index, index + 5) });
                index += 5;
                continue;
            }
            if (lower.startsWith("a/p", index)) {
                tokens.push({ kind: "ampm", text: section.slice(index, index + 3) });
                index += 3;
                continue;
            }
            const count = countRun(lower, index, "a");
            if (count !== 3 && count !== 4) return null;
            tokens.push({ kind: "weekday", count, japanese: true });
            index += count;
            continue;
        }
        if ("ymdhsge".includes(lowerChar)) {
            const count = countRun(lower, index, lowerChar);
            index += count;
            switch (lowerChar) {
                case "y":
                    if (count > 4) return null;
                    tokens.push({ kind: "year", digits: count <= 2 ? 2 : 4 });
                    break;
                case "m":
                    if (count > 5) return null;
                    tokens.push({ kind: "month", count });
                    break;
                case "d":
                    if (count > 4) return null;
                    tokens.push({ kind: "day", count });
                    break;
                case "h":
                    if (count > 2) return null;
                    tokens.push({ kind: "hour", count, elapsed: false });
                    break;
                case "s":
                    if (count > 2) return null;
                    tokens.push({ kind: "second", count, elapsed: false });
                    break;
                case "g":
                    if (count > 3) return null;
                    tokens.push({ kind: "era", count });
                    break;
                default:
                    if (count > 2) return null;
                    tokens.push({ kind: "eraYear", count });
            }
            continue;
        }
        if (char === ".") {
            const previous = tokens[tokens.length - 1];
            const zeros = countRun(section, index + 1, "0");
            if (previous && previous.kind === "second" && zeros > 0) {
                tokens.push({ kind: "fraction", digits: zeros });
                index += 1 + zeros;
                continue;
            }
        }
        if (BARE_LITERALS.has(char) || char.charCodeAt(0) > 0x7f) {
            tokens.push({ kind: "literal", text: char });
            index += 1;
            continue;
        }
        // Digits, #, ?, %, @, E, and any other letter are not part of a date format.
        return null;
    }
    // An era year needs an era token in the same section.
    const hasEra = tokens.some((token) => token.kind === "era");
    if (tokens.some((token) => token.kind === "eraYear") && !hasEra) return null;
    return { tokens: resolveMinutes(tokens), locale, system: null };
}

function countRun(text: string, start: number, char: string): number {
    let count = 0;
    while (text[start + count] === char) count += 1;
    return count;
}

/** "m" means minutes right after an hour or right before a second. Otherwise it means month. */
function resolveMinutes(tokens: Token[]): Token[] {
    const significant = (from: number, step: number): Token | null => {
        for (let index = from; index >= 0 && index < tokens.length; index += step) {
            if (tokens[index].kind !== "literal") return tokens[index];
        }
        return null;
    };
    return tokens.map((token, index) => {
        if (token.kind !== "month" || token.count > 2) return token;
        const before = significant(index - 1, -1);
        const after = significant(index + 1, 1);
        if ((before && before.kind === "hour") || (after && after.kind === "second")) {
            return { kind: "minute", count: token.count, elapsed: false };
        }
        return token;
    });
}

function pad(value: number, width: number): string {
    return String(value).padStart(width, "0");
}

function findEra(parts: DateParts): { name: string; initial: string; year: number } | null {
    for (const era of JA_ERAS) {
        const [year, month, day] = era.start;
        const onOrAfter =
            parts.year > year ||
            (parts.year === year && (parts.month > month || (parts.month === month && parts.day >= day)));
        if (onOrAfter) return { name: era.name, initial: era.initial, year: parts.year - year + 1 };
    }
    return null;
}

function localNames(locale: string, kind: "month" | "weekday", width: "long" | "short" | "narrow"): string[] {
    const key = `${locale}|${kind}|${width}`;
    const cached = nameCache.get(key);
    if (cached) return cached;
    const format = new Intl.DateTimeFormat(locale, { [kind]: width, timeZone: "UTC" });
    // 2006-01-01 is a Sunday, so day offsets match getUTCDay().
    const names =
        kind === "month"
            ? Array.from({ length: 12 }, (_, month) => format.format(new Date(Date.UTC(2006, month, 1))))
            : Array.from({ length: 7 }, (_, day) => format.format(new Date(Date.UTC(2006, 0, 1 + day))));
    nameCache.set(key, names);
    return names;
}

function renderSystem(kind: "date" | "time", parts: DateParts, locale: string | null): string | null {
    if (!locale) return null;
    const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second));
    const options: Intl.DateTimeFormatOptions =
        kind === "date" ? { dateStyle: "full", timeZone: "UTC" } : { timeStyle: "medium", timeZone: "UTC" };
    try {
        return new Intl.DateTimeFormat(locale, options).format(date);
    } catch {
        return null;
    }
}

function render(parsed: ParsedSection, serial: number, parts: DateParts, defaultLocale: string | null): string | null {
    if (parsed.system) return renderSystem(parsed.system, parts, defaultLocale);
    const locale = parsed.locale || defaultLocale || DEFAULT_LOCALE;
    const twelveHour = parsed.tokens.some((token) => token.kind === "ampm");
    const totalSeconds = Math.round(serial * 86400);
    const weekday = new Date(Date.UTC(parts.year, parts.month - 1, parts.day)).getUTCDay();
    const needsEra = parsed.tokens.some((token) => token.kind === "era" || token.kind === "eraYear");
    const era = needsEra ? findEra(parts) : null;
    if (needsEra && !era) return null;

    let output = "";
    for (const token of parsed.tokens) {
        switch (token.kind) {
            case "literal":
                output += token.text;
                break;
            case "year":
                output += token.digits === 2 ? pad(parts.year % 100, 2) : pad(parts.year, 4);
                break;
            case "month":
                if (token.count <= 2) output += pad(parts.month, token.count);
                else {
                    const width = token.count === 3 ? "short" : token.count === 4 ? "long" : "narrow";
                    output += localNames(locale, "month", width)[parts.month - 1];
                }
                break;
            case "day":
                if (token.count <= 2) output += pad(parts.day, token.count);
                else output += localNames(locale, "weekday", token.count === 3 ? "short" : "long")[weekday];
                break;
            case "weekday":
                // aaa and aaaa are Japanese weekday codes in every edition.
                output += localNames("ja-JP", "weekday", token.count === 3 ? "short" : "long")[weekday];
                break;
            case "era":
                if (!era) return null;
                output += token.count === 1 ? era.initial : token.count === 2 ? era.name[0] : era.name;
                break;
            case "eraYear":
                if (!era) return null;
                output += pad(era.year, token.count);
                break;
            case "hour": {
                const value = token.elapsed
                    ? Math.floor(totalSeconds / 3600)
                    : twelveHour
                      ? parts.hour % 12 || 12
                      : parts.hour;
                output += pad(value, token.count);
                break;
            }
            case "minute":
                output += pad(token.elapsed ? Math.floor(totalSeconds / 60) : parts.minute, token.count);
                break;
            case "second":
                output += pad(token.elapsed ? totalSeconds : parts.second, token.count);
                break;
            case "ampm": {
                const isPm = parts.hour >= 12;
                if (token.am && token.pm) {
                    output += isPm ? token.pm : token.am;
                    break;
                }
                const full = token.text.length === 5;
                const text = full ? (isPm ? "PM" : "AM") : isPm ? "P" : "A";
                output += token.text === token.text.toLowerCase() ? text.toLowerCase() : text;
                break;
            }
            case "fraction": {
                const dayFraction = serial - Math.floor(serial);
                const millis = Math.round(dayFraction * 86400000) % 1000;
                output += `.${pad(millis, 3).padEnd(token.digits, "0").slice(0, token.digits)}`;
                break;
            }
        }
    }
    return output;
}

/**
 * Formats an Excel serial number with a date/time format code.
 * Returns null when the code is not a date format that this renderer supports.
 * A code with several sections uses the first section, and the third for zero.
 */
export function formatDateByFormatCode(
    formatCode: string,
    serial: number,
    toParts: (serial: number) => DateParts | null,
    defaultLocale: string | null = null
): string | null {
    if (!Number.isFinite(serial) || serial < 0) return null;
    const sections = splitSections(String(formatCode || ""));
    const section = serial === 0 && sections.length >= 3 ? sections[2] : sections[0];
    const parsed = parseSection(section);
    if (!parsed) return null;
    const hasDateTimeToken = parsed.system !== null || parsed.tokens.some((token) => token.kind !== "literal");
    if (!hasDateTimeToken) return null;
    const parts = toParts(serial);
    if (!parts) return null;
    return render(parsed, serial, parts, defaultLocale);
}
