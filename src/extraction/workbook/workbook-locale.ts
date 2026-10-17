/** Workbook locale from the file itself, and the built-in number formats that depend on it. */

const LCID_LOCALES: Record<number, string> = {
    0x0401: "ar-SA",
    0x0402: "bg-BG",
    0x0403: "ca-ES",
    0x0404: "zh-TW",
    0x0405: "cs-CZ",
    0x0406: "da-DK",
    0x0407: "de-DE",
    0x0408: "el-GR",
    0x0409: "en-US",
    0x040a: "es-ES",
    0x040b: "fi-FI",
    0x040c: "fr-FR",
    0x040d: "he-IL",
    0x040e: "hu-HU",
    0x040f: "is-IS",
    0x0410: "it-IT",
    0x0411: "ja-JP",
    0x0412: "ko-KR",
    0x0413: "nl-NL",
    0x0414: "nb-NO",
    0x0415: "pl-PL",
    0x0416: "pt-BR",
    0x0418: "ro-RO",
    0x0419: "ru-RU",
    0x041a: "hr-HR",
    0x041b: "sk-SK",
    0x041d: "sv-SE",
    0x041e: "th-TH",
    0x041f: "tr-TR",
    0x0421: "id-ID",
    0x0422: "uk-UA",
    0x0424: "sl-SI",
    0x0425: "et-EE",
    0x0426: "lv-LV",
    0x0427: "lt-LT",
    0x042a: "vi-VN",
    0x0439: "hi-IN",
    0x043e: "ms-MY",
    0x0804: "zh-CN",
    0x0807: "de-CH",
    0x0809: "en-GB",
    0x080a: "es-MX",
    0x080c: "fr-BE",
    0x0810: "it-CH",
    0x0813: "nl-BE",
    0x0816: "pt-PT",
    0x0c04: "zh-HK",
    0x0c07: "de-AT",
    0x0c09: "en-AU",
    0x0c0a: "es-ES",
    0x0c0c: "fr-CA",
    0x1004: "zh-SG",
    0x1009: "en-CA",
    0x100c: "fr-CH",
    0x1409: "en-NZ",
    0x1809: "en-IE",
    0x1c09: "en-ZA",
    0x2c0a: "es-AR",
    0x4009: "en-IN",
    0x4809: "en-SG",
};

/** Primary language ids, for LCIDs missing from the table above. */
const LANGUAGE_IDS: Record<number, string> = {
    0x01: "ar",
    0x02: "bg",
    0x03: "ca",
    0x04: "zh",
    0x05: "cs",
    0x06: "da",
    0x07: "de",
    0x08: "el",
    0x09: "en",
    0x0a: "es",
    0x0b: "fi",
    0x0c: "fr",
    0x0d: "he",
    0x0e: "hu",
    0x0f: "is",
    0x10: "it",
    0x11: "ja",
    0x12: "ko",
    0x13: "nl",
    0x14: "nb",
    0x15: "pl",
    0x16: "pt",
    0x18: "ro",
    0x19: "ru",
    0x1a: "hr",
    0x1b: "sk",
    0x1d: "sv",
    0x1e: "th",
    0x1f: "tr",
    0x21: "id",
    0x22: "uk",
    0x24: "sl",
    0x25: "et",
    0x26: "lv",
    0x27: "lt",
    0x2a: "vi",
    0x39: "hi",
    0x3e: "ms",
};

/** Default fonts that only East Asian editions of Excel write. Western fonts say nothing about the language. */
const FONT_LOCALES: Array<[RegExp, string]> = [
    [
        /^(游ゴシック|游明朝|ＭＳ Ｐゴシック|ＭＳ ゴシック|ＭＳ Ｐ明朝|ＭＳ 明朝|メイリオ|Yu Gothic|MS PGothic|MS Gothic|Meiryo)/i,
        "ja-JP",
    ],
    [/^(맑은 고딕|굴림|돋움|바탕|Malgun Gothic|Gulim|Dotum|Batang)/i, "ko-KR"],
    [/^(等线|宋体|微软雅黑|DengXian|SimSun|Microsoft YaHei)/i, "zh-CN"],
    [/^(新細明體|微軟正黑體|PMingLiU|Microsoft JhengHei)/i, "zh-TW"],
];

const CURRENCY_BY_REGION: Record<string, string> = {
    US: "$",
    CA: "$",
    AU: "$",
    NZ: "$",
    SG: "$",
    GB: "£",
    IE: "€",
    IN: "₹",
    ZA: "R",
    JP: "¥",
    CN: "¥",
    TW: "NT$",
    HK: "HK$",
    KR: "₩",
    VN: "₫",
    TH: "฿",
    ID: "Rp",
    MY: "RM",
    RU: "₽",
    UA: "₴",
    PL: "zł",
    CZ: "Kč",
    HU: "Ft",
    TR: "₺",
    BR: "R$",
    MX: "$",
    AR: "$",
    CH: "CHF",
    SE: "kr",
    NO: "kr",
    DK: "kr",
    IS: "kr",
    IL: "₪",
    SA: "ر.س.‏",
    BG: "лв.",
    RO: "lei",
    HR: "€",
};
const EURO_REGIONS = new Set([
    "DE",
    "FR",
    "IT",
    "ES",
    "NL",
    "BE",
    "AT",
    "PT",
    "FI",
    "GR",
    "SK",
    "SI",
    "EE",
    "LV",
    "LT",
]);

/** Built-in formats that every edition shares (ECMA-376 Part 1, 18.8.30). */
const INVARIANT_FORMATS: Record<number, string> = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mm:ss.0",
    48: "##0.0E+0",
    49: "@",
};

/** Built-ins 27-36 and 50-58 exist only in East Asian editions (ECMA-376 Part 1, 18.8.30). */
const EAST_ASIAN_FORMATS: Record<string, Record<number, string>> = {
    ja: {
        27: "[$-411]ge.m.d",
        28: '[$-411]ggge"年"m"月"d"日"',
        29: '[$-411]ggge"年"m"月"d"日"',
        30: "m/d/yy",
        31: 'yyyy"年"m"月"d"日"',
        32: 'h"時"mm"分"',
        33: 'h"時"mm"分"ss"秒"',
        34: 'yyyy"年"m"月"',
        35: 'm"月"d"日"',
        36: "[$-411]ge.m.d",
        50: "[$-411]ge.m.d",
        51: '[$-411]ggge"年"m"月"d"日"',
        52: 'yyyy"年"m"月"',
        53: 'm"月"d"日"',
        54: '[$-411]ggge"年"m"月"d"日"',
        55: 'yyyy"年"m"月"',
        56: 'm"月"d"日"',
        57: "[$-411]ge.m.d",
        58: '[$-411]ggge"年"m"月"d"日"',
    },
    "zh-CN": {
        27: 'yyyy"年"m"月"',
        28: 'm"月"d"日"',
        29: 'm"月"d"日"',
        30: "m-d-yy",
        31: 'yyyy"年"m"月"d"日"',
        32: 'h"时"mm"分"',
        33: 'h"时"mm"分"ss"秒"',
        34: '上午/下午h"时"mm"分"',
        35: '上午/下午h"时"mm"分"ss"秒"',
        36: 'yyyy"年"m"月"',
        50: 'yyyy"年"m"月"',
        51: 'm"月"d"日"',
        52: 'yyyy"年"m"月"',
        53: 'm"月"d"日"',
        54: 'm"月"d"日"',
        55: '上午/下午h"时"mm"分"',
        56: '上午/下午h"时"mm"分"ss"秒"',
        57: 'yyyy"年"m"月"',
        58: 'm"月"d"日"',
    },
    "zh-TW": {
        27: "[$-404]e/m/d",
        28: '[$-404]e"年"m"月"d"日"',
        29: '[$-404]e"年"m"月"d"日"',
        30: "m/d/yy",
        31: 'yyyy"年"m"月"d"日"',
        32: 'hh"時"mm"分"',
        33: 'hh"時"mm"分"ss"秒"',
        34: '上午/下午hh"時"mm"分"',
        35: '上午/下午hh"時"mm"分"ss"秒"',
        36: "[$-404]e/m/d",
        50: "[$-404]e/m/d",
        51: '[$-404]e"年"m"月"d"日"',
        52: '上午/下午hh"時"mm"分"',
        53: '上午/下午hh"時"mm"分"ss"秒"',
        54: '[$-404]e"年"m"月"d"日"',
        55: '上午/下午hh"時"mm"分"',
        56: '上午/下午hh"時"mm"分"ss"秒"',
        57: "[$-404]e/m/d",
        58: '[$-404]e"年"m"月"d"日"',
    },
    ko: {
        27: 'yyyy"年" mm"月" dd"日"',
        28: "mm-dd",
        29: "mm-dd",
        30: "mm-dd-yy",
        31: 'yyyy"년" mm"월" dd"일"',
        32: 'h"시" mm"분"',
        33: 'h"시" mm"분" ss"초"',
        34: "yyyy-mm-dd",
        35: "yyyy-mm-dd",
        36: 'yyyy"年" mm"月" dd"日"',
        50: 'yyyy"年" mm"月" dd"日"',
        51: "mm-dd",
        52: "yyyy-mm-dd",
        53: "yyyy-mm-dd",
        54: "mm-dd",
        55: "yyyy-mm-dd",
        56: "yyyy-mm-dd",
        57: 'yyyy"年" mm"月" dd"日"',
        58: "mm-dd",
    },
};

const ISO_DATE = "yyyy-mm-dd";

function canonical(tag: string): string | null {
    try {
        return Intl.getCanonicalLocales(tag)[0] || null;
    } catch {
        return null;
    }
}

/** Reads the id of a `[$-id]` tag: a hexadecimal LCID or a language tag. Returns null for system and unknown ids. */
function localeFromTagId(id: string): string | null {
    const value = id.trim();
    if (/^[0-9a-f]+$/i.test(value)) {
        const lcid = parseInt(value, 16) & 0xffff;
        if (lcid === 0 || lcid >= 0xf000) return null;
        return LCID_LOCALES[lcid] || LANGUAGE_IDS[lcid & 0x3ff] || null;
    }
    if (/^x-/i.test(value)) return null;
    return canonical(value);
}

/** Collects locales from `[$...-id]` tags in the custom number formats. */
function localesFromFormatCodes(formatCodes: string[]): string[] {
    const locales: string[] = [];
    for (const code of formatCodes) {
        for (const match of code.matchAll(/\[\$[^\]-]*-([^\]]+)\]/g)) {
            const locale = localeFromTagId(match[1]);
            if (locale) locales.push(locale);
        }
    }
    return locales;
}

/**
 * Picks the workbook locale from the file. Locale tags in number formats win, then an East Asian default font.
 * Returns null when the file does not say, so callers do not guess.
 */
function inferWorkbookLocale(formatCodes: string[], defaultFontName: string): string | null {
    const counts = new Map<string, number>();
    for (const locale of localesFromFormatCodes(formatCodes)) counts.set(locale, (counts.get(locale) || 0) + 1);
    let best: string | null = null;
    for (const [locale, count] of counts) {
        if (best === null || count > (counts.get(best) || 0)) best = locale;
    }
    if (best) return best;
    const font = defaultFontName.trim();
    for (const [pattern, locale] of FONT_LOCALES) {
        if (pattern.test(font)) return locale;
    }
    return null;
}

function regionOf(locale: string): string {
    try {
        return new Intl.Locale(locale).maximize().region || "";
    } catch {
        return "";
    }
}

function currencySymbolFor(locale: string | null): string {
    if (!locale) return "";
    const region = regionOf(locale);
    if (EURO_REGIONS.has(region)) return "€";
    return CURRENCY_BY_REGION[region] || "";
}

function quoteLiteral(text: string): string {
    return /^[\s\-/.:,]+$/.test(text) ? text : `"${text.replace(/"/g, "")}"`;
}

/** Builds the short date format code of a locale from CLDR data, as the Excel edition of that locale shows it. */
function shortDateFormatCode(locale: string | null): string {
    if (!locale) return ISO_DATE;
    try {
        const parts = new Intl.DateTimeFormat(locale, {
            year: "numeric",
            month: "numeric",
            day: "numeric",
            timeZone: "UTC",
        }).formatToParts(new Date(Date.UTC(2006, 0, 2)));
        let code = "";
        for (const part of parts) {
            if (part.type === "year") code += part.value.length <= 2 ? "yy" : "yyyy";
            else if (part.type === "month") code += part.value.length === 2 ? "mm" : "m";
            else if (part.type === "day") code += part.value.length === 2 ? "dd" : "d";
            else if (part.type === "literal") code += quoteLiteral(part.value);
            else return ISO_DATE;
        }
        return code || ISO_DATE;
    } catch {
        return ISO_DATE;
    }
}

function eastAsianTable(locale: string | null): Record<number, string> | null {
    if (!locale) return null;
    if (/^ja\b/i.test(locale)) return EAST_ASIAN_FORMATS.ja;
    if (/^ko\b/i.test(locale)) return EAST_ASIAN_FORMATS.ko;
    if (/^zh\b/i.test(locale)) {
        const region = regionOf(locale);
        return region === "TW" || region === "HK" || region === "MO"
            ? EAST_ASIAN_FORMATS["zh-TW"]
            : EAST_ASIAN_FORMATS["zh-CN"];
    }
    return null;
}

function currencyFormat(id: number, symbol: string): string {
    const sym = symbol ? `"${symbol}"` : "";
    switch (id) {
        case 5:
            return `${sym}#,##0_);(${sym}#,##0)`;
        case 6:
            return `${sym}#,##0_);[Red](${sym}#,##0)`;
        case 7:
            return `${sym}#,##0.00_);(${sym}#,##0.00)`;
        case 8:
            return `${sym}#,##0.00_);[Red](${sym}#,##0.00)`;
        case 42:
            return `_(${sym}* #,##0_);_(${sym}* \\(#,##0\\);_(${sym}* "-"_);_(@_)`;
        default:
            return `_(${sym}* #,##0.00_);_(${sym}* \\(#,##0.00\\);_(${sym}* "-"??_);_(@_)`;
    }
}

/**
 * Format code of a built-in id. The file stores no code for these ids, and some ids depend on the edition locale.
 * Returns undefined for ids that are not built in.
 */
function builtinFormatCode(id: number, locale: string | null): string | undefined {
    if (id in INVARIANT_FORMATS) return INVARIANT_FORMATS[id];
    if (id === 14) return shortDateFormatCode(locale);
    if (id === 22) return `${shortDateFormatCode(locale)} h:mm`;
    if ((id >= 5 && id <= 8) || id === 42 || id === 44) return currencyFormat(id, currencySymbolFor(locale));
    if ((id >= 27 && id <= 36) || (id >= 50 && id <= 58)) {
        const table = eastAsianTable(locale);
        if (table) return table[id];
        return id === 32 ? "h:mm" : id === 33 ? "h:mm:ss" : shortDateFormatCode(locale);
    }
    return undefined;
}

export { builtinFormatCode, currencySymbolFor, inferWorkbookLocale, localeFromTagId, shortDateFormatCode };
