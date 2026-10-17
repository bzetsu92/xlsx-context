/**
 * Cell texts that mark a calendar layout: weekday names and ISO dates.
 *
 * The names live in one table. To support another language, add one entry.
 * Match is exact after trimming. Case does not matter.
 */

const WEEKDAY_NAMES: Record<string, readonly string[]> = {
    ja: [
        "日",
        "月",
        "火",
        "水",
        "木",
        "金",
        "土",
        "日曜日",
        "月曜日",
        "火曜日",
        "水曜日",
        "木曜日",
        "金曜日",
        "土曜日",
    ],
    en: [
        "Sun",
        "Mon",
        "Tue",
        "Wed",
        "Thu",
        "Fri",
        "Sat",
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ],
};

const WEEKDAY_SET = new Set(
    Object.values(WEEKDAY_NAMES)
        .flat()
        .map((name) => name.toLowerCase())
);

/** True when the text is a weekday name in any supported language. */
export function isWeekdayToken(value: string): boolean {
    return WEEKDAY_SET.has(
        String(value || "")
            .trim()
            .toLowerCase()
    );
}

/** True when the text is an ISO date such as 2024-03-15. */
export function isIsoDateToken(value: string): boolean {
    return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "").trim());
}
