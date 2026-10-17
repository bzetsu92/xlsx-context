/**
 * Month-title tokens that mark a calendar header (e.g. a planner template).
 *
 * Each entry is one locale's "year + month" title. To support another
 * language, add one pattern here. Match is exact after trimming.
 */

const MONTH_TITLE_PATTERNS: readonly RegExp[] = [
    /^\d{4}年\d{1,2}月$/, // Japanese "YYYY年M月"
];

/** True when the text is a calendar month title in any supported locale. */
export function isMonthTitleToken(value: string): boolean {
    const text = String(value || "").trim();
    if (!text) return false;
    return MONTH_TITLE_PATTERNS.some((pattern) => pattern.test(text));
}
