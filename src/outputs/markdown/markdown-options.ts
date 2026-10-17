type MarkdownOptions = {
    treatFirstRowAsHeader?: boolean;
    trimText?: boolean;
    removeEmptyRows?: boolean;
    removeEmptyColumns?: boolean;
    includeShapeDetails?: boolean;
    outputMode?: "display" | "raw" | "both" | string;
    formattingMode?: "plain" | "gfm" | "github" | string;
    tableDetectionMode?: "balanced" | "border" | "planner-aware" | string;
};

type ResolvedMarkdownOptions = {
    treatFirstRowAsHeader: boolean;
    trimText: boolean;
    removeEmptyRows: boolean;
    removeEmptyColumns: boolean;
    includeShapeDetails: boolean;
    outputMode: "display" | "raw" | "both";
    formattingMode: "plain" | "gfm";
    tableDetectionMode: "balanced" | "border" | "planner-aware";
};

const OUTPUT_MODES = ["display", "raw", "both"] as const;
const FORMATTING_MODES = ["plain", "gfm"] as const;
// "github" was the mode name up to 1.3. Keep it so existing callers do not break.
const FORMATTING_MODE_ALIASES: Record<string, "plain" | "gfm"> = {
    github: "gfm",
};
const TABLE_DETECTION_MODES = ["balanced", "border", "planner-aware"] as const;
const TABLE_DETECTION_MODE_ALIASES: Record<string, "balanced" | "border" | "planner-aware"> = {
    "border-priority": "border",
};

/** Matches option text case-insensitively with alias support. Falls back on empty or unknown values. */
function normalizeEnum<T extends string>(
    value: string | null | undefined,
    allowedValues: readonly T[],
    fallback: T,
    aliases: Record<string, T> = {}
): T {
    const normalizedInput = String(value || "")
        .trim()
        .toLowerCase();
    if (!normalizedInput) {
        return fallback;
    }
    const normalizedValue = aliases[normalizedInput] || normalizedInput;
    return (allowedValues as readonly string[]).includes(normalizedValue) ? (normalizedValue as T) : fallback;
}

/** Falls back on nullish values. */
function resolveBoolean(value: boolean | null | undefined, fallback: boolean): boolean {
    return value === undefined || value === null ? fallback : value !== false;
}

function normalizeOutputMode(value?: string | null): "display" | "raw" | "both" {
    return normalizeEnum(value, OUTPUT_MODES, "display");
}

function normalizeFormattingMode(value?: string | null): "plain" | "gfm" {
    return normalizeEnum(value, FORMATTING_MODES, "plain", FORMATTING_MODE_ALIASES);
}

function normalizeTableDetectionMode(value?: string | null): "balanced" | "border" | "planner-aware" {
    return normalizeEnum(value, TABLE_DETECTION_MODES, "balanced", TABLE_DETECTION_MODE_ALIASES);
}

/** Fills option defaults. Unknown modes fall back to display, plain, and balanced. */
function resolveMarkdownOptions(options: MarkdownOptions = {}): ResolvedMarkdownOptions {
    return {
        treatFirstRowAsHeader: resolveBoolean(options.treatFirstRowAsHeader, true),
        trimText: resolveBoolean(options.trimText, true),
        removeEmptyRows: resolveBoolean(options.removeEmptyRows, true),
        removeEmptyColumns: resolveBoolean(options.removeEmptyColumns, true),
        includeShapeDetails: resolveBoolean(options.includeShapeDetails, true),
        outputMode: normalizeOutputMode(options.outputMode),
        formattingMode: normalizeFormattingMode(options.formattingMode),
        tableDetectionMode: normalizeTableDetectionMode(options.tableDetectionMode),
    };
}

/** Normalizes Markdown options with safe fallbacks. */
export {
    OUTPUT_MODES,
    FORMATTING_MODES,
    FORMATTING_MODE_ALIASES,
    TABLE_DETECTION_MODES,
    TABLE_DETECTION_MODE_ALIASES,
    normalizeOutputMode,
    normalizeFormattingMode,
    normalizeTableDetectionMode,
    resolveMarkdownOptions,
};
