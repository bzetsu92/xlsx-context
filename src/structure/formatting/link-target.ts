// Schemes that run code or embed content when a viewer opens the link.
const UNSAFE_LINK_SCHEME = /^(?:javascript|vbscript|data):/i;

// Characters that end or break a Markdown link destination.
const LINK_TARGET_ESCAPES: Record<string, string> = {
    " ": "%20",
    "(": "%28",
    ")": "%29",
    "<": "%3C",
    ">": "%3E",
};

/**
 * Prepares a hyperlink target from the file for a Markdown link destination.
 * Returns null for an empty target or an unsafe scheme. Then write the target as plain text.
 * Workbook links come from untrusted files, so a target must not close the link or run script.
 */
function toMarkdownLinkTarget(target: string): string | null {
    const trimmed = String(target || "").trim();
    // Browsers drop control characters and spaces inside a scheme, so check the compacted form.
    if (!trimmed || UNSAFE_LINK_SCHEME.test(trimmed.replace(/[\s\u0000-\u001f]+/g, ""))) return null;
    return trimmed.replace(/[\s()<>]/g, (ch) => LINK_TARGET_ESCAPES[ch] ?? encodeURIComponent(ch));
}

/** Makes hyperlink targets safe for Markdown link destinations. */
export { toMarkdownLinkTarget };
