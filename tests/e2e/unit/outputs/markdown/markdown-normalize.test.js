import { describe, expect, it } from "vitest";
import * as markdownNormalize from "../../../../../src/outputs/markdown/markdown-normalize.ts";

describe("xlsx-context markdown normalize", () => {
    it("replaces line breaks, tabs, and control characters with spaces without trimming", () => {
        const api = markdownNormalize;

        expect(api.normalizeMarkdownText(" a\r\nb\tc\u0007d ")).toBe(" a b c d ");
    });

    it("normalizes markdown newlines without changing other characters", () => {
        const api = markdownNormalize;

        expect(api.normalizeMarkdownNewlines("a\r\nb\rc\nd")).toBe("a\nb\nc\nd");
        expect(api.normalizeMarkdownNewlines("a\r\nb", " / ")).toBe("a / b");
    });

    it("replaces unsafe unicode format and separator characters with spaces", () => {
        const api = markdownNormalize;

        expect(api.normalizeMarkdownText("a\u0085b\u200Bc\u2028d\u202Ee\u2060f\uFEFFg\u00ADh")).toBe("a b c d e f g h");
    });

    it("escapes pipes in table cells after normalization", () => {
        const api = markdownNormalize;

        expect(api.normalizeMarkdownTableCell("A|\nB")).toBe("A\\| B");
    });

    it("does not escape a pipe that literal escaping already escaped", () => {
        const api = markdownNormalize;

        expect(api.normalizeMarkdownTableCell("a \\| b")).toBe("a \\| b");
        // An escaped backslash before a raw pipe still needs the pipe escaped.
        expect(api.normalizeMarkdownTableCell("a\\\\|b")).toBe("a\\\\\\|b");
    });

    it("preserves surrounding spaces while normalizing tabs and pipes in table cells", () => {
        const api = markdownNormalize;

        expect(api.normalizeMarkdownTableCell("  A\t|\tB  ")).toBe("  A \\| B  ");
    });

    it("removes heading and list markers from normalized text", () => {
        const api = markdownNormalize;

        expect(api.normalizeMarkdownHeadingText("### heading")).toBe("heading");
        expect(api.normalizeMarkdownListItemText("- item")).toBe("item");
        expect(api.normalizeMarkdownListItemText("1. item")).toBe("item");
    });
});
