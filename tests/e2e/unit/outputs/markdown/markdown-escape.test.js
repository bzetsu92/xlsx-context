import { describe, expect, it } from "vitest";
import * as markdownEscape from "../../../../../src/outputs/markdown/markdown-escape.ts";

describe("xlsx-context markdown escape", () => {
    it("escapes inline markdown control characters and html-sensitive brackets", () => {
        const api = markdownEscape;

        expect(api.escapeMarkdownLiteralText("a*b _c_ [x](y) ![z](w) <tag> a | b `q` ~")).toBe(
            "a\\*b \\_c\\_ \\[x\\]\\(y\\) \\!\\[z\\]\\(w\\) &lt;tag&gt; a \\| b \\`q\\` \\~"
        );
        expect(api.escapeMarkdownLiteralParts("a*b <tag>")).toEqual([
            { kind: "text", text: "a", rawText: "a" },
            { kind: "escaped", text: "\\*", rawText: "*" },
            { kind: "text", text: "b ", rawText: "b " },
            { kind: "escaped", text: "&lt;", rawText: "<" },
            { kind: "text", text: "tag", rawText: "tag" },
            { kind: "escaped", text: "&gt;", rawText: ">" },
        ]);
    });

    it("escapes line-start markdown markers line by line", () => {
        const api = markdownEscape;

        expect(api.escapeMarkdownLineStart("# h\n- item\r\n1. num\r> quote")).toBe(
            "\\# h\n\\- item\n1\\. num\n\\> quote"
        );
        expect(api.escapeMarkdownLiteralText("# h\n- item\n1. num\n> quote")).toBe(
            "\\# h\n\\- item\n1\\. num\n&gt; quote"
        );
    });

    it("escapes additional list markers and ampersands", () => {
        const api = markdownEscape;

        expect(api.escapeMarkdownLiteralText("+ plus\n* star\na & b")).toBe("\\+ plus\n\\* star\na &amp; b");
        expect(api.escapeMarkdownLiteralParts("a & b")).toEqual([
            { kind: "text", text: "a ", rawText: "a " },
            { kind: "escaped", text: "&amp;", rawText: "&" },
            { kind: "text", text: " b", rawText: " b" },
        ]);
    });

    it("escapes heading names without hurting readability", () => {
        const escape = markdownEscape.escapeMarkdownHeadingText;
        expect(escape("A B-東京&大阪.01")).toBe("A B-東京&大阪.01");
        expect(escape("Q1_2024 (draft)")).toBe("Q1\\_2024 (draft)");
        expect(escape("<img src=x onerror=alert(1)>")).toBe("\\<img src=x onerror=alert(1)\\>");
        expect(escape("[click](javascript:x)")).toBe("\\[click\\](javascript:x)");
        expect(escape("a &copy; b")).toBe("a &amp;copy; b");
        expect(escape("Plan ##")).toBe("Plan \\#\\#");
        expect(escape("line\nbreak")).toBe("line break");
    });
});
