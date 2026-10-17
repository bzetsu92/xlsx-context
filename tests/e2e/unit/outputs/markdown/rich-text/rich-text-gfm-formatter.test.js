import { describe, expect, it } from "vitest";
import * as richTextGfmFormatter from "../../../../../../src/outputs/markdown/rich-text/rich-text-gfm-formatter.ts";

function bootRichTextGfmFormatter() {
    return richTextGfmFormatter.createRichTextGfmFormatter();
}

describe("xlsx-context rich text gfm formatter", () => {
    it("applies gfm-compatible wrappers in stable order", () => {
        const api = bootRichTextGfmFormatter();

        expect(
            api.applyTextStyle("text", {
                bold: true,
                italic: true,
                strike: true,
                underline: true,
            })
        ).toBe("***~~<ins>text</ins>~~***");
    });

    it("renders gfm tokens with <br> and styled parts", () => {
        const api = bootRichTextGfmFormatter();

        expect(
            api.renderGfmTokens([
                { kind: "text", text: "plain " },
                {
                    kind: "styledText",
                    parts: [{ kind: "text", text: "a\\*b", rawText: "a*b" }],
                    style: { bold: true, italic: false, strike: false, underline: false },
                },
                { kind: "lineBreak" },
                {
                    kind: "styledText",
                    parts: [{ kind: "text", text: "next", rawText: "next" }],
                    style: { bold: false, italic: true, strike: false, underline: false },
                },
            ])
        ).toBe("plain **a\\*b**<br>*next*");
    });

    it("renders line breaks across combined styles in stable wrapper order", () => {
        const api = bootRichTextGfmFormatter();

        expect(
            api.renderGfmTokens([
                {
                    kind: "styledText",
                    parts: [{ kind: "text", text: "top", rawText: "top" }],
                    style: { bold: true, italic: false, strike: false, underline: true },
                },
                { kind: "lineBreak" },
                {
                    kind: "styledText",
                    parts: [{ kind: "text", text: "next", rawText: "next" }],
                    style: { bold: false, italic: true, strike: true, underline: false },
                },
            ])
        ).toBe("**<ins>top</ins>**<br>*~~next~~*");
    });

    it("collapses repeated spaces after joining gfm tokens", () => {
        const api = bootRichTextGfmFormatter();

        expect(
            api.renderGfmTokens([
                { kind: "text", text: "a  " },
                {
                    kind: "styledText",
                    parts: [{ kind: "text", text: "b", rawText: "b" }],
                    style: { bold: true, italic: false, strike: false, underline: false },
                },
                { kind: "text", text: "   c" },
            ])
        ).toBe("a **b** c");
    });

    it("renders escaped ampersands and image-like text inside combined styles", () => {
        const api = bootRichTextGfmFormatter();

        expect(
            api.renderGfmTokens([
                {
                    kind: "styledText",
                    parts: [
                        { kind: "text", text: "a ", rawText: "a " },
                        { kind: "escaped", text: "&amp;", rawText: "&" },
                        { kind: "text", text: " b", rawText: " b" },
                    ],
                    style: { bold: true, italic: false, strike: false, underline: false },
                },
                { kind: "text", text: " " },
                {
                    kind: "styledText",
                    parts: [
                        { kind: "escaped", text: "\\!", rawText: "!" },
                        { kind: "escaped", text: "\\[", rawText: "[" },
                        { kind: "text", text: "alt", rawText: "alt" },
                        { kind: "escaped", text: "\\]", rawText: "]" },
                        { kind: "escaped", text: "\\(", rawText: "(" },
                        { kind: "text", text: "img.png", rawText: "img.png" },
                        { kind: "escaped", text: "\\)", rawText: ")" },
                    ],
                    style: { bold: false, italic: true, strike: true, underline: true },
                },
            ])
        ).toBe("**a &amp; b** *~~<ins>\\!\\[alt\\]\\(img.png\\)</ins>~~*");
    });

    it("renders escaped plus and star markers literally inside styled runs", () => {
        const api = bootRichTextGfmFormatter();

        expect(
            api.renderGfmTokens([
                {
                    kind: "styledText",
                    parts: [
                        { kind: "escaped", text: "\\+", rawText: "+" },
                        { kind: "text", text: " plus", rawText: " plus" },
                    ],
                    style: { bold: true, italic: false, strike: false, underline: false },
                },
                { kind: "text", text: " " },
                {
                    kind: "styledText",
                    parts: [
                        { kind: "escaped", text: "\\*", rawText: "*" },
                        { kind: "text", text: " star", rawText: " star" },
                    ],
                    style: { bold: false, italic: true, strike: false, underline: false },
                },
            ])
        ).toBe("**\\+ plus** *\\* star*");
    });

    it("renders escaped line-start markers inside styled gfm tokens", () => {
        const api = bootRichTextGfmFormatter();

        expect(
            api.renderGfmTokens([
                {
                    kind: "styledText",
                    parts: [
                        { kind: "escaped", text: "\\#", rawText: "#" },
                        { kind: "text", text: " head", rawText: " head" },
                    ],
                    style: { bold: false, italic: true, strike: false, underline: false },
                },
            ])
        ).toBe("*\\# head*");
    });

    it("shares open markers across adjacent runs so nesting stays valid", () => {
        const api = bootRichTextGfmFormatter();
        const run = (text, style) => ({
            kind: "styledText",
            parts: [{ kind: "text", text, rawText: text }],
            style: { bold: false, italic: false, strike: false, underline: false, ...style },
        });

        expect(
            api.renderGfmTokens([run("italic+", { italic: true }), run("strike", { italic: true, strike: true })])
        ).toBe("*italic+~~strike~~*");
        expect(api.renderGfmTokens([run("a", { bold: true, italic: true }), run("b", { bold: true })])).toBe(
            "***a*b**"
        );
    });

    it("falls back to HTML when CommonMark would not parse the Markdown delimiters", () => {
        const api = bootRichTextGfmFormatter();
        const run = (text, style = {}) => ({
            kind: "styledText",
            parts: [{ kind: "text", text, rawText: text }],
            style: { bold: false, italic: false, strike: false, underline: false, ...style },
        });

        // Opener between a letter and punctuation.
        expect(api.renderGfmTokens([run("a"), run("\\*b", { bold: true })])).toBe("a<strong>\\*b</strong>");
        // Closer between punctuation and a letter, common in Japanese text.
        expect(api.renderGfmTokens([run("「強調」", { bold: true }), run("です")])).toBe(
            "<strong>「強調」</strong>です"
        );
        // Closer after a space.
        expect(api.renderGfmTokens([run("bold ", { bold: true }), run("x")])).toBe("<strong>bold </strong>x");
        // Valid delimiters stay Markdown.
        expect(api.renderGfmTokens([run("重要、"), run("強調", { bold: true }), run("。")])).toBe("重要、**強調**。");
    });
});
