import { describe, expect, it } from "vitest";
import * as richTextPlainFormatter from "../../../../../../src/outputs/markdown/rich-text/rich-text-plain-formatter.ts";

function bootRichTextPlainFormatter() {
    return richTextPlainFormatter.createRichTextPlainFormatter();
}

describe("xlsx-context rich text plain formatter", () => {
    it("renders plain tokens without markdown wrappers", () => {
        const api = bootRichTextPlainFormatter();

        expect(
            api.renderPlainTokens([
                { kind: "text", text: "\\# heading" },
                { kind: "lineBreak" },
                {
                    kind: "styledText",
                    parts: [{ kind: "text", text: "bold", rawText: "bold" }],
                    style: { bold: true, italic: false, strike: false, underline: false },
                },
            ])
        ).toBe("\\# heading bold");
    });

    it("renders styled parts as plain text", () => {
        const api = bootRichTextPlainFormatter();

        expect(
            api.renderStyledTextParts([
                { kind: "text", text: "a\\*b", rawText: "a*b" },
                { kind: "escaped", text: "\\_", rawText: "_" },
                { kind: "text", text: " c", rawText: " c" },
            ])
        ).toBe("a\\*b\\_ c");
    });

    it("joins line breaks and repeated spaces conservatively in plain mode", () => {
        const api = bootRichTextPlainFormatter();

        expect(
            api.renderPlainTokens([
                { kind: "text", text: "a  " },
                { kind: "lineBreak" },
                {
                    kind: "styledText",
                    parts: [{ kind: "text", text: "b", rawText: "b" }],
                    style: { bold: false, italic: true, strike: false, underline: false },
                },
                { kind: "text", text: "   c" },
            ])
        ).toBe("a b c");
    });

    it("collapses consecutive line breaks into single-space joins in plain mode", () => {
        const api = bootRichTextPlainFormatter();

        expect(
            api.renderPlainTokens([
                { kind: "text", text: "top" },
                { kind: "lineBreak" },
                { kind: "lineBreak" },
                {
                    kind: "styledText",
                    parts: [{ kind: "text", text: "next", rawText: "next" }],
                    style: { bold: false, italic: false, strike: false, underline: true },
                },
            ])
        ).toBe("top next");
    });

    it("keeps escaped ampersands and image-like text as plain joined text", () => {
        const api = bootRichTextPlainFormatter();

        expect(
            api.renderPlainTokens([
                {
                    kind: "styledText",
                    parts: [
                        { kind: "text", text: "a ", rawText: "a " },
                        { kind: "escaped", text: "&amp;", rawText: "&" },
                        { kind: "text", text: " b", rawText: " b" },
                    ],
                    style: { bold: true, italic: false, strike: false, underline: false },
                },
                { kind: "lineBreak" },
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
                    style: { bold: false, italic: false, strike: false, underline: false },
                },
            ])
        ).toBe("a &amp; b \\!\\[alt\\]\\(img.png\\)");
    });
});
