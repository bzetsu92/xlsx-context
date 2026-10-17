import { describe, expect, it } from "vitest";
import { expectModeResults } from "../../../../../helpers/mode-assertions.js";
import { createRichTextRenderer } from "../../../../../../src/outputs/markdown/rich-text/rich-text-renderer.ts";

function normalizeMarkdownText(text) {
    return String(text || "")
        .replace(/\r\n?|\n/g, " ")
        .replace(/\t/g, " ");
}

function bootRichTextRenderer() {
    return createRichTextRenderer({ normalizeMarkdownText });
}

describe("xlsx-context rich text renderer", () => {
    it("tokenizes rich runs into styledText and lineBreak tokens for gfm mode", () => {
        const api = bootRichTextRenderer();

        expect(
            api.tokenizeGfmRichTextRuns([
                { text: "plain ", bold: false, italic: false, strike: false, underline: false },
                { text: "bold\nnext", bold: true, italic: false, strike: false, underline: false },
            ])
        ).toEqual([
            {
                kind: "styledText",
                parts: [{ kind: "text", text: "plain ", rawText: "plain " }],
                style: { bold: false, italic: false, strike: false, underline: false },
            },
            {
                kind: "styledText",
                parts: [{ kind: "text", text: "bold", rawText: "bold" }],
                style: { bold: true, italic: false, strike: false, underline: false },
            },
            { kind: "lineBreak" },
            {
                kind: "styledText",
                parts: [{ kind: "text", text: "next", rawText: "next" }],
                style: { bold: true, italic: false, strike: false, underline: false },
            },
        ]);
    });

    it("renders tokenized rich text with gfm formatting", () => {
        const api = bootRichTextRenderer();
        const cell = {
            outputValue: "plain bold\nnext",
            textStyle: { bold: false, italic: false, strike: false, underline: false },
            richTextRuns: [
                { text: "plain ", bold: false, italic: false, strike: false, underline: false },
                { text: "bold\nnext", bold: true, italic: false, strike: false, underline: false },
            ],
        };

        expect(api.renderCellDisplayText(cell, "gfm")).toBe("plain **bold**<br>**next**");
    });

    it("falls back to plain tokens in plain mode", () => {
        const api = bootRichTextRenderer();

        expect(
            api.tokenizeCellDisplayText(
                {
                    outputValue: "# head\nline2",
                    textStyle: { bold: true, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
                "plain"
            )
        ).toEqual([{ kind: "text", text: "\\# head line2" }]);
        expect(
            api.renderCellDisplayText(
                {
                    outputValue: "# head\nline2",
                    textStyle: { bold: true, italic: false, strike: false, underline: false },
                    richTextRuns: null,
                },
                "plain"
            )
        ).toBe("\\# head line2");
    });

    it("renders styled text parts before applying wrappers", () => {
        const api = bootRichTextRenderer();

        expect(
            api.renderStyledTextParts([
                { kind: "text", text: "a\\*b" },
                { kind: "text", text: " c" },
            ])
        ).toBe("a\\*b c");
    });

    it("keeps escaped list markers literal inside gfm-styled text", () => {
        const api = bootRichTextRenderer();
        const cell = {
            outputValue: "- item",
            textStyle: { bold: true, italic: false, strike: false, underline: false },
            richTextRuns: null,
        };

        expect(api.renderCellDisplayText(cell, "gfm")).toBe("**\\- item**");
    });

    it("keeps escaped ordered-list and quote markers literal inside gfm-styled text", () => {
        const api = bootRichTextRenderer();

        expect(
            api.renderCellDisplayText(
                {
                    outputValue: "1. item",
                    textStyle: { bold: false, italic: true, strike: false, underline: false },
                    richTextRuns: null,
                },
                "gfm"
            )
        ).toBe("*1\\. item*");

        expect(
            api.renderCellDisplayText(
                {
                    outputValue: "> quote",
                    textStyle: { bold: false, italic: false, strike: true, underline: false },
                    richTextRuns: null,
                },
                "gfm"
            )
        ).toBe("~~&gt; quote~~");
    });

    it("renders escaped markdown link-like text safely inside styled runs", () => {
        const api = bootRichTextRenderer();
        const cell = {
            outputValue: "[x](y) `code` <tag>",
            textStyle: { bold: false, italic: false, strike: false, underline: false },
            richTextRuns: [
                { text: "[x](y) ", bold: true, italic: false, strike: false, underline: false },
                { text: "`code` ", bold: false, italic: true, strike: false, underline: false },
                { text: "<tag>", bold: false, italic: false, strike: false, underline: true },
            ],
        };

        expect(api.renderCellDisplayText(cell, "gfm")).toBe(
            "<strong>\\[x\\]\\(y\\) </strong><em>\\`code\\` </em><ins>&lt;tag&gt;</ins>"
        );
    });

    it("renders escaped ampersands and image-like text across styled rich runs", () => {
        const api = bootRichTextRenderer();
        const cell = {
            outputValue: "a & b ![alt](img.png)",
            textStyle: { bold: false, italic: false, strike: false, underline: false },
            richTextRuns: [
                { text: "a & b ", bold: true, italic: false, strike: false, underline: false },
                { text: "![alt](img.png)", bold: false, italic: true, strike: true, underline: true },
            ],
        };

        expect(api.renderCellDisplayText(cell, "gfm")).toBe(
            "<strong>a &amp; b </strong>*~~<ins>\\!\\[alt\\]\\(img.png\\)</ins>~~*"
        );
    });

    it("renders consecutive line breaks across styled rich runs", () => {
        const api = bootRichTextRenderer();
        const cell = {
            outputValue: "top\n\nnext",
            textStyle: { bold: false, italic: false, strike: false, underline: false },
            richTextRuns: [
                { text: "top\n", bold: true, italic: false, strike: false, underline: false },
                { text: "\nnext", bold: false, italic: true, strike: false, underline: true },
            ],
        };

        expect(api.renderCellDisplayText(cell, "gfm")).toBe("**top**<br><br>*<ins>next</ins>*");
    });

    it("renders plus and star markers literally across styled rich runs", () => {
        const api = bootRichTextRenderer();
        const cell = {
            outputValue: "+ plus * star",
            textStyle: { bold: false, italic: false, strike: false, underline: false },
            richTextRuns: [
                { text: "+ plus ", bold: true, italic: false, strike: false, underline: false },
                { text: "* star", bold: false, italic: true, strike: false, underline: false },
            ],
        };

        expect(api.renderCellDisplayText(cell, "gfm")).toBe("<strong>\\+ plus </strong>*\\* star*");
    });
});
