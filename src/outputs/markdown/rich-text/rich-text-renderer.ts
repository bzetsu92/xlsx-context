import type { RichTextRun, RichTextStyle } from "../../../model/index.js";
import { createRichTextGfmFormatter } from "./rich-text-gfm-formatter.js";
import { createRichTextParser } from "./rich-text-parser.js";
import { createRichTextPlainFormatter } from "./rich-text-plain-formatter.js";

type ParsedCellLike = {
    outputValue: string;
    textStyle: RichTextStyle;
    richTextRuns: RichTextRun[] | null;
};

type RichTextRendererDeps = {
    normalizeMarkdownText?: (text: string) => string;
};

type RichTextToken =
    | {
          kind: "text";
          text: string;
      }
    | {
          kind: "lineBreak";
      }
    | {
          kind: "styledText";
          parts: Array<{
              kind: "text" | "escaped";
              text: string;
              rawText: string;
          }>;
          style: RichTextStyle;
      };

function createRichTextRenderer(deps: RichTextRendererDeps = {}) {
    const richTextParser = createRichTextParser({
        normalizeMarkdownText: deps.normalizeMarkdownText,
    });
    const plainFormatter = createRichTextPlainFormatter();
    const gfmFormatter = createRichTextGfmFormatter();

    function normalizeGfmSegment(text: string): string {
        return gfmFormatter.renderGfmTokens(richTextParser.splitTextWithLineBreaks(text));
    }

    function normalizeGfmCellText(text: string): string {
        return normalizeGfmSegment(text).replace(/ {2,}/g, " ").trim();
    }

    function renderTokens(tokens: RichTextToken[], formattingMode: "plain" | "gfm"): string {
        if (!tokens.length) return "";
        if (formattingMode !== "gfm") {
            return plainFormatter.renderPlainTokens(tokens);
        }
        return gfmFormatter.renderGfmTokens(tokens);
    }

    function tokenizeCellDisplayText(
        cell: ParsedCellLike | undefined,
        formattingMode: "plain" | "gfm" = "plain"
    ): RichTextToken[] {
        return richTextParser.tokenizeCellDisplayText(cell, formattingMode);
    }

    function renderCellDisplayText(
        cell: ParsedCellLike | undefined,
        formattingMode: "plain" | "gfm" = "plain"
    ): string {
        return renderTokens(tokenizeCellDisplayText(cell, formattingMode), formattingMode);
    }

    return {
        compactText: richTextParser.compactText,
        normalizeGfmSegment,
        normalizeGfmCellText,
        applyTextStyle: gfmFormatter.applyTextStyle,
        renderStyledTextParts: plainFormatter.renderStyledTextParts,
        splitTextWithLineBreaks: richTextParser.splitTextWithLineBreaks,
        tokenizePlainCellText: richTextParser.tokenizePlainCellText,
        tokenizeGfmCellText: richTextParser.tokenizeGfmCellText,
        tokenizeGfmRichTextRuns: richTextParser.tokenizeGfmRichTextRuns,
        tokenizeCellDisplayText,
        renderPlainTokens: plainFormatter.renderPlainTokens,
        renderGfmTokens: gfmFormatter.renderGfmTokens,
        renderTokens,
        renderCellDisplayText,
    };
}

/** Renders rich-text tokens in plain or GFM mode. */
export { createRichTextRenderer };
