import type { RichTextStyle } from "../../../model/index.js";

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

type StyleMarker = {
    key: keyof RichTextStyle;
    open: string;
    close: string;
    /** Delimiter character for CommonMark flanking checks. Null for HTML markers. */
    delimiter: "*" | "~" | null;
    /** HTML pair used when the Markdown delimiters would not parse. */
    html: [string, string] | null;
};

// Orders markers outermost first to match applyTextStyle.
const STYLE_MARKERS: readonly StyleMarker[] = [
    { key: "bold", open: "**", close: "**", delimiter: "*", html: ["<strong>", "</strong>"] },
    { key: "italic", open: "*", close: "*", delimiter: "*", html: ["<em>", "</em>"] },
    { key: "strike", open: "~~", close: "~~", delimiter: "~", html: ["<del>", "</del>"] },
    { key: "underline", open: "<ins>", close: "</ins>", delimiter: null, html: null },
];

type RenderPiece =
    { kind: "text"; text: string } | { kind: "marker"; marker: StyleMarker; side: "open" | "close"; useHtml: boolean };

const WHITESPACE_REGEX = /\s/u;
const PUNCTUATION_REGEX = /[\p{P}\p{S}]/u;

function isWhitespaceChar(ch: string): boolean {
    return ch === "" || WHITESPACE_REGEX.test(ch);
}

function isPunctuationChar(ch: string): boolean {
    return PUNCTUATION_REGEX.test(ch);
}

function isLeftFlanking(before: string, after: string): boolean {
    return (
        !isWhitespaceChar(after) && (!isPunctuationChar(after) || isWhitespaceChar(before) || isPunctuationChar(before))
    );
}

function isRightFlanking(before: string, after: string): boolean {
    return (
        !isWhitespaceChar(before) && (!isPunctuationChar(before) || isWhitespaceChar(after) || isPunctuationChar(after))
    );
}

/** Nearest character next to pieces[index], skipping markers that join the same delimiter run. */
function neighborChar(pieces: RenderPiece[], index: number, step: -1 | 1, delimiter: "*" | "~"): string {
    for (let cursor = index + step; cursor >= 0 && cursor < pieces.length; cursor += step) {
        const piece = pieces[cursor];
        const text =
            piece.kind === "text" ? piece.text : piece.side === "open" ? piece.marker.open : piece.marker.close;
        if (piece.kind === "marker" && piece.marker.delimiter === delimiter) continue;
        if (!text) continue;
        return step < 0 ? text[text.length - 1] : text[0];
    }
    return "";
}

/**
 * Switches a marker pair to HTML when CommonMark rejects it as emphasis.
 * Covers letter-punctuation boundaries and CJK text.
 */
function applyHtmlFallback(pieces: RenderPiece[]): void {
    const stack: number[] = [];
    pieces.forEach((piece, index) => {
        if (piece.kind !== "marker" || !piece.marker.delimiter) return;
        if (piece.side === "open") {
            stack.push(index);
            return;
        }
        const openIndex = stack.pop() as number;
        const delimiter = piece.marker.delimiter;
        const opens = isLeftFlanking(
            neighborChar(pieces, openIndex, -1, delimiter),
            neighborChar(pieces, openIndex, 1, delimiter)
        );
        const closes = isRightFlanking(
            neighborChar(pieces, index, -1, delimiter),
            neighborChar(pieces, index, 1, delimiter)
        );
        if (!opens || !closes) {
            (pieces[openIndex] as Extract<RenderPiece, { kind: "marker" }>).useHtml = true;
            piece.useHtml = true;
        }
    });
}

function renderPiece(piece: RenderPiece): string {
    if (piece.kind === "text") return piece.text;
    const { marker, side, useHtml } = piece;
    if (useHtml && marker.html) return side === "open" ? marker.html[0] : marker.html[1];
    return side === "open" ? marker.open : marker.close;
}

function createRichTextGfmFormatter() {
    function applyTextStyle(text: string, style: RichTextStyle): string {
        if (!text) return "";
        let result = text;
        if (style.underline) result = `<ins>${result}</ins>`;
        if (style.strike) result = `~~${result}~~`;
        if (style.italic) result = `*${result}*`;
        if (style.bold) result = `**${result}**`;
        return result;
    }

    function renderStyledTextPart(part: { kind: "text" | "escaped"; text: string; rawText: string }): string {
        if (part.kind === "escaped") {
            return part.text;
        }
        return part.text;
    }

    function renderStyledTextParts(parts: Array<{ kind: "text" | "escaped"; text: string; rawText: string }>): string {
        return parts.map((part) => renderStyledTextPart(part)).join("");
    }

    // Adjacent runs share open markers, so "*a*" + "*~~b~~*" renders "*a~~b~~*" instead of a stray "**".
    // Markers close innermost-first, so the output stays well nested. Line breaks and plain text close all markers.
    function renderGfmTokens(tokens: RichTextToken[]): string {
        if (!tokens.length) return "";
        const pieces: RenderPiece[] = [];
        const openMarkers: StyleMarker[] = [];
        const closeMarkersFrom = (index: number) => {
            while (openMarkers.length > index) {
                pieces.push({
                    kind: "marker",
                    marker: openMarkers.pop() as StyleMarker,
                    side: "close",
                    useHtml: false,
                });
            }
        };
        for (const token of tokens) {
            if (token.kind !== "styledText") {
                closeMarkersFrom(0);
                pieces.push({ kind: "text", text: token.kind === "lineBreak" ? "<br>" : token.text });
                continue;
            }
            const text = renderStyledTextParts(token.parts);
            if (!text) continue;
            const firstUnwanted = openMarkers.findIndex((marker) => !token.style[marker.key]);
            if (firstUnwanted >= 0) closeMarkersFrom(firstUnwanted);
            for (const marker of STYLE_MARKERS) {
                if (token.style[marker.key] && !openMarkers.includes(marker)) {
                    pieces.push({ kind: "marker", marker, side: "open", useHtml: false });
                    openMarkers.push(marker);
                }
            }
            pieces.push({ kind: "text", text });
        }
        closeMarkersFrom(0);
        applyHtmlFallback(pieces);
        return pieces.map(renderPiece).join("").replace(/ {2,}/g, " ").trim();
    }

    return {
        applyTextStyle,
        renderStyledTextPart,
        renderStyledTextParts,
        renderGfmTokens,
    };
}

/** Formats rich-text tokens as GFM with CommonMark-safe emphasis. */
export { createRichTextGfmFormatter };
