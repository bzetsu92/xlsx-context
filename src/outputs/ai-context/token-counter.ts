import type { TextEncoderLike, TokenCounter } from "./types.js";

/**
 * Counts whitespace-separated tokens in text.
 * Use only for rough chunking. It is not a model token count.
 */
function countWhitespaceTokens(text: string): number {
    const matches = String(text || "").match(/\S+/g);
    return matches ? matches.length : 0;
}

/** Creates the built-in whitespace counter. */
function createWhitespaceCounter(): TokenCounter {
    return {
        name: "whitespace",
        count: countWhitespaceTokens,
    };
}

// Covers scripts without spaces between words: CJK ideographs, kana, hangul.
const DENSE_SCRIPT = /[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uf900-\ufaff\uff66-\uff9f]/g;

/**
 * Estimates tokens from characters: one per CJK/kana/hangul character,
 * one per four other non-space characters. Unlike the whitespace count,
 * it does not collapse unspaced Japanese or Chinese text into one token.
 * Still an estimate, not a model token count.
 */
function countCharEstimateTokens(text: string): number {
    const value = String(text || "");
    const dense = (value.match(DENSE_SCRIPT) || []).length;
    const other = value.replace(DENSE_SCRIPT, "").replace(/\s+/g, "").length;
    return dense + Math.ceil(other / 4);
}

/** Creates the built-in character counter. */
function createCharEstimateCounter(): TokenCounter {
    return {
        name: "char-estimate",
        count: countCharEstimateTokens,
    };
}

/** Adapts any tokenizer encoder into a TokenCounter. The count is the encoded length. */
function createEncoderCounter(encoder: TextEncoderLike): TokenCounter {
    return {
        name: encoder.name,
        count: (text: string) => encoder.encode(String(text || "")).length,
    };
}

/** Counts tokens with whitespace and character estimators. Adapts model encoders. */
export {
    countWhitespaceTokens,
    createWhitespaceCounter,
    countCharEstimateTokens,
    createCharEstimateCounter,
    createEncoderCounter,
};
