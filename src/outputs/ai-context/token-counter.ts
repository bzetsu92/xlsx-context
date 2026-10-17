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

// Splits text the way BPE tokenizers roughly do: letter runs, digit runs, and single symbols.
const TOKEN_PIECE = /[\p{L}\p{M}]+|\p{N}+|[^\s\p{L}\p{M}\p{N}]/gu;

/**
 * Estimates model tokens without a tokenizer. Counts one token per CJK, kana, or hangul
 * character, per four other letters, per three digits, per symbol, and per line break.
 * Symbols and line breaks matter: pipe tables are mostly "|" and short cells.
 * On the test fixtures it reads about 1.1x the o200k_base count, and 1.0x cl100k_base,
 * so chunks rarely go over budget. Still an estimate, not a model token count.
 */
function countCharEstimateTokens(text: string): number {
    const value = String(text || "");
    let count = (value.match(/\n/g) || []).length;
    for (const piece of value.match(TOKEN_PIECE) || []) {
        const dense = (piece.match(DENSE_SCRIPT) || []).length;
        if (dense > 0) {
            count += dense + Math.ceil((piece.length - dense) / 4);
        } else if (/^\p{N}/u.test(piece)) {
            count += Math.ceil(piece.length / 3);
        } else if (/^[\p{L}\p{M}]/u.test(piece)) {
            count += Math.ceil(piece.length / 4);
        } else {
            count += 1;
        }
    }
    return count;
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
