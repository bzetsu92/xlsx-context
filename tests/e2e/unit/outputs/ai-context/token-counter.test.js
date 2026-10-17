import { describe, expect, it } from "vitest";
import * as aiTokenCounter from "../../../../../src/outputs/ai-context/token-counter.ts";

describe("ai token counter", () => {
    it("adapts any encoder by encoded length", () => {
        const api = aiTokenCounter;
        const counter = api.createEncoderCounter({
            name: "chars",
            encode: (text) => Array.from(text),
        });

        expect(counter.name).toBe("chars");
        expect(counter.count("日本語 ok")).toBe(6);
        expect(counter.count("")).toBe(0);
    });
});

describe("char estimate counter", () => {
    it("counts unspaced CJK text per character instead of collapsing it", () => {
        const api = aiTokenCounter;
        expect(api.countCharEstimateTokens("画像とグラフの共存サンプル")).toBe(13);
        expect(api.countWhitespaceTokens("画像とグラフの共存サンプル")).toBe(1);
        expect(api.countCharEstimateTokens("hello world")).toBe(4);
        expect(api.countCharEstimateTokens("")).toBe(0);
        expect(api.createCharEstimateCounter().name).toBe("char-estimate");
    });
});

describe("char estimate counter on table text", () => {
    it("counts symbols, digit groups, and line breaks, as BPE tokenizers split them", () => {
        const count = aiTokenCounter.countCharEstimateTokens;
        // "|" x3 + "1" + "Taro"
        expect(count("| 1 | Taro |")).toBe(5);
        // Digit runs split into groups of three: 1234567 -> 3 tokens; "," and "." are symbols.
        expect(count("1234567")).toBe(3);
        expect(count("13,568.25")).toBe(5);
        // One token per line break.
        expect(count("| a |\n| b |")).toBe(count("| a |") + count("| b |") + 1);
        // Empty cells still cost their pipes.
        expect(count("|  |  |  |")).toBe(4);
    });
});
