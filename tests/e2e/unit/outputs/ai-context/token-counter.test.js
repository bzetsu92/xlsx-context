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
        expect(api.countCharEstimateTokens("hello world")).toBe(3);
        expect(api.countCharEstimateTokens("")).toBe(0);
        expect(api.createCharEstimateCounter().name).toBe("char-estimate");
    });
});
