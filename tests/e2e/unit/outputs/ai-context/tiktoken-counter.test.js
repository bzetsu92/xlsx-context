import { afterEach, describe, expect, it, vi } from "vitest";
import { SRC_ROOT } from "../../../../helpers/paths.js";

const TIKTOKEN_MODULE = `${SRC_ROOT}/outputs/ai-context/tiktoken-counter.ts`;

describe("tiktoken counter", () => {
    afterEach(() => {
        vi.doUnmock("js-tiktoken");
        vi.resetModules();
    });

    it("counts exact tiktoken tokens and caches the counter per encoding", async () => {
        const { createTiktokenCounter } = await import(TIKTOKEN_MODULE);
        const counter = await createTiktokenCounter("cl100k_base");

        expect(counter.name).toBe("tiktoken:cl100k_base");
        expect(counter.count("hello world")).toBe(2);
        expect(await createTiktokenCounter("cl100k_base")).toBe(counter);
    });
});
