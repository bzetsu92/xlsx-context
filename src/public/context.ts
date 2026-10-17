import type { ParsedWorkbook } from "../model/index.js";
import type {
    AiContext,
    AiContextOptions,
    ChunkOptions,
    ContextChunk,
    WorkbookIR,
} from "../outputs/ai-context/types.js";
import { planContextChunks } from "../outputs/ai-context/chunk-planner.js";
import { renderCompactMarkdown } from "../outputs/ai-context/compact-renderer.js";
import {
    createCharEstimateCounter,
    createEncoderCounter,
    createWhitespaceCounter,
} from "../outputs/ai-context/token-counter.js";
import { createTiktokenCounter } from "../outputs/ai-context/tiktoken-counter.js";
import { buildWorkbookIR } from "../outputs/ai-context/workbook-ir.js";

/** Builds the AI-context IR from a parsed workbook. Never parses Markdown. */
export { buildWorkbookIR };

/** Renders the IR as compact Markdown for model context. */
export { renderCompactMarkdown };

/**
 * Creates the built-in whitespace token estimator.
 * It is not a model token count. Pass a model-specific TokenCounter for budgets.
 */
export { createWhitespaceCounter };

/**
 * Splits the IR into retrieval chunks with provenance, within a token budget.
 * Table chunks hold whole rows and repeat the header row.
 */
export function createContextChunks(ir: WorkbookIR, options: ChunkOptions = {}): ContextChunk[] {
    return planContextChunks(ir, options);
}

/**
 * Creates the built-in character token estimator, the default for chunking.
 * Counts CJK, kana, and hangul per character. Still an estimate. Pass a model tokenizer for exact budgets.
 */
export { createCharEstimateCounter };

/**
 * Prepares a parsed workbook for a model in one call: IR, full compact text, and retrieval chunks.
 * maxTokens is the budget per chunk. tokens is the size of the full text under the same counter,
 * so callers can send text whole when it fits their prompt budget, or retrieve chunks otherwise.
 */
export function toAIContext(workbook: ParsedWorkbook, options: AiContextOptions = {}): AiContext {
    const ir = buildWorkbookIR(workbook, options);
    const counter = options.counter || createCharEstimateCounter();
    const text = renderCompactMarkdown(ir, options);
    return { ir, text, tokens: counter.count(text), chunks: createContextChunks(ir, { ...options, counter }) };
}

/**
 * Adapts any model tokenizer into a TokenCounter, for exact chunk budgets.
 * Example: createEncoderCounter({ name: "my-model", encode: (text) => tokenizer.encode(text) }).
 */
export { createEncoderCounter };

/**
 * Creates an exact counter for an OpenAI tiktoken encoding (default o200k_base).
 * Needs the optional peer dependency js-tiktoken, loaded on first call. The result is cached per encoding.
 */
export { createTiktokenCounter };
