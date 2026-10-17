import type { TiktokenEncodingName, TokenCounter } from "./types.js";
import { createEncoderCounter } from "./token-counter.js";

type TiktokenModule = { getEncoding: (encoding: string) => { encode: (text: string) => number[] } };

// Caches one loaded encoder per encoding. Building the rank table is the expensive part.
const counters = new Map<string, Promise<TokenCounter>>();

async function loadTiktoken(): Promise<TiktokenModule> {
    try {
        return (await import("js-tiktoken")) as unknown as TiktokenModule;
    } catch (_error) {
        throw new Error(
            "createTiktokenCounter needs the optional peer dependency js-tiktoken. Run `npm install js-tiktoken`."
        );
    }
}

/**
 * Creates an exact token counter for an OpenAI tiktoken encoding, for example o200k_base.
 * Loads js-tiktoken on first use, so it stays out of the bundle for callers that never ask for it.
 * The returned counter is synchronous and cached per encoding.
 */
function createTiktokenCounter(encoding: TiktokenEncodingName = "o200k_base"): Promise<TokenCounter> {
    let counter = counters.get(encoding);
    if (!counter) {
        counter = loadTiktoken().then((tiktoken) => {
            const encoder = tiktoken.getEncoding(encoding);
            return createEncoderCounter({
                name: `tiktoken:${encoding}`,
                encode: (text) => encoder.encode(text),
            });
        });
        // A failed load is not cached, so installing the package later works without a restart.
        counter.catch(() => counters.delete(encoding));
        counters.set(encoding, counter);
    }
    return counter;
}

/** Creates exact token counters from tiktoken encodings. */
export { createTiktokenCounter };
