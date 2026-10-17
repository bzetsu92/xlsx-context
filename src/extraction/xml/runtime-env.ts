const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

let sharedDomParser: DOMParser | null = null;

function getDomParser(): DOMParser {
    if (sharedDomParser) return sharedDomParser;
    if (typeof DOMParser !== "function") {
        throw new Error("This environment does not provide DOMParser.");
    }
    sharedDomParser = new DOMParser();
    return sharedDomParser;
}

function xmlToDocument(xmlText: string): Document {
    return getDomParser().parseFromString(xmlText, "application/xml");
}

/**
 * Returns the Node.js require hook installed by the Node entry, or null elsewhere.
 * Read at call time so load order of entry modules does not matter.
 */
export function getNodeRequire(): ((id: string) => unknown) | null {
    const candidate = (globalThis as { __xlsxContextNodeRequire?: unknown }).__xlsxContextNodeRequire;
    return typeof candidate === "function" ? (candidate as (id: string) => unknown) : null;
}

/** Holds DOM constants and one shared DOM parser. */
export { ELEMENT_NODE, TEXT_NODE, xmlToDocument };
