import { DOMParser, Document, Element, Node, XMLSerializer } from "@xmldom/xmldom";

/*
 * Workers have no DOMParser, so install @xmldom/xmldom only when the host lacks one.
 * Blob streams and DecompressionStream exist in workers, so no require hook is needed.
 */
const host = globalThis as Record<string, unknown>;
if (typeof host.DOMParser !== "function") {
    host.DOMParser = DOMParser;
    host.Node = Node;
    host.Document = Document;
    host.Element = Element;
    host.XMLSerializer ??= XMLSerializer;
}
