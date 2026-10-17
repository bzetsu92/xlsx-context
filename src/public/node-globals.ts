import { createRequire } from "node:module";

import { DOMParser, Document, Element, Node, XMLSerializer } from "@xmldom/xmldom";

/*
 * Node has no DOMParser, so install @xmldom/xmldom only when the host lacks one.
 * The require hook lets the core use node:zlib instead of the slower stream path.
 */
const host = globalThis as Record<string, unknown>;
if (typeof host.DOMParser !== "function") {
    host.DOMParser = DOMParser;
    host.Node = Node;
    host.Document = Document;
    host.Element = Element;
    host.XMLSerializer ??= XMLSerializer;
}
host.__xlsxContextNodeRequire ??= createRequire(import.meta.url);
