import { ELEMENT_NODE, xmlToDocument } from "./runtime-env.js";

const textDecoder = new TextDecoder("utf-8");

/** Descendant elements named `localName` in document order, collected in one pass without the live NodeList. */
function getElementsByLocalName(root: ParentNode | null, localName: string): Element[] {
    const found: Element[] = [];
    if (!root) return found;
    const stack: Node[] = [];
    let node: Node | null = (root as Node).lastChild;
    while (node) {
        stack.push(node);
        node = node.previousSibling;
    }
    while (stack.length > 0) {
        const current = stack.pop() as Node;
        if (current.nodeType !== ELEMENT_NODE) continue;
        if ((current as Element).localName === localName) found.push(current as Element);
        let child: Node | null = current.lastChild;
        while (child) {
            stack.push(child);
            child = child.previousSibling;
        }
    }
    return found;
}

/** Returns the first descendant element named `localName` in document order. */
function getFirstDescendantByLocalName(root: ParentNode | null, localName: string): Element | null {
    if (!root) return null;
    const stack: Node[] = [];
    let node: Node | null = (root as Node).lastChild;
    while (node) {
        stack.push(node);
        node = node.previousSibling;
    }
    while (stack.length > 0) {
        const current = stack.pop() as Node;
        if (current.nodeType !== ELEMENT_NODE) continue;
        if ((current as Element).localName === localName) return current as Element;
        let child: Node | null = current.lastChild;
        while (child) {
            stack.push(child);
            child = child.previousSibling;
        }
    }
    return null;
}

function getDirectChildByLocalName(root: ParentNode | null, localName: string): Element | null {
    if (!root) return null;
    for (let node: Node | null = (root as Node).firstChild; node; node = node.nextSibling) {
        if (node.nodeType === ELEMENT_NODE && (node as Element).localName === localName) {
            return node as Element;
        }
    }
    return null;
}

/**
 * Reads an OOXML boolean flag such as `<b/>`.
 * Returns true when the flag exists and no value turns it off.
 */
function hasEnabledBooleanValue(node: Element | null | undefined): boolean {
    if (!node) return false;
    const value = (node.getAttribute("val") || "").trim().toLowerCase();
    return value !== "false" && value !== "0" && value !== "none";
}

function decodeXmlText(bytes: Uint8Array): string {
    return textDecoder.decode(bytes);
}

function getTextContent(node: Element | null | undefined): string {
    return (node?.textContent || "").replace(/\r\n/g, "\n");
}

/** Decodes XML bytes and finds elements by local name. */
export {
    xmlToDocument,
    getElementsByLocalName,
    getFirstDescendantByLocalName,
    getDirectChildByLocalName,
    decodeXmlText,
    getTextContent,
    hasEnabledBooleanValue,
};
