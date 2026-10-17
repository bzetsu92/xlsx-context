/*
 * Reads one large, repetitive XML container without a generic DOM parser.
 * Generic parsers cost too much here: xmldom builds one RegExp per end tag.
 * Items need only a small DOM subset, so this reader builds exactly that subset
 * (getAttribute, getElementsByTagName, childNodes, textContent) and returns the
 * XML with the container emptied.
 * It mirrors the generic parser on line endings, attribute whitespace, entities,
 * CDATA, comments, and PIs. Outside that envelope it returns null, and the caller
 * uses the generic DOM path instead.
 */

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

const NO_CHILDREN: readonly never[] = Object.freeze([]);

class LiteText {
    readonly nodeType = TEXT_NODE;
    // DOM text nodes expose an empty childNodes list. Tree walkers rely on it.
    readonly childNodes = NO_CHILDREN;
    readonly localName = null;
    constructor(readonly data: string) {}
    get nodeValue(): string {
        return this.data;
    }
    get textContent(): string {
        return this.data;
    }
}

type LiteNode = LiteElement | LiteText;

class LiteElement {
    readonly nodeType = ELEMENT_NODE;
    readonly childNodes: LiteNode[] = [];
    readonly localName: string;

    // Stores attributes as a flat name-value list. Cells carry 1-3 attributes, so a scan beats a Map.
    constructor(
        readonly tagName: string,
        private readonly attributes: string[] | null
    ) {
        const colon = tagName.indexOf(":");
        this.localName = colon >= 0 ? tagName.slice(colon + 1) : tagName;
    }

    get nodeName(): string {
        return this.tagName;
    }

    getAttribute(name: string): string | null {
        const attributes = this.attributes;
        if (!attributes) return null;
        for (let index = 0; index < attributes.length; index += 2) {
            if (attributes[index] === name) return attributes[index + 1];
        }
        return null;
    }

    /** Descendants matching the qualified name (or "*"), in document order, like Element.getElementsByTagName. */
    getElementsByTagName(name: string): LiteElement[] {
        const found: LiteElement[] = [];
        collectByTagName(this, name, found);
        return found;
    }

    get textContent(): string {
        const children = this.childNodes;
        if (children.length === 1) return children[0].textContent;
        let text = "";
        for (const child of children) text += child.textContent;
        return text;
    }
}

function collectByTagName(element: LiteElement, name: string, found: LiteElement[]): void {
    for (const child of element.childNodes) {
        if (child.nodeType !== ELEMENT_NODE) continue;
        if (name === "*" || child.tagName === name) found.push(child);
        if (child.childNodes.length > 0) collectByTagName(child, name, found);
    }
}

const LINE_ENDING_CHARS = /[\r\u0085\u2028\u2029]/;
const NUMERIC_REFERENCE = /^#(?:x[0-9a-fA-F]+|[0-9]+)$/;
// Callers pass fixed OOXML element names. Fixed names need no regex escaping.
const tagPatternCache = new Map<string, { start: RegExp; anywhere: RegExp }>();

function getTagPatterns(name: string): { start: RegExp; anywhere: RegExp } {
    let patterns = tagPatternCache.get(name);
    if (!patterns) {
        patterns = { start: new RegExp(`<${name}(?=[\\s/>])`), anywhere: new RegExp(`<${name}[\\s/>]`) };
        tagPatternCache.set(name, patterns);
    }
    return patterns;
}

// Applies the same normalization as the generic parser.
function normalizeLineEndings(input: string): string {
    return input.replace(/\r[\n\u0085]/g, "\n").replace(/[\r\u0085\u2028\u2029]/g, "\n");
}

function isXmlChar(code: number): boolean {
    return (
        code === 0x9 ||
        code === 0xa ||
        code === 0xd ||
        (code >= 0x20 && code <= 0xd7ff) ||
        (code >= 0xe000 && code <= 0xfffd) ||
        (code >= 0x10000 && code <= 0x10ffff)
    );
}

/** Decodes predefined and numeric entities. Returns null for anything else, and the caller falls back. */
function decodeEntities(text: string): string | null {
    let ampIndex = text.indexOf("&");
    if (ampIndex < 0) return text;
    let out = "";
    let cursor = 0;
    while (ampIndex >= 0) {
        const semiIndex = text.indexOf(";", ampIndex);
        if (semiIndex < 0) return null;
        out += text.slice(cursor, ampIndex);
        const ref = text.slice(ampIndex + 1, semiIndex);
        if (ref === "lt") out += "<";
        else if (ref === "gt") out += ">";
        else if (ref === "amp") out += "&";
        else if (ref === "quot") out += '"';
        else if (ref === "apos") out += "'";
        else if (NUMERIC_REFERENCE.test(ref)) {
            const code = ref[1] === "x" ? parseInt(ref.slice(2), 16) : parseInt(ref.slice(1), 10);
            if (!isXmlChar(code)) return null;
            out += String.fromCodePoint(code);
        } else {
            return null;
        }
        cursor = semiIndex + 1;
        ampIndex = text.indexOf("&", cursor);
    }
    return out + text.slice(cursor);
}

function isWhitespace(code: number): boolean {
    return code === 0x20 || code === 0x9 || code === 0xa || code === 0xd;
}

type StartTag = { tagName: string; attributes: string[] | null; selfClosing: boolean; end: number };

/** Parses a start tag beginning at the given offset. Returns null for malformed tags. */
function readStartTag(xml: string, start: number): StartTag | null {
    const length = xml.length;
    let index = start + 1;
    while (index < length) {
        const code = xml.charCodeAt(index);
        if (isWhitespace(code) || code === 0x2f /* / */ || code === 0x3e /* > */) break;
        index += 1;
    }
    const tagName = xml.slice(start + 1, index);
    if (!tagName) return null;
    let attributes: string[] | null = null;
    while (index < length) {
        let code = xml.charCodeAt(index);
        while (isWhitespace(code)) code = xml.charCodeAt(++index);
        if (code === 0x3e) return { tagName, attributes, selfClosing: false, end: index + 1 };
        if (code === 0x2f) {
            if (xml.charCodeAt(index + 1) !== 0x3e) return null;
            return { tagName, attributes, selfClosing: true, end: index + 2 };
        }
        const nameStart = index;
        while (index < length) {
            code = xml.charCodeAt(index);
            if (code === 0x3d /* = */ || isWhitespace(code) || code === 0x3e || code === 0x2f) break;
            index += 1;
        }
        const name = xml.slice(nameStart, index);
        while (isWhitespace(code)) code = xml.charCodeAt(++index);
        if (!name || code !== 0x3d) return null;
        code = xml.charCodeAt(++index);
        while (isWhitespace(code)) code = xml.charCodeAt(++index);
        if (code !== 0x22 /* " */ && code !== 0x27 /* ' */) return null;
        const valueEnd = xml.indexOf(code === 0x22 ? '"' : "'", index + 1);
        if (valueEnd < 0) return null;
        const rawValue = xml.slice(index + 1, valueEnd);
        if (rawValue.indexOf("<") >= 0) return null;
        const value = decodeEntities(rawValue.replace(/[\t\n\r]/g, " "));
        if (value === null) return null;
        (attributes ||= []).push(name, value);
        index = valueEnd + 1;
        code = xml.charCodeAt(index);
        // Attributes must be separated by whitespace.
        if (!isWhitespace(code) && code !== 0x3e && code !== 0x2f) return null;
    }
    return null;
}

function appendText(parent: LiteElement, text: string): void {
    if (text) parent.childNodes.push(new LiteText(text));
}

export type LiteContainerReadResult = {
    /** Every item element inside the container, in document order (like getElementsByTagName). */
    items: LiteElement[];
    /** The source XML with the container emptied, for the generic parser. */
    xmlWithoutContainer: string;
    /** The container element. Its direct children are available, for example rows around cells. */
    container: LiteElement;
};

/**
 * Reads the first `<containerName>` element and collects its `<itemName>` descendants.
 * Returns null when the XML is outside the supported envelope or an item exists outside the container.
 */
function readContainer(sourceXml: string, containerName: string, itemName: string): LiteContainerReadResult | null {
    if (sourceXml.indexOf("<!DOCTYPE") >= 0) return null;
    const xml = LINE_ENDING_CHARS.test(sourceXml) ? normalizeLineEndings(sourceXml) : sourceXml;
    const startMatch = getTagPatterns(containerName).start.exec(xml);
    if (!startMatch) return null;
    const containerStart = startMatch.index;
    const containerTag = readStartTag(xml, containerStart);
    if (!containerTag) return null;

    const items: LiteElement[] = [];
    let containerEnd = containerTag.end;
    const root = new LiteElement(containerName, null);
    if (!containerTag.selfClosing) {
        const stack: LiteElement[] = [root];
        let index = containerTag.end;
        let closed = false;
        while (index < xml.length) {
            const parent = stack[stack.length - 1];
            const tagStart = xml.indexOf("<", index);
            if (tagStart < 0) return null;
            if (tagStart > index) {
                const text = decodeEntities(xml.slice(index, tagStart));
                if (text === null) return null;
                appendText(parent, text);
            }
            const next = xml.charCodeAt(tagStart + 1);
            if (next === 0x2f /* / */) {
                const tagEnd = xml.indexOf(">", tagStart + 2);
                if (tagEnd < 0) return null;
                const name = xml.slice(tagStart + 2, tagEnd).trimEnd();
                index = tagEnd + 1;
                if (stack.length === 1) {
                    if (name !== containerName) return null;
                    closed = true;
                    containerEnd = index;
                    break;
                }
                if (name !== parent.tagName) return null;
                stack.pop();
            } else if (next === 0x21 /* ! */) {
                if (xml.startsWith("<!--", tagStart)) {
                    const commentEnd = xml.indexOf("-->", tagStart + 4);
                    if (commentEnd < 0) return null;
                    index = commentEnd + 3;
                } else if (xml.startsWith("<![CDATA[", tagStart)) {
                    const cdataEnd = xml.indexOf("]]>", tagStart + 9);
                    if (cdataEnd < 0) return null;
                    appendText(parent, xml.slice(tagStart + 9, cdataEnd));
                    index = cdataEnd + 3;
                } else {
                    return null;
                }
            } else if (next === 0x3f /* ? */) {
                const piEnd = xml.indexOf("?>", tagStart + 2);
                if (piEnd < 0) return null;
                index = piEnd + 2;
            } else {
                const tag = readStartTag(xml, tagStart);
                if (!tag) return null;
                const element = new LiteElement(tag.tagName, tag.attributes);
                parent.childNodes.push(element);
                if (tag.tagName === itemName) items.push(element);
                if (!tag.selfClosing) stack.push(element);
                index = tag.end;
            }
        }
        if (!closed) return null;
    }

    const xmlWithoutContainer = xml.slice(0, containerStart) + `<${containerName}/>` + xml.slice(containerEnd);
    // The generic path would also pick up an item outside the container. This check keeps that behavior.
    if (getTagPatterns(itemName).anywhere.test(xmlWithoutContainer)) return null;
    return { items, xmlWithoutContainer, container: root };
}

/** Reads one large OOXML container without a generic DOM parser. */
export { readContainer };
