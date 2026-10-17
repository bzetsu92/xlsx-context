import type { RichTextRun, RichTextStyle, SharedStringEntry } from "../../model/index.js";
import { readContainer } from "../xml/lite-xml-reader.js";
import { ELEMENT_NODE } from "../xml/runtime-env.js";
import { decodeXmlText, getTextContent, hasEnabledBooleanValue, xmlToDocument } from "../xml/xml-utils.js";

/**
 * Rich-text run as the worksheet parser needs it.
 * `ownStyle` is true when the run has `rPr`. Excel then ignores the cell font for that run.
 */
type SharedStringRun = RichTextRun & { ownStyle: boolean };

/** Shared-string entry with run font sources. `toSharedStringEntry` gives the public shape. */
type ParsedSharedString = {
    text: string;
    runs: SharedStringRun[] | null;
};

function hasTextStyle(style: RichTextStyle): boolean {
    return style.bold || style.italic || style.strike || style.underline;
}

function hasSameTextStyle(left: RichTextStyle, right: RichTextStyle): boolean {
    return (
        left.bold === right.bold &&
        left.italic === right.italic &&
        left.strike === right.strike &&
        left.underline === right.underline
    );
}

/**
 * Merges adjacent runs with equal style and font source.
 * Returns null when no run has `rPr`, because the cell font then applies to all text.
 */
function parseRichTextRuns(item: Element): SharedStringRun[] | null {
    const runElements = Array.from(item.childNodes).filter(
        (node): node is Element => node.nodeType === ELEMENT_NODE && (node as Element).localName === "r"
    );
    if (runElements.length === 0) {
        return null;
    }
    const runs: SharedStringRun[] = [];
    for (const runElement of runElements) {
        const text = Array.from(runElement.getElementsByTagName("t"))
            .map((node) => getTextContent(node))
            .join("");
        if (!text) continue;
        const properties = runElement.getElementsByTagName("rPr")[0] || null;
        const run: SharedStringRun = {
            text,
            ownStyle: properties !== null,
            bold: hasEnabledBooleanValue(properties?.getElementsByTagName("b")[0]),
            italic: hasEnabledBooleanValue(properties?.getElementsByTagName("i")[0]),
            strike: hasEnabledBooleanValue(properties?.getElementsByTagName("strike")[0]),
            underline: hasEnabledBooleanValue(properties?.getElementsByTagName("u")[0]),
        };
        const previous = runs[runs.length - 1];
        if (previous && previous.ownStyle === run.ownStyle && hasSameTextStyle(previous, run)) {
            previous.text += run.text;
        } else {
            runs.push(run);
        }
    }
    return runs.some((run) => run.ownStyle) ? runs : null;
}

/** Prefers styled runs, else joins plain text. Skips phonetic hints. */
function parseSharedStringEntry(item: Element): ParsedSharedString {
    const runs = parseRichTextRuns(item);
    if (runs) {
        return {
            text: runs.map((run) => run.text).join(""),
            runs,
        };
    }
    const parts: string[] = [];
    const walk = (node: Node): void => {
        if (node.nodeType === ELEMENT_NODE) {
            const element = node as Element;
            if (element.localName === "rPh" || element.localName === "phoneticPr") {
                return;
            }
            if (element.localName === "t") {
                parts.push(getTextContent(element));
                return;
            }
        }
        for (const child of Array.from(node.childNodes)) {
            walk(child);
        }
    };
    walk(item);
    return {
        text: parts.join(""),
        runs: null,
    };
}

/** Reads the shared-string table. Returns an empty list when the part is missing. */
function parseSharedStrings(files: Map<string, Uint8Array>): ParsedSharedString[] {
    const sharedStringsBytes = files.get("xl/sharedStrings.xml");
    if (!sharedStringsBytes) {
        return [];
    }
    const xmlText = decodeXmlText(sharedStringsBytes);
    // The fast path skips the DOM parser entirely. Lite items implement the Element subset used here.
    const fast = readContainer(xmlText, "sst", "si");
    const items = fast
        ? (fast.items as unknown as Element[])
        : Array.from(xmlToDocument(xmlText).getElementsByTagName("si"));
    return items.map((item) => parseSharedStringEntry(item));
}

/** Drops font sources and unstyled runs, so `ParsedWorkbook.sharedStrings` keeps its public shape. */
function toSharedStringEntry(entry: ParsedSharedString): SharedStringEntry {
    if (!entry.runs || !entry.runs.some(hasTextStyle)) {
        return { text: entry.text, runs: null };
    }
    const runs: RichTextRun[] = [];
    for (const { ownStyle: _ownStyle, ...run } of entry.runs) {
        const previous = runs[runs.length - 1];
        if (previous && hasSameTextStyle(previous, run)) {
            previous.text += run.text;
        } else {
            runs.push(run);
        }
    }
    return { text: entry.text, runs };
}

/** Reads the shared-string table into text entries with rich-text runs. */
export { parseSharedStringEntry, parseSharedStrings, toSharedStringEntry };
export type { ParsedSharedString, SharedStringRun };
