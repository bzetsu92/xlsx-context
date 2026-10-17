import type { BorderFlags, CellAlignment, CellStyleInfo } from "../../model/index.js";
import { decodeXmlText, hasEnabledBooleanValue, xmlToDocument } from "../xml/xml-utils.js";
import { builtinFormatCode, inferWorkbookLocale } from "./workbook-locale.js";

const EMPTY_BORDERS: BorderFlags = {
    top: false,
    bottom: false,
    left: false,
    right: false,
};
const EMPTY_TEXT_STYLE = {
    bold: false,
    italic: false,
    strike: false,
    underline: false,
};

/** A border side counts when it carries a style attribute or child elements. */
function hasBorderSide(side: Element | null): boolean {
    if (!side) return false;
    return side.hasAttribute("style") || side.children.length > 0;
}

function parseFontStyle(fontElement: Element | null | undefined): typeof EMPTY_TEXT_STYLE {
    return {
        bold: hasEnabledBooleanValue(fontElement?.getElementsByTagName("b")[0]),
        italic: hasEnabledBooleanValue(fontElement?.getElementsByTagName("i")[0]),
        strike: hasEnabledBooleanValue(fontElement?.getElementsByTagName("strike")[0]),
        underline: hasEnabledBooleanValue(fontElement?.getElementsByTagName("u")[0]),
    };
}

/** Reads the alignment child of one xf. Returns undefined when the xf has none. */
function parseAlignment(xfElement: Element): CellAlignment | undefined {
    const alignment = xfElement.getElementsByTagName("alignment")[0];
    if (!alignment) return undefined;
    const flag = (name: string) => {
        const raw = alignment.getAttribute(name);
        return raw === "1" || raw === "true";
    };
    return {
        horizontal: alignment.getAttribute("horizontal") || "general",
        contained: flag("wrapText") || flag("shrinkToFit"),
    };
}

type ParsedStyles = { styles: CellStyleInfo[]; locale: string | null };

const parsedStylesCache = new WeakMap<Uint8Array, ParsedStyles>();

/** Reads borders, fonts, and number formats. Falls back to one General style when style data is missing. */
function parseCellStyles(files: Map<string, Uint8Array>): CellStyleInfo[] {
    return readStyles(files).styles;
}

/** Workbook locale that the styles part shows. Null when the file does not say. */
function readWorkbookLocale(files: Map<string, Uint8Array>): string | null {
    return readStyles(files).locale;
}

function readStyles(files: Map<string, Uint8Array>): ParsedStyles {
    const stylesBytes = files.get("xl/styles.xml");
    if (!stylesBytes) return { styles: parseStyleBytes(null).styles, locale: null };
    let parsed = parsedStylesCache.get(stylesBytes);
    if (!parsed) {
        parsed = parseStyleBytes(stylesBytes);
        parsedStylesCache.set(stylesBytes, parsed);
    }
    return parsed;
}

function parseStyleBytes(stylesBytes: Uint8Array | null): ParsedStyles {
    if (!stylesBytes) {
        return {
            locale: null,
            styles: [
                {
                    borders: EMPTY_BORDERS,
                    numFmtId: 0,
                    formatCode: "General",
                    textStyle: EMPTY_TEXT_STYLE,
                },
            ],
        };
    }
    const doc = xmlToDocument(decodeXmlText(stylesBytes));
    const borderElements = Array.from(doc.getElementsByTagName("border"));
    const borders = borderElements.map((borderElement) => {
        const top = borderElement.getElementsByTagName("top")[0] || null;
        const bottom = borderElement.getElementsByTagName("bottom")[0] || null;
        const left = borderElement.getElementsByTagName("left")[0] || null;
        const right = borderElement.getElementsByTagName("right")[0] || null;
        return {
            top: hasBorderSide(top),
            bottom: hasBorderSide(bottom),
            left: hasBorderSide(left),
            right: hasBorderSide(right),
        };
    });
    const fontElements = Array.from(doc.getElementsByTagName("font"));
    const fontStyles = fontElements.map((fontElement) => parseFontStyle(fontElement));
    const defaultFontName = fontElements[0]?.getElementsByTagName("name")[0]?.getAttribute("val") || "";
    const numFmtMap = new Map<number, string>();
    const numFmtParent = doc.getElementsByTagName("numFmts")[0];
    if (numFmtParent) {
        for (const numFmtElement of Array.from(numFmtParent.getElementsByTagName("numFmt"))) {
            const numFmtId = Number(numFmtElement.getAttribute("numFmtId") || 0);
            const formatCode = numFmtElement.getAttribute("formatCode") || "";
            if (!Number.isNaN(numFmtId) && formatCode) {
                numFmtMap.set(numFmtId, formatCode);
            }
        }
    }

    const locale = inferWorkbookLocale([...numFmtMap.values()], defaultFontName);
    const xfsParent = doc.getElementsByTagName("cellXfs")[0];
    if (!xfsParent) {
        return {
            locale,
            styles: [
                {
                    borders: borders[0] || EMPTY_BORDERS,
                    numFmtId: 0,
                    formatCode: "General",
                    textStyle: fontStyles[0] || EMPTY_TEXT_STYLE,
                },
            ],
        };
    }
    const xfElements = Array.from(xfsParent.getElementsByTagName("xf"));
    const styles = xfElements.map((xfElement) => {
        const borderId = Number(xfElement.getAttribute("borderId") || 0);
        const numFmtId = Number(xfElement.getAttribute("numFmtId") || 0);
        const fontId = Number(xfElement.getAttribute("fontId") || 0);
        return {
            borders: borders[borderId] || EMPTY_BORDERS,
            numFmtId,
            formatCode: numFmtMap.get(numFmtId) || builtinFormatCode(numFmtId, locale) || "General",
            textStyle: fontStyles[fontId] || EMPTY_TEXT_STYLE,
            alignment: parseAlignment(xfElement),
        };
    });
    return {
        locale,
        styles:
            styles.length > 0
                ? styles
                : [
                      {
                          borders: EMPTY_BORDERS,
                          numFmtId: 0,
                          formatCode: "General",
                          textStyle: EMPTY_TEXT_STYLE,
                      },
                  ],
    };
}

/** Reads borders, fonts, and number formats from styles.xml. */
export { EMPTY_BORDERS, EMPTY_TEXT_STYLE, hasBorderSide, parseFontStyle, parseCellStyles, readWorkbookLocale };
