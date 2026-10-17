import { readContainer } from "../xml/lite-xml-reader.js";
import { hasEnabledBooleanValue } from "../xml/xml-utils.js";
import { buildSheetGrid, readNormalFont } from "./sheet-grid.js";
import type { ParseAssetDependencies } from "../assets/sheet-assets.js";
import type {
    BorderFlags,
    CachedValueState,
    CellStyleInfo,
    FormulaResolutionSource,
    FormulaResolutionStatus,
    MergeRange,
    ParsedCell,
    ParsedCellComment,
    ParsedChartAsset,
    ParsedImageAsset,
    ParsedShapeAsset,
    ParsedSheet,
    ParsedTable,
    RichTextRun,
    RichTextStyle,
} from "../../model/index.js";
import type { ParsedSharedString, SharedStringRun } from "./shared-strings.js";

type ExtractedCellOutput = {
    valueType: string;
    rawValue: string;
    outputValue: string;
    formulaText: string;
    resolutionStatus: FormulaResolutionStatus;
    resolutionSource: FormulaResolutionSource;
    cachedValueState: CachedValueState;
    richTextRuns: RichTextRun[] | null;
};

type WorksheetParserDependencies = {
    EMPTY_BORDERS: BorderFlags;
    xmlToDocument: (xmlText: string) => Document;
    decodeXmlText: (bytes: Uint8Array) => string;
    getTextContent: (node: Element | null | undefined) => string;
    parseCellAddress: (address: string) => { row: number; col: number };
    parseRangeRef: (ref: string) => MergeRange;
    parseWorksheetTables: (
        files: Map<string, Uint8Array>,
        worksheetDoc: Document,
        sheetName: string,
        sheetPath: string
    ) => ParsedTable[];
    parseDrawingImages: (
        files: Map<string, Uint8Array>,
        sheetName: string,
        sheetPath: string,
        deps: ParseAssetDependencies
    ) => ParsedImageAsset[];
    parseDrawingCharts: (
        files: Map<string, Uint8Array>,
        sheetName: string,
        sheetPath: string,
        deps: ParseAssetDependencies
    ) => ParsedChartAsset[];
    parseDrawingShapes: (
        files: Map<string, Uint8Array>,
        sheetName: string,
        sheetPath: string,
        deps: ParseAssetDependencies
    ) => ParsedShapeAsset[];
    parseRelationshipEntries: (
        files: Map<string, Uint8Array>,
        relsPath: string,
        sourcePath: string
    ) => Map<
        string,
        {
            target: string;
            targetMode: string;
            type: string;
        }
    >;
    buildRelsPath: (sourcePath: string) => string;
    formatCellDisplayValue: (
        rawValue: string,
        cellStyle: CellStyleInfo,
        context?: { locale?: string | null; date1904?: boolean }
    ) => string | null;
    buildAssetDeps: () => ParseAssetDependencies;
    parseShapes?: boolean;
};

/**
 * Expands a range reference into single-cell addresses in row-major order.
 *
 * @param ref A range reference such as A1:B2, or one address.
 */
function expandRangeAddresses(
    ref: string,
    deps: Pick<WorksheetParserDependencies, "parseRangeRef"> & { colToLetters: (col: number) => string }
): string[] {
    const range = deps.parseRangeRef(ref);
    const addresses: string[] = [];
    for (let row = Math.max(1, range.startRow); row <= Math.max(range.startRow, range.endRow); row += 1) {
        for (let col = Math.max(1, range.startCol); col <= Math.max(range.startCol, range.endCol); col += 1) {
            addresses.push(`${deps.colToLetters(col)}${row}`);
        }
    }
    return addresses;
}

type CellPosition = { address: string; row: number; col: number };

// Caps the cell checks for all hyperlink ranges on one sheet. Real files stay far below it.
const MAX_HYPERLINK_CELL_CHECKS = 10_000_000;

/**
 * Collects hyperlinks for every covered cell that exists in the sheet.
 * Reads external targets from relationships and keeps in-sheet locations internal.
 * Visits at most min(range area, cell count) cells per link, so one wide ref
 * such as A1:XFD1048576 cannot expand into billions of addresses.
 *
 * @param getCells Cell positions of the sheet. Called only when the sheet has hyperlinks.
 */
function parseWorksheetHyperlinks(
    files: Map<string, Uint8Array>,
    worksheetDoc: Document,
    sheetPath: string,
    deps: Pick<WorksheetParserDependencies, "parseRelationshipEntries" | "buildRelsPath" | "parseRangeRef"> & {
        colToLetters: (col: number) => string;
    },
    getCells: () => CellPosition[]
): Map<string, ParsedCell["hyperlink"]> {
    const hyperlinks = new Map<string, ParsedCell["hyperlink"]>();
    const hyperlinkNodes = Array.from(worksheetDoc.getElementsByTagName("hyperlink"));
    if (hyperlinkNodes.length === 0) return hyperlinks;
    const relsPath = deps.buildRelsPath(sheetPath);
    const relEntries = deps.parseRelationshipEntries(files, relsPath, sheetPath);
    const cells = getCells();
    const cellAddresses = new Set(cells.map((cell) => cell.address));
    let cellChecks = 0;
    for (const node of hyperlinkNodes) {
        const ref = (node.getAttribute("ref") || "").trim();
        if (!ref) continue;
        const relId = (node.getAttribute("r:id") || node.getAttribute("id") || "").trim();
        const relEntry = relId ? relEntries.get(relId) : null;
        const display = (node.getAttribute("display") || "").trim();
        const tooltip = (node.getAttribute("tooltip") || "").trim();
        const location = (node.getAttribute("location") || "").trim().replace(/^#/, "");
        const rawTarget = relEntry?.target?.trim() || "";
        const kind: "external" | "internal" | null =
            relEntry?.targetMode.toLowerCase() === "external"
                ? "external"
                : location
                  ? "internal"
                  : rawTarget.startsWith("#")
                    ? "internal"
                    : rawTarget
                      ? "external"
                      : null;
        if (!kind) continue;
        const target = kind === "internal" ? location || rawTarget.replace(/^#/, "") : rawTarget;
        if (!target) continue;
        const hyperlink = {
            kind,
            target,
            location: location || (kind === "internal" ? target : ""),
            tooltip,
            display,
        } satisfies NonNullable<ParsedCell["hyperlink"]>;
        const range = deps.parseRangeRef(ref);
        const startRow = Math.max(1, range.startRow);
        const endRow = Math.max(range.startRow, range.endRow);
        const startCol = Math.max(1, range.startCol);
        const endCol = Math.max(range.startCol, range.endCol);
        const area = Math.max(0, endRow - startRow + 1) * Math.max(0, endCol - startCol + 1);
        cellChecks += Math.min(area, cells.length);
        if (cellChecks > MAX_HYPERLINK_CELL_CHECKS) {
            throw new Error(`${sheetPath}: Hyperlink ranges cover too many cells.`);
        }
        if (area <= cells.length) {
            for (const address of expandRangeAddresses(ref, deps)) {
                if (cellAddresses.has(address)) hyperlinks.set(address, hyperlink);
            }
        } else {
            for (const cell of cells) {
                if (cell.row >= startRow && cell.row <= endRow && cell.col >= startCol && cell.col <= endCol) {
                    hyperlinks.set(cell.address, hyperlink);
                }
            }
        }
    }
    return hyperlinks;
}

function extractRichTextPlainText(
    element: Element | null | undefined,
    deps: Pick<WorksheetParserDependencies, "getTextContent">
): string {
    if (!element) return "";
    const textNodes = Array.from(element.getElementsByTagName("t"));
    if (textNodes.length > 0) {
        return textNodes.map((node) => deps.getTextContent(node)).join("");
    }
    return deps.getTextContent(element);
}

/**
 * Reads notes (Excel stores them in comments parts). Drops notes without address or text.
 */
function parseNotesXml(
    xmlText: string,
    deps: Pick<WorksheetParserDependencies, "xmlToDocument" | "getTextContent">
): ParsedCellComment[] {
    const doc = deps.xmlToDocument(xmlText);
    const authors = Array.from(doc.getElementsByTagName("author")).map((node) => deps.getTextContent(node).trim());
    return Array.from(doc.getElementsByTagName("comment"))
        .map((commentElement) => {
            const authorId = Number(commentElement.getAttribute("authorId") || 0);
            return {
                address: (commentElement.getAttribute("ref") || "").trim(),
                kind: "note",
                author: authors[authorId] || "",
                text: extractRichTextPlainText(commentElement.getElementsByTagName("text")[0] || null, deps).trim(),
                dateTime: "",
            } satisfies ParsedCellComment;
        })
        .filter((comment) => !!comment.address && !!comment.text);
}

/** Maps person ids to display names across person parts. */
function parsePersonDisplayNames(
    files: Map<string, Uint8Array>,
    deps: Pick<WorksheetParserDependencies, "xmlToDocument" | "decodeXmlText">
): Map<string, string> {
    const persons = new Map<string, string>();
    for (const [path, bytes] of files.entries()) {
        if (!/^xl\/persons\/person[^/]*\.xml$/u.test(path)) continue;
        const doc = deps.xmlToDocument(deps.decodeXmlText(bytes));
        for (const personElement of Array.from(doc.getElementsByTagName("person"))) {
            const id = (personElement.getAttribute("id") || "").trim();
            const displayName = (personElement.getAttribute("displayName") || "").trim();
            if (id && displayName) {
                persons.set(id, displayName);
            }
        }
    }
    return persons;
}

/** Reads threaded comments. Resolves authors through the person map. */
function parseThreadedCommentsXml(
    xmlText: string,
    persons: Map<string, string>,
    deps: Pick<WorksheetParserDependencies, "xmlToDocument" | "getTextContent">
): ParsedCellComment[] {
    const doc = deps.xmlToDocument(xmlText);
    return Array.from(doc.getElementsByTagName("threadedComment"))
        .map((commentElement) => {
            const personId = (commentElement.getAttribute("personId") || "").trim();
            return {
                address: (commentElement.getAttribute("ref") || "").trim(),
                kind: "threaded",
                author: persons.get(personId) || personId,
                text: extractRichTextPlainText(commentElement, deps).trim(),
                dateTime: (commentElement.getAttribute("dT") || "").trim(),
            } satisfies ParsedCellComment;
        })
        .filter((comment) => !!comment.address && !!comment.text);
}

function isNoteRelationship(type: string): boolean {
    return /\/comments$/u.test(type);
}

function isThreadedCommentRelationship(type: string): boolean {
    return /\/threadedComment$/u.test(type) || /\/threadedComments$/u.test(type);
}

/**
 * Merges notes and threaded comments from all related parts.
 * Sorts by cell address, then by comment kind.
 */
function parseWorksheetComments(
    files: Map<string, Uint8Array>,
    sheetPath: string,
    deps: Pick<
        WorksheetParserDependencies,
        | "xmlToDocument"
        | "decodeXmlText"
        | "getTextContent"
        | "parseRelationshipEntries"
        | "buildRelsPath"
        | "parseCellAddress"
    >
): ParsedCellComment[] {
    const comments: ParsedCellComment[] = [];
    const relsPath = deps.buildRelsPath(sheetPath);
    const relEntries = deps.parseRelationshipEntries(files, relsPath, sheetPath);
    const persons = parsePersonDisplayNames(files, deps);
    for (const entry of relEntries.values()) {
        const bytes = files.get(entry.target);
        if (!bytes) continue;
        const xmlText = deps.decodeXmlText(bytes);
        if (isNoteRelationship(entry.type)) {
            comments.push(...parseNotesXml(xmlText, deps));
        } else if (isThreadedCommentRelationship(entry.type)) {
            comments.push(...parseThreadedCommentsXml(xmlText, persons, deps));
        }
    }
    return comments.sort((left, right) => {
        const leftPos = deps.parseCellAddress(left.address);
        const rightPos = deps.parseCellAddress(right.address);
        if (leftPos.row !== rightPos.row) return leftPos.row - rightPos.row;
        if (leftPos.col !== rightPos.col) return leftPos.col - rightPos.col;
        return left.kind.localeCompare(right.kind);
    });
}

/** Hidden column spans stop here. Excel writes max="16384" for "hide every column to the right". */
const MAX_HIDDEN_SPAN_COL = 16384;

function isTrueFlag(value: string | null): boolean {
    return value === "1" || value === "true";
}

function hasTextStyle(style: RichTextStyle): boolean {
    return style.bold || style.italic || style.strike || style.underline;
}

function parseRichTextStyle(runProperties: Element | null | undefined): RichTextStyle {
    return {
        bold: hasEnabledBooleanValue(runProperties?.getElementsByTagName("b")[0]),
        italic: hasEnabledBooleanValue(runProperties?.getElementsByTagName("i")[0]),
        strike: hasEnabledBooleanValue(runProperties?.getElementsByTagName("strike")[0]),
        underline: hasEnabledBooleanValue(runProperties?.getElementsByTagName("u")[0]),
    };
}

/** Merges adjacent runs with equal style. Returns null when no run carries style. */
function mergeAdjacentRuns(runs: RichTextRun[]): RichTextRun[] | null {
    const merged: RichTextRun[] = [];
    for (const run of runs) {
        if (!run.text) continue;
        const previous = merged[merged.length - 1];
        if (
            previous &&
            previous.bold === run.bold &&
            previous.italic === run.italic &&
            previous.strike === run.strike &&
            previous.underline === run.underline
        ) {
            previous.text += run.text;
        } else {
            merged.push({ ...run });
        }
    }
    return merged.length > 0 && merged.some((run) => hasTextStyle(run)) ? merged : null;
}

function createStyledRuns(text: string, style: RichTextStyle): RichTextRun[] | null {
    if (!text || !hasTextStyle(style)) {
        return null;
    }
    return [
        {
            text,
            ...style,
        },
    ];
}

/**
 * Applies the cell font only to runs without `rPr`, as Excel does.
 * Returns one plain run when every run sets its own plain font, so the cell style cannot apply later.
 */
function resolveRunStyles(runs: SharedStringRun[], cellStyle: RichTextStyle): RichTextRun[] | null {
    const resolved = runs.map((run) => ({
        text: run.text,
        bold: run.ownStyle ? run.bold : cellStyle.bold,
        italic: run.ownStyle ? run.italic : cellStyle.italic,
        strike: run.ownStyle ? run.strike : cellStyle.strike,
        underline: run.ownStyle ? run.underline : cellStyle.underline,
    }));
    const merged = mergeAdjacentRuns(resolved);
    const text = resolved.map((run) => run.text).join("");
    if (merged || !text || !hasTextStyle(cellStyle)) {
        return merged;
    }
    return [{ text, bold: false, italic: false, strike: false, underline: false }];
}

function parseInlineRichTextRuns(
    cellElement: Element,
    cellStyle: RichTextStyle,
    deps: Pick<WorksheetParserDependencies, "getTextContent">
): RichTextRun[] | null {
    const inlineStringElement = cellElement.getElementsByTagName("is")[0] || null;
    if (!inlineStringElement) {
        return null;
    }
    const runElements = Array.from(inlineStringElement.childNodes).filter(
        (node): node is Element => node.nodeType === Node.ELEMENT_NODE && (node as Element).localName === "r"
    );
    if (runElements.length === 0) {
        return null;
    }
    return resolveRunStyles(
        runElements.map((runElement) => {
            const properties = runElement.getElementsByTagName("rPr")[0] || null;
            return {
                text: Array.from(runElement.getElementsByTagName("t"))
                    .map((node) => deps.getTextContent(node))
                    .join(""),
                ownStyle: properties !== null,
                ...parseRichTextStyle(properties),
            };
        }),
        cellStyle
    );
}

/**
 * Splits one cell into raw value, display value, and formula state.
 * Prefers cached formula values and falls back to formula text without them.
 * Marks formulas that point at external workbooks as unsupported.
 *
 * @param formulaOverride Translated shared formula. Wins over the cell formula text.
 */
function extractCellOutputValue(
    cellElement: Element,
    sharedStrings: ParsedSharedString[],
    cellStyle: CellStyleInfo,
    deps: Pick<WorksheetParserDependencies, "getTextContent" | "formatCellDisplayValue">,
    formulaOverride = ""
): ExtractedCellOutput {
    const type = (cellElement.getAttribute("t") || "").trim();
    const valueNode = cellElement.getElementsByTagName("v")[0] || null;
    const valueText = deps.getTextContent(valueNode);
    const formulaText = formulaOverride || deps.getTextContent(cellElement.getElementsByTagName("f")[0]);
    const cachedValueState: CachedValueState = !formulaText
        ? null
        : !valueNode
          ? "absent"
          : valueText === ""
            ? "present_empty"
            : "present_nonempty";
    if (formulaText) {
        const normalizedFormula = formulaText.startsWith("=") ? formulaText : `=${formulaText}`;
        if (/\[[^\]]+\.xlsx\]/i.test(normalizedFormula)) {
            return {
                valueType: type || "formula",
                rawValue: valueText || normalizedFormula,
                outputValue: normalizedFormula,
                formulaText: normalizedFormula,
                resolutionStatus: "unsupported_external",
                resolutionSource: "external_unsupported",
                cachedValueState,
                richTextRuns: null,
            };
        }
        if (valueNode) {
            const formattedValue = deps.formatCellDisplayValue(valueText, cellStyle);
            return {
                valueType: type || "formula",
                rawValue: valueText,
                outputValue: formattedValue ?? valueText,
                formulaText: normalizedFormula,
                resolutionStatus: "resolved",
                resolutionSource: "cached_value",
                cachedValueState,
                richTextRuns: null,
            };
        }
        return {
            valueType: type || "formula",
            rawValue: normalizedFormula,
            outputValue: normalizedFormula,
            formulaText: normalizedFormula,
            resolutionStatus: "fallback_formula",
            resolutionSource: "formula_text",
            cachedValueState,
            richTextRuns: null,
        };
    }

    if (type === "s") {
        const sharedIndex = Number(valueText || 0);
        const sharedEntry = sharedStrings[sharedIndex] || { text: "", runs: null };
        return {
            valueType: type,
            rawValue: valueText,
            outputValue: sharedEntry.text,
            formulaText: "",
            resolutionStatus: null,
            resolutionSource: null,
            cachedValueState: null,
            richTextRuns: sharedEntry.runs
                ? resolveRunStyles(sharedEntry.runs, cellStyle.textStyle)
                : createStyledRuns(sharedEntry.text, cellStyle.textStyle),
        };
    }
    if (type === "inlineStr") {
        const inlineText = Array.from(cellElement.getElementsByTagName("t"))
            .map((node) => deps.getTextContent(node))
            .join("");
        return {
            valueType: type,
            rawValue: inlineText,
            outputValue: inlineText,
            formulaText: "",
            resolutionStatus: null,
            resolutionSource: null,
            cachedValueState: null,
            richTextRuns:
                parseInlineRichTextRuns(cellElement, cellStyle.textStyle, deps) ||
                createStyledRuns(inlineText, cellStyle.textStyle),
        };
    }
    if (type === "b") {
        return {
            valueType: type,
            rawValue: valueText,
            outputValue: valueText === "1" ? "TRUE" : "FALSE",
            formulaText: "",
            resolutionStatus: null,
            resolutionSource: null,
            cachedValueState: null,
            richTextRuns: createStyledRuns(valueText === "1" ? "TRUE" : "FALSE", cellStyle.textStyle),
        };
    }
    if (type === "str" || type === "e") {
        return {
            valueType: type,
            rawValue: valueText,
            outputValue: valueText,
            formulaText: "",
            resolutionStatus: null,
            resolutionSource: null,
            cachedValueState: null,
            richTextRuns: createStyledRuns(valueText, cellStyle.textStyle),
        };
    }
    if (valueText) {
        const formattedValue = deps.formatCellDisplayValue(valueText, cellStyle);
        if (formattedValue !== null) {
            return {
                valueType: type,
                rawValue: valueText,
                outputValue: formattedValue,
                formulaText: "",
                resolutionStatus: null,
                resolutionSource: null,
                cachedValueState: null,
                richTextRuns: createStyledRuns(formattedValue, cellStyle.textStyle),
            };
        }
    }
    return {
        valueType: type,
        rawValue: valueText,
        outputValue: valueText,
        formulaText: "",
        resolutionStatus: null,
        resolutionSource: null,
        cachedValueState: null,
        richTextRuns: createStyledRuns(valueText, cellStyle.textStyle),
    };
}

/** Shifts one A1 address by row and column offsets. Keeps $ parts fixed and clamps at row 1, column 1. */
function shiftReferenceAddress(
    addressText: string,
    rowOffset: number,
    colOffset: number,
    deps: { lettersToCol: (letters: string) => number; colToLetters: (col: number) => string }
): string {
    const match = String(addressText || "").match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i);
    if (!match) return addressText;
    const colAbsolute = match[1] === "$";
    const rowAbsolute = match[3] === "$";
    const baseCol = deps.lettersToCol(match[2]);
    const baseRow = Number(match[4]);
    const shiftedCol = colAbsolute ? baseCol : baseCol + colOffset;
    const shiftedRow = rowAbsolute ? baseRow : baseRow + rowOffset;
    const safeCol = Math.max(1, shiftedCol);
    const safeRow = Math.max(1, shiftedRow);
    return `${colAbsolute ? "$" : ""}${deps.colToLetters(safeCol)}${rowAbsolute ? "$" : ""}${safeRow}`;
}

/** Rewrites a shared base formula for a target cell. Shifts relative references and keeps absolute ones. */
function translateSharedFormula(
    baseFormulaText: string,
    baseAddress: string,
    targetAddress: string,
    deps: {
        parseCellAddress: (address: string) => { row: number; col: number };
        lettersToCol: (letters: string) => number;
        colToLetters: (col: number) => string;
    }
): string {
    const basePos = deps.parseCellAddress(baseAddress);
    const targetPos = deps.parseCellAddress(targetAddress);
    if (!basePos.row || !basePos.col || !targetPos.row || !targetPos.col) {
        return baseFormulaText;
    }
    const rowOffset = targetPos.row - basePos.row;
    const colOffset = targetPos.col - basePos.col;
    const normalized = String(baseFormulaText || "").replace(/^=/, "");
    const translated = normalized.replace(
        /(?:'((?:[^']|'')+)'|([A-Za-z0-9_ ]+))!(\$?[A-Z]+\$?\d+)|(\$?[A-Z]+\$?\d+)/g,
        (full, quotedSheet, plainSheet, qualifiedAddress, localAddress) => {
            const address = qualifiedAddress || localAddress;
            if (!address) return full;
            const shifted = shiftReferenceAddress(address, rowOffset, colOffset, deps);
            if (qualifiedAddress) {
                const sheetPrefix = quotedSheet ? `'${quotedSheet}'` : plainSheet;
                return `${sheetPrefix}!${shifted}`;
            }
            return shifted;
        }
    );
    return translated.startsWith("=") ? translated : `=${translated}`;
}

/**
 * Parses one worksheet part into cells, tables, assets, and comments.
 * Reads cell data through the fast container path when possible.
 *
 * @param sheetIndex 1-based sheet position.
 */
function parseWorksheet(
    files: Map<string, Uint8Array>,
    sheetName: string,
    sheetPath: string,
    sheetIndex: number,
    sharedStrings: ParsedSharedString[],
    cellStyles: CellStyleInfo[],
    sheetDeps: WorksheetParserDependencies & {
        lettersToCol: (letters: string) => number;
        colToLetters: (col: number) => string;
        /** False drops cells in hidden rows and columns. */
        includeHidden?: boolean;
    },
    formatContext: { locale?: string | null; date1904?: boolean } = {}
): ParsedSheet {
    const deps = {
        ...sheetDeps,
        formatCellDisplayValue: (rawValue: string, cellStyle: CellStyleInfo) =>
            sheetDeps.formatCellDisplayValue(rawValue, cellStyle, formatContext),
    };
    const bytes = files.get(sheetPath);
    if (!bytes) {
        throw new Error(`Sheet XML not found: ${sheetPath}`);
    }
    const xmlText = deps.decodeXmlText(bytes);
    const sheetData = readContainer(xmlText, "sheetData", "c");
    const doc = deps.xmlToDocument(sheetData ? sheetData.xmlWithoutContainer : xmlText);
    const cellElements = sheetData
        ? (sheetData.items as unknown as Element[])
        : Array.from(doc.getElementsByTagName("c"));
    const sharedFormulaMap = new Map<string, { address: string; formulaText: string }>();
    const hyperlinks = parseWorksheetHyperlinks(files, doc, sheetPath, deps, () =>
        cellElements.map((cellElement) => {
            const address = cellElement.getAttribute("r") || "";
            return { address, ...deps.parseCellAddress(address) };
        })
    );
    const parsedCells = cellElements.map((cellElement) => {
        const address = cellElement.getAttribute("r") || "";
        const position = deps.parseCellAddress(address);
        const styleIndex = Number(cellElement.getAttribute("s") || 0);
        const cellStyle = cellStyles[styleIndex] || {
            borders: deps.EMPTY_BORDERS,
            numFmtId: 0,
            formatCode: "General",
            textStyle: {
                bold: false,
                italic: false,
                strike: false,
                underline: false,
            },
        };
        let formulaOverride = "";
        const formulaElement = cellElement.getElementsByTagName("f")[0] || null;
        const formulaType = formulaElement?.getAttribute("t") || "";
        const spillRef = formulaElement?.getAttribute("ref") || "";
        const sharedIndex = formulaElement?.getAttribute("si") || "";
        const formulaText = deps.getTextContent(formulaElement);
        if (formulaType === "shared" && sharedIndex) {
            if (formulaText) {
                const normalizedFormula = formulaText.startsWith("=") ? formulaText : `=${formulaText}`;
                sharedFormulaMap.set(sharedIndex, { address, formulaText: normalizedFormula });
                formulaOverride = normalizedFormula;
            } else {
                const sharedBase = sharedFormulaMap.get(sharedIndex);
                if (sharedBase) {
                    formulaOverride = translateSharedFormula(sharedBase.formulaText, sharedBase.address, address, deps);
                }
            }
        }
        const output = extractCellOutputValue(cellElement, sharedStrings, cellStyle, deps, formulaOverride);
        return {
            address,
            row: position.row,
            col: position.col,
            valueType: output.valueType,
            rawValue: output.rawValue,
            outputValue: output.outputValue,
            formulaText: output.formulaText,
            resolutionStatus: output.resolutionStatus,
            resolutionSource: output.resolutionSource,
            cachedValueState: output.cachedValueState,
            styleIndex,
            borders: cellStyle.borders,
            numFmtId: cellStyle.numFmtId,
            formatCode: cellStyle.formatCode,
            textStyle: cellStyle.textStyle,
            ...(cellStyle.alignment ? { alignment: cellStyle.alignment } : {}),
            richTextRuns: output.richTextRuns,
            formulaType,
            spillRef,
            hyperlink: hyperlinks.get(address) || null,
        } satisfies ParsedCell;
    });
    const colNodes = Array.from(doc.getElementsByTagName("col"));
    const hiddenCols: number[] = [];
    for (const node of colNodes) {
        if (!isTrueFlag(node.getAttribute("hidden"))) continue;
        const min = Number(node.getAttribute("min") || 0);
        const max = Number(node.getAttribute("max") || min);
        for (let col = min; col >= 1 && col <= max && col <= MAX_HIDDEN_SPAN_COL; col += 1) hiddenCols.push(col);
    }
    const merges = Array.from(doc.getElementsByTagName("mergeCell")).map((mergeElement) =>
        deps.parseRangeRef(mergeElement.getAttribute("ref") || "")
    );
    const tables = deps.parseWorksheetTables(files, doc, sheetName, sheetPath);
    // The fast path keeps rows inside the container. The generic path keeps them in the document.
    const rowNodes = sheetData
        ? (sheetData.container.childNodes.filter(
              (node) => node.nodeType === 1 && node.localName === "row"
          ) as unknown as Element[])
        : Array.from(doc.getElementsByTagName("row"));
    const hiddenRows: number[] = [];
    let previousRow = 0;
    for (const node of rowNodes) {
        const row = Number(node.getAttribute("r")) || previousRow + 1;
        previousRow = row;
        if (isTrueFlag(node.getAttribute("hidden"))) hiddenRows.push(row);
    }
    const hiddenRowSet = new Set(hiddenRows);
    const hiddenColSet = new Set(hiddenCols);
    const cells =
        deps.includeHidden === false && (hiddenRows.length > 0 || hiddenCols.length > 0)
            ? parsedCells.filter((cell) => !hiddenRowSet.has(cell.row) && !hiddenColSet.has(cell.col))
            : parsedCells;
    const grid = buildSheetGrid(
        doc.getElementsByTagName("sheetFormatPr")[0] || null,
        colNodes,
        rowNodes,
        readNormalFont(files, deps.decodeXmlText)
    );
    const assetDeps = { ...deps.buildAssetDeps(), grid };
    const images = deps.parseDrawingImages(files, sheetName, sheetPath, assetDeps);
    const charts = deps.parseDrawingCharts(files, sheetName, sheetPath, assetDeps);
    const shapes = deps.parseShapes === false ? [] : deps.parseDrawingShapes(files, sheetName, sheetPath, assetDeps);
    const comments = parseWorksheetComments(files, sheetPath, deps);
    let maxRow = 0;
    let maxCol = 0;
    for (const cell of cells) {
        if (cell.row > maxRow) maxRow = cell.row;
        if (cell.col > maxCol) maxCol = cell.col;
    }
    for (const merge of merges) {
        if (merge.endRow > maxRow) maxRow = merge.endRow;
        if (merge.endCol > maxCol) maxCol = merge.endCol;
    }
    return {
        name: sheetName,
        index: sheetIndex,
        path: sheetPath,
        cells,
        merges,
        tables,
        images,
        charts,
        shapes,
        comments,
        maxRow,
        maxCol,
        grid,
        hiddenRows: hiddenRows.filter((row) => row <= maxRow),
        hiddenCols: hiddenCols.filter((col) => col <= maxCol),
    };
}

/** Reads one worksheet part into cells, tables, assets, and comments. */
export {
    extractCellOutputValue,
    expandRangeAddresses,
    parseWorksheetHyperlinks,
    parseNotesXml,
    parsePersonDisplayNames,
    parseThreadedCommentsXml,
    parseWorksheetComments,
    shiftReferenceAddress,
    translateSharedFormula,
    parseWorksheet,
};
