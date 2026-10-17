import type { CellStyleInfo, ParsedSheet, ParsedWorkbook, SheetState } from "../../model/index.js";
import { toSharedStringEntry } from "./shared-strings.js";
import type { ParsedSharedString } from "./shared-strings.js";
type FormatContext = { locale: string | null; date1904: boolean };

type WorkbookLoaderDependencies = {
    unzipEntries: (arrayBuffer: ArrayBuffer) => Promise<Map<string, Uint8Array>>;
    parseSharedStrings: (files: Map<string, Uint8Array>) => ParsedSharedString[];
    parseCellStyles: (files: Map<string, Uint8Array>) => CellStyleInfo[];
    parseRelationships: (files: Map<string, Uint8Array>, relsPath: string, sourcePath: string) => Map<string, string>;
    xmlToDocument: (xmlText: string) => Document;
    decodeXmlText: (bytes: Uint8Array) => string;
    getTextContent: (node: Element | null | undefined) => string;
    parseWorksheet: (
        files: Map<string, Uint8Array>,
        sheetName: string,
        sheetPath: string,
        sheetIndex: number,
        sharedStrings: ParsedSharedString[],
        cellStyles: CellStyleInfo[],
        formatContext: FormatContext
    ) => ParsedSheet;
    readWorkbookLocale?: (files: Map<string, Uint8Array>) => string | null;
    /** False drops hidden sheets. */
    includeHidden?: boolean;
    postProcessWorkbook?: (workbook: ParsedWorkbook) => void;
};

/** Skips Excel-internal `_xlnm` names. Adds a leading `=` when the formula lacks one. */
function parseDefinedNames(
    workbookDoc: Document,
    sheetNames: string[],
    getTextContent: WorkbookLoaderDependencies["getTextContent"]
): {
    name: string;
    formulaText: string;
    localSheetName: string | null;
}[] {
    const result: {
        name: string;
        formulaText: string;
        localSheetName: string | null;
    }[] = [];
    const definedNameElements = Array.from(workbookDoc.getElementsByTagName("definedName"));
    for (const element of definedNameElements) {
        const name = element.getAttribute("name") || "";
        if (!name || name.startsWith("_xlnm.")) continue;
        const formulaText = getTextContent(element).trim();
        if (!formulaText) continue;
        const localSheetIdText = element.getAttribute("localSheetId");
        const localSheetId =
            localSheetIdText == null || localSheetIdText === "" ? Number.NaN : Number(localSheetIdText);
        result.push({
            name,
            formulaText: formulaText.startsWith("=") ? formulaText : `=${formulaText}`,
            localSheetName: Number.isNaN(localSheetId) ? null : sheetNames[localSheetId] || null,
        });
    }
    return result;
}

async function parseWorkbook(
    arrayBuffer: ArrayBuffer,
    workbookName: string,
    deps: WorkbookLoaderDependencies
): Promise<ParsedWorkbook> {
    const files = await deps.unzipEntries(arrayBuffer);
    const workbookBytes = files.get("xl/workbook.xml");
    if (!workbookBytes) {
        throw new Error("xl/workbook.xml was not found.");
    }
    const sharedStrings = deps.parseSharedStrings(files);
    const cellStyles = deps.parseCellStyles(files);
    const rels = deps.parseRelationships(files, "xl/_rels/workbook.xml.rels", "xl/workbook.xml");
    const workbookDoc = deps.xmlToDocument(deps.decodeXmlText(workbookBytes));
    const sheetNodes = Array.from(workbookDoc.getElementsByTagName("sheet"));
    const sheetNames = sheetNodes.map((sheetNode, index) => sheetNode.getAttribute("name") || `Sheet${index + 1}`);
    const definedNames = parseDefinedNames(workbookDoc, sheetNames, deps.getTextContent);
    const date1904Flag = workbookDoc.getElementsByTagName("workbookPr")[0]?.getAttribute("date1904") || "";
    const formatContext: FormatContext = {
        locale: deps.readWorkbookLocale?.(files) ?? null,
        date1904: date1904Flag === "1" || date1904Flag === "true",
    };
    const sheets = sheetNodes.flatMap((sheetNode, index) => {
        const name = sheetNode.getAttribute("name") || `Sheet${index + 1}`;
        const relId = sheetNode.getAttribute("r:id") || "";
        const sheetPath = rels.get(relId) || "";
        const rawState = sheetNode.getAttribute("state") || "";
        const state: SheetState = rawState === "hidden" || rawState === "veryHidden" ? rawState : "visible";
        if (state !== "visible" && deps.includeHidden === false) return [];
        const sheet = deps.parseWorksheet(files, name, sheetPath, index + 1, sharedStrings, cellStyles, formatContext);
        return [{ ...sheet, state }];
    });
    const workbook: ParsedWorkbook = {
        name: workbookName,
        sheets,
        sharedStrings: sharedStrings.map(toSharedStringEntry),
        definedNames,
        dateSystem: formatContext.date1904 ? "1904" : "1900",
        locale: formatContext.locale,
    };
    deps.postProcessWorkbook?.(workbook);
    return workbook;
}

/** Unzips one workbook package into sheets, strings, styles, and defined names. */
export { parseDefinedNames, parseWorkbook };
