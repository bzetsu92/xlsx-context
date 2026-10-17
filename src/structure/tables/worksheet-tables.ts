import { createRelsParser } from "../../extraction/workbook/rels-parser.js";
import { decodeXmlText, getElementsByLocalName, xmlToDocument } from "../../extraction/xml/xml-utils.js";
import { parseRangeAddress } from "../addresses/address-utils.js";

type ParsedTable = {
    sheetName: string;
    name: string;
    displayName: string;
    start: string;
    end: string;
    columns: string[];
    headerRowCount: number;
    totalsRowCount: number;
};

const { parseRelationships, buildRelsPath } = createRelsParser({ xmlToDocument, decodeXmlText });

/** Normalizes table names for case-insensitive matching. */
function normalizeStructuredTableKey(value: string): string {
    return String(value || "")
        .normalize("NFKC")
        .trim()
        .toUpperCase();
}

function readCount(value: string | null, fallback: number): number {
    const count = Number(value);
    return value === null || value === "" || !Number.isFinite(count) ? fallback : count;
}

/** Reads declared Excel tables through sheet relationships. Defaults to one header row and no totals row. */
function parseWorksheetTables(
    files: Map<string, Uint8Array>,
    worksheetDoc: Document,
    sheetName: string,
    sheetPath: string
): ParsedTable[] {
    const sheetRels = parseRelationships(files, buildRelsPath(sheetPath), sheetPath);
    const tablePartElements = getElementsByLocalName(worksheetDoc, "tablePart");
    const tables: ParsedTable[] = [];

    for (const tablePartElement of tablePartElements) {
        const relId = tablePartElement.getAttribute("r:id") || tablePartElement.getAttribute("id") || "";
        if (!relId) continue;
        const tablePath = sheetRels.get(relId) || "";
        if (!tablePath) continue;
        const tableBytes = files.get(tablePath);
        if (!tableBytes) continue;
        const tableDoc = xmlToDocument(decodeXmlText(tableBytes));
        const tableElement = getElementsByLocalName(tableDoc, "table")[0] || null;
        if (!tableElement) continue;
        const ref = tableElement.getAttribute("ref") || "";
        const range = parseRangeAddress(ref);
        if (!range) continue;
        const columns = getElementsByLocalName(tableElement, "tableColumn")
            .map((columnElement) => String(columnElement.getAttribute("name") || "").trim())
            .filter(Boolean);
        tables.push({
            sheetName,
            name: tableElement.getAttribute("name") || "",
            displayName: tableElement.getAttribute("displayName") || tableElement.getAttribute("name") || "",
            start: range.start,
            end: range.end,
            columns,
            headerRowCount: readCount(tableElement.getAttribute("headerRowCount"), 1),
            totalsRowCount: readCount(tableElement.getAttribute("totalsRowCount"), 0),
        });
    }

    return tables;
}

/** Reads declared Excel table objects with columns and header counts. */
export { normalizeStructuredTableKey, parseWorksheetTables };
