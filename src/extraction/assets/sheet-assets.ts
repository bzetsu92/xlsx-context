import type { SheetGrid } from "../../model/worksheet.js";

type DrawingHelper = {
    renderShapeSvg?: (
        shapeNode: Element,
        anchor: Element,
        sheetName: string,
        shapeIndex: number
    ) => {
        filename: string;
        path: string;
        data: Uint8Array;
        warning: string | null;
    } | null;
} | null;

type ParseAssetDependencies = {
    parseRelationships: (files: Map<string, Uint8Array>, relsPath: string, sourcePath: string) => Map<string, string>;
    buildRelsPath: (sourcePath: string) => string;
    xmlToDocument: (xmlText: string) => Document;
    decodeXmlText: (bytes: Uint8Array) => string;
    getElementsByLocalName: (root: ParentNode, localName: string) => Element[];
    getFirstDescendantByLocalName: (root: ParentNode, localName: string) => Element | null;
    getDirectChildByLocalName: (root: ParentNode | null, localName: string) => Element | null;
    getTextContent: (node: Element | null | undefined) => string;
    colToLetters: (col: number) => string;
    drawingHelper: DrawingHelper;
    defaultCellWidthEmu: number;
    defaultCellHeightEmu: number;
    shapeBlockGapXEmu: number;
    shapeBlockGapYEmu: number;
    /** Grid of the sheet that owns the drawing. Without it, every cell has the default size. */
    grid?: SheetGrid;
};

/** Replaces unsafe filename characters with underscores. Falls back to `Sheet` for empty names. */
function createSafeSheetAssetDir(sheetName: string): string {
    return sheetName.replace(/[\\/:*?"<>|]+/g, "_").trim() || "Sheet";
}

const ANCHOR_NAMES = new Set(["oneCellAnchor", "twoCellAnchor"]);

/** Collects one-cell anchors before two-cell anchors. Keeps document order within each kind. */
function collectAnchors(drawingDoc: Document): Element[] {
    const one: Element[] = [];
    const two: Element[] = [];
    const stack: Node[] = [drawingDoc];
    while (stack.length > 0) {
        const current = stack.pop() as Node;
        if (current.nodeType === 1 && ANCHOR_NAMES.has((current as Element).localName)) {
            ((current as Element).localName === "oneCellAnchor" ? one : two).push(current as Element);
        }
        for (let child = current.lastChild; child; child = child.previousSibling) stack.push(child);
    }
    return one.concat(two);
}

/** Shared drawing-anchor helpers. Asset kinds live in sheet-images, sheet-charts, and sheet-shapes. */
export { collectAnchors, createSafeSheetAssetDir };
export type { ParseAssetDependencies };
