import type { ParsedShapeAsset, ParsedShapeConnection, SheetGrid } from "../../model/index.js";
import { columnAtEmu, createUniformGrid, rowAtEmu } from "../workbook/sheet-grid.js";
import type { ParseAssetDependencies } from "./sheet-assets.js";
import { ELEMENT_NODE, TEXT_NODE } from "../xml/runtime-env.js";
import { collectAnchors } from "./sheet-assets.js";
import {
    applyGeometryToShape,
    calloutTip,
    connectorEndpoints,
    fitBoxToAnchor,
    mapBoxThroughGroups,
    normalizeShapeGeometry,
    readChildBox,
    readGroupFrame,
} from "./drawing-geometry.js";
import type { GroupFrame } from "./drawing-geometry.js";
import { partitionShapeRegions } from "./shape-regions.js";
import { readShapeLook, readThemeColors } from "./shape-look.js";

type ShapeBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    shapeIndexes: number[];
};

/** Labels shapes by type: connectors, text boxes, rectangles, or preset geometry. */
function parseShapeKind(shapeNode: Element | null, deps: ParseAssetDependencies): string {
    if (!shapeNode) return "Shape";
    if (shapeNode.localName === "cxnSp") {
        const geomNode = deps.getFirstDescendantByLocalName(
            deps.getFirstDescendantByLocalName(shapeNode, "spPr") || shapeNode,
            "prstGeom"
        );
        const prst = String(geomNode?.getAttribute("prst") || "").trim();
        if (prst === "straightConnector1") {
            return "Straight Arrow Connector";
        }
        return prst ? `Connector (${prst})` : "Connector";
    }
    if (shapeNode.localName !== "sp") {
        return "Shape";
    }
    const nvSpPr = deps.getFirstDescendantByLocalName(shapeNode, "nvSpPr");
    const cNvSpPr = deps.getFirstDescendantByLocalName(nvSpPr || shapeNode, "cNvSpPr");
    if (cNvSpPr?.getAttribute("txBox") === "1") {
        return "Text Box";
    }
    const geomNode = deps.getFirstDescendantByLocalName(
        deps.getFirstDescendantByLocalName(shapeNode, "spPr") || shapeNode,
        "prstGeom"
    );
    const prst = String(geomNode?.getAttribute("prst") || "").trim();
    if (prst === "rect") {
        return "Rectangle";
    }
    return prst ? `Shape (${prst})` : "Shape";
}

/** True when a line end element draws an arrowhead. */
function hasArrowHead(endNode: Element | null): boolean {
    const type = String(endNode?.getAttribute("type") || "").trim();
    return type !== "" && type !== "none";
}

// Presets that draw a line. Older files and some writers draw arrows as plain shapes with these presets.
const LINE_PRESETS = new Set([
    "line",
    "straightConnector1",
    "bentConnector2",
    "bentConnector3",
    "bentConnector4",
    "bentConnector5",
    "curvedConnector2",
    "curvedConnector3",
    "curvedConnector4",
    "curvedConnector5",
]);

/** True for a plain shape that draws a line. Such a shape has no connection ids. */
function isLineShape(shapeNode: Element, deps: ParseAssetDependencies): boolean {
    const geomNode = deps.getDirectChildByLocalName(deps.getDirectChildByLocalName(shapeNode, "spPr"), "prstGeom");
    return LINE_PRESETS.has(String(geomNode?.getAttribute("prst") || "").trim());
}

/** Reads the shapes a connector or line joins and its arrowheads. Returns null for other shapes. */
function parseShapeConnection(shapeNode: Element | null, deps: ParseAssetDependencies): ParsedShapeConnection | null {
    if (!shapeNode) return null;
    if (shapeNode.localName !== "cxnSp" && !(shapeNode.localName === "sp" && isLineShape(shapeNode, deps))) return null;
    const nvCxnSpPr = deps.getDirectChildByLocalName(shapeNode, "nvCxnSpPr");
    const cNvCxnSpPr = deps.getDirectChildByLocalName(nvCxnSpPr, "cNvCxnSpPr");
    const readId = (localName: string): string | null =>
        String(deps.getDirectChildByLocalName(cNvCxnSpPr, localName)?.getAttribute("id") || "").trim() || null;
    const line = deps.getDirectChildByLocalName(deps.getDirectChildByLocalName(shapeNode, "spPr"), "ln");
    // DrawingML puts headEnd at the line start and tailEnd at the line end.
    const atStart = hasArrowHead(deps.getDirectChildByLocalName(line, "headEnd"));
    const atEnd = hasArrowHead(deps.getDirectChildByLocalName(line, "tailEnd"));
    return {
        startId: readId("stCxn"),
        endId: readId("endCxn"),
        arrow: atStart && atEnd ? "both" : atStart ? "start" : atEnd ? "end" : "none",
    };
}

function isStruck(rPr: Element | null): boolean | null {
    const value = rPr?.getAttribute("strike");
    if (!value) return null;
    return value !== "noStrike";
}

/**
 * Reads shape text by paragraph. Runs join directly, and paragraphs and breaks join with line breaks.
 * The marked form wraps struck runs in ~~. Writers strike text to show removed content.
 */
function parseShapeText(shapeNode: Element | null, deps: ParseAssetDependencies): { text: string; marked: string } {
    if (!shapeNode) return { text: "", marked: "" };
    const paragraphs = deps.getElementsByLocalName(shapeNode, "p");
    if (paragraphs.length === 0) {
        const text = deps
            .getElementsByLocalName(shapeNode, "t")
            .map((node) => deps.getTextContent(node))
            .filter(Boolean)
            .join("")
            .trim();
        return { text, marked: text };
    }
    // A run without its own strike flag takes the default of its paragraph, then of the text body.
    const txBody = deps.getFirstDescendantByLocalName(shapeNode, "txBody");
    const level1 = deps.getDirectChildByLocalName(deps.getDirectChildByLocalName(txBody, "lstStyle"), "lvl1pPr");
    const bodyDefault = isStruck(deps.getDirectChildByLocalName(level1, "defRPr"));
    const plainLines: string[] = [];
    const markedLines: string[] = [];
    for (const paragraph of paragraphs) {
        const paragraphDefault =
            isStruck(deps.getDirectChildByLocalName(deps.getDirectChildByLocalName(paragraph, "pPr"), "defRPr")) ??
            bodyDefault;
        let plain = "";
        let marked = "";
        let struckRun = "";
        const flushStruck = () => {
            if (!struckRun) return;
            // Markdown strike marks must touch the text, so spaces stay outside.
            const match = /^(\s*)(.*?)(\s*)$/s.exec(struckRun) as RegExpExecArray;
            marked += match[2] ? `${match[1]}~~${match[2]}~~${match[3]}` : struckRun;
            struckRun = "";
        };
        for (const child of Array.from(paragraph.childNodes)) {
            const element = child as Element;
            if (element.nodeType !== 1) continue;
            if (element.localName === "br") {
                flushStruck();
                plain += "\n";
                marked += "\n";
                continue;
            }
            if (element.localName !== "r" && element.localName !== "fld") continue;
            const runText = deps
                .getElementsByLocalName(element, "t")
                .map((node) => deps.getTextContent(node))
                .join("");
            plain += runText;
            if (isStruck(deps.getDirectChildByLocalName(element, "rPr")) ?? paragraphDefault) {
                struckRun += runText;
            } else {
                flushStruck();
                marked += runText;
            }
        }
        flushStruck();
        plainLines.push(plain);
        markedLines.push(marked);
    }
    return { text: plainLines.join("\n").trim(), marked: markedLines.join("\n").trim() };
}

/** Reads shape size in EMUs. Returns null for missing dimensions. */
function parseShapeExt(
    anchor: Element,
    shapeNode: Element | null,
    deps: ParseAssetDependencies
): { widthEmu: number | null; heightEmu: number | null } {
    const extNode =
        deps.getDirectChildByLocalName(anchor, "ext") ||
        deps.getDirectChildByLocalName(
            deps.getDirectChildByLocalName(
                deps.getDirectChildByLocalName(shapeNode || anchor, "spPr") || shapeNode || anchor,
                "xfrm"
            ),
            "ext"
        );
    const widthEmu = Number(extNode?.getAttribute("cx") || "");
    const heightEmu = Number(extNode?.getAttribute("cy") || "");
    return {
        widthEmu: Number.isFinite(widthEmu) ? widthEmu : null,
        heightEmu: Number.isFinite(heightEmu) ? heightEmu : null,
    };
}

/**
 * Flattens one XML subtree into path-keyed entries.
 * Attributes use `@`, direct text uses `#text`.
 */
function flattenXmlNodeEntries(
    node: Element | null,
    deps: ParseAssetDependencies,
    path = "",
    entries: { key: string; value: string }[] = []
): { key: string; value: string }[] {
    if (!node) return entries;
    const nodeName = node.tagName || node.nodeName || node.localName || "node";
    const currentPath = path ? `${path}/${nodeName}` : nodeName;

    for (const attribute of Array.from(node.attributes)) {
        entries.push({
            key: `${currentPath}@${attribute.name}`,
            value: attribute.value,
        });
    }

    const directText = Array.from(node.childNodes)
        .filter((child) => child.nodeType === TEXT_NODE)
        .map((child) => (child.textContent || "").trim())
        .filter(Boolean)
        .join(" ");
    if (directText) {
        entries.push({
            key: `${currentPath}#text`,
            value: directText,
        });
    }

    for (const child of Array.from(node.childNodes)) {
        if (child.nodeType === ELEMENT_NODE) {
            flattenXmlNodeEntries(child as Element, deps, currentPath, entries);
        }
    }
    return entries;
}

function parseShapeRawEntries(anchor: Element, deps: ParseAssetDependencies): { key: string; value: string }[] {
    const entries: { key: string; value: string }[] = [];
    return flattenXmlNodeEntries(anchor, deps, "", entries);
}

/** Renders flat entries as indented Markdown bullets. Rebuilds the tree from slash-separated paths. */
function renderHierarchicalRawEntries(entries: { key: string; value: string }[]): string[] {
    type RawTreeNode = {
        children: Map<string, RawTreeNode>;
        value: string | null;
    };

    const root: RawTreeNode = {
        children: new Map<string, RawTreeNode>(),
        value: null,
    };

    for (const entry of entries) {
        const parts = entry.key.split("/").filter(Boolean);
        let current = root;
        for (const part of parts) {
            if (!current.children.has(part)) {
                current.children.set(part, {
                    children: new Map<string, RawTreeNode>(),
                    value: null,
                });
            }
            current = current.children.get(part)!;
        }
        current.value = entry.value;
    }

    const lines: string[] = [];

    function visit(node: RawTreeNode, depth: number): void {
        for (const [key, child] of node.children.entries()) {
            const indent = " ".repeat(depth * 4);
            if (child.value !== null) {
                lines.push(`${indent}- \`${key}\`: \`${child.value}\``);
            } else {
                lines.push(`${indent}- \`${key}\``);
            }
            visit(child, depth + 1);
        }
    }

    visit(root, 0);
    return lines;
}

/** Names the group from its header. Falls back to the drawing id. */
function groupDisplayName(groupNode: Element, deps: ParseAssetDependencies, fallback: string): string {
    const header =
        deps.getFirstDescendantByLocalName(groupNode, "nvGrpSpPr") ||
        deps.getFirstDescendantByLocalName(groupNode, "nvPr");
    const cNvPr = deps.getFirstDescendantByLocalName(header || groupNode, "cNvPr");
    const name = String(cNvPr?.getAttribute("name") || "").trim();
    if (name) return name;
    const id = String(cNvPr?.getAttribute("id") || "").trim();
    return id ? `group-${id}` : fallback;
}

/** Collects shape nodes at any depth. Group shapes contribute their path to each child. */
function collectShapeNodes(
    root: Element,
    deps: ParseAssetDependencies,
    groupPath: string[] = [],
    frames: (GroupFrame | null)[] = [],
    found: { node: Element; groupPath: string[]; frames: (GroupFrame | null)[] }[] = []
): { node: Element; groupPath: string[]; frames: (GroupFrame | null)[] }[] {
    let groupCounter = 0;
    for (let child = root.firstChild; child; child = child.nextSibling) {
        if (child.nodeType !== ELEMENT_NODE) continue;
        const element = child as Element;
        if (element.localName === "grpSp") {
            groupCounter += 1;
            const name = groupDisplayName(element, deps, `group-${groupCounter}`);
            collectShapeNodes(element, deps, [...groupPath, name], [...frames, readGroupFrame(element)], found);
            continue;
        }
        if (element.localName === "sp" || element.localName === "cxnSp") {
            found.push({ node: element, groupPath, frames });
        }
    }
    return found;
}

/** Collects one-cell and two-cell anchors in document order, then absolute anchors. */
function collectShapeAnchors(drawingDoc: Document): Element[] {
    const anchors = collectAnchors(drawingDoc);
    const absolute: Element[] = [];
    const stack: Node[] = [drawingDoc];
    while (stack.length > 0) {
        const current = stack.pop() as Node;
        if (current.nodeType === 1 && (current as Element).localName === "absoluteAnchor") {
            absolute.push(current as Element);
        }
        for (let child = current.lastChild; child; child = child.previousSibling) stack.push(child);
    }
    return [...anchors, ...absolute];
}

/**
 * Groups shapes into blocks without losing members.
 * Declared connector links decide the grouping first.
 * Distance only groups shapes without declared links.
 */
function extractShapeBlocks(
    shapes: ParsedShapeAsset[],
    deps: Pick<
        ParseAssetDependencies,
        "defaultCellWidthEmu" | "defaultCellHeightEmu" | "shapeBlockGapXEmu" | "shapeBlockGapYEmu" | "grid"
    >
): ShapeBlock[] {
    return partitionShapeRegions(shapes, deps);
}

/** Reads the anchor cell for labels. Absolute anchors derive the cell from EMU geometry. */
function parseAnchorCell(
    anchor: Element,
    geometry: { left: number; top: number },
    grid: SheetGrid,
    deps: ParseAssetDependencies
): { col: number; row: number } | null {
    const from = deps.getFirstDescendantByLocalName(anchor, "from");
    const colNode = deps.getFirstDescendantByLocalName(from || anchor, "col");
    const rowNode = deps.getFirstDescendantByLocalName(from || anchor, "row");
    const fromCol = Number(deps.getTextContent(colNode));
    const fromRow = Number(deps.getTextContent(rowNode));
    if (Number.isFinite(fromCol) && Number.isFinite(fromRow) && from !== null) {
        return { col: fromCol + 1, row: fromRow + 1 };
    }
    if (anchor.localName === "absoluteAnchor") {
        return { col: columnAtEmu(grid, geometry.left), row: rowAtEmu(grid, geometry.top) };
    }
    return null;
}

/** Collects shape anchors and skips image and chart anchors. Renders each shape to SVG when supported. */
function parseDrawingShapes(
    files: Map<string, Uint8Array>,
    sheetName: string,
    sheetPath: string,
    deps: ParseAssetDependencies
): ParsedShapeAsset[] {
    const sheetRels = deps.parseRelationships(files, deps.buildRelsPath(sheetPath), sheetPath);
    const shapes: ParsedShapeAsset[] = [];
    let shapeCounter = 1;
    const grid = deps.grid ?? createUniformGrid(deps.defaultCellWidthEmu, deps.defaultCellHeightEmu);
    const theme = readThemeColors(files, deps.decodeXmlText);

    for (const drawingPath of sheetRels.values()) {
        if (!/\/drawings\/.+\.xml$/i.test(drawingPath)) continue;
        const drawingBytes = files.get(drawingPath);
        if (!drawingBytes) continue;
        const drawingDoc = deps.xmlToDocument(deps.decodeXmlText(drawingBytes));
        const anchors = collectShapeAnchors(drawingDoc);

        for (const anchor of anchors) {
            const found = collectShapeNodes(anchor, deps);
            if (found.length === 0) continue;

            for (const { node: shapeNode, groupPath, frames } of found) {
                const cNvPr = deps.getFirstDescendantByLocalName(
                    deps.getFirstDescendantByLocalName(
                        shapeNode,
                        shapeNode.localName === "sp" ? "nvSpPr" : "nvCxnSpPr"
                    ) || shapeNode,
                    "cNvPr"
                );
                const drawingId = String(cNvPr?.getAttribute("id") || "").trim();
                const geometry = normalizeShapeGeometry(anchor, shapeNode, grid);
                // Group children carry their own box in child coords. Map it outward, then into the anchor box.
                // Fall back to the anchor box when frames or child extents are unusable.
                if (frames.length > 0) {
                    const outward = mapBoxThroughGroups(readChildBox(shapeNode), frames);
                    const mapped = outward && fitBoxToAnchor(outward, frames[0] as GroupFrame, geometry);
                    if (mapped) {
                        geometry.left = mapped.left;
                        geometry.top = mapped.top;
                        geometry.right = mapped.right;
                        geometry.bottom = mapped.bottom;
                    }
                }
                const cell = parseAnchorCell(anchor, geometry, grid, deps);
                if (!cell || cell.col <= 0 || cell.row <= 0) {
                    continue;
                }
                const { widthEmu, heightEmu } = parseShapeExt(anchor, shapeNode, deps);
                const shapeText = parseShapeText(shapeNode, deps);
                const svgAsset =
                    deps.drawingHelper?.renderShapeSvg?.(shapeNode, anchor, sheetName, shapeCounter) || null;
                const groupLabel = groupPath.length > 0 ? groupPath.join("/") : null;
                const shape: ParsedShapeAsset = {
                    sheetName,
                    anchor: `${deps.colToLetters(cell.col)}${cell.row}`,
                    drawingId,
                    stableId: `${drawingPath}#${groupLabel ? `${groupLabel}/` : ""}${drawingId || `shape-${shapeCounter}`}`,
                    groupPath: groupLabel,
                    anchorKind: geometry.anchorKind,
                    geometryConfidence: geometry.confidence,
                    rotationDeg: geometry.rotationDeg,
                    flipH: geometry.flipH,
                    flipV: geometry.flipV,
                    name: String(cNvPr?.getAttribute("name") || "").trim() || "Shape",
                    kind: parseShapeKind(shapeNode, deps),
                    text: shapeText.text,
                    ...(shapeText.marked !== shapeText.text ? { markedText: shapeText.marked } : {}),
                    look: readShapeLook(shapeNode, theme),
                    connection: parseShapeConnection(shapeNode, deps),
                    widthEmu,
                    heightEmu,
                    elementName: `xdr:${shapeNode.localName}`,
                    anchorElementName: anchor.tagName || anchor.nodeName || anchor.localName || "anchor",
                    rawEntries: parseShapeRawEntries(anchor, deps),
                    bbox: { left: 0, top: 0, right: 0, bottom: 0 },
                    svgFilename: svgAsset?.filename || null,
                    svgPath: svgAsset?.path || null,
                    svgData: svgAsset?.data || null,
                    svgWarning: svgAsset?.warning || null,
                };
                applyGeometryToShape(shape, geometry);
                if (shape.connection) {
                    Object.assign(shape.connection, connectorEndpoints(shapeNode, shape.bbox, geometry));
                } else {
                    shape.calloutTip = calloutTip(shapeNode, shape.bbox, geometry);
                }
                shapes.push(shape);
                shapeCounter += 1;
            }
        }
    }

    return shapes;
}

export { extractShapeBlocks, parseDrawingShapes, renderHierarchicalRawEntries };
