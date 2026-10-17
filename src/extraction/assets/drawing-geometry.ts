import type { ParsedShapeAsset, SheetGrid } from "../../model/worksheet.js";
import { columnLeftEmu, rowTopEmu } from "../workbook/sheet-grid.js";
import { getDirectChildByLocalName, getFirstDescendantByLocalName, getTextContent } from "../xml/xml-utils.js";

type GeometryConfidence = "EXACT" | "ESTIMATED" | "UNKNOWN";

type AnchorKind = "twoCell" | "oneCell" | "absolute" | "unknown";

type NormalizedGeometry = {
    left: number;
    top: number;
    right: number;
    bottom: number;
    anchorKind: AnchorKind;
    confidence: GeometryConfidence;
    rotationDeg: number | null;
    flipH: boolean;
    flipV: boolean;
};

/** Names the anchor flavor. Absolute anchors carry EMU positions, not cells. */
function anchorKindOf(anchor: Element | null): AnchorKind {
    const name = anchor?.localName || "";
    if (name === "twoCellAnchor") return "twoCell";
    if (name === "oneCellAnchor") return "oneCell";
    if (name === "absoluteAnchor") return "absolute";
    return "unknown";
}

/** Reads one integer inside a named parent. Returns null without a finite value. */
function readInt(root: Element | null, parentName: string, childName: string): number | null {
    if (!root) return null;
    const parent = getFirstDescendantByLocalName(root, parentName);
    if (!parent) return null;
    const child = getDirectChildByLocalName(parent, childName);
    const value = Number(getTextContent(child));
    return Number.isFinite(value) ? value : null;
}

/** Reads an EMU attribute. Returns null when the attribute is missing or not finite. Zero is valid. */
function readEmu(node: Element | null, name: string): number | null {
    const raw = node?.getAttribute(name);
    if (raw === null || raw === undefined || raw.trim() === "") return null;
    const value = Number(raw);
    return Number.isFinite(value) ? value : null;
}

/**
 * Reads rotation and flips from the shape transform.
 * Rotation arrives in 60000ths of a degree. Group transforms accumulate later.
 */
function readShapeTransform(shapeNode: Element | null): {
    rotationDeg: number | null;
    flipH: boolean;
    flipV: boolean;
} {
    const xfrm = getDirectChildByLocalName(getDirectChildByLocalName(shapeNode, "spPr"), "xfrm");
    const raw = Number(xfrm?.getAttribute("rot") || "");
    return {
        rotationDeg: Number.isFinite(raw) ? raw / 60000 : null,
        flipH: xfrm?.getAttribute("flipH") === "1" || xfrm?.getAttribute("flipH") === "true",
        flipV: xfrm?.getAttribute("flipV") === "1" || xfrm?.getAttribute("flipV") === "true",
    };
}

/** Reads the anchor-level extent used by one-cell anchors. */
function readAnchorExt(anchor: Element | null): { widthEmu: number | null; heightEmu: number | null } {
    const ext = getDirectChildByLocalName(anchor, "ext");
    return { widthEmu: readEmu(ext, "cx"), heightEmu: readEmu(ext, "cy") };
}

/**
 * Builds the shape box in sheet EMU space.
 * Absolute anchors map EMUs directly and stay exact.
 * Cell anchors add their offsets to the cell edges of the grid, so they are as exact as the grid.
 */
function normalizeShapeGeometry(
    anchor: Element | null,
    shapeNode: Element | null,
    grid: SheetGrid
): NormalizedGeometry {
    const anchorKind = anchorKindOf(anchor);
    const transform = readShapeTransform(shapeNode);
    const fromCol = readInt(anchor, "from", "col") || 0;
    const fromRow = readInt(anchor, "from", "row") || 0;
    const fromColOff = readInt(anchor, "from", "colOff") || 0;
    const fromRowOff = readInt(anchor, "from", "rowOff") || 0;

    if (anchorKind === "absolute") {
        const pos = getDirectChildByLocalName(anchor, "pos");
        const ext = getDirectChildByLocalName(anchor, "ext");
        const left = readEmu(pos, "x");
        const top = readEmu(pos, "y");
        const width = readEmu(ext, "cx");
        const height = readEmu(ext, "cy");
        if (left !== null && top !== null && width !== null && height !== null) {
            return {
                left,
                top,
                right: left + Math.max(1, width),
                bottom: top + Math.max(1, height),
                anchorKind,
                confidence: "EXACT",
                ...transform,
            };
        }
        return {
            ...cellFallback(fromCol, fromRow, fromColOff, fromRowOff, grid),
            anchorKind,
            confidence: "UNKNOWN",
            ...transform,
        };
    }

    const toCol = readInt(anchor, "to", "col");
    const toRow = readInt(anchor, "to", "row");
    if (toCol !== null && toRow !== null) {
        const toColOff = readInt(anchor, "to", "colOff") || 0;
        const toRowOff = readInt(anchor, "to", "rowOff") || 0;
        // Anchor columns and rows are 0-based. Grid indexes are 1-based.
        const left = columnLeftEmu(grid, fromCol + 1) + fromColOff;
        const top = rowTopEmu(grid, fromRow + 1) + fromRowOff;
        const right = columnLeftEmu(grid, toCol + 1) + toColOff;
        const bottom = rowTopEmu(grid, toRow + 1) + toRowOff;
        return {
            left,
            top,
            // A writer can give a "to" point before the "from" point. Keep the box valid.
            right: Math.max(right, left + 1),
            bottom: Math.max(bottom, top + 1),
            anchorKind,
            confidence: grid.confidence,
            ...transform,
        };
    }

    const ext = readAnchorExt(anchor);
    if (ext.widthEmu === null || ext.heightEmu === null) {
        return {
            ...cellFallback(fromCol, fromRow, fromColOff, fromRowOff, grid),
            anchorKind,
            confidence: "UNKNOWN",
            ...transform,
        };
    }
    return {
        ...cellFallback(fromCol, fromRow, fromColOff, fromRowOff, grid, ext.widthEmu, ext.heightEmu),
        anchorKind,
        confidence: grid.confidence,
        ...transform,
    };
}

/** Places the box from the anchor cell plus offsets. Without an extent, the box fills one default cell. */
function cellFallback(
    fromCol: number,
    fromRow: number,
    fromColOff: number,
    fromRowOff: number,
    grid: SheetGrid,
    widthEmu?: number,
    heightEmu?: number
): { left: number; top: number; right: number; bottom: number } {
    const left = columnLeftEmu(grid, fromCol + 1) + fromColOff;
    const top = rowTopEmu(grid, fromRow + 1) + fromRowOff;
    return {
        left,
        top,
        right: left + Math.max(1, widthEmu || grid.defaultColWidthEmu),
        bottom: top + Math.max(1, heightEmu || grid.defaultRowHeightEmu),
    };
}

/** Copies normalized geometry onto a parsed shape asset. Keeps bbox in sync. */
function applyGeometryToShape(shape: ParsedShapeAsset, geometry: NormalizedGeometry): void {
    shape.bbox = { left: geometry.left, top: geometry.top, right: geometry.right, bottom: geometry.bottom };
    shape.anchorKind = geometry.anchorKind;
    shape.geometryConfidence = geometry.confidence;
    shape.rotationDeg = geometry.rotationDeg;
    shape.flipH = geometry.flipH;
    shape.flipV = geometry.flipV;
}

/** One group frame. Child boxes live in child coords and map into the parent space. */
type GroupFrame = {
    offX: number;
    offY: number;
    extCx: number;
    extCy: number;
    chOffX: number;
    chOffY: number;
    chExtCx: number;
    chExtCy: number;
};

type EmuBox = { left: number; top: number; right: number; bottom: number };

/** Reads the group frame from grpSpPr/xfrm. Returns null without usable extents. */
function readGroupFrame(groupNode: Element | null): GroupFrame | null {
    const xfrm = getDirectChildByLocalName(groupNode, "grpSpPr");
    const inner = getDirectChildByLocalName(xfrm, "xfrm");
    if (!inner) return null;
    const off = getDirectChildByLocalName(inner, "off");
    const ext = getDirectChildByLocalName(inner, "ext");
    const chOff = getDirectChildByLocalName(inner, "chOff");
    const chExt = getDirectChildByLocalName(inner, "chExt");
    const frame = {
        offX: readEmu(off, "x"),
        offY: readEmu(off, "y"),
        extCx: readEmu(ext, "cx"),
        extCy: readEmu(ext, "cy"),
        chOffX: readEmu(chOff, "x"),
        chOffY: readEmu(chOff, "y"),
        chExtCx: readEmu(chExt, "cx"),
        chExtCy: readEmu(chExt, "cy"),
    };
    if (
        frame.offX === null ||
        frame.offY === null ||
        frame.extCx === null ||
        frame.extCy === null ||
        frame.chOffX === null ||
        frame.chOffY === null ||
        frame.chExtCx === null ||
        frame.chExtCy === null
    ) {
        return null;
    }
    return frame as GroupFrame;
}

/** Reads the child box from spPr/xfrm in group child coords. Returns null when absent. */
function readChildBox(shapeNode: Element | null): EmuBox | null {
    const xfrm = getDirectChildByLocalName(getDirectChildByLocalName(shapeNode, "spPr"), "xfrm");
    const off = getDirectChildByLocalName(xfrm, "off");
    const ext = getDirectChildByLocalName(xfrm, "ext");
    const left = readEmu(off, "x");
    const top = readEmu(off, "y");
    const width = readEmu(ext, "cx");
    const height = readEmu(ext, "cy");
    if (left === null || top === null || width === null || height === null) return null;
    return { left, top, right: left + Math.max(1, width), bottom: top + Math.max(1, height) };
}

/**
 * Maps one child box into its parent group space.
 * Returns null when the child extent is zero, since scaling has no meaning.
 */
function mapChildBoxToParent(box: EmuBox, frame: GroupFrame): EmuBox | null {
    if (frame.chExtCx <= 0 || frame.chExtCy <= 0) return null;
    const scaleX = frame.extCx / frame.chExtCx;
    const scaleY = frame.extCy / frame.chExtCy;
    const left = frame.offX + (box.left - frame.chOffX) * scaleX;
    const top = frame.offY + (box.top - frame.chOffY) * scaleY;
    const right = frame.offX + (box.right - frame.chOffX) * scaleX;
    const bottom = frame.offY + (box.bottom - frame.chOffY) * scaleY;
    return {
        left: Math.min(left, right),
        top: Math.min(top, bottom),
        right: Math.max(left, right),
        bottom: Math.max(top, bottom),
    };
}

/**
 * Folds a child box through nested group frames, innermost first.
 * Returns null when any frame or the child box is unusable.
 */
function mapBoxThroughGroups(box: EmuBox | null, frames: (GroupFrame | null)[]): EmuBox | null {
    if (!box) return null;
    let current: EmuBox | null = box;
    for (let index = frames.length - 1; index >= 0 && current; index -= 1) {
        const frame = frames[index];
        if (!frame) return null;
        current = mapChildBoxToParent(current, frame);
    }
    return current;
}

type Point = { x: number; y: number };

type ShapeTransform = { rotationDeg: number | null; flipH: boolean; flipV: boolean };

/**
 * Returns a function that maps a point of the unrotated shape into the anchor box.
 * The point is relative to the shape center, in the units of the shape transform.
 * Flips mirror the point. Rotation turns it clockwise around the center.
 * The anchor box is the visual box after rotation, so the point maps into it by fraction.
 */
function shapePointMapper(
    shapeNode: Element | null,
    box: EmuBox,
    transform: ShapeTransform
): { halfW: number; halfH: number; place: (localX: number, localY: number) => Point } {
    const xfrm = getDirectChildByLocalName(getDirectChildByLocalName(shapeNode, "spPr"), "xfrm");
    const ext = getDirectChildByLocalName(xfrm, "ext");
    const halfW = (readEmu(ext, "cx") ?? box.right - box.left) / 2;
    const halfH = (readEmu(ext, "cy") ?? box.bottom - box.top) / 2;
    const radians = ((transform.rotationDeg || 0) * Math.PI) / 180;
    const cos = Math.cos(radians);
    const sin = Math.sin(radians);
    const boundW = Math.abs(halfW * cos) + Math.abs(halfH * sin);
    const boundH = Math.abs(halfW * sin) + Math.abs(halfH * cos);
    const place = (localX: number, localY: number): Point => {
        const flippedX = transform.flipH ? -localX : localX;
        const flippedY = transform.flipV ? -localY : localY;
        const x = flippedX * cos - flippedY * sin;
        const y = flippedX * sin + flippedY * cos;
        const u = boundW > 0 ? (x + boundW) / (2 * boundW) : 0.5;
        const v = boundH > 0 ? (y + boundH) / (2 * boundH) : 0.5;
        return { x: box.left + u * (box.right - box.left), y: box.top + v * (box.bottom - box.top) };
    };
    return { halfW, halfH, place };
}

/**
 * Finds the line start and end of a connector inside its box.
 * The unrotated line runs from the top-left to the bottom-right corner of the transform.
 */
function connectorEndpoints(
    shapeNode: Element | null,
    box: EmuBox,
    transform: ShapeTransform
): { startPoint: Point; endPoint: Point } {
    const { halfW, halfH, place } = shapePointMapper(shapeNode, box, transform);
    return { startPoint: place(-halfW, -halfH), endPoint: place(halfW, halfH) };
}

/**
 * Tip guides of callout presets, from presetShapeDefinitions.xml in ECMA-376 Part 1.
 * "wedge" tips sit at (adjX, adjY) from the center. "line" tips sit at (adjX, adjY) from the top-left corner.
 * Values are in 1/100000 of the shape width or height.
 */
const CALLOUT_TIPS: Record<
    string,
    { from: "center" | "corner"; x: string; y: string; defaults: Record<string, number> }
> = {
    wedgeRectCallout: { from: "center", x: "adj1", y: "adj2", defaults: { adj1: -20833, adj2: 62500 } },
    wedgeRoundRectCallout: { from: "center", x: "adj1", y: "adj2", defaults: { adj1: -20833, adj2: 62500 } },
    wedgeEllipseCallout: { from: "center", x: "adj1", y: "adj2", defaults: { adj1: -20833, adj2: 62500 } },
    cloudCallout: { from: "center", x: "adj1", y: "adj2", defaults: { adj1: -20833, adj2: 62500 } },
    ...lineCalloutTips(1, { adj1: 18750, adj2: -8333, adj3: 112500, adj4: -38333 }),
    ...lineCalloutTips(2, { adj1: 18750, adj2: -8333, adj3: 18750, adj4: -16667, adj5: 112500, adj6: -46667 }),
    ...lineCalloutTips(3, {
        adj1: 18750,
        adj2: -8333,
        adj3: 18750,
        adj4: -16667,
        adj5: 100000,
        adj6: -16667,
        adj7: 112963,
        adj8: -8333,
    }),
};

/** Line callouts of one level share guides. The last guide pair is the tip, with y before x. */
function lineCalloutTips(
    level: 1 | 2 | 3,
    defaults: Record<string, number>
): Record<string, { from: "corner"; x: string; y: string; defaults: Record<string, number> }> {
    const tip = { from: "corner" as const, x: `adj${level * 2 + 2}`, y: `adj${level * 2 + 1}`, defaults };
    const names = ["callout", "accentCallout", "borderCallout", "accentBorderCallout"];
    return Object.fromEntries(names.map((name) => [`${name}${level}`, tip]));
}

/** Reads the adjust values of a preset. Missing guides keep their defaults. */
function readAdjustValues(geomNode: Element | null, defaults: Record<string, number>): Record<string, number> {
    const values = { ...defaults };
    const avLst = getDirectChildByLocalName(geomNode, "avLst");
    for (const gd of Array.from(avLst?.childNodes || [])) {
        const element = gd as Element;
        if (element.localName !== "gd") continue;
        const match = /^val\s+(-?\d+(?:\.\d+)?)$/.exec(String(element.getAttribute("fmla") || "").trim());
        const name = element.getAttribute("name") || "";
        if (match && name) values[name] = Number(match[1]);
    }
    return values;
}

/**
 * Finds the point that a callout points at, in the space of the anchor box.
 * Returns null for shapes that are not callouts.
 */
function calloutTip(shapeNode: Element | null, box: EmuBox, transform: ShapeTransform): Point | null {
    const geomNode = getDirectChildByLocalName(getDirectChildByLocalName(shapeNode, "spPr"), "prstGeom");
    const preset = CALLOUT_TIPS[String(geomNode?.getAttribute("prst") || "")];
    if (!preset) return null;
    const values = readAdjustValues(geomNode, preset.defaults);
    const { halfW, halfH, place } = shapePointMapper(shapeNode, box, transform);
    const fx = (values[preset.x] ?? 0) / 100000;
    const fy = (values[preset.y] ?? 0) / 100000;
    return preset.from === "center"
        ? place(fx * 2 * halfW, fy * 2 * halfH)
        : place(fx * 2 * halfW - halfW, fy * 2 * halfH - halfH);
}

export {
    normalizeShapeGeometry,
    applyGeometryToShape,
    anchorKindOf,
    readShapeTransform,
    connectorEndpoints,
    calloutTip,
};
/**
 * Places a box from the outer group frame into the anchor box by fraction.
 * The outer frame holds the sheet EMUs that the writer computed. The anchor box uses the grid of this parse.
 * The two can differ, for example when the writer used another font. The anchor box wins,
 * so grouped and ungrouped shapes stay in the same space.
 * Returns null when the frame has no extent.
 */
function fitBoxToAnchor(box: EmuBox, outerFrame: GroupFrame, anchorBox: EmuBox): EmuBox | null {
    if (outerFrame.extCx <= 0 || outerFrame.extCy <= 0) return null;
    const scaleX = (anchorBox.right - anchorBox.left) / outerFrame.extCx;
    const scaleY = (anchorBox.bottom - anchorBox.top) / outerFrame.extCy;
    return {
        left: anchorBox.left + (box.left - outerFrame.offX) * scaleX,
        top: anchorBox.top + (box.top - outerFrame.offY) * scaleY,
        right: anchorBox.left + (box.right - outerFrame.offX) * scaleX,
        bottom: anchorBox.top + (box.bottom - outerFrame.offY) * scaleY,
    };
}

export { readGroupFrame, readChildBox, mapChildBoxToParent, mapBoxThroughGroups, fitBoxToAnchor };
export type { NormalizedGeometry, GeometryConfidence, AnchorKind, GroupFrame, EmuBox, Point };
