import { describe, expect, it } from "vitest";
import {
    calloutTip,
    connectorEndpoints,
    fitBoxToAnchor,
    mapBoxThroughGroups,
    mapChildBoxToParent,
    normalizeShapeGeometry,
    readGroupFrame,
    readShapeTransform,
} from "../../../../../src/extraction/assets/drawing-geometry.ts";
import { getFirstDescendantByLocalName, xmlToDocument } from "../../../../../src/extraction/xml/xml-utils.ts";
import {
    buildSheetGrid,
    columnLeftEmu,
    createUniformGrid,
    rowTopEmu,
} from "../../../../../src/extraction/workbook/sheet-grid.ts";

const CELL_W = 609600;
const CELL_H = 190500;
const METRICS = createUniformGrid(CELL_W, CELL_H);

function parseDrawing(xml) {
    return xmlToDocument(xml);
}

function firstAnchor(doc) {
    return (
        getFirstDescendantByLocalName(doc, "twoCellAnchor") ||
        getFirstDescendantByLocalName(doc, "oneCellAnchor") ||
        getFirstDescendantByLocalName(doc, "absoluteAnchor")
    );
}

function firstShape(anchor) {
    return getFirstDescendantByLocalName(anchor, "sp") || getFirstDescendantByLocalName(anchor, "cxnSp");
}

describe("drawing geometry", () => {
    it("maps two-cell anchors with estimated confidence", () => {
        const doc = parseDrawing(`<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:twoCellAnchor>
            <xdr:from><xdr:col>1</xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>20</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>3</xdr:col><xdr:colOff>30</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>40</xdr:rowOff></xdr:to>
            <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="R"/></xdr:nvSpPr><xdr:spPr/></xdr:sp>
        </xdr:twoCellAnchor></xdr:wsDr>`);
        const anchor = firstAnchor(doc);
        const geometry = normalizeShapeGeometry(anchor, firstShape(anchor), METRICS);
        expect(geometry.anchorKind).toBe("twoCell");
        expect(geometry.confidence).toBe("ESTIMATED");
        expect(geometry.left).toBe(1 * CELL_W + 10);
        expect(geometry.top).toBe(2 * CELL_H + 20);
        expect(geometry.right).toBe(3 * CELL_W + 30);
        expect(geometry.bottom).toBe(4 * CELL_H + 40);
    });

    it("places two-cell anchors on real column widths and row heights", () => {
        const sheet = xmlToDocument(
            `<worksheet><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="3" width="2.5"/></cols>` +
                `<sheetData><row r="2" ht="40"/></sheetData></worksheet>`
        );
        const grid = buildSheetGrid(
            sheet.getElementsByTagName("sheetFormatPr")[0],
            Array.from(sheet.getElementsByTagName("col")),
            Array.from(sheet.getElementsByTagName("row")),
            { name: "Calibri", sizePt: 11 }
        );
        const doc = parseDrawing(`<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:twoCellAnchor>
            <xdr:from><xdr:col>1</xdr:col><xdr:colOff>10</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>20</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>4</xdr:col><xdr:colOff>30</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>40</xdr:rowOff></xdr:to>
            <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="R"/></xdr:nvSpPr><xdr:spPr/></xdr:sp>
        </xdr:twoCellAnchor></xdr:wsDr>`);
        const anchor = firstAnchor(doc);
        const geometry = normalizeShapeGeometry(anchor, firstShape(anchor), grid);
        expect(geometry.confidence).toBe("EXACT");
        expect(geometry.left).toBe(columnLeftEmu(grid, 2) + 10);
        expect(geometry.right).toBe(columnLeftEmu(grid, 5) + 30);
        // Row 2 is 40 pt high, so row 3 starts 55 pt down.
        expect(geometry.top).toBe(rowTopEmu(grid, 3) + 20);
        expect(rowTopEmu(grid, 3)).toBe(55 * 12700);
        expect(geometry.bottom).toBe(rowTopEmu(grid, 4) + 40);
    });

    it("sizes one-cell anchors from the anchor extent", () => {
        const doc = parseDrawing(`<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:oneCellAnchor>
            <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:ext cx="1000" cy="2000"/>
            <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="R"/></xdr:nvSpPr><xdr:spPr/></xdr:sp>
        </xdr:oneCellAnchor></xdr:wsDr>`);
        const anchor = firstAnchor(doc);
        const geometry = normalizeShapeGeometry(anchor, firstShape(anchor), METRICS);
        expect(geometry.anchorKind).toBe("oneCell");
        expect(geometry.confidence).toBe("ESTIMATED");
        expect(geometry.right - geometry.left).toBe(1000);
        expect(geometry.bottom - geometry.top).toBe(2000);
    });

    it("keeps absolute anchors exact in EMU space", () => {
        const doc = parseDrawing(`<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:absoluteAnchor>
            <xdr:pos x="1000" y="2000"/>
            <xdr:ext cx="5000" cy="6000"/>
            <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="R"/></xdr:nvSpPr><xdr:spPr/></xdr:sp>
        </xdr:absoluteAnchor></xdr:wsDr>`);
        const anchor = firstAnchor(doc);
        const geometry = normalizeShapeGeometry(anchor, firstShape(anchor), METRICS);
        expect(geometry.anchorKind).toBe("absolute");
        expect(geometry.confidence).toBe("EXACT");
        expect([geometry.left, geometry.top, geometry.right, geometry.bottom]).toEqual([1000, 2000, 6000, 8000]);
    });

    it("marks incomplete anchors unknown instead of guessing", () => {
        const doc = parseDrawing(`<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:absoluteAnchor>
            <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="R"/></xdr:nvSpPr><xdr:spPr/></xdr:sp>
        </xdr:absoluteAnchor></xdr:wsDr>`);
        const anchor = firstAnchor(doc);
        const geometry = normalizeShapeGeometry(anchor, firstShape(anchor), METRICS);
        expect(geometry.confidence).toBe("UNKNOWN");
    });

    it("reads rotation and flips from the shape transform", () => {
        const doc = parseDrawing(`<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:twoCellAnchor>
            <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
            <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="1" name="R"/></xdr:nvSpPr>
            <xdr:spPr><a:xfrm rot="5400000" flipH="1"><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm></xdr:spPr></xdr:sp>
        </xdr:twoCellAnchor></xdr:wsDr>`);
        const transform = readShapeTransform(firstShape(firstAnchor(doc)));
        expect(transform.rotationDeg).toBe(90);
        expect(transform.flipH).toBe(true);
        expect(transform.flipV).toBe(false);
    });
});

describe("group frames", () => {
    const frameDoc = (inner) =>
        xmlToDocument(
            `<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="10" name="G"/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="2000" cy="1000"/><a:chOff x="100" y="200"/><a:chExt cx="200" cy="100"/></a:xfrm></xdr:grpSpPr>${inner}</xdr:grpSp></xdr:wsDr>`
        );

    it("maps child boxes through the frame scale", () => {
        const frame = readGroupFrame(getFirstDescendantByLocalName(frameDoc(""), "grpSp"));
        expect(frame).not.toBe(null);
        const mapped = mapChildBoxToParent({ left: 100, top: 200, right: 300, bottom: 300 }, frame);
        expect(mapped).toEqual({ left: 1000, top: 2000, right: 3000, bottom: 3000 });
    });

    it("rejects zero child extents instead of scaling by infinity", () => {
        const frame = {
            offX: 0,
            offY: 0,
            extCx: 10,
            extCy: 10,
            chOffX: 0,
            chOffY: 0,
            chExtCx: 0,
            chExtCy: 10,
        };
        expect(mapChildBoxToParent({ left: 0, top: 0, right: 5, bottom: 5 }, frame)).toBe(null);
        expect(mapBoxThroughGroups({ left: 0, top: 0, right: 5, bottom: 5 }, [frame])).toBe(null);
        expect(mapBoxThroughGroups(null, [frame])).toBe(null);
    });

    it("places connector ends from flips and rotation", () => {
        const box = { left: 0, top: 0, right: 100, bottom: 50 };
        const connector = (xfrm) =>
            getFirstDescendantByLocalName(
                parseDrawing(
                    `<cxnSp xmlns:a="a"><spPr><a:xfrm${xfrm}><a:off x="0" y="0"/><a:ext cx="100" cy="50"/></a:xfrm></spPr></cxnSp>`
                ),
                "cxnSp"
            );
        const plain = { rotationDeg: null, flipH: false, flipV: false };
        expect(connectorEndpoints(connector(""), box, plain)).toEqual({
            startPoint: { x: 0, y: 0 },
            endPoint: { x: 100, y: 50 },
        });
        expect(connectorEndpoints(connector(""), box, { ...plain, flipV: true })).toEqual({
            startPoint: { x: 0, y: 50 },
            endPoint: { x: 100, y: 0 },
        });
        // A horizontal line turned 90 degrees clockwise runs from the top to the bottom of its visual box.
        const line = getFirstDescendantByLocalName(
            parseDrawing(
                `<cxnSp xmlns:a="a"><spPr><a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="200" cy="0"/></a:xfrm></spPr></cxnSp>`
            ),
            "cxnSp"
        );
        const ends = connectorEndpoints(
            line,
            { left: 10, top: 0, right: 10, bottom: 200 },
            { ...plain, rotationDeg: 90 }
        );
        expect(ends.startPoint.y).toBeCloseTo(0);
        expect(ends.endPoint.y).toBeCloseTo(200);
    });

    it("fits an outer group frame into the anchor box", () => {
        const frame = { offX: 1000, offY: 0, extCx: 2000, extCy: 1000, chOffX: 0, chOffY: 0, chExtCx: 1, chExtCy: 1 };
        const anchorBox = { left: 0, top: 0, right: 400, bottom: 100 };
        expect(fitBoxToAnchor({ left: 2000, top: 500, right: 3000, bottom: 1000 }, frame, anchorBox)).toEqual({
            left: 200,
            top: 50,
            right: 400,
            bottom: 100,
        });
        expect(fitBoxToAnchor(anchorBox, { ...frame, extCx: 0 }, anchorBox)).toBe(null);
    });
});

describe("callout tips", () => {
    function calloutNode(prst, avLst = "", xfrmAttrs = "") {
        const doc = parseDrawing(`<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:sp><xdr:spPr>
            <a:xfrm${xfrmAttrs}><a:off x="0" y="0"/><a:ext cx="1000" cy="400"/></a:xfrm>
            <a:prstGeom prst="${prst}"><a:avLst>${avLst}</a:avLst></a:prstGeom>
        </xdr:spPr></xdr:sp></xdr:wsDr>`);
        return getFirstDescendantByLocalName(doc, "sp");
    }
    const box = { left: 0, top: 0, right: 1000, bottom: 400 };
    const plain = { rotationDeg: null, flipH: false, flipV: false };

    it("uses the preset defaults of a wedge callout", () => {
        const tip = calloutTip(calloutNode("wedgeRectCallout"), box, plain);
        expect(tip.x).toBeCloseTo(500 - 208.33, 1);
        expect(tip.y).toBeCloseTo(200 + 250, 1);
    });

    it("reads adjust values of a wedge callout", () => {
        const av = '<a:gd name="adj1" fmla="val -100000"/><a:gd name="adj2" fmla="val 0"/>';
        expect(calloutTip(calloutNode("wedgeRoundRectCallout", av), box, plain)).toEqual({ x: -500, y: 200 });
    });

    it("places line callout tips from the top-left corner", () => {
        const tip = calloutTip(calloutNode("borderCallout1"), box, plain);
        expect(tip.x).toBeCloseTo(-383.33, 1);
        expect(tip.y).toBeCloseTo(450, 1);
    });

    it("mirrors the tip of a flipped callout", () => {
        const av = '<a:gd name="adj1" fmla="val -100000"/><a:gd name="adj2" fmla="val 0"/>';
        expect(calloutTip(calloutNode("wedgeRectCallout", av), box, { ...plain, flipH: true })).toEqual({
            x: 1500,
            y: 200,
        });
    });
});
