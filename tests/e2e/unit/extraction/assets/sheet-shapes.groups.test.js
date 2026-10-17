import { describe, expect, it } from "vitest";
import { parseDrawingShapes } from "../../../../../src/extraction/assets/sheet-shapes.ts";
import {
    decodeXmlText,
    getDirectChildByLocalName,
    getElementsByLocalName,
    getFirstDescendantByLocalName,
    getTextContent,
    xmlToDocument,
} from "../../../../../src/extraction/xml/xml-utils.ts";
import { colToLetters } from "../../../../../src/structure/addresses/address-utils.ts";

const DRAWING_PATH = "xl/drawings/drawing1.xml";
const encoder = new TextEncoder();

function shapeXml(id, name, text) {
    return `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/></xdr:nvSpPr><xdr:spPr/><xdr:txBody><a:p xmlns:a="y"><a:r><a:t>${text}</a:t></a:r></a:p></xdr:txBody></xdr:sp>`;
}

function deps() {
    return {
        parseRelationships: () => new Map([["rId1", DRAWING_PATH]]),
        buildRelsPath: () => "xl/worksheets/_rels/sheet1.xml.rels",
        xmlToDocument,
        decodeXmlText,
        getElementsByLocalName,
        getFirstDescendantByLocalName,
        getDirectChildByLocalName,
        getTextContent,
        colToLetters,
        drawingHelper: null,
        defaultCellWidthEmu: 609600,
        defaultCellHeightEmu: 190500,
        shapeBlockGapXEmu: 609600 * 4,
        shapeBlockGapYEmu: 190500 * 6,
    };
}

function filesFor(drawingXml) {
    return new Map([[DRAWING_PATH, encoder.encode(drawingXml)]]);
}

describe("grouped and absolute shapes (fixture C)", () => {
    it("parses shapes nested in groups with their group path", () => {
        const drawing = `<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:twoCellAnchor>
            <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
            <xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="10" name="Group 1"/></xdr:nvGrpSpPr><xdr:grpSpPr/>
            ${shapeXml("11", "Inner A", "A")}
            ${shapeXml("12", "Inner B", "B")}
            </xdr:grpSp>
            <xdr:clientData/>
        </xdr:twoCellAnchor></xdr:wsDr>`;
        const shapes = parseDrawingShapes(filesFor(drawing), "Sheet1", "xl/worksheets/sheet1.xml", deps());
        expect(shapes.length).toBe(2);
        expect(shapes.map((shape) => shape.drawingId).sort()).toEqual(["11", "12"]);
        for (const shape of shapes) {
            expect(shape.groupPath).toBe("Group 1");
            expect(shape.stableId).toContain(DRAWING_PATH);
            expect(shape.stableId).toContain("Group 1");
        }
    });

    it("parses absolute anchors as exact geometry", () => {
        const drawing = `<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:absoluteAnchor>
            <xdr:pos x="1000" y="2000"/>
            <xdr:ext cx="5000" cy="6000"/>
            ${shapeXml("21", "Free", "F")}
        </xdr:absoluteAnchor></xdr:wsDr>`;
        const shapes = parseDrawingShapes(filesFor(drawing), "Sheet1", "xl/worksheets/sheet1.xml", deps());
        expect(shapes.length).toBe(1);
        expect(shapes[0].anchorKind).toBe("absolute");
        expect(shapes[0].geometryConfidence).toBe("EXACT");
        expect(shapes[0].bbox).toEqual({ left: 1000, top: 2000, right: 6000, bottom: 8000 });
    });

    it("scopes stable ids to one drawing file", () => {
        const drawing = `<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:twoCellAnchor>
            <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
            ${shapeXml("31", "Solo", "S")}
            <xdr:clientData/>
        </xdr:twoCellAnchor></xdr:wsDr>`;
        const shapes = parseDrawingShapes(filesFor(drawing), "Sheet1", "xl/worksheets/sheet1.xml", deps());
        expect(shapes[0].stableId).toBe(`${DRAWING_PATH}#31`);
        expect(shapes[0].groupPath).toBe(null);
    });

    it("joins shape paragraphs with line breaks, not empty strings", () => {
        const text = `<xdr:txBody><a:p><a:r><a:t>DA10</a:t></a:r></a:p><a:p><a:r><a:t>Shipment Info</a:t></a:r></a:p></xdr:txBody>`;
        const drawing = `<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:twoCellAnchor>
            <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
            <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="41" name="Box"/></xdr:nvSpPr><xdr:spPr/>${text}</xdr:sp>
            <xdr:clientData/>
        </xdr:twoCellAnchor></xdr:wsDr>`;
        const shapes = parseDrawingShapes(filesFor(drawing), "Sheet1", "xl/worksheets/sheet1.xml", deps());
        expect(shapes[0].text).toBe("DA10\nShipment Info");
    });

    it("maps nested boxes through group frames instead of sharing one box", () => {
        const child = (id, x, w) =>
            `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${id}" name="C${id}"/></xdr:nvSpPr>` +
            `<xdr:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="${w}" cy="50"/></a:xfrm></xdr:spPr></xdr:sp>`;
        const drawing = `<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:twoCellAnchor>
            <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
            <xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="10" name="Group 1"/></xdr:nvGrpSpPr>
            <xdr:grpSpPr><a:xfrm><a:off x="1000" y="0"/><a:ext cx="2000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></xdr:grpSpPr>
            ${child("51", 0, 50)}
            ${child("52", 50, 50)}
            </xdr:grpSp>
            <xdr:clientData/>
        </xdr:twoCellAnchor></xdr:wsDr>`;
        const shapes = parseDrawingShapes(filesFor(drawing), "Sheet1", "xl/worksheets/sheet1.xml", deps());
        // The group frame fills the anchor box A1:J10, so each half of the frame is half the anchor box.
        expect(shapes.length).toBe(2);
        expect(shapes[0].bbox).toEqual({ left: 0, top: 0, right: 2743200, bottom: 857250 });
        expect(shapes[1].bbox).toEqual({ left: 2743200, top: 0, right: 5486400, bottom: 857250 });
    });

    it("keeps grouped shapes in the same space as ungrouped shapes", () => {
        const anchor = (from, to, body) =>
            `<xdr:twoCellAnchor><xdr:from><xdr:col>${from}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${from}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
            `<xdr:to><xdr:col>${to}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${to}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>${body}<xdr:clientData/></xdr:twoCellAnchor>`;
        const grouped =
            `<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="10" name="Group 1"/></xdr:nvGrpSpPr>` +
            `<xdr:grpSpPr><a:xfrm><a:off x="400000" y="400000"/><a:ext cx="200000" cy="200000"/><a:chOff x="400000" y="400000"/><a:chExt cx="200000" cy="200000"/></a:xfrm></xdr:grpSpPr>` +
            `<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="11" name="In"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="400000" y="400000"/><a:ext cx="200000" cy="200000"/></a:xfrm></xdr:spPr></xdr:sp></xdr:grpSp>`;
        const drawing = `<xdr:wsDr xmlns:xdr="x" xmlns:a="y">${anchor(2, 4, grouped)}${anchor(2, 4, shapeXml("21", "Out", "O"))}</xdr:wsDr>`;
        const shapes = parseDrawingShapes(filesFor(drawing), "Sheet1", "xl/worksheets/sheet1.xml", deps());
        expect(shapes.length).toBe(2);
        expect(shapes[0].bbox).toEqual(shapes[1].bbox);
    });
});

describe("line shapes", () => {
    it("reads a plain shape with a line preset as a loose connector", () => {
        const drawing = `<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:twoCellAnchor>
            <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
            <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="5" name="Line 5"/></xdr:nvSpPr>
            <xdr:spPr><a:prstGeom prst="line"/><a:ln><a:tailEnd type="triangle"/></a:ln></xdr:spPr></xdr:sp>
            <xdr:clientData/>
        </xdr:twoCellAnchor></xdr:wsDr>`;
        const [line] = parseDrawingShapes(filesFor(drawing), "Sheet1", "xl/worksheets/sheet1.xml", deps());
        expect(line.connection).toMatchObject({ startId: null, endId: null, arrow: "end" });
        expect(line.connection.endPoint.x).toBeGreaterThan(line.connection.startPoint.x);
        expect(line.calloutTip).toBeUndefined();
    });
});

describe("shape text marks", () => {
    function textShape(body) {
        return `<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:twoCellAnchor>
            <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
            <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
            <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="5" name="Note"/></xdr:nvSpPr><xdr:spPr/><xdr:txBody>${body}</xdr:txBody></xdr:sp>
            <xdr:clientData/>
        </xdr:twoCellAnchor></xdr:wsDr>`;
    }

    it("marks struck runs and keeps plain text unchanged", () => {
        const body = `<a:p><a:r><a:rPr strike="noStrike"/><a:t>keep </a:t></a:r><a:r><a:rPr strike="sngStrike"/><a:t>old rule </a:t></a:r>
            <a:r><a:t>end</a:t></a:r></a:p><a:p><a:r><a:rPr strike="dblStrike"/><a:t>gone</a:t></a:r></a:p>`;
        const [shape] = parseDrawingShapes(filesFor(textShape(body)), "Sheet1", "xl/worksheets/sheet1.xml", deps());
        expect(shape.text).toBe("keep old rule end\ngone");
        expect(shape.markedText).toBe("keep ~~old rule~~ end\n~~gone~~");
    });

    it("takes the strike default of the paragraph and splits lines at breaks", () => {
        const body = `<a:p><a:pPr><a:defRPr strike="sngStrike"/></a:pPr><a:r><a:t>a</a:t></a:r><a:br/><a:r><a:rPr strike="noStrike"/><a:t>b</a:t></a:r></a:p>`;
        const [shape] = parseDrawingShapes(filesFor(textShape(body)), "Sheet1", "xl/worksheets/sheet1.xml", deps());
        expect(shape.text).toBe("a\nb");
        expect(shape.markedText).toBe("~~a~~\nb");
    });

    it("leaves markedText out when nothing is struck", () => {
        const [shape] = parseDrawingShapes(
            filesFor(textShape(`<a:p><a:r><a:t>plain</a:t></a:r></a:p>`)),
            "Sheet1",
            "xl/worksheets/sheet1.xml",
            deps()
        );
        expect(shape.markedText).toBeUndefined();
    });
});
