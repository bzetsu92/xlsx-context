import { describe, expect, it } from "vitest";
import { renderShapeSvg } from "../../../../../src/extraction/assets/office-drawing.ts";
import { getFirstDescendantByLocalName, xmlToDocument } from "../../../../../src/extraction/xml/xml-utils.ts";

const decoder = new TextDecoder();

function connectorXml(lineInner) {
    return `<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:twoCellAnchor>
        <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:cxnSp><xdr:nvCxnSpPr><xdr:cNvPr id="5" name="Link"/></xdr:nvCxnSpPr>
        <xdr:spPr><a:ln><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>${lineInner}</a:ln></xdr:spPr>
        </xdr:cxnSp>
    </xdr:twoCellAnchor></xdr:wsDr>`;
}

function shapeXml(preset) {
    return `<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:twoCellAnchor>
        <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
        <xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
        <xdr:sp><xdr:nvSpPr><xdr:cNvPr id="7" name="Box"/></xdr:nvSpPr>
        <xdr:spPr><a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>
    </xdr:twoCellAnchor></xdr:wsDr>`;
}

function renderFirstShape(drawingXml, localName) {
    const doc = xmlToDocument(drawingXml);
    const anchor = getFirstDescendantByLocalName(doc, "twoCellAnchor");
    const node = getFirstDescendantByLocalName(anchor, localName);
    const result = renderShapeSvg(node, anchor, "Sheet1", 1);
    return decoder.decode(result.data);
}

describe("connector svg arrowheads", () => {
    it("draws no arrow without declared ends", () => {
        const svg = renderFirstShape(connectorXml(""), "cxnSp");
        expect(svg).not.toContain("marker-start");
        expect(svg).not.toContain("marker-end");
    });

    it("draws one arrow at the declared end", () => {
        const svg = renderFirstShape(connectorXml(`<a:tailEnd type="triangle"/>`), "cxnSp");
        expect(svg).not.toContain("marker-start");
        expect(svg).toContain('marker-end="url(#arrow-end)"');
    });

    it("draws arrows at both declared ends", () => {
        const svg = renderFirstShape(connectorXml(`<a:headEnd type="triangle"/><a:tailEnd type="triangle"/>`), "cxnSp");
        expect(svg).toContain('marker-start="url(#arrow-start)"');
        expect(svg).toContain('marker-end="url(#arrow-end)"');
    });
});

describe("preset svg warnings", () => {
    it("marks unsupported presets as approximate", () => {
        const doc = xmlToDocument(shapeXml("diamond"));
        const anchor = getFirstDescendantByLocalName(doc, "twoCellAnchor");
        const node = getFirstDescendantByLocalName(anchor, "sp");
        const result = renderShapeSvg(node, anchor, "Sheet1", 1);
        expect(result.warning).toContain("diamond");
        expect(decoder.decode(result.data)).toContain("data-prst=");
    });
});
