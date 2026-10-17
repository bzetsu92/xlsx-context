import { describe, expect, it } from "vitest";
import * as officeDrawing from "../../../../../src/extraction/assets/office-drawing.ts";

function parseXml(xmlText) {
    return new DOMParser().parseFromString(xmlText, "application/xml");
}

describe("xlsx-context office drawing", () => {
    it("renders textbox shapes as svg assets with sanitized sheet directories", () => {
        const api = officeDrawing;
        const doc = parseXml(`
      <root xmlns:xdr="xdr" xmlns:a="a">
        <xdr:twoCellAnchor>
          <xdr:ext cx="1905000" cy="476250"/>
          <xdr:sp>
            <xdr:nvSpPr>
              <xdr:cNvSpPr txBox="1"/>
            </xdr:nvSpPr>
            <xdr:spPr>
              <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
              <a:ln w="9525"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln>
            </xdr:spPr>
            <xdr:txBody>
              <a:p><a:r><a:t>Hello</a:t></a:r></a:p>
            </xdr:txBody>
          </xdr:sp>
        </xdr:twoCellAnchor>
      </root>
    `);
        const anchor = doc.getElementsByTagName("xdr:twoCellAnchor")[0];
        const shape = doc.getElementsByTagName("xdr:sp")[0];

        const result = api.renderShapeSvg(shape, anchor, "A/B:Sheet", 1);
        const svgText = new TextDecoder().decode(result.data);

        expect(result.filename).toBe("shape_001.svg");
        expect(result.path).toBe("assets/A_B_Sheet/shape_001.svg");
        expect(svgText).toContain("<svg");
        expect(svgText).toContain("Hello");
        expect(svgText).toContain("#FF0000");
        expect(svgText).toContain("#00FF00");
    });

    it("renders connector shapes with arrow markers", () => {
        const api = officeDrawing;
        const doc = parseXml(`
      <root xmlns:xdr="xdr" xmlns:a="a">
        <xdr:twoCellAnchor>
          <xdr:ext cx="2857500" cy="190500"/>
          <xdr:cxnSp>
            <xdr:spPr>
              <a:ln w="19050"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>
            </xdr:spPr>
          </xdr:cxnSp>
        </xdr:twoCellAnchor>
      </root>
    `);
        const anchor = doc.getElementsByTagName("xdr:twoCellAnchor")[0];
        const shape = doc.getElementsByTagName("xdr:cxnSp")[0];

        const result = api.renderShapeSvg(shape, anchor, "Sheet1", 2);
        const svgText = new TextDecoder().decode(result.data);

        expect(result.path).toBe("assets/Sheet1/shape_002.svg");
        expect(svgText).toContain("<marker");
        expect(svgText).toContain('marker-end="url(#arrow-end)"');
        expect(svgText).toContain("#123456");
    });

    it("renders a dashed fallback box with the preset name for presets without an exact outline", () => {
        const api = officeDrawing;
        const doc = parseXml(`
      <root xmlns:xdr="xdr" xmlns:a="a">
        <xdr:twoCellAnchor>
          <xdr:sp>
            <xdr:spPr>
              <a:prstGeom prst="ellipse"/>
            </xdr:spPr>
            <xdr:txBody><a:p><a:r><a:t>A &amp; B</a:t></a:r></a:p></xdr:txBody>
          </xdr:sp>
        </xdr:twoCellAnchor>
      </root>
    `);
        const anchor = doc.getElementsByTagName("xdr:twoCellAnchor")[0];
        const shape = doc.getElementsByTagName("xdr:sp")[0];

        const result = api.renderShapeSvg(shape, anchor, "Sheet1", 3);
        const svg = new TextDecoder().decode(result.data);

        expect(result.filename).toBe("shape_003.svg");
        expect(svg).toContain('stroke-dasharray="6 3" data-prst="ellipse"');
        expect(svg).toContain(">A &amp; B</text>");
    });

    it("names custom geometry in the fallback box", () => {
        const api = officeDrawing;
        const doc = parseXml(`
      <root xmlns:xdr="xdr" xmlns:a="a">
        <xdr:twoCellAnchor>
          <xdr:sp>
            <xdr:spPr><a:custGeom/></xdr:spPr>
          </xdr:sp>
        </xdr:twoCellAnchor>
      </root>
    `);
        const anchor = doc.getElementsByTagName("xdr:twoCellAnchor")[0];
        const shape = doc.getElementsByTagName("xdr:sp")[0];

        const svg = new TextDecoder().decode(api.renderShapeSvg(shape, anchor, "Sheet1", 1).data);

        expect(svg).toContain('data-prst="custom"');
    });
});
