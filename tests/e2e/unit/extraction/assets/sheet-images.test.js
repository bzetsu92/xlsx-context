import { describe, expect, it } from "vitest";
import * as sheetImages from "../../../../../src/extraction/assets/sheet-images.ts";

function createDeps() {
    return {
        parseRelationships(files, relsPath, sourcePath) {
            const relBytes = files.get(relsPath);
            const relations = new Map();
            if (!relBytes) return relations;
            const doc = new DOMParser().parseFromString(new TextDecoder().decode(relBytes), "application/xml");
            for (const node of Array.from(doc.getElementsByTagName("Relationship"))) {
                const id = node.getAttribute("Id") || "";
                const target = node.getAttribute("Target") || "";
                if (!id || !target) continue;
                const baseDirParts = sourcePath.split("/").slice(0, -1);
                const inputParts = target.split("/");
                const parts = target.startsWith("/") ? [] : baseDirParts;
                for (const part of inputParts) {
                    if (!part || part === ".") continue;
                    if (part === "..") {
                        parts.pop();
                    } else {
                        parts.push(part);
                    }
                }
                relations.set(id, parts.join("/"));
            }
            return relations;
        },
        buildRelsPath(sourcePath) {
            const parts = sourcePath.split("/");
            const fileName = parts.pop() || "";
            const dir = parts.join("/");
            return `${dir}/_rels/${fileName}.rels`;
        },
        xmlToDocument(xmlText) {
            return new DOMParser().parseFromString(xmlText, "application/xml");
        },
        decodeXmlText(bytes) {
            return new TextDecoder().decode(bytes);
        },
        getElementsByLocalName(root, localName) {
            return Array.from(root.getElementsByTagName("*")).filter((element) => element.localName === localName);
        },
        getFirstDescendantByLocalName(root, localName) {
            return this.getElementsByLocalName(root, localName)[0] || null;
        },
        getDirectChildByLocalName(root, localName) {
            if (!root) return null;
            for (const node of Array.from(root.childNodes)) {
                if (node.nodeType === Node.ELEMENT_NODE && node.localName === localName) {
                    return node;
                }
            }
            return null;
        },
        getTextContent(node) {
            return (node?.textContent || "").replace(/\r\n/g, "\n");
        },
        colToLetters(col) {
            let current = col;
            let result = "";
            while (current > 0) {
                const remainder = (current - 1) % 26;
                result = String.fromCharCode(65 + remainder) + result;
                current = Math.floor((current - 1) / 26);
            }
            return result;
        },
        drawingHelper: null,
        defaultCellWidthEmu: 100,
        defaultCellHeightEmu: 20,
        shapeBlockGapXEmu: 250,
        shapeBlockGapYEmu: 60,
    };
}

describe("xlsx-context sheet images", () => {
    it("parses drawing images through drawing relationships", () => {
        const api = sheetImages;
        const deps = createDeps();
        const worksheetRels = `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Target="../drawings/drawing1.xml"/>
      </Relationships>`;
        const drawingXml = `<?xml version="1.0" encoding="UTF-8"?>
      <xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <xdr:oneCellAnchor>
          <xdr:from><xdr:col>1</xdr:col><xdr:row>2</xdr:row></xdr:from>
          <xdr:pic><xdr:blipFill><a:blip r:embed="rIdImg1"/></xdr:blipFill></xdr:pic>
        </xdr:oneCellAnchor>
      </xdr:wsDr>`;
        const drawingRels = `<?xml version="1.0" encoding="UTF-8"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rIdImg1" Target="../media/image1.png"/>
      </Relationships>`;
        const files = new Map([
            ["xl/worksheets/_rels/sheet1.xml.rels", new TextEncoder().encode(worksheetRels)],
            ["xl/drawings/drawing1.xml", new TextEncoder().encode(drawingXml)],
            ["xl/drawings/_rels/drawing1.xml.rels", new TextEncoder().encode(drawingRels)],
            ["xl/media/image1.png", new Uint8Array([1, 2, 3])],
        ]);

        const images = api.parseDrawingImages(files, "Sheet/A", "xl/worksheets/sheet1.xml", deps);

        expect(images).toEqual([
            {
                sheetName: "Sheet/A",
                filename: "image_001.png",
                path: "assets/Sheet_A/image_001.png",
                anchor: "B3",
                data: new Uint8Array([1, 2, 3]),
                mediaPath: "xl/media/image1.png",
                // The anchor has no extent, so the box fills one default cell at B3.
                bbox: { left: 100, top: 40, right: 200, bottom: 60 },
            },
        ]);
    });
});
