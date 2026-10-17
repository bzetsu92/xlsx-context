import { describe, expect, it } from "vitest";
import { createRelsParser } from "../../../../../src/extraction/workbook/rels-parser.ts";
import * as xmlUtils from "../../../../../src/extraction/xml/xml-utils.ts";

function bootRelsParser() {
    return createRelsParser(xmlUtils);
}

describe("xlsx-context rels parser", () => {
    it("parses relationship targets into a map", () => {
        const api = bootRelsParser();
        const files = new Map([
            [
                "xl/_rels/workbook.xml.rels",
                new TextEncoder().encode(
                    '<?xml version="1.0" encoding="UTF-8"?>' +
                        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                        '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/>' +
                        '<Relationship Id="rId2" Target="../sharedStrings.xml"/>' +
                        "</Relationships>"
                ),
            ],
        ]);

        const rels = api.parseRelationships(files, "xl/_rels/workbook.xml.rels", "xl/workbook.xml");

        expect(rels.get("rId1")).toBe("xl/worksheets/sheet1.xml");
        expect(rels.get("rId2")).toBe("sharedStrings.xml");
    });

    it("keeps external relationship targets as-is", () => {
        const api = bootRelsParser();
        const files = new Map([
            [
                "xl/worksheets/_rels/sheet1.xml.rels",
                new TextEncoder().encode(
                    '<?xml version="1.0" encoding="UTF-8"?>' +
                        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                        '<Relationship Id="rId1" Target="https://example.com/docs?a=1&amp;b=2" TargetMode="External"/>' +
                        "</Relationships>"
                ),
            ],
        ]);

        const rels = api.parseRelationshipEntries(
            files,
            "xl/worksheets/_rels/sheet1.xml.rels",
            "xl/worksheets/sheet1.xml"
        );

        expect(rels.get("rId1")).toEqual({
            target: "https://example.com/docs?a=1&b=2",
            targetMode: "External",
            type: "",
        });
    });
});
