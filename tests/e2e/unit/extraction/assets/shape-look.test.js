import { describe, expect, it } from "vitest";
import { colorName, readShapeLook, readThemeColors } from "../../../../../src/extraction/assets/shape-look.ts";
import { getFirstDescendantByLocalName, xmlToDocument } from "../../../../../src/extraction/xml/xml-utils.ts";

const decode = (bytes) => new TextDecoder().decode(bytes);
const THEME = readThemeColors(new Map(), decode);

function shapeNode(inner) {
    const doc = xmlToDocument(`<xdr:wsDr xmlns:xdr="x" xmlns:a="y"><xdr:sp>${inner}</xdr:sp></xdr:wsDr>`);
    return getFirstDescendantByLocalName(doc, "sp");
}

describe("shape look", () => {
    it("reads RGB fill, line color, dash, and width", () => {
        const look = readShapeLook(
            shapeNode(`<xdr:spPr><a:solidFill><a:srgbClr val="0070c0"/></a:solidFill>
                <a:ln w="15875"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></xdr:spPr>`),
            THEME
        );
        expect(look).toEqual({ fill: "#0070C0", line: "#FF0000", dash: "sysDash", lineWidthEmu: 15875 });
    });

    it("resolves theme colors with luminance transforms", () => {
        const look = readShapeLook(
            shapeNode(`<xdr:spPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr></a:solidFill>
                <a:ln><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></xdr:spPr>`),
            THEME
        );
        // Office accent1 at 60% luminance plus 40% is the "lighter 40%" swatch.
        expect(look.fill).toBe("#95B3D7");
        expect(look.line).toBe("#000000");
    });

    it("takes missing fill and line from the style element", () => {
        const look = readShapeLook(
            shapeNode(`<xdr:spPr/><xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>
                <a:fillRef idx="1"><a:schemeClr val="accent2"/></a:fillRef></xdr:style>`),
            THEME
        );
        expect(look.fill).toBe("#C0504D");
        expect(look.line).toBe("#28415F");
    });

    it("keeps no fill and no line apart from unknown values", () => {
        expect(
            readShapeLook(shapeNode(`<xdr:spPr><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr>`), THEME)
        ).toMatchObject({
            fill: "none",
            line: "none",
        });
        expect(readShapeLook(shapeNode(`<xdr:spPr/><xdr:style><a:fillRef idx="0"/></xdr:style>`), THEME)).toMatchObject(
            {
                fill: "none",
                line: null,
            }
        );
        expect(
            readShapeLook(shapeNode(`<xdr:spPr><a:blipFill/><a:ln><a:custDash/></a:ln></xdr:spPr>`), THEME)
        ).toMatchObject({
            fill: "picture",
            dash: "custom",
        });
    });

    it("reads the first gradient stop, system colors, and preset colors", () => {
        const gradient = readShapeLook(
            shapeNode(`<xdr:spPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="00B050"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst></a:gradFill>
                <a:ln><a:solidFill><a:sysClr val="windowText" lastClr="000000"/></a:solidFill></a:ln></xdr:spPr>`),
            THEME
        );
        expect(gradient).toMatchObject({ fill: "#00B050", line: "#000000" });
        const preset = readShapeLook(
            shapeNode(`<xdr:spPr><a:solidFill><a:prstClr val="red"/></a:solidFill></xdr:spPr>`),
            THEME
        );
        expect(preset.fill).toBe("#FF0000");
    });

    it("reads theme colors from the workbook theme", () => {
        const theme = `<a:theme xmlns:a="y"><a:themeElements><a:clrScheme name="X"><a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1>
            <a:accent1><a:srgbClr val="123456"/></a:accent1></a:clrScheme></a:themeElements></a:theme>`;
        const colors = readThemeColors(new Map([["xl/theme/theme1.xml", new TextEncoder().encode(theme)]]), decode);
        expect(colors.dk1).toBe("111111");
        expect(colors.accent1).toBe("123456");
        expect(colors.accent2).toBe("C0504D");
    });
});
