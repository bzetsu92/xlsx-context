import { createSafeSheetAssetDir } from "./sheet-assets.js";
import { getDirectChildByLocalName, getElementsByLocalName, getTextContent } from "../xml/xml-utils.js";

type SvgRenderResult = {
    filename: string;
    path: string;
    data: Uint8Array;
    warning: string | null;
} | null;

const textEncoder = new TextEncoder();

function escapeXml(text: string): string {
    return String(text || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

/** Converts EMUs to pixels at 96 dpi. Falls back when the value is missing or not positive. */
function emuToPx(emu: number | null, fallback: number): number {
    if (!Number.isFinite(emu as number) || (emu as number) <= 0) return fallback;
    return Math.max(1, Math.round((emu as number) / 9525));
}

/** Reads sRGB colors first, then Office scheme colors. Returns null without a known color. */
function parseHexColor(root: Element | null): string | null {
    const srgb = getElementsByLocalName(root, "srgbClr")[0] || null;
    if (srgb?.getAttribute("val")) {
        return `#${String(srgb.getAttribute("val")).trim()}`;
    }
    const scheme = getElementsByLocalName(root, "schemeClr")[0] || null;
    const schemeVal = String(scheme?.getAttribute("val") || "").trim();
    const schemeMap: Record<string, string> = {
        accent1: "#4472C4",
        accent2: "#ED7D31",
        accent3: "#A5A5A5",
        accent4: "#FFC000",
        accent5: "#5B9BD5",
        accent6: "#70AD47",
        tx1: "#000000",
        tx2: "#44546A",
        lt1: "#FFFFFF",
        lt2: "#E7E6E6",
    };
    return schemeMap[schemeVal] || null;
}

/** Reads shape text by paragraph. Runs join directly, paragraphs join with line breaks. */
function parseShapeText(shapeNode: Element | null): string {
    const paragraphs = getElementsByLocalName(shapeNode, "p");
    if (paragraphs.length === 0) {
        return getElementsByLocalName(shapeNode, "t")
            .map((node) => getTextContent(node).trim())
            .filter(Boolean)
            .join("")
            .trim();
    }
    return paragraphs
        .map((paragraph) =>
            getElementsByLocalName(paragraph, "t")
                .map((node) => getTextContent(node))
                .join("")
        )
        .join("\n")
        .trim();
}

/**
 * Detects text boxes, rectangles, and connectors.
 * Other preset or custom geometry is "preset". Callers draw its box only.
 * Returns null for non-shape nodes.
 */
function parseShapeKind(shapeNode: Element | null): "textbox" | "rect" | "connector" | "preset" | null {
    if (!shapeNode) return null;
    if (shapeNode.localName === "cxnSp") {
        return "connector";
    }
    if (shapeNode.localName !== "sp") {
        return null;
    }
    const nvSpPr = getDirectChildByLocalName(shapeNode, "nvSpPr");
    const cNvSpPr = getDirectChildByLocalName(nvSpPr, "cNvSpPr");
    if (cNvSpPr?.getAttribute("txBox") === "1") {
        return "textbox";
    }
    const spPr = getDirectChildByLocalName(shapeNode, "spPr");
    const prstGeom = getDirectChildByLocalName(spPr, "prstGeom");
    if (String(prstGeom?.getAttribute("prst") || "").trim() === "rect") {
        return "rect";
    }
    return "preset";
}

/** Reads the preset name, or "custom" for custom geometry. */
function parseShapePreset(shapeNode: Element): string {
    const spPr = getDirectChildByLocalName(shapeNode, "spPr");
    const prst = String(getDirectChildByLocalName(spPr, "prstGeom")?.getAttribute("prst") || "").trim();
    if (prst) return prst;
    return getDirectChildByLocalName(spPr, "custGeom") ? "custom" : "unknown";
}

/** Reads shape size in pixels. Falls back to 160 by 48. */
function parseShapeDimensions(anchor: Element, shapeNode: Element | null): { widthPx: number; heightPx: number } {
    const extNode =
        getDirectChildByLocalName(anchor, "ext") ||
        getDirectChildByLocalName(
            getDirectChildByLocalName(getDirectChildByLocalName(shapeNode || anchor, "spPr"), "xfrm"),
            "ext"
        );
    const widthEmu = Number(extNode?.getAttribute("cx") || "");
    const heightEmu = Number(extNode?.getAttribute("cy") || "");
    return {
        widthPx: emuToPx(widthEmu, 160),
        heightPx: emuToPx(heightEmu, 48),
    };
}

/**
 * Draws a box with the shape fill, line, and text.
 * A set preset means the real outline is not drawn. A dashed line and data-prst mark that.
 */
function renderRectLikeSvg(
    shapeNode: Element,
    anchor: Element,
    text: string,
    treatAsTextbox: boolean,
    preset: string | null = null
): string {
    const { widthPx, heightPx } = parseShapeDimensions(anchor, shapeNode);
    const spPr = getDirectChildByLocalName(shapeNode, "spPr");
    const fillColor =
        parseHexColor(getDirectChildByLocalName(spPr, "solidFill")) || (treatAsTextbox ? "#FFFFFF" : "#F3F3F3");
    const lineNode = getDirectChildByLocalName(spPr, "ln");
    const strokeColor = parseHexColor(lineNode) || "#333333";
    const strokeWidth = Math.max(1, Math.round(Number(lineNode?.getAttribute("w") || "") / 9525) || 1);
    const safeText = escapeXml(text);
    const textMarkup = safeText
        ? `<text x="${Math.round(widthPx / 2)}" y="${Math.round(heightPx / 2)}" text-anchor="middle" dominant-baseline="middle" font-size="14" font-family="sans-serif" fill="#000000">${safeText}</text>`
        : "";
    return [
        `<svg xmlns="http://www.w3.org/2000/svg" width="${widthPx}" height="${heightPx}" viewBox="0 0 ${widthPx} ${heightPx}">`,
        `  <rect x="1" y="1" width="${Math.max(1, widthPx - 2)}" height="${Math.max(1, heightPx - 2)}" fill="${fillColor}" stroke="${strokeColor}" stroke-width="${strokeWidth}"${preset ? ` stroke-dasharray="6 3" data-prst="${escapeXml(preset)}"` : ""}/>`,
        textMarkup ? `  ${textMarkup}` : "",
        `</svg>`,
    ]
        .filter(Boolean)
        .join("\n");
}

/** True when a line end element draws an arrowhead. Absent ends draw no arrow. */
function hasArrowHead(endNode: Element | null): boolean {
    const type = String(endNode?.getAttribute("type") || "").trim();
    return type !== "" && type !== "none";
}

/** Draws the connector line with exactly the arrowheads the XML declares. */
function renderConnectorSvg(shapeNode: Element, anchor: Element): string {
    const { widthPx, heightPx } = parseShapeDimensions(anchor, shapeNode);
    const spPr = getDirectChildByLocalName(shapeNode, "spPr");
    const lineNode = getDirectChildByLocalName(spPr, "ln");
    const strokeColor = parseHexColor(lineNode) || "#333333";
    const strokeWidth = Math.max(1, Math.round(Number(lineNode?.getAttribute("w") || "") / 9525) || 1);
    // DrawingML puts headEnd at the line start and tailEnd at the line end.
    const atStart = hasArrowHead(getDirectChildByLocalName(lineNode, "headEnd"));
    const atEnd = hasArrowHead(getDirectChildByLocalName(lineNode, "tailEnd"));
    const effectiveHeight = Math.max(heightPx, 24);
    const y = Math.round(effectiveHeight / 2);
    const markers = [
        `  <defs>`,
        `    <marker id="arrow-start" markerWidth="10" markerHeight="10" refX="1" refY="3" orient="auto" markerUnits="strokeWidth">`,
        `      <path d="M9,0 L9,6 L0,3 z" fill="${strokeColor}"/>`,
        `    </marker>`,
        `    <marker id="arrow-end" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto" markerUnits="strokeWidth">`,
        `      <path d="M0,0 L0,6 L9,3 z" fill="${strokeColor}"/>`,
        `    </marker>`,
        `  </defs>`,
    ].join("\n");
    const ends = [atStart ? `marker-start="url(#arrow-start)"` : "", atEnd ? `marker-end="url(#arrow-end)"` : ""]
        .filter(Boolean)
        .join(" ");
    return [
        `<svg xmlns="http://www.w3.org/2000/svg" width="${widthPx}" height="${effectiveHeight}" viewBox="0 0 ${widthPx} ${effectiveHeight}">`,
        markers,
        `  <line x1="2" y1="${y}" x2="${Math.max(2, widthPx - 4)}" y2="${y}" stroke="${strokeColor}" stroke-width="${strokeWidth}"${ends ? ` ${ends}` : ""}/>`,
        `</svg>`,
    ].join("\n");
}

/**
 * Renders shapes and connectors as SVG assets.
 * Other presets get a dashed fallback box. Returns null for non-shape nodes.
 *
 * @param shapeIndex 1-based shape number used in the file name.
 */
function renderShapeSvg(shapeNode: Element, anchor: Element, sheetName: string, shapeIndex: number): SvgRenderResult {
    const kind = parseShapeKind(shapeNode);
    if (!kind) return null;
    let svg = "";
    let warning: string | null = null;
    if (kind === "connector") {
        svg = renderConnectorSvg(shapeNode, anchor);
    } else {
        const preset = kind === "preset" ? parseShapePreset(shapeNode) : null;
        svg = renderRectLikeSvg(shapeNode, anchor, parseShapeText(shapeNode), kind === "textbox", preset);
        if (preset) {
            warning = `Shape outline is approximate; preset '${preset}' has no exact renderer.`;
        }
    }
    const safeDir = createSafeSheetAssetDir(sheetName);
    const filename = `shape_${String(shapeIndex).padStart(3, "0")}.svg`;
    return {
        filename,
        path: `assets/${safeDir}/${filename}`,
        data: textEncoder.encode(`${svg}\n`),
        warning,
    };
}

/** Renders drawing shapes as SVG assets. */
export { renderShapeSvg };
