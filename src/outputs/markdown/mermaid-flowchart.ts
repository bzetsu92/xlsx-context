import type { ParsedShapeAsset } from "../../model/index.js";
import type { ShapeGraph } from "../../structure/layout/shape-graph.js";

/** Mermaid node brackets per DrawingML preset. Other presets use a plain box. */
const NODE_BRACKETS: Record<string, [string, string]> = {
    roundRect: ["(", ")"],
    flowChartAlternateProcess: ["(", ")"],
    ellipse: ["((", "))"],
    flowChartConnector: ["((", "))"],
    diamond: ["{", "}"],
    flowChartDecision: ["{", "}"],
    flowChartTerminator: ["([", "])"],
    parallelogram: ["[/", "/]"],
    flowChartInputOutput: ["[/", "/]"],
    can: ["[(", ")]"],
    flowChartMagneticDisk: ["[(", ")]"],
    flowChartPredefinedProcess: ["[[", "]]"],
    hexagon: ["{{", "}}"],
    flowChartPreparation: ["{{", "}}"],
};

const EDGE_ARROWS: Record<"none" | "end" | "both", string> = { none: "---", end: "-->", both: "<-->" };
const DASHED_EDGE_ARROWS: Record<"none" | "end" | "both", string> = { none: "-.-", end: "-.->", both: "<-.->" };

/** True when the shape line uses a dash preset other than solid. */
function isDashed(shape: ParsedShapeAsset | undefined): boolean {
    const dash = shape?.look?.dash;
    return !!dash && dash !== "solid";
}

/** Reads the preset name from kind labels such as `Shape (flowChartProcess)`. */
function presetOf(shape: ParsedShapeAsset): string {
    if (shape.kind === "Rectangle") return "rect";
    return /\(([A-Za-z0-9]+)\)$/.exec(shape.kind || "")?.[1] || "";
}

/**
 * Escapes text for a quoted Mermaid label. Mermaid reads `#name;` as an entity, so `#` goes first.
 * Line breaks become <br>, so a label never starts a new line and never closes the code fence.
 */
function escapeMermaidLabel(text: string): string {
    return String(text || "")
        .replace(/#/g, "#35;")
        .replace(/"/g, "#quot;")
        .replace(/</g, "#lt;")
        .replace(/>/g, "#gt;")
        .replace(/\r\n?|\n/g, "<br>")
        .trim();
}

function shapeNodeId(shapeIndex: number): string {
    return `sh${String(shapeIndex + 1).padStart(3, "0")}`;
}

/** Renders one shape node with brackets matching its preset. Falls back to a plain box. */
function renderShapeNode(shapes: ParsedShapeAsset[], shapeIndex: number): string {
    const shape = shapes[shapeIndex];
    const [open, close] = NODE_BRACKETS[presetOf(shape)] || ["[", "]"];
    const label = escapeMermaidLabel(shape.text || shape.name || shapeNodeId(shapeIndex));
    return `    ${shapeNodeId(shapeIndex)}${open}"${label}"${close}`;
}

/**
 * Copies the fill, line color, and dash of a node, so the drawn chart looks like the sheet.
 * Returns null when the shape keeps default colors.
 */
function renderNodeStyle(shape: ParsedShapeAsset | undefined, shapeIndex: number): string | null {
    const look = shape?.look;
    if (!look) return null;
    const parts: string[] = [];
    if (look.fill?.startsWith("#")) parts.push(`fill:${look.fill}`);
    if (look.line?.startsWith("#")) parts.push(`stroke:${look.line}`);
    if (isDashed(shape)) parts.push("stroke-dasharray:5 5");
    // White boxes with black lines are the default look. Styling them only adds noise.
    const plainFill = !look.fill || look.fill === "none" || look.fill === "#FFFFFF";
    const plainLine = !look.line || look.line === "none" || look.line === "#000000";
    if (parts.length === 0 || (plainFill && plainLine && !isDashed(shape))) return null;
    return `    style ${shapeNodeId(shapeIndex)} ${parts.join(",")}`;
}

/**
 * Returns the fenced Mermaid block.
 * Returns an empty string when the set holds no edges.
 */
function renderMermaidFlowchart(shapes: ParsedShapeAsset[], diagram: ShapeGraph): string {
    // Loose connectors stay out. Item lines already say where their ends point.
    if (diagram.edges.length === 0) return "";
    const lines = ["```mermaid", `flowchart ${diagram.direction}`];
    // Node styles help only when nodes differ. One shared look tells nothing.
    const styles = diagram.nodes.map((index) => renderNodeStyle(shapes[index], index));
    const looks = new Set(styles.map((style) => (style ? style.replace(/^\s*style \S+ /, "") : "")));
    diagram.nodes.forEach((index, position) => {
        lines.push(renderShapeNode(shapes, index));
        const style = styles[position];
        if (style && looks.size > 1) lines.push(style);
    });
    for (const edge of diagram.edges) {
        const edgeText = escapeMermaidLabel(shapes[edge.connector]?.text || "");
        const arrow = (isDashed(shapes[edge.connector]) ? DASHED_EDGE_ARROWS : EDGE_ARROWS)[edge.arrow];
        const label = edgeText ? `|"${edgeText}"|` : "";
        lines.push(`    ${shapeNodeId(edge.from)} ${arrow}${label} ${shapeNodeId(edge.to)}`);
    }
    lines.push("```");
    return lines.join("\n");
}

export { renderMermaidFlowchart, shapeNodeId };
