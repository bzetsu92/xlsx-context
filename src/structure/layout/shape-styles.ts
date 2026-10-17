/**
 * Groups shapes that look the same.
 * Writers often give one look one meaning, for example red dashed boxes for changes.
 * The groups let readers find such conventions without guessing them from colors.
 */
import type { ParsedShapeAsset, ShapeLook } from "../../model/index.js";

/** Shapes with one look. `id` is S1, S2, and so on, in first-seen order. */
export type ShapeStyle = {
    id: string;
    /** Presets of the member shapes, in first-seen order. */
    presets: string[];
    look: ShapeLook;
    /** True when the line is clearly heavier than normal lines. */
    thick: boolean;
    /** True for connectors and line shapes. They have no fill. */
    isLine: boolean;
    shapeIndexes: number[];
};

// Lines above 2 pt stand out. Thinner differences, such as 0.75 pt against 1.25 pt, carry no meaning.
const THICK_LINE_EMU = 2 * 12700;

/** Reads the preset name from kind labels such as `Shape (flowChartProcess)`. */
function presetOf(shape: ParsedShapeAsset): string {
    if (shape.kind === "Rectangle") return "rect";
    if (shape.kind === "Text Box") return "textBox";
    if (shape.kind === "Straight Arrow Connector") return "straightConnector1";
    return /\(([A-Za-z0-9]+)\)$/.exec(shape.kind || "")?.[1] || shape.kind || "shape";
}

/**
 * Groups shapes by look: fill, line color, dash, and thick lines.
 * The shape kind is not part of the look, because item lines already name it.
 * Shapes without style data stay out, so they never get a style.
 */
function groupShapeStyles(shapes: ParsedShapeAsset[]): ShapeStyle[] {
    const styles = new Map<string, ShapeStyle>();
    shapes.forEach((shape, index) => {
        const look = shape?.look;
        if (!look || (look.fill === null && look.line === null && look.dash === null)) return;
        const preset = presetOf(shape);
        const isLine = !!shape.connection;
        const thick = (look.lineWidthEmu ?? 0) > THICK_LINE_EMU;
        const key = JSON.stringify([isLine, isLine ? null : look.fill, look.line, look.dash || "solid", thick]);
        const style = styles.get(key);
        if (style) {
            style.shapeIndexes.push(index);
            if (!style.presets.includes(preset)) style.presets.push(preset);
        } else {
            styles.set(key, {
                id: `S${styles.size + 1}`,
                presets: [preset],
                look,
                thick,
                isLine,
                shapeIndexes: [index],
            });
        }
    });
    return [...styles.values()];
}

export { groupShapeStyles, presetOf };
