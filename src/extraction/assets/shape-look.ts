import type { ShapeLook } from "../../model/worksheet.js";
import { getDirectChildByLocalName, getElementsByLocalName, xmlToDocument } from "../xml/xml-utils.js";

type ThemeColors = Record<string, string>;

// Spreadsheets map text and background names to the dark and light theme slots.
const SCHEME_ALIASES: Record<string, string> = { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" };

// The Office 2007 theme. Used when the workbook has no theme part.
const DEFAULT_THEME: ThemeColors = {
    dk1: "000000",
    lt1: "FFFFFF",
    dk2: "1F497D",
    lt2: "EEECE1",
    accent1: "4F81BD",
    accent2: "C0504D",
    accent3: "9BBB59",
    accent4: "8064A2",
    accent5: "4BACC6",
    accent6: "F79646",
    hlink: "0000FF",
    folHlink: "800080",
};

// A subset of the DrawingML preset colors. Unknown names resolve to null.
const PRESET_COLORS: Record<string, string> = {
    black: "000000",
    white: "FFFFFF",
    red: "FF0000",
    green: "008000",
    lime: "00FF00",
    blue: "0000FF",
    yellow: "FFFF00",
    cyan: "00FFFF",
    magenta: "FF00FF",
    orange: "FFA500",
    gray: "808080",
    grey: "808080",
    darkGray: "A9A9A9",
    lightGray: "D3D3D3",
    navy: "000080",
    purple: "800080",
};

const COLOR_ELEMENTS = new Set(["srgbClr", "schemeClr", "sysClr", "prstClr", "scrgbClr", "hslClr"]);

const themeCache = new WeakMap<Map<string, Uint8Array>, ThemeColors>();

/** Reads the color scheme of the first theme. Falls back to the Office theme. */
function readThemeColors(files: Map<string, Uint8Array>, decodeXmlText: (bytes: Uint8Array) => string): ThemeColors {
    const cached = themeCache.get(files);
    if (cached) return cached;
    const colors = { ...DEFAULT_THEME };
    const themePath = [...files.keys()].filter((path) => /^xl\/theme\/theme\d+\.xml$/.test(path)).sort()[0];
    const bytes = themePath ? files.get(themePath) : undefined;
    const scheme = bytes ? getElementsByLocalName(xmlToDocument(decodeXmlText(bytes)), "clrScheme")[0] : undefined;
    for (const slot of Array.from(scheme?.childNodes || [])) {
        const element = slot as Element;
        if (element.nodeType !== 1) continue;
        const color = resolveColorElement(firstColorElement(element), DEFAULT_THEME);
        if (color) colors[element.localName] = color.slice(1);
    }
    themeCache.set(files, colors);
    return colors;
}

function firstColorElement(parent: Element | null | undefined): Element | null {
    for (const child of Array.from(parent?.childNodes || [])) {
        const element = child as Element;
        if (element.nodeType === 1 && COLOR_ELEMENTS.has(element.localName)) return element;
    }
    return null;
}

function hexToRgb(hex: string): [number, number, number] | null {
    const match = /^#?([0-9a-f]{6})$/i.exec(hex.trim());
    if (!match) return null;
    const value = parseInt(match[1] as string, 16);
    return [(value >> 16) & 255, (value >> 8) & 255, value & 255];
}

function rgbToHex([r, g, b]: number[]): string {
    const part = (value: number) =>
        Math.round(Math.min(255, Math.max(0, value)))
            .toString(16)
            .padStart(2, "0");
    return `#${part(r as number)}${part(g as number)}${part(b as number)}`.toUpperCase();
}

function rgbToHsl([r, g, b]: number[]): [number, number, number] {
    const [rn, gn, bn] = [(r as number) / 255, (g as number) / 255, (b as number) / 255];
    const max = Math.max(rn, gn, bn);
    const min = Math.min(rn, gn, bn);
    const l = (max + min) / 2;
    if (max === min) return [0, 0, l];
    const d = max - min;
    const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    const h = max === rn ? (gn - bn) / d + (gn < bn ? 6 : 0) : max === gn ? (bn - rn) / d + 2 : (rn - gn) / d + 4;
    return [h * 60, s, l];
}

function hslToRgb([h, s, l]: number[]): [number, number, number] {
    const hue = (h as number) / 360;
    if (s === 0) return [(l as number) * 255, (l as number) * 255, (l as number) * 255];
    const q =
        (l as number) < 0.5
            ? (l as number) * (1 + (s as number))
            : (l as number) + (s as number) - (l as number) * (s as number);
    const p = 2 * (l as number) - q;
    const channel = (t: number) => {
        const x = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
        if (x < 1 / 6) return p + (q - p) * 6 * x;
        if (x < 1 / 2) return q;
        if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
        return p;
    };
    return [channel(hue + 1 / 3) * 255, channel(hue) * 255, channel(hue - 1 / 3) * 255];
}

/** Applies the color transforms that change the visible color. Alpha and others stay out. */
function applyTransforms(rgb: [number, number, number], element: Element): [number, number, number] {
    let current = rgb;
    for (const child of Array.from(element.childNodes)) {
        const transform = child as Element;
        if (transform.nodeType !== 1) continue;
        const value = Number(transform.getAttribute("val")) / 100000;
        if (!Number.isFinite(value)) continue;
        if (transform.localName === "lumMod" || transform.localName === "lumOff") {
            const [h, s, l] = rgbToHsl(current);
            const next = transform.localName === "lumMod" ? l * value : l + value;
            current = hslToRgb([h, s, Math.min(1, Math.max(0, next))]);
        } else if (transform.localName === "shade") {
            current = current.map((channel) => channel * value) as [number, number, number];
        } else if (transform.localName === "tint") {
            current = current.map((channel) => channel + (255 - channel) * (1 - value)) as [number, number, number];
        }
    }
    return current;
}

function resolveColorElement(element: Element | null, theme: ThemeColors): string | null {
    if (!element) return null;
    let base: string | null = null;
    const name = element.localName;
    if (name === "srgbClr") base = element.getAttribute("val");
    else if (name === "sysClr")
        base = element.getAttribute("lastClr") || (element.getAttribute("val") === "window" ? "FFFFFF" : "000000");
    else if (name === "prstClr") base = PRESET_COLORS[element.getAttribute("val") || ""] || null;
    else if (name === "schemeClr") {
        const key = element.getAttribute("val") || "";
        base = theme[SCHEME_ALIASES[key] || key] || null;
    } else if (name === "scrgbClr") {
        const percent = (attr: string) => (Number(element.getAttribute(attr)) / 100000) * 255;
        return rgbToHex(applyTransforms([percent("r"), percent("g"), percent("b")], element));
    } else if (name === "hslClr") {
        const h = Number(element.getAttribute("hue")) / 60000;
        const s = Number(element.getAttribute("sat")) / 100000;
        const l = Number(element.getAttribute("lum")) / 100000;
        return rgbToHex(applyTransforms(hslToRgb([h, s, l]), element));
    }
    const rgb = base ? hexToRgb(base) : null;
    return rgb ? rgbToHex(applyTransforms(rgb, element)) : null;
}

/** Reads one fill element. Returns undefined when the parent has no fill element. */
function readFillElement(parent: Element | null, theme: ThemeColors): string | null | undefined {
    for (const child of Array.from(parent?.childNodes || [])) {
        const element = child as Element;
        if (element.nodeType !== 1) continue;
        switch (element.localName) {
            case "noFill":
                return "none";
            case "solidFill":
                return resolveColorElement(firstColorElement(element), theme);
            case "gradFill": {
                // The first stop stands for the gradient.
                const stop = getElementsByLocalName(element, "gs")[0];
                return resolveColorElement(firstColorElement(stop), theme);
            }
            case "pattFill":
                return resolveColorElement(firstColorElement(getDirectChildByLocalName(element, "fgClr")), theme);
            case "blipFill":
                return "picture";
            case "grpFill":
                return null;
        }
    }
    return undefined;
}

/** Reads the theme reference of a style. Index 0 means no fill or no line. */
function readStyleRef(style: Element | null, refName: string, theme: ThemeColors): string | null {
    const ref = getDirectChildByLocalName(style, refName);
    if (!ref) return null;
    if (ref.getAttribute("idx") === "0") return "none";
    return resolveColorElement(firstColorElement(ref), theme);
}

/**
 * Reads the look of one shape.
 * The shape properties win. The style element of the shape fills the gaps.
 */
function readShapeLook(shapeNode: Element | null, theme: ThemeColors): ShapeLook {
    const spPr = getDirectChildByLocalName(shapeNode, "spPr");
    const style = getDirectChildByLocalName(shapeNode, "style");
    const fill = readFillElement(spPr, theme);
    const ln = getDirectChildByLocalName(spPr, "ln");
    const lineFill = readFillElement(ln, theme);
    const prstDash = getDirectChildByLocalName(ln, "prstDash")?.getAttribute("val") || null;
    const width = Number(ln?.getAttribute("w"));
    return {
        fill: fill !== undefined ? fill : readStyleRef(style, "fillRef", theme),
        line: lineFill !== undefined ? lineFill : readStyleRef(style, "lnRef", theme),
        dash: getDirectChildByLocalName(ln, "custDash") ? "custom" : prstDash,
        lineWidthEmu: ln?.hasAttribute("w") && Number.isFinite(width) ? width : null,
    };
}

/** Names a color in plain words, for example "light blue". Readers match words better than hex codes. */
function colorName(hex: string | null): string {
    if (!hex) return "default";
    if (!hex.startsWith("#")) return hex;
    const rgb = hexToRgb(hex);
    if (!rgb) return hex;
    const [h, s, l] = rgbToHsl(rgb);
    if (s < 0.15 || l > 0.97 || l < 0.06) {
        if (l > 0.93) return "white";
        if (l > 0.7) return "light gray";
        if (l > 0.35) return "gray";
        if (l > 0.15) return "dark gray";
        return "black";
    }
    const hues: [number, string][] = [
        [15, "red"],
        [45, "orange"],
        [70, "yellow"],
        [165, "green"],
        [195, "cyan"],
        [255, "blue"],
        [290, "purple"],
        [345, "pink"],
        [360, "red"],
    ];
    const hue = (hues.find(([limit]) => h < limit) as [number, string])[1];
    if (l > 0.75) return `light ${hue}`;
    if (l < 0.3) return `dark ${hue}`;
    return hue;
}

export { colorName, readShapeLook, readThemeColors };
export type { ThemeColors };
