import { getNodeRequire } from "../../extraction/xml/runtime-env.js";
const utf8Encoder = new TextEncoder();

type MarkdownEncoding = "utf-8" | "shift_jis" | "utf-16le" | "utf-16be" | "utf-32le" | "utf-32be";
type MarkdownBomMode = "off" | "on";
type MarkdownEncodingOptions = {
    encoding?: MarkdownEncoding | string | null;
    bom?: MarkdownBomMode | string | null;
};

function normalizeEncoding(value?: MarkdownEncoding | string | null): MarkdownEncoding {
    const normalized = String(value || "utf-8").toLowerCase();
    if (
        normalized === "utf-8" ||
        normalized === "shift_jis" ||
        normalized === "utf-16le" ||
        normalized === "utf-16be" ||
        normalized === "utf-32le" ||
        normalized === "utf-32be"
    ) {
        return normalized;
    }
    throw new Error(`Unsupported encoding: ${String(value)}`);
}

function normalizeBom(value?: MarkdownBomMode | string | null): MarkdownBomMode {
    const normalized = String(value || "off").toLowerCase();
    if (normalized === "off" || normalized === "on") {
        return normalized;
    }
    throw new Error(`Unsupported BOM mode: ${String(value)}`);
}

function concatBytes(parts: Uint8Array[]): Uint8Array {
    const totalLength = parts.reduce((sum, part) => sum + part.length, 0);
    const result = new Uint8Array(totalLength);
    let offset = 0;
    for (const part of parts) {
        result.set(part, offset);
        offset += part.length;
    }
    return result;
}

function getBomBytes(encoding: MarkdownEncoding): Uint8Array {
    if (encoding === "utf-8") return new Uint8Array([0xef, 0xbb, 0xbf]);
    if (encoding === "utf-16le") return new Uint8Array([0xff, 0xfe]);
    if (encoding === "utf-16be") return new Uint8Array([0xfe, 0xff]);
    if (encoding === "utf-32le") return new Uint8Array([0xff, 0xfe, 0x00, 0x00]);
    if (encoding === "utf-32be") return new Uint8Array([0x00, 0x00, 0xfe, 0xff]);
    throw new Error(`Encoding does not support BOM: ${encoding}`);
}

function encodeUtf16(text: string, littleEndian: boolean): Uint8Array {
    const result = new Uint8Array(text.length * 2);
    for (let index = 0; index < text.length; index += 1) {
        const codeUnit = text.charCodeAt(index);
        const offset = index * 2;
        if (littleEndian) {
            result[offset] = codeUnit & 0xff;
            result[offset + 1] = codeUnit >>> 8;
        } else {
            result[offset] = codeUnit >>> 8;
            result[offset + 1] = codeUnit & 0xff;
        }
    }
    return result;
}

function encodeUtf32(text: string, littleEndian: boolean): Uint8Array {
    const codePoints: number[] = [];
    for (let index = 0; index < text.length; index += 1) {
        const first = text.charCodeAt(index);
        if (first >= 0xd800 && first <= 0xdbff && index + 1 < text.length) {
            const second = text.charCodeAt(index + 1);
            if (second >= 0xdc00 && second <= 0xdfff) {
                codePoints.push(((first - 0xd800) << 10) + (second - 0xdc00) + 0x10000);
                index += 1;
                continue;
            }
        }
        codePoints.push(first);
    }
    const result = new Uint8Array(codePoints.length * 4);
    codePoints.forEach((codePoint, index) => {
        const offset = index * 4;
        if (littleEndian) {
            result[offset] = codePoint & 0xff;
            result[offset + 1] = (codePoint >>> 8) & 0xff;
            result[offset + 2] = (codePoint >>> 16) & 0xff;
            result[offset + 3] = (codePoint >>> 24) & 0xff;
        } else {
            result[offset] = (codePoint >>> 24) & 0xff;
            result[offset + 1] = (codePoint >>> 16) & 0xff;
            result[offset + 2] = (codePoint >>> 8) & 0xff;
            result[offset + 3] = codePoint & 0xff;
        }
    });
    return result;
}

type IconvLite = { encode: (value: string, encodingName: string) => Uint8Array };

/** Loads the optional iconv-lite peer. Returns null outside Node.js or when it is not installed. */
function loadIconvLite(): IconvLite | null {
    const nodeRequire = getNodeRequire();
    if (!nodeRequire) return null;
    try {
        return nodeRequire("iconv-lite") as IconvLite;
    } catch (_error) {
        return null;
    }
}

/** Encodes text with optional BOM. Rejects BOM for shift_jis. Shift_jis needs iconv-lite in Node.js. */
function encodeText(text: string, options: MarkdownEncodingOptions = {}): Uint8Array {
    const encoding = normalizeEncoding(options.encoding);
    const bom = normalizeBom(options.bom);
    if (encoding === "shift_jis") {
        if (bom === "on") {
            throw new Error("BOM cannot be enabled for shift_jis.");
        }
        const iconvLite = loadIconvLite();
        if (!iconvLite) {
            throw new Error(
                "Shift_JIS encoding needs the optional peer dependency iconv-lite in Node.js. Run `npm install iconv-lite`."
            );
        }
        return Uint8Array.from(iconvLite.encode(text, "shift_jis"));
    }
    const body =
        encoding === "utf-8"
            ? utf8Encoder.encode(text)
            : encoding === "utf-16le"
              ? encodeUtf16(text, true)
              : encoding === "utf-16be"
                ? encodeUtf16(text, false)
                : encoding === "utf-32le"
                  ? encodeUtf32(text, true)
                  : encodeUtf32(text, false);
    if (bom === "off") {
        return body;
    }
    return concatBytes([getBomBytes(encoding), body]);
}

/** Reports whether an encoding can run here. Shift_jis needs iconv-lite. */
function isEncodingAvailable(value?: MarkdownEncoding | string | null): boolean {
    const encoding = normalizeEncoding(value);
    if (encoding === "shift_jis") {
        return !!loadIconvLite();
    }
    return true;
}

function createTextMimeType(options: MarkdownEncodingOptions = {}): string {
    return `text/markdown;charset=${normalizeEncoding(options.encoding)}`;
}

/** Encodes Markdown text to bytes with BOM support. */
export { normalizeEncoding, normalizeBom, getBomBytes, isEncodingAvailable, encodeText, createTextMimeType };
