import { describe, expect, it } from "vitest";
import * as zipIo from "../../../../../src/extraction/zip/zip-io.ts";

// Builds a one-entry deflated ZIP. declaredSize may lie, as in a zip bomb.
function createDeflatedZip(name, compressed, declaredSize) {
    const nameBytes = new TextEncoder().encode(name);
    const local = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(local.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(8, 8, true);
    localView.setUint32(18, compressed.length, true);
    localView.setUint32(22, declaredSize, true);
    localView.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    const central = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(central.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(10, 8, true);
    centralView.setUint32(20, compressed.length, true);
    centralView.setUint32(24, declaredSize, true);
    centralView.setUint16(28, nameBytes.length, true);
    centralView.setUint32(42, 0, true);
    central.set(nameBytes, 46);
    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    eocdView.setUint32(0, 0x06054b50, true);
    eocdView.setUint16(8, 1, true);
    eocdView.setUint16(10, 1, true);
    eocdView.setUint32(12, central.length, true);
    eocdView.setUint32(16, local.length + compressed.length, true);
    const zip = new Uint8Array(local.length + compressed.length + central.length + eocd.length);
    zip.set(local, 0);
    zip.set(compressed, local.length);
    zip.set(central, local.length + compressed.length);
    zip.set(eocd, local.length + compressed.length + central.length);
    return zip.buffer;
}

describe("xlsx-context zip io", () => {
    it("round-trips stored zip entries", async () => {
        const api = zipIo;
        const encoder = new TextEncoder();
        const zipBytes = api.createStoredZip([
            { name: "output/test.md", data: encoder.encode("# Test\n") },
            { name: "output/assets/icon.txt", data: encoder.encode("asset") },
        ]);

        const extracted = await api.unzipEntries(
            zipBytes.buffer.slice(zipBytes.byteOffset, zipBytes.byteOffset + zipBytes.byteLength)
        );

        expect(Array.from(extracted.keys()).sort()).toEqual(["output/assets/icon.txt", "output/test.md"]);
        expect(new TextDecoder().decode(extracted.get("output/test.md"))).toBe("# Test\n");
        expect(new TextDecoder().decode(extracted.get("output/assets/icon.txt"))).toBe("asset");
    });

    it("supports empty file payloads", async () => {
        const api = zipIo;
        const zipBytes = api.createStoredZip([{ name: "empty.txt", data: new Uint8Array([]) }]);

        const extracted = await api.unzipEntries(
            zipBytes.buffer.slice(zipBytes.byteOffset, zipBytes.byteOffset + zipBytes.byteLength)
        );

        expect(extracted.has("empty.txt")).toBe(true);
        expect(extracted.get("empty.txt")).toEqual(new Uint8Array([]));
    });

    it("writes a fixed reproducible ZIP entry timestamp", () => {
        const api = zipIo;
        const encoder = new TextEncoder();
        const zipBytes = api.createStoredZip([{ name: "output/test.md", data: encoder.encode("# Test\n") }]);
        const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength);

        expect(view.getUint32(0, true)).toBe(0x04034b50);
        expect(view.getUint16(10, true)).toBe(api.fixedZipEntryTimestamp.dosTime);
        expect(view.getUint16(12, true)).toBe(api.fixedZipEntryTimestamp.dosDate);

        const localNameLength = view.getUint16(26, true);
        const centralOffset = 30 + localNameLength + encoder.encode("# Test\n").length;
        expect(view.getUint32(centralOffset, true)).toBe(0x02014b50);
        expect(view.getUint16(centralOffset + 12, true)).toBe(api.fixedZipEntryTimestamp.dosTime);
        expect(view.getUint16(centralOffset + 14, true)).toBe(api.fixedZipEntryTimestamp.dosDate);
    });

    it("does not mark ASCII-only file names with the UTF-8 flag", () => {
        const api = zipIo;
        const encoder = new TextEncoder();
        const zipBytes = api.createStoredZip([{ name: "output/test.md", data: encoder.encode("# Test\n") }]);
        const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength);

        expect(view.getUint16(6, true)).toBe(0);

        const localNameLength = view.getUint16(26, true);
        const centralOffset = 30 + localNameLength + encoder.encode("# Test\n").length;
        expect(view.getUint16(centralOffset + 8, true)).toBe(0);
    });

    it("marks UTF-8 file names so non-ASCII entries unzip correctly", async () => {
        const api = zipIo;
        const encoder = new TextEncoder();
        const zipBytes = api.createStoredZip([{ name: "output/日本語.md", data: encoder.encode("# 日本語\n") }]);
        const view = new DataView(zipBytes.buffer, zipBytes.byteOffset, zipBytes.byteLength);

        expect(view.getUint16(6, true)).toBe(0x0800);

        const localNameLength = view.getUint16(26, true);
        const centralOffset = 30 + localNameLength + encoder.encode("# 日本語\n").length;
        expect(view.getUint16(centralOffset + 8, true)).toBe(0x0800);

        const extracted = await api.unzipEntries(
            zipBytes.buffer.slice(zipBytes.byteOffset, zipBytes.byteOffset + zipBytes.byteLength)
        );
        expect(Array.from(extracted.keys())).toEqual(["output/日本語.md"]);
        expect(new TextDecoder().decode(extracted.get("output/日本語.md"))).toBe("# 日本語\n");
    });

    it("normalizes entry paths to OPC part paths", async () => {
        const api = zipIo;
        const zipBytes = api.createStoredZip([{ name: "/xl/./worksheets/../workbook.xml", data: new Uint8Array([1]) }]);

        const extracted = await api.unzipEntries(
            zipBytes.buffer.slice(zipBytes.byteOffset, zipBytes.byteOffset + zipBytes.byteLength)
        );

        expect(Array.from(extracted.keys())).toEqual(["xl/workbook.xml"]);
    });

    it("rejects entry paths that escape the package root", async () => {
        const api = zipIo;
        const zipBytes = api.createStoredZip([{ name: "../evil.xml", data: new Uint8Array([1]) }]);

        await expect(
            api.unzipEntries(zipBytes.buffer.slice(zipBytes.byteOffset, zipBytes.byteOffset + zipBytes.byteLength))
        ).rejects.toThrow("OPC part path escapes package root: ../evil.xml");
    });

    it("throws for invalid zip input", async () => {
        const api = zipIo;

        await expect(api.unzipEntries(new Uint8Array([1, 2, 3, 4]).buffer)).rejects.toThrow(
            "ZIP end-of-central-directory record was not found."
        );
    });

    it("falls back to node zlib inflateRawSync when deflate-raw streams are unavailable", async () => {
        const api = zipIo;
        const encoder = new TextEncoder();
        const zlib = globalThis.__xlsxContextNodeRequire("node:zlib");
        const fileName = "output/test.txt";
        const fileNameBytes = encoder.encode(fileName);
        const original = encoder.encode("fallback");
        const compressed = Uint8Array.from(zlib.deflateRawSync(original));
        const crc = 0xa87e4381;

        const localHeader = new Uint8Array(30 + fileNameBytes.length);
        const localView = new DataView(localHeader.buffer);
        localView.setUint32(0, 0x04034b50, true);
        localView.setUint16(4, 20, true);
        localView.setUint16(6, 0, true);
        localView.setUint16(8, 8, true);
        localView.setUint16(10, 0, true);
        localView.setUint16(12, 0, true);
        localView.setUint32(14, crc, true);
        localView.setUint32(18, compressed.length, true);
        localView.setUint32(22, original.length, true);
        localView.setUint16(26, fileNameBytes.length, true);
        localView.setUint16(28, 0, true);
        localHeader.set(fileNameBytes, 30);

        const centralHeader = new Uint8Array(46 + fileNameBytes.length);
        const centralView = new DataView(centralHeader.buffer);
        centralView.setUint32(0, 0x02014b50, true);
        centralView.setUint16(4, 20, true);
        centralView.setUint16(6, 20, true);
        centralView.setUint16(8, 0, true);
        centralView.setUint16(10, 8, true);
        centralView.setUint16(12, 0, true);
        centralView.setUint16(14, 0, true);
        centralView.setUint32(16, crc, true);
        centralView.setUint32(20, compressed.length, true);
        centralView.setUint32(24, original.length, true);
        centralView.setUint16(28, fileNameBytes.length, true);
        centralView.setUint16(30, 0, true);
        centralView.setUint16(32, 0, true);
        centralView.setUint16(34, 0, true);
        centralView.setUint16(36, 0, true);
        centralView.setUint32(38, 0, true);
        centralView.setUint32(42, 0, true);
        centralHeader.set(fileNameBytes, 46);

        const eocd = new Uint8Array(22);
        const eocdView = new DataView(eocd.buffer);
        eocdView.setUint32(0, 0x06054b50, true);
        eocdView.setUint16(8, 1, true);
        eocdView.setUint16(10, 1, true);
        eocdView.setUint32(12, centralHeader.length, true);
        eocdView.setUint32(16, localHeader.length + compressed.length, true);

        const zipBytes = new Uint8Array(localHeader.length + compressed.length + centralHeader.length + eocd.length);
        zipBytes.set(localHeader, 0);
        zipBytes.set(compressed, localHeader.length);
        zipBytes.set(centralHeader, localHeader.length + compressed.length);
        zipBytes.set(eocd, localHeader.length + compressed.length + centralHeader.length);

        const originalDecompressionStream = globalThis.DecompressionStream;
        globalThis.DecompressionStream = class UnsupportedDecompressionStream {
            constructor(format) {
                throw new Error(`unsupported: ${format}`);
            }
        };

        try {
            const extracted = await api.unzipEntries(
                zipBytes.buffer.slice(zipBytes.byteOffset, zipBytes.byteOffset + zipBytes.byteLength)
            );
            expect(new TextDecoder().decode(extracted.get(fileName))).toBe("fallback");
        } finally {
            globalThis.DecompressionStream = originalDecompressionStream;
        }
    });
    it("stops inflating an entry at its declared size (zip bomb)", async () => {
        const zlib = globalThis.__xlsxContextNodeRequire("node:zlib");
        const compressed = Uint8Array.from(zlib.deflateRawSync(new Uint8Array(64 * 1024 * 1024)));
        const before = process.memoryUsage().rss;

        await expect(zipIo.unzipEntries(createDeflatedZip("xl/media/pad.bin", compressed, 1))).rejects.toThrow(
            "xl/media/pad.bin: Deflated ZIP entry size mismatch."
        );
        expect(process.memoryUsage().rss - before).toBeLessThan(32 * 1024 * 1024);
    });

    it("stops a deflate-raw stream at the declared size when node zlib is unavailable", async () => {
        const zlib = globalThis.__xlsxContextNodeRequire("node:zlib");
        const size = 128 * 1024 * 1024;
        const compressed = Uint8Array.from(zlib.deflateRawSync(new Uint8Array(size)));
        const originalRequire = globalThis.__xlsxContextNodeRequire;
        globalThis.__xlsxContextNodeRequire = undefined;

        try {
            const before = process.memoryUsage().rss;
            await expect(zipIo.unzipEntries(createDeflatedZip("a.bin", compressed, 10))).rejects.toThrow(
                "a.bin: Deflated ZIP entry size mismatch."
            );
            expect(process.memoryUsage().rss - before).toBeLessThan(32 * 1024 * 1024);
            const small = Uint8Array.from(zlib.deflateRawSync(new TextEncoder().encode("exact")));
            const exact = await zipIo.unzipEntries(createDeflatedZip("b.txt", small, 5));
            expect(new TextDecoder().decode(exact.get("b.txt"))).toBe("exact");
        } finally {
            globalThis.__xlsxContextNodeRequire = originalRequire;
        }
    });
});
