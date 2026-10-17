import { describe, expect, it } from "vitest";
import * as textEncoding from "../../../../../src/outputs/markdown/text-encoding.ts";

describe("xlsx-context text encoding", () => {
    it("encodes UTF-8 without BOM by default", () => {
        const api = textEncoding;
        expect(Array.from(api.encodeText("A"))).toEqual([0x41]);
    });

    it("encodes UTF-16LE with BOM when requested", () => {
        const api = textEncoding;
        expect(Array.from(api.encodeText("A", { encoding: "utf-16le", bom: "on" }))).toEqual([0xff, 0xfe, 0x41, 0x00]);
    });

    it("encodes UTF-32BE without BOM", () => {
        const api = textEncoding;
        expect(Array.from(api.encodeText("A", { encoding: "utf-32be", bom: "off" }))).toEqual([0x00, 0x00, 0x00, 0x41]);
    });

    it("encodes shift_jis in the Node-backed runtime", () => {
        const api = textEncoding;
        expect(Array.from(api.encodeText("あA", { encoding: "shift_jis", bom: "off" }))).toEqual([0x82, 0xa0, 0x41]);
    });

    it("rejects BOM for shift_jis", () => {
        const api = textEncoding;
        expect(() => api.encodeText("A", { encoding: "shift_jis", bom: "on" })).toThrow(
            "BOM cannot be enabled for shift_jis."
        );
    });
});
