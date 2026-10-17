import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";

import { describe, expect, it, vi } from "vitest";
import * as core from "../../../src/core.ts";
import * as liteXmlReader from "../../../src/extraction/xml/lite-xml-reader.ts";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

describe("readContainer() fast path", () => {
    function listFixtures(dir = FIXTURE_ROOT) {
        const out = [];
        for (const entry of readdirSync(dir, { withFileTypes: true })) {
            const full = path.join(dir, entry.name);
            if (entry.isDirectory()) out.push(...listFixtures(full));
            else if (entry.name.endsWith(".xlsx")) out.push(full);
        }
        return out.sort();
    }

    // Parses every fixture twice, so it needs more than the default 5 s on a busy machine.
    it("parses every fixture workbook identically to the DOM path", async () => {
        const api = core;
        let fastSheets = 0;
        for (const fixturePath of listFixtures()) {
            const bytes = readFileSync(fixturePath);
            const buffer = () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
            const name = path.basename(fixturePath);
            const probe = vi.spyOn(liteXmlReader, "readContainer");
            const fast = await api.parseWorkbook(buffer(), name);
            fastSheets += probe.mock.results.filter((entry) => entry.value !== null).length;
            probe.mockRestore();
            const spy = vi.spyOn(liteXmlReader, "readContainer").mockReturnValue(null);
            const dom = await api.parseWorkbook(buffer(), name);
            spy.mockRestore();
            expect(fast.sheets, name).toEqual(dom.sheets);
        }
        // Every fixture sheet should take the fast path; a drop means the envelope regressed.
        expect(fastSheets).toBeGreaterThanOrEqual(listFixtures().length);
    }, 30000);
});
