import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";

import { describe, expect, it } from "vitest";
import { unzipEntries } from "../../../src/extraction/zip/zip-io.ts";
import { FIXTURE_ROOT } from "../../helpers/paths.js";

// Excel writes personal data into every saved file. Fixtures are public test data, so they must hold none.
const ABS_PATH = /x15ac:absPath|absPath url=/u;
const CORE_PROFILE = /<dc:creator>|<cp:lastModifiedBy>|<dcterms:created\b|<dcterms:modified\b/u;
const APP_PROFILE = /<Application>|<AppVersion>/u;
const SENSITIVE_PART =
    /^(xl\/comments\d*\.xml|xl\/threadedComments\/|xl\/persons\/|docProps\/custom\.xml|xl\/externalLinks\/|xl\/connections\.xml|xl\/embeddings\/|xl\/vbaProject\.bin|xl\/printerSettings\/)/u;

function listFixtures(dir = FIXTURE_ROOT) {
    const out = [];
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) out.push(...listFixtures(full));
        else if (entry.name.endsWith(".xlsx")) out.push(full);
    }
    return out.sort();
}

async function readFixtureParts(fixturePath) {
    const bytes = readFileSync(fixturePath);
    const entries = await unzipEntries(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength));
    const text = (name) => (entries.has(name) ? new TextDecoder().decode(entries.get(name)) : "");
    return {
        name: path.relative(FIXTURE_ROOT, fixturePath),
        partNames: [...entries.keys()],
        workbookXml: text("xl/workbook.xml"),
        coreXml: text("docProps/core.xml"),
        appXml: text("docProps/app.xml"),
    };
}

describe("fixture hygiene", () => {
    it("keeps local paths, author profile, and sensitive parts out of every fixture", async () => {
        const fixtures = await Promise.all(listFixtures().map(readFixtureParts));
        const offenders = fixtures.flatMap((fixture) => [
            ...(ABS_PATH.test(fixture.workbookXml) ? [`${fixture.name}: local path in xl/workbook.xml`] : []),
            ...(CORE_PROFILE.test(fixture.coreXml) ? [`${fixture.name}: author or dates in docProps/core.xml`] : []),
            ...(APP_PROFILE.test(fixture.appXml) ? [`${fixture.name}: application in docProps/app.xml`] : []),
            ...fixture.partNames
                .filter((partName) => SENSITIVE_PART.test(partName))
                .map((partName) => `${fixture.name}: sensitive part ${partName}`),
        ]);

        expect(fixtures.length).toBeGreaterThan(0);
        expect(offenders).toEqual([]);
    });
});
