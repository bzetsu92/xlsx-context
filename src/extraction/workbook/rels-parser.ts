type RelsParserDeps = {
    xmlToDocument: (xmlText: string) => Document;
    decodeXmlText: (bytes: Uint8Array) => string;
};

type RelationshipEntry = {
    target: string;
    targetMode: string;
    type: string;
};

function createRelsParser(deps: RelsParserDeps) {
    /** Keeps external targets as-is. Resolves internal targets against the source path. */
    function normalizeRelationshipTarget(baseFilePath: string, targetPath: string, targetMode = ""): string {
        if ((targetMode || "").toLowerCase() === "external") {
            return targetPath;
        }
        return normalizeZipPath(baseFilePath, targetPath);
    }

    /** Resolves a relative target against its source directory. A leading slash starts at the package root. */
    function normalizeZipPath(baseFilePath: string, targetPath: string): string {
        const baseDirParts = baseFilePath.split("/").slice(0, -1);
        const inputParts = targetPath.split("/");
        const parts = targetPath.startsWith("/") ? [] : baseDirParts;
        for (const part of inputParts) {
            if (!part || part === ".") continue;
            if (part === "..") {
                parts.pop();
            } else {
                parts.push(part);
            }
        }
        return parts.join("/");
    }

    /** Reads one .rels file into entries keyed by relationship id. Returns an empty map when the file is missing. */
    function parseRelationshipEntries(
        files: Map<string, Uint8Array>,
        relsPath: string,
        sourcePath: string
    ): Map<string, RelationshipEntry> {
        const relBytes = files.get(relsPath);
        const relations = new Map<string, RelationshipEntry>();
        if (!relBytes) {
            return relations;
        }
        const doc = deps.xmlToDocument(deps.decodeXmlText(relBytes));
        const nodes = Array.from(doc.getElementsByTagName("Relationship"));
        for (const node of nodes) {
            const id = node.getAttribute("Id") || "";
            const target = node.getAttribute("Target") || "";
            if (!id || !target) continue;
            const targetMode = node.getAttribute("TargetMode") || "";
            relations.set(id, {
                target: normalizeRelationshipTarget(sourcePath, target, targetMode),
                targetMode,
                type: node.getAttribute("Type") || "",
            });
        }
        return relations;
    }

    /** Reduces relationship entries to id-to-target maps. */
    function parseRelationships(
        files: Map<string, Uint8Array>,
        relsPath: string,
        sourcePath: string
    ): Map<string, string> {
        const relations = new Map<string, string>();
        const entries = parseRelationshipEntries(files, relsPath, sourcePath);
        for (const [id, entry] of entries.entries()) {
            relations.set(id, entry.target);
        }
        return relations;
    }

    /** Builds the .rels path for a source part, following the OPC `_rels` convention. */
    function buildRelsPath(sourcePath: string): string {
        const parts = sourcePath.split("/");
        const fileName = parts.pop() || "";
        const dir = parts.join("/");
        return `${dir}/_rels/${fileName}.rels`;
    }

    return {
        normalizeRelationshipTarget,
        normalizeZipPath,
        parseRelationshipEntries,
        parseRelationships,
        buildRelsPath,
    };
}

/** Reads OPC relationship files into id-to-target maps. */
export { createRelsParser };
