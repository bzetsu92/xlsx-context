import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { build } from "esbuild";

/*
 * Builds the npm package:
 * - dist/lib: ESM bundles from src/public. Code splitting puts the shared core into
 *   one chunk, so index.js and node.js load the same module instances.
 * - dist/types: declarations, pruned to the files reachable from the public entries
 *   so internal module types do not leak into the package.
 */

const ROOT = process.cwd();
const TYPES_DIR = path.resolve(ROOT, "dist/types");
const TYPE_ENTRIES = ["public/index.d.ts", "public/node.d.ts", "public/worker.d.ts"];

await build({
    entryPoints: {
        index: path.resolve(ROOT, "src/public/index.ts"),
        node: path.resolve(ROOT, "src/public/node.ts"),
        worker: path.resolve(ROOT, "src/public/worker.ts"),
    },
    outdir: path.resolve(ROOT, "dist/lib"),
    bundle: true,
    splitting: true,
    format: "esm",
    platform: "neutral",
    target: "es2020",
    external: ["@xmldom/xmldom", "js-tiktoken", "node:*"],
    chunkNames: "chunks/[name]-[hash]",
    legalComments: "inline",
    logLevel: "warning",
});

execFileSync(process.execPath, [path.resolve(ROOT, "node_modules/typescript/bin/tsc"), "-p", "tsconfig.lib.json"], {
    stdio: "inherit",
});
pruneDeclarations();

console.log("[build:lib] generated dist/lib and dist/types");

function pruneDeclarations() {
    const reachable = new Set();
    const queue = TYPE_ENTRIES.map((entry) => path.join(TYPES_DIR, entry));
    while (queue.length > 0) {
        const file = queue.pop();
        if (reachable.has(file)) continue;
        reachable.add(file);
        const source = fs.readFileSync(file, "utf8");
        for (const match of source.matchAll(/(?:from|import)\s*\(?\s*"(\.{1,2}\/[^"]+)"/g)) {
            queue.push(path.join(path.dirname(file), match[1].replace(/\.js$/, ".d.ts")));
        }
    }
    for (const file of listFiles(TYPES_DIR)) {
        if (!reachable.has(file)) fs.rmSync(file);
    }
    removeEmptyDirs(TYPES_DIR);
}

function listFiles(dir) {
    return fs
        .readdirSync(dir, { withFileTypes: true })
        .flatMap((entry) =>
            entry.isDirectory() ? listFiles(path.join(dir, entry.name)) : [path.join(dir, entry.name)]
        );
}

function removeEmptyDirs(dir) {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        if (entry.isDirectory()) removeEmptyDirs(path.join(dir, entry.name));
    }
    if (dir !== TYPES_DIR && fs.readdirSync(dir).length === 0) fs.rmdirSync(dir);
}
