import path from "node:path";
import { fileURLToPath } from "node:url";

// Single source for repository paths, so tests never count "../" segments.
export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
export const SRC_ROOT = path.join(ROOT, "src");
export const FIXTURE_ROOT = path.join(ROOT, "tests", "fixtures");
export const BIN = path.join(ROOT, "bin", "xlsx-context.mjs");

export function fixture(...parts) {
    return path.join(FIXTURE_ROOT, ...parts);
}
