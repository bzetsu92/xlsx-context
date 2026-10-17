import type { ParsedImageAsset } from "../../model/index.js";
import type { ParseAssetDependencies } from "./sheet-assets.js";
import { collectAnchors, createSafeSheetAssetDir } from "./sheet-assets.js";
import { normalizeShapeGeometry } from "./drawing-geometry.js";
import { createUniformGrid } from "../workbook/sheet-grid.js";

/** Reads the file extension in lowercase. Falls back to `bin`. */
function getImageExtension(mediaPath: string): string {
    const match = mediaPath.match(/\.([a-z0-9]+)$/i);
    return match ? match[1].toLowerCase() : "bin";
}

/** Collects anchored images with numbered file names. Skips anchors without resolvable media. */
function parseDrawingImages(
    files: Map<string, Uint8Array>,
    sheetName: string,
    sheetPath: string,
    deps: ParseAssetDependencies
): ParsedImageAsset[] {
    const sheetRels = deps.parseRelationships(files, deps.buildRelsPath(sheetPath), sheetPath);
    const imageAssets: ParsedImageAsset[] = [];
    let imageCounter = 1;
    const grid = deps.grid ?? createUniformGrid(deps.defaultCellWidthEmu, deps.defaultCellHeightEmu);

    for (const drawingPath of sheetRels.values()) {
        if (!/\/drawings\/.+\.xml$/i.test(drawingPath)) continue;
        const drawingBytes = files.get(drawingPath);
        if (!drawingBytes) continue;
        const drawingDoc = deps.xmlToDocument(deps.decodeXmlText(drawingBytes));
        const drawingRels = deps.parseRelationships(files, deps.buildRelsPath(drawingPath), drawingPath);
        const anchors = collectAnchors(drawingDoc);

        for (const anchor of anchors) {
            const from = deps.getFirstDescendantByLocalName(anchor, "from");
            const colNode = deps.getFirstDescendantByLocalName(from || anchor, "col");
            const rowNode = deps.getFirstDescendantByLocalName(from || anchor, "row");
            const col = Number(deps.getTextContent(colNode)) + 1;
            const row = Number(deps.getTextContent(rowNode)) + 1;
            if (!Number.isFinite(col) || !Number.isFinite(row) || col <= 0 || row <= 0) {
                continue;
            }

            const blip = deps.getElementsByLocalName(anchor, "blip")[0] || null;
            const embedId = blip?.getAttribute("r:embed") || blip?.getAttribute("embed") || "";
            const mediaPath = drawingRels.get(embedId) || "";
            if (!mediaPath) continue;
            const mediaBytes = files.get(mediaPath);
            if (!mediaBytes) continue;

            const extension = getImageExtension(mediaPath);
            const safeDir = createSafeSheetAssetDir(sheetName);
            const filename = `image_${String(imageCounter).padStart(3, "0")}.${extension}`;
            // The anchor box places the picture in the same space as shapes and cells.
            const geometry = normalizeShapeGeometry(
                anchor,
                deps.getElementsByLocalName(anchor, "pic")[0] || null,
                grid
            );
            imageAssets.push({
                sheetName,
                filename,
                path: `assets/${safeDir}/${filename}`,
                anchor: `${deps.colToLetters(col)}${row}`,
                data: new Uint8Array(mediaBytes),
                mediaPath,
                bbox: { left: geometry.left, top: geometry.top, right: geometry.right, bottom: geometry.bottom },
            });
            imageCounter += 1;
        }
    }

    return imageAssets;
}

export { parseDrawingImages };
