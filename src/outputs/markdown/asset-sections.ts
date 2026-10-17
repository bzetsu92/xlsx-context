import type {
    MarkdownFile,
    ParsedCellComment,
    ParsedChartAsset,
    ParsedImageAsset,
    ParsedShapeAsset,
    ParsedSheet,
} from "../../model/index.js";
import type { LayoutEntry } from "../../structure/layout/sheet-layout.js";
import { buildShapeGraphForBlock, type ShapeGraph } from "../../structure/layout/shape-graph.js";
import type { NearbyText, ShapeTarget } from "../../structure/layout/shape-targets.js";
import { groupShapeStyles, type ShapeStyle } from "../../structure/layout/shape-styles.js";
import { colorName } from "../../extraction/assets/shape-look.js";
import { renderMermaidFlowchart } from "./mermaid-flowchart.js";
import type { SheetMarkdownDeps } from "./sheet-markdown.js";

export type ShapeBlock = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    shapeIndexes: number[];
};

/** Asset section dependencies. Only range, address, table, and raw-entry rendering is injected. */
type AssetSectionDeps = Pick<
    SheetMarkdownDeps,
    "renderHierarchicalRawEntries" | "formatRange" | "colToLetters" | "parseCellAddress" | "renderMarkdownTable"
>;

/** A rendered chart, image, or shape block, placed at its anchor in the sheet reading order. */
export type AssetEntry = LayoutEntry & { kind: "asset"; markdown: string };

/** Where a chart takes its data from. Labels name the Markdown tables, for example `Table 001 (B3-E8)`. */
export type ChartSource = { tableLabels: string[]; complete: boolean };

/** Renders image, chart, shape, and comment sections. */
export function createAssetSections(deps: AssetSectionDeps) {
    function createDefaultShapeSvgFilename(shapeIndex: number): string {
        return `shape_${String(shapeIndex + 1).padStart(3, "0")}.svg`;
    }

    function createImageSectionEntry(image: ParsedImageAsset, index: number): string {
        return [
            `### Image: ${String(index + 1).padStart(3, "0")} (${image.anchor})`,
            `- File: ${image.path}`,
            "",
            `![${image.filename}](${image.path})`,
        ].join("\n");
    }

    function createChartSeriesLines(chart: ParsedChartAsset): string[] {
        if (chart.series.length === 0) {
            return [];
        }
        const lines = ["- Series:"];
        for (const series of chart.series) {
            lines.push(`  - ${series.name}`);
            if (series.axis === "secondary") lines.push("    - Axis: secondary");
            if (series.categoriesRef) lines.push(`    - categories: ${series.categoriesRef}`);
            if (series.valuesRef) lines.push(`    - values: ${series.valuesRef}`);
        }
        return lines;
    }

    /** Builds the cached data grid: one category column, then one column per series. */
    function createChartDataTable(chart: ParsedChartAsset): string {
        const series = chart.series.filter((entry) => entry.cachedValues?.length > 0);
        if (series.length === 0) return "";
        const categories = series.find((entry) => entry.cachedCategories?.length > 0)?.cachedCategories || [];
        const rowCount = Math.max(categories.length, ...series.map((entry) => entry.cachedValues.length));
        const rows = [["Category", ...series.map((entry) => entry.name)]];
        for (let row = 0; row < rowCount; row += 1) {
            rows.push([
                categories[row] ?? String(row + 1),
                ...series.map((entry) => {
                    const value = entry.cachedValues[row];
                    return value === null || value === undefined ? "" : String(value);
                }),
            ]);
        }
        return deps.renderMarkdownTable(rows, true);
    }

    /**
     * Points to the source tables. Adds the cached data only when no table on this sheet holds every series.
     * Then the reader never has to resolve a range by hand.
     */
    function createChartSectionEntry(chart: ParsedChartAsset, index: number, source?: ChartSource): string {
        const lines = [
            `### Chart: ${String(index + 1).padStart(3, "0")} (${chart.anchor})`,
            `- Title: ${chart.title || "(none)"}`,
            `- Type: ${chart.chartType}`,
        ];
        if (source && source.tableLabels.length > 0) lines.push(`- Source: ${source.tableLabels.join(", ")}`);
        lines.push(...createChartSeriesLines(chart));
        const dataTable = source?.complete ? "" : createChartDataTable(chart);
        if (dataTable) lines.push("- Data (chart cache):", "", dataTable);
        return lines.join("\n");
    }

    function renderShapeDetails(shape: ParsedShapeAsset, shapeIndex: number): string {
        const lines = [
            `#### Shape: ${String(shapeIndex + 1).padStart(3, "0")} (${shape.anchor})`,
            ...deps.renderHierarchicalRawEntries(shape.rawEntries),
        ];
        if (shape.svgWarning) {
            lines.push(`- Warning: ${shape.svgWarning}`);
        }
        if (shape.svgPath) {
            lines.push(`- SVG: ${shape.svgPath}`);
            lines.push("");
            lines.push(`![${shape.svgFilename || createDefaultShapeSvgFilename(shapeIndex)}](${shape.svgPath})`);
        }
        return lines.join("\n");
    }

    function createShapeBlockSummaryLine(shapeIndexes: number[]): string {
        return shapeIndexes.map((shapeIndex) => shapeLabel(shapeIndex)).join(", ");
    }

    function shapeLabel(shapeIndex: number): string {
        return `Shape ${String(shapeIndex + 1).padStart(3, "0")}`;
    }

    /** Quotes cell text on one line. Long text keeps its start. */
    function quoteText(address: string, text: string): string {
        const flat = text.replace(/\s*\r?\n\s*/g, " / ").trim();
        return `${address} "${flat.length > 40 ? `${flat.slice(0, 40)}…` : flat}"`;
    }

    function describeAlso(nearby: NearbyText[]): string {
        return nearby.length > 0
            ? `; also ${nearby.map((entry) => quoteText(entry.address, entry.text)).join(", ")}`
            : "";
    }

    function imageLabel(imageIndex: number): string {
        return `Image ${String(imageIndex + 1).padStart(3, "0")}`;
    }

    /** Describes the place a callout or loose line end points at: the cell, a shape or picture there, and the nearest text. */
    function describePoint(target: ShapeTarget): string {
        const parts: string[] = [];
        if (target.targetShapeIndex !== null) parts.push(`in ${shapeLabel(target.targetShapeIndex)}`);
        if (target.targetImageIndex !== null) parts.push(`in ${imageLabel(target.targetImageIndex)}`);
        if (target.textAddress && target.text !== null) {
            parts.push(`text ${quoteText(target.textAddress, target.text)}${describeAlso(target.nearby)}`);
        } else if (target.targetShapeIndex === null && target.targetImageIndex === null) {
            parts.push("no text nearby");
        }
        return `${target.cell} (${parts.join(", ")})`;
    }

    /** Names what a marker labels: a picture, text in a shape, or cell text. */
    function describeMarkerItem(target: ShapeTarget): string {
        if (target.targetImageIndex !== null) return `a point in ${imageLabel(target.targetImageIndex)}`;
        if (!target.textAddress || target.text === null) return "";
        const text = quoteText(target.textAddress, target.text);
        return target.targetShapeIndex !== null ? `${shapeLabel(target.targetShapeIndex)} ${text}` : text;
    }

    /** Short form for the item list. The Markers section keeps the other nearby texts. */
    function describeMarker(target: ShapeTarget): string {
        const item = describeMarkerItem(target);
        if (!item) return `marks ${target.cell} (no text nearby)`;
        return `marks ${item} (${target.method})`;
    }

    /** Names both ends of a connector with a loose end. Attached ends name their shape. */
    function describeLooseLine(shape: ParsedShapeAsset, targets: ShapeTarget[]): string | null {
        const start = targets.find((target) => target.role === "lineStart");
        const end = targets.find((target) => target.role === "lineEnd");
        if (!start && !end) return null;
        const attached = (start ?? end)?.connectedShapeIndex ?? null;
        const startText = start ? describePoint(start) : attached !== null ? shapeLabel(attached) : "?";
        const endText = end ? describePoint(end) : attached !== null ? shapeLabel(attached) : "?";
        const arrow = shape.connection?.arrow || "none";
        // A start-only arrowhead points back at the start. Write it as an arrow from end to start.
        if (arrow === "start") return `${endText} → ${startText}`;
        return `${startText} ${arrow === "both" ? "↔" : arrow === "end" ? "→" : "—"} ${endText}`;
    }

    /** One line per shape: kind and text. Connectors name what they join. Callouts and markers name what they point at. */
    function createShapeItemLines(
        shapes: ParsedShapeAsset[],
        shapeIndexes: number[],
        diagram: ShapeGraph,
        targetsByShape: Map<number, ShapeTarget[]>,
        styleByShape: Map<number, string>
    ): string[] {
        const edgeByConnector = new Map(diagram.edges.map((edge) => [edge.connector, edge]));
        const arrows = { none: "—", end: "→", both: "↔" } as const;
        return shapeIndexes.flatMap((shapeIndex) => {
            const shape = shapes[shapeIndex];
            if (!shape) return [];
            const tags = [shape.kind, styleByShape.get(shapeIndex)].filter(Boolean).join(", ");
            const head = tags ? `${shapeLabel(shapeIndex)} (${tags})` : shapeLabel(shapeIndex);
            const text = String(shape.markedText ?? shape.text ?? "")
                .replace(/\s*\r?\n\s*/g, " / ")
                .trim();
            const edge = edgeByConnector.get(shapeIndex);
            const targets = targetsByShape.get(shapeIndex) || [];
            let body = text;
            if (edge) {
                const link = `${shapeLabel(edge.from)} ${arrows[edge.arrow]} ${shapeLabel(edge.to)}`;
                body = text ? `${link} (${text})` : link;
            } else if (shape.connection) {
                const link = describeLooseLine(shape, targets) || "not attached";
                body = text ? `${link} (${text})` : link;
            } else {
                const callout = targets.find((target) => target.role === "callout");
                const marker = targets.find((target) => target.role === "marker");
                if (callout) body = `${text} → points to ${describePoint(callout)}`;
                else if (marker) body = `${text} → ${describeMarker(marker)}`;
            }
            return [`  - ${head}${body ? `: ${body}` : ""}`];
        });
    }

    function describeColor(value: string | null): string {
        if (value === null) return "default";
        if (!value.startsWith("#")) return value;
        return `${colorName(value)} ${value}`;
    }

    function describeLine(style: ShapeStyle): string {
        const { line, dash, lineWidthEmu } = style.look;
        if (line === "none") return "none";
        const parts = [describeColor(line)];
        if (dash && dash !== "solid") parts.push(`dashed (${dash})`);
        if (style.thick && lineWidthEmu !== null)
            parts.push(`thick ${Math.round((lineWidthEmu / 12700) * 100) / 100} pt`);
        return parts.join(", ");
    }

    /**
     * Lists the looks of the shapes on one sheet. Item lines name their style, for example S2.
     * Each row shows the first texts of its shapes as examples.
     */
    function createStyleSection(styles: ShapeStyle[], shapes: ParsedShapeAsset[]): string {
        const rows = styles.map((style) => {
            const examples = style.shapeIndexes
                .map((index) =>
                    String(shapes[index]?.text || "")
                        .replace(/\s+/g, " ")
                        .trim()
                )
                .filter(Boolean)
                .slice(0, 3)
                .map((text) => `"${text.length > 24 ? `${text.slice(0, 24)}…` : text}"`)
                .join(", ");
            return [
                style.id,
                style.presets.join(", "),
                style.isLine ? "—" : describeColor(style.look.fill),
                describeLine(style),
                String(style.shapeIndexes.length),
                examples,
            ];
        });
        return [
            "### Shape Styles",
            "- Shapes with one look share a style. Writers often give each style one meaning.",
            "",
            "| Style | Shape | Fill | Line | Count | Examples |",
            "| --- | --- | --- | --- | --- | --- |",
            ...rows.map((cells) => `| ${cells.map(escapeTableCell).join(" | ")} |`),
        ].join("\n");
    }

    function escapeTableCell(value: string): string {
        return value.replace(/\s*\r?\n\s*/g, " / ").replace(/\|/g, "\\|");
    }

    /**
     * Lists every marker of a sheet in reading order, with the text it labels.
     * Forms often repeat one marker number in several copies, so each row keeps its cell.
     */
    function createMarkerSection(markers: ShapeTarget[], shapes: ParsedShapeAsset[]): string {
        const rows = markers.map((target) => {
            const shape = shapes[target.shapeIndex] as ParsedShapeAsset;
            const text = describeMarkerItem(target);
            const also = target.nearby.map((entry) => quoteText(entry.address, entry.text)).join(", ");
            return [
                String(shape.text || "").trim(),
                target.cell,
                text,
                target.method,
                also,
                shapeLabel(target.shapeIndex),
            ];
        });
        return [
            "### Markers",
            "- Each marker is a small numbered shape. It labels the text next to it.",
            "- Match: inside = on the text cell or shape, right = text starts after the marker, image = on a picture, overflow = on text that runs past its cell, nearest = closest text or picture.",
            "",
            "| Marker | Cell | Text | Match | Also near | Shape |",
            "| --- | --- | --- | --- | --- | --- |",
            ...rows.map((cells) => `| ${cells.map(escapeTableCell).join(" | ")} |`),
        ].join("\n");
    }

    /**
     * Lists the shapes with their text and links, then draws joined shapes as a Mermaid flowchart.
     * Raw drawing details follow only when includeShapeDetails is on.
     */
    function createShapeGroupBody(
        shapes: ParsedShapeAsset[],
        shapeIndexes: number[],
        includeShapeDetails: boolean,
        targetsByShape: Map<number, ShapeTarget[]>,
        styleByShape: Map<number, string>
    ): string[] {
        const diagram = buildShapeGraphForBlock(shapes, shapeIndexes);
        const lines = [
            "- Items:",
            ...createShapeItemLines(shapes, shapeIndexes, diagram, targetsByShape, styleByShape),
        ];
        const flowchart = renderMermaidFlowchart(shapes, diagram);
        if (flowchart) lines.push("", flowchart);
        if (includeShapeDetails) {
            const details = shapeIndexes
                .map((shapeIndex) => (shapes[shapeIndex] ? renderShapeDetails(shapes[shapeIndex], shapeIndex) : ""))
                .filter(Boolean)
                .join("\n\n");
            if (details) lines.push("", details);
        }
        return lines;
    }

    function createShapeBlockEntry(
        block: ShapeBlock,
        blockIndex: number,
        shapes: ParsedShapeAsset[],
        includeShapeDetails: boolean,
        targetsByShape: Map<number, ShapeTarget[]>,
        styleByShape: Map<number, string>
    ): string {
        return [
            `### Shape Block: ${String(blockIndex + 1).padStart(3, "0")} (${deps.formatRange(block.startRow, block.startCol, block.endRow, block.endCol)})`,
            `- Shapes: ${createShapeBlockSummaryLine(block.shapeIndexes)}`,
            `- anchorRange: ${deps.colToLetters(block.startCol)}${block.startRow}-${deps.colToLetters(block.endCol)}${block.endRow}`,
            ...createShapeGroupBody(shapes, block.shapeIndexes, includeShapeDetails, targetsByShape, styleByShape),
        ].join("\n");
    }

    function collectUngroupedShapes(
        shapes: ParsedShapeAsset[],
        shapeBlocks: ShapeBlock[]
    ): Array<{ shape: ParsedShapeAsset; index: number }> {
        const grouped = new Set(shapeBlocks.flatMap((block) => block.shapeIndexes));
        return shapes.map((shape, index) => ({ shape, index })).filter(({ index }) => !grouped.has(index));
    }

    function createAssetEntry(anchor: { row: number; col: number }, markdown: string): AssetEntry {
        return { sortRow: anchor.row, sortCol: anchor.col, kind: "asset", markdown: `${markdown}\n` };
    }

    /**
     * Renders charts, images, and shape blocks as entries at their anchors.
     * The sheet layout then places each one next to the tables and text around it.
     * chartSources follows the chart order.
     */
    function createAssetEntries(
        sheet: ParsedSheet,
        charts: ParsedChartAsset[],
        shapes: ParsedShapeAsset[],
        shapeBlocks: ShapeBlock[],
        options: { includeShapeDetails: boolean; chartSources?: ChartSource[]; shapeTargets?: ShapeTarget[] }
    ): AssetEntry[] {
        const entries: AssetEntry[] = [];
        const targetsByShape = new Map<number, ShapeTarget[]>();
        for (const target of options.shapeTargets || []) {
            const list = targetsByShape.get(target.shapeIndex);
            if (list) list.push(target);
            else targetsByShape.set(target.shapeIndex, [target]);
        }
        // One look tells nothing. Styles show only when shapes differ from shapes, or lines from lines.
        const styles = groupShapeStyles(shapes);
        const styleByShape = new Map<number, string>();
        const shapeLooks = styles.filter((style) => !style.isLine).length;
        const lineLooks = styles.length - shapeLooks;
        if (shapeLooks > 1 || lineLooks > 1) {
            for (const style of styles) for (const index of style.shapeIndexes) styleByShape.set(index, style.id);
        }
        charts.forEach((chart, index) => {
            const markdown = createChartSectionEntry(chart, index, options.chartSources?.[index]);
            entries.push(createAssetEntry(deps.parseCellAddress(chart.anchor), markdown));
        });
        (sheet.images || []).forEach((image, index) => {
            entries.push(createAssetEntry(deps.parseCellAddress(image.anchor), createImageSectionEntry(image, index)));
        });
        shapeBlocks.forEach((block, blockIndex) => {
            const markdown = createShapeBlockEntry(
                block,
                blockIndex,
                shapes,
                options.includeShapeDetails,
                targetsByShape,
                styleByShape
            );
            entries.push(createAssetEntry({ row: block.startRow, col: block.startCol }, markdown));
        });
        const ungrouped = collectUngroupedShapes(shapes, shapeBlocks);
        if (ungrouped.length > 0) {
            const indexes = ungrouped.map(({ index }) => index);
            const anchor = deps.parseCellAddress(ungrouped[0].shape.anchor);
            const markdown = [
                "### Ungrouped Shapes",
                ...createShapeGroupBody(shapes, indexes, options.includeShapeDetails, targetsByShape, styleByShape),
            ].join("\n");
            entries.push(createAssetEntry(anchor, markdown));
        }
        if (styleByShape.size > 0) {
            // The legend goes before every shape of the sheet.
            const first = shapes
                .filter((shape) => shape?.anchor)
                .map((shape) => deps.parseCellAddress(shape.anchor))
                .reduce((a, b) => (b.row < a.row || (b.row === a.row && b.col < a.col) ? b : a));
            entries.push(createAssetEntry(first, createStyleSection(styles, shapes)));
        }
        const markers = (options.shapeTargets || [])
            .filter((target) => target.role === "marker")
            .map((target) => ({ target, at: deps.parseCellAddress(target.cell) }))
            .sort((a, b) => a.at.row - b.at.row || a.at.col - b.at.col);
        if (markers.length > 0) {
            const markdown = createMarkerSection(
                markers.map(({ target }) => target),
                shapes
            );
            entries.push(createAssetEntry((markers[0] as (typeof markers)[number]).at, markdown));
        }
        return entries;
    }

    function createCommentSectionEntry(comment: ParsedCellComment, index: number): string {
        const label = `comment-${index + 1}`;
        const metadata = [
            comment.address,
            comment.kind,
            comment.author ? `author=${comment.author}` : "",
            comment.dateTime ? `date=${comment.dateTime}` : "",
        ]
            .filter(Boolean)
            .join("; ");
        return `- [${label}] (${metadata}) ${comment.text}`;
    }

    function renderCommentSection(comments: ParsedCellComment[] = []): string {
        return comments.length > 0
            ? [
                  "",
                  "### Comments",
                  "",
                  ...comments.map((comment, index) => createCommentSectionEntry(comment, index)),
              ].join("\n\n")
            : "";
    }

    function createFormulaDiagnostics(sheet: ParsedSheet): MarkdownFile["summary"]["formulaDiagnostics"] {
        return sheet.cells
            .filter((cell) => !!cell.formulaText && cell.resolutionStatus !== null)
            .map((cell) => ({
                address: cell.address,
                formulaText: cell.formulaText,
                status: cell.resolutionStatus,
                source: cell.resolutionSource,
                outputValue: cell.outputValue,
            }));
    }

    return {
        createAssetEntries,
        renderCommentSection,
        createFormulaDiagnostics,
    };
}
