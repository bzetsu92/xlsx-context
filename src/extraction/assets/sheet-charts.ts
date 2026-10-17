import type { ParsedChartAsset } from "../../model/index.js";
import type { ParseAssetDependencies } from "./sheet-assets.js";
import { collectAnchors } from "./sheet-assets.js";

/** Groups descendant elements by local name in one pass. Skips the root. Keeps document order per group. */
function groupByLocalName(root: Node, names: readonly string[]): Map<string, Element[]> {
    const buckets = new Map<string, Element[]>(names.map((name) => [name, []]));
    const stack: Node[] = [];
    for (let child = root.lastChild; child; child = child.previousSibling) stack.push(child);
    while (stack.length > 0) {
        const current = stack.pop() as Node;
        if (current.nodeType === 1) {
            buckets.get((current as Element).localName)?.push(current as Element);
        }
        for (let child = current.lastChild; child; child = child.previousSibling) stack.push(child);
    }
    return buckets;
}

/** Plot element names with their labels. 3D and special types keep the same series layout as 2D types. */
const CHART_TYPES: ReadonlyArray<{ localName: string; label: string }> = [
    { localName: "barChart", label: "Bar Chart" },
    { localName: "bar3DChart", label: "3D Bar Chart" },
    { localName: "lineChart", label: "Line Chart" },
    { localName: "line3DChart", label: "3D Line Chart" },
    { localName: "pieChart", label: "Pie Chart" },
    { localName: "pie3DChart", label: "3D Pie Chart" },
    { localName: "ofPieChart", label: "Pie of Pie Chart" },
    { localName: "doughnutChart", label: "Doughnut Chart" },
    { localName: "areaChart", label: "Area Chart" },
    { localName: "area3DChart", label: "3D Area Chart" },
    { localName: "scatterChart", label: "Scatter Chart" },
    { localName: "radarChart", label: "Radar Chart" },
    { localName: "bubbleChart", label: "Bubble Chart" },
    { localName: "stockChart", label: "Stock Chart" },
    { localName: "surfaceChart", label: "Surface Chart" },
    { localName: "surface3DChart", label: "3D Surface Chart" },
];

const CHART_TYPE_NAMES = CHART_TYPES.map((entry) => entry.localName);

const GROUPING_LABELS: Record<string, string> = { stacked: "stacked", percentStacked: "100% stacked" };

/**
 * Adds the layout that changes how values read: horizontal bars and stacked values.
 * Clustered and standard layouts add nothing.
 */
function describePlotLayout(plotNode: Element, deps: ParseAssetDependencies): string[] {
    const barDir = deps.getDirectChildByLocalName(plotNode, "barDir")?.getAttribute("val");
    const grouping = deps.getDirectChildByLocalName(plotNode, "grouping")?.getAttribute("val") || "";
    const ofPieType = deps.getDirectChildByLocalName(plotNode, "ofPieType")?.getAttribute("val");
    return [
        barDir === "bar" ? "horizontal" : "",
        GROUPING_LABELS[grouping] || "",
        ofPieType === "bar" ? "bar of pie" : "",
    ].filter(Boolean);
}

/** Labels the chart by its plot types. Joins several types as one combined label. Falls back to `Chart`. */
function parseChartType(chartDoc: Document, deps: ParseAssetDependencies): string {
    const buckets = groupByLocalName(chartDoc, CHART_TYPE_NAMES);
    const matched = CHART_TYPES.flatMap((entry) => {
        const plotNode = (buckets.get(entry.localName) as Element[])[0];
        if (!plotNode) return [];
        const layout = describePlotLayout(plotNode, deps);
        return [layout.length > 0 ? `${entry.label} (${layout.join(", ")})` : entry.label];
    });
    if (matched.length === 0) return "Chart";
    if (matched.length === 1) return matched[0];
    return `${matched.join(" + ")} (Combined)`;
}

/**
 * Reads the chart title only, never axis or data label titles.
 * A rich title joins its paragraphs with spaces. A title linked to a cell uses the cached text.
 * Returns an empty string without a title.
 */
function parseChartTitle(chartDoc: Document, deps: ParseAssetDependencies): string {
    const chartNode = deps.getFirstDescendantByLocalName(chartDoc, "chart");
    const titleText = deps.getDirectChildByLocalName(deps.getDirectChildByLocalName(chartNode, "title"), "tx");
    if (!titleText) return "";
    const rich = deps.getDirectChildByLocalName(titleText, "rich");
    if (rich) {
        return deps
            .getElementsByLocalName(rich, "p")
            .map((paragraph) =>
                deps
                    .getElementsByLocalName(paragraph, "t")
                    .map((node) => deps.getTextContent(node))
                    .join("")
                    .trim()
            )
            .filter(Boolean)
            .join(" ");
    }
    const cachedText = deps.getElementsByLocalName(titleText, "v").map((node) => deps.getTextContent(node));
    return cachedText.join("").trim();
}

/** Upper bound for cached points per series. */
const MAX_CHART_POINTS = 100_000;

/**
 * Reads one series data element: the range formula and the cached points.
 * Takes the first matching child, so scatter charts use xVal and yVal.
 * Multi-level categories keep the innermost level only.
 */
function readSeriesData(
    seriesNode: Element,
    localNames: string[],
    deps: ParseAssetDependencies
): { ref: string; points: (string | null)[] } {
    const dataNode = localNames
        .map((localName) => deps.getDirectChildByLocalName(seriesNode, localName))
        .find((node): node is Element => !!node);
    if (!dataNode) return { ref: "", points: [] };
    const ref = deps.getTextContent(deps.getFirstDescendantByLocalName(dataNode, "f"));
    const pointRoot = deps.getFirstDescendantByLocalName(dataNode, "lvl") || dataNode;
    const pointNodes = deps.getElementsByLocalName(pointRoot, "pt");
    // ptCount is the series length. Without it, the highest idx decides. The cap guards against corrupt counts.
    const declared = Number(deps.getFirstDescendantByLocalName(dataNode, "ptCount")?.getAttribute("val") ?? NaN);
    const highest = pointNodes.reduce((max, node) => Math.max(max, Number(node.getAttribute("idx")) || 0), -1);
    const count = Math.min(Number.isInteger(declared) && declared >= 0 ? declared : highest + 1, MAX_CHART_POINTS);
    const points: (string | null)[] = new Array(count).fill(null);
    for (const node of pointNodes) {
        const index = Number(node.getAttribute("idx") || "");
        if (!Number.isInteger(index) || index < 0 || index >= count) continue;
        points[index] = deps.getTextContent(deps.getFirstDescendantByLocalName(node, "v"));
    }
    return { ref, points };
}

/** Reads each series name with its category and value references. Marks right-axis series as secondary. */
function parseChartSeries(chartDoc: Document, deps: ParseAssetDependencies): ParsedChartAsset["series"] {
    const plotArea = deps.getFirstDescendantByLocalName(chartDoc, "plotArea") || chartDoc.documentElement;
    const axisPositionById = new Map<string, string>();
    for (const axisNode of deps.getElementsByLocalName(plotArea, "valAx")) {
        const axisIdNode = deps.getFirstDescendantByLocalName(axisNode, "axId");
        const axisPosNode = deps.getFirstDescendantByLocalName(axisNode, "axPos");
        const axisId = axisIdNode?.getAttribute("val") || deps.getTextContent(axisIdNode);
        const axisPos = axisPosNode?.getAttribute("val") || deps.getTextContent(axisPosNode);
        if (axisId) {
            axisPositionById.set(axisId, axisPos || "");
        }
    }

    const series: ParsedChartAsset["series"] = [];

    const containers = groupByLocalName(plotArea, CHART_TYPE_NAMES);
    for (const localName of CHART_TYPE_NAMES) {
        for (const chartNode of containers.get(localName) as Element[]) {
            const axisIds = deps
                .getElementsByLocalName(chartNode, "axId")
                .map((node) => node.getAttribute("val") || deps.getTextContent(node))
                .filter(Boolean);
            const isSecondary = axisIds.some((axisId) => axisPositionById.get(axisId) === "r");

            for (const seriesNode of deps.getElementsByLocalName(chartNode, "ser")) {
                const txNode = deps.getFirstDescendantByLocalName(seriesNode, "tx") || seriesNode;
                const nameRef = deps.getFirstDescendantByLocalName(txNode, "f");
                const nameValue = deps.getFirstDescendantByLocalName(txNode, "v");
                const nameText = deps
                    .getElementsByLocalName(txNode, "t")
                    .map((node) => deps.getTextContent(node))
                    .join("")
                    .trim();
                const categories = readSeriesData(seriesNode, ["cat", "xVal"], deps);
                const values = readSeriesData(seriesNode, ["val", "yVal"], deps);
                series.push({
                    name: nameText || deps.getTextContent(nameValue) || deps.getTextContent(nameRef) || "Series",
                    categoriesRef: categories.ref,
                    valuesRef: values.ref,
                    cachedCategories: categories.points.map((point) => point ?? ""),
                    cachedValues: values.points.map((point) => {
                        const value = point === null || point.trim() === "" ? NaN : Number(point);
                        return Number.isFinite(value) ? value : null;
                    }),
                    axis: isSecondary ? "secondary" : "primary",
                });
            }
        }
    }

    return series;
}

/** Collects anchored charts with type, title, and series. */
function parseDrawingCharts(
    files: Map<string, Uint8Array>,
    sheetName: string,
    sheetPath: string,
    deps: ParseAssetDependencies
): ParsedChartAsset[] {
    const sheetRels = deps.parseRelationships(files, deps.buildRelsPath(sheetPath), sheetPath);
    const charts: ParsedChartAsset[] = [];

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

            const chartNode = deps.getFirstDescendantByLocalName(anchor, "graphicFrame");
            const chartRef = deps.getElementsByLocalName(chartNode || anchor, "chart")[0] || null;
            const relId = chartRef?.getAttribute("r:id") || chartRef?.getAttribute("id") || "";
            const chartPath = drawingRels.get(relId) || "";
            if (!chartPath) continue;
            const chartBytes = files.get(chartPath);
            if (!chartBytes) continue;
            const chartDoc = deps.xmlToDocument(deps.decodeXmlText(chartBytes));

            charts.push({
                sheetName,
                anchor: `${deps.colToLetters(col)}${row}`,
                chartPath,
                title: parseChartTitle(chartDoc, deps),
                chartType: parseChartType(chartDoc, deps),
                series: parseChartSeries(chartDoc, deps),
            });
        }
    }

    return charts;
}

export { parseDrawingCharts };
