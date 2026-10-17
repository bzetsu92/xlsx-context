import type { ParsedShapeAsset } from "../../model/index.js";
import { resolveShapeConnections, type ResolvedShapeConnection } from "../../extraction/assets/shape-connections.js";

/** One connector between two shapes. Indexes point into the sheet shape list. */
export type ShapeGraphEdge = {
    from: number;
    to: number;
    connector: number;
    /** "end" draws the arrowhead at `to`. */
    arrow: "none" | "end" | "both";
};

export type ShapeGraph = {
    /** Shapes that at least one edge touches, in shape order. */
    nodes: number[];
    edges: ShapeGraphEdge[];
    /** Connectors with one or two ends not attached to a shape in the set. */
    looseConnectors: number[];
    /** Loose connectors with the reason each end stayed unresolved. */
    unresolved?: { connector: number; reason: string | null }[];
    /** Main flow direction, from the positions of joined shapes. */
    direction: "TD" | "LR";
};

function centerOf(shape: ParsedShapeAsset): { x: number; y: number } {
    const box = shape.bbox;
    return { x: (box.left + box.right) / 2, y: (box.top + box.bottom) / 2 };
}

/** Connector ends as shape indexes. `to` is the arrowhead end. A null end is loose. */
export type OrientedConnection = { from: number | null; to: number | null; arrow: "none" | "end" | "both" };

/** Flips a start-only arrowhead, so `to` is always the arrow target. */
function orientConnection(entry: ResolvedShapeConnection): OrientedConnection {
    if (entry.arrowAtStart && !entry.arrowAtEnd) {
        return { from: entry.endIndex, to: entry.startIndex, arrow: "end" };
    }
    return {
        from: entry.startIndex,
        to: entry.endIndex,
        arrow: entry.arrowAtStart && entry.arrowAtEnd ? "both" : entry.arrowAtEnd ? "end" : "none",
    };
}

/** Turns one resolved connection into an edge. Loose and self-joined connectors give no edge. */
function toEdge(entry: ResolvedShapeConnection): ShapeGraphEdge | null {
    const { from, to, arrow } = orientConnection(entry);
    if (from === null || to === null || from === to) return null;
    return { from, to, connector: entry.connectorIndex, arrow };
}

/** Builds a graph from resolved connections. Unresolved ends stay loose but remain listed. */
function buildGraphFromResolved(shapes: ParsedShapeAsset[], resolved: ResolvedShapeConnection[]): ShapeGraph {
    const edges: ShapeGraphEdge[] = [];
    const looseConnectors: number[] = [];
    const unresolved: { connector: number; reason: string | null }[] = [];
    for (const entry of resolved) {
        const edge = toEdge(entry);
        if (!edge) {
            looseConnectors.push(entry.connectorIndex);
            unresolved.push({ connector: entry.connectorIndex, reason: entry.reason });
            continue;
        }
        edges.push(edge);
    }

    const touched = new Set(edges.flatMap((edge) => [edge.from, edge.to]));
    const nodes = shapes.map((_shape, index) => index).filter((index) => touched.has(index));

    let horizontal = 0;
    let vertical = 0;
    for (const edge of edges) {
        const from = centerOf(shapes[edge.from]);
        const to = centerOf(shapes[edge.to]);
        horizontal += Math.abs(to.x - from.x);
        vertical += Math.abs(to.y - from.y);
    }

    return { nodes, edges, looseConnectors, unresolved, direction: horizontal > vertical ? "LR" : "TD" };
}

/** Maps each connector index to its oriented ends over the full shape list. */
function resolveConnectorEnds(shapes: ParsedShapeAsset[]): Map<number, OrientedConnection> {
    return new Map(resolveShapeConnections(shapes).map((entry) => [entry.connectorIndex, orientConnection(entry)]));
}

/** Builds the graph over the full shape list. Cross-block links survive partitioning. */
function buildShapeGraph(shapes: ParsedShapeAsset[]): ShapeGraph {
    return buildGraphFromResolved(shapes, resolveShapeConnections(shapes));
}

/**
 * Builds the graph slice for one block with globally accurate reasons.
 * Edges stay inside the block. Loose connectors keep drawing-wide resolve reasons,
 * so an endpoint in another block never reports as unknown.
 */
function buildShapeGraphForBlock(shapes: ParsedShapeAsset[], shapeIndexes: number[]): ShapeGraph {
    const scope = new Set(shapeIndexes);
    const edges: ShapeGraphEdge[] = [];
    const looseConnectors: number[] = [];
    const unresolved: { connector: number; reason: string | null }[] = [];
    for (const entry of resolveShapeConnections(shapes)) {
        if (!scope.has(entry.connectorIndex)) continue;
        const edge = toEdge(entry);
        if (edge && scope.has(edge.from) && scope.has(edge.to)) {
            edges.push(edge);
            continue;
        }
        looseConnectors.push(entry.connectorIndex);
        unresolved.push({ connector: entry.connectorIndex, reason: entry.reason });
    }

    const touched = new Set(edges.flatMap((edge) => [edge.from, edge.to]));
    const nodes = shapeIndexes.filter((index) => touched.has(index));

    let horizontal = 0;
    let vertical = 0;
    for (const edge of edges) {
        const from = centerOf(shapes[edge.from]);
        const to = centerOf(shapes[edge.to]);
        horizontal += Math.abs(to.x - from.x);
        vertical += Math.abs(to.y - from.y);
    }

    return { nodes, edges, looseConnectors, unresolved, direction: horizontal > vertical ? "LR" : "TD" };
}

export { buildShapeGraph, buildShapeGraphForBlock, resolveConnectorEnds };
