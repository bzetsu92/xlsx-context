import type { ParsedShapeAsset } from "../../model/worksheet.js";

/** CONFIRMED: both ends by declared id. INFERRED: one or both ends by geometry. */
type ConnectionResolution = "CONFIRMED" | "INFERRED" | "UNRESOLVED";

type FlowEvidence = "CONFIRMED" | "INFERRED" | "UNRESOLVED";

type ResolvedShapeConnection = {
    connectorIndex: number;
    connectorId: string;
    startIndex: number | null;
    endIndex: number | null;
    resolution: ConnectionResolution;
    reason: string | null;
    arrowAtStart: boolean;
    arrowAtEnd: boolean;
    flowFrom: number | null;
    flowTo: number | null;
    flowEvidence: FlowEvidence;
};

/**
 * A line end may stop this far outside a shape box and still touch it: one default row (0.2 inch).
 * Boxes use estimated cell sizes, which can stretch a real gap, so the margin stays generous.
 */
const ENDPOINT_TOLERANCE_EMU = 190500;

type Point = { x: number; y: number };

/** Drawing ids are unique per drawing part only. The stable id starts with that part path. */
function drawingPartOf(shape: ParsedShapeAsset): string {
    const stableId = shape.stableId || "";
    const hash = stableId.indexOf("#");
    return hash >= 0 ? stableId.slice(0, hash) : "";
}

/** Distance from a point to the box edge. Zero inside the box. */
function outsideDistance(box: ParsedShapeAsset["bbox"], point: Point): number {
    const dx = Math.max(box.left - point.x, 0, point.x - box.right);
    const dy = Math.max(box.top - point.y, 0, point.y - box.bottom);
    return Math.hypot(dx, dy);
}

/** Distance from a point to the nearest box side, from inside or outside. */
function edgeDistance(box: ParsedShapeAsset["bbox"], point: Point): number {
    const outside = outsideDistance(box, point);
    if (outside > 0) return outside;
    return Math.min(point.x - box.left, box.right - point.x, point.y - box.top, box.bottom - point.y);
}

/**
 * Finds the shape that a loose line end touches.
 * Many files draw connectors next to shapes without gluing them, so the XML has no id.
 * The nearest box side wins. Connectors and shapes from other drawing parts never match.
 */
function findShapeAtPoint(
    shapes: ParsedShapeAsset[],
    connectorIndex: number,
    point: Point | null | undefined,
    scope: Set<number> | null
): number | null {
    if (!point) return null;
    const part = drawingPartOf(shapes[connectorIndex]);
    let best: number | null = null;
    let bestDistance = Number.POSITIVE_INFINITY;
    shapes.forEach((shape, index) => {
        if (scope && !scope.has(index)) return;
        if (!shape || shape.connection || drawingPartOf(shape) !== part) return;
        if (outsideDistance(shape.bbox, point) > ENDPOINT_TOLERANCE_EMU) return;
        const distance = edgeDistance(shape.bbox, point);
        if (distance < bestDistance) {
            best = index;
            bestDistance = distance;
        }
    });
    return best;
}

/** Maps non-connector shape indexes by drawing id. First occurrence wins. */
function buildIndexById(shapes: ParsedShapeAsset[], scope: Set<number> | null): Map<string, number> {
    const indexById = new Map<string, number>();
    shapes.forEach((shape, index) => {
        if (scope && !scope.has(index)) return;
        if (!shape || shape.connection || !shape.drawingId) return;
        if (!indexById.has(shape.drawingId)) indexById.set(shape.drawingId, index);
    });
    return indexById;
}

/** Splits the arrow label into per-end arrowhead flags. XML declares hardware, not flow. */
function readArrowEnds(arrow: string): { arrowAtStart: boolean; arrowAtEnd: boolean } {
    return {
        arrowAtStart: arrow === "start" || arrow === "both",
        arrowAtEnd: arrow === "end" || arrow === "both",
    };
}

/**
 * Derives business flow from arrowheads only.
 * A single arrowhead implies direction. Zero or two arrowheads declare nothing.
 */
function deriveFlow(
    startIndex: number | null,
    endIndex: number | null,
    arrowAtStart: boolean,
    arrowAtEnd: boolean
): { flowFrom: number | null; flowTo: number | null; flowEvidence: FlowEvidence; reason: string | null } {
    if (startIndex === null || endIndex === null) {
        return { flowFrom: null, flowTo: null, flowEvidence: "UNRESOLVED", reason: "connector end is unresolved" };
    }
    if (arrowAtStart && !arrowAtEnd) {
        return {
            flowFrom: endIndex,
            flowTo: startIndex,
            flowEvidence: "INFERRED",
            reason: "single arrowhead at start",
        };
    }
    if (arrowAtEnd && !arrowAtStart) {
        return { flowFrom: startIndex, flowTo: endIndex, flowEvidence: "INFERRED", reason: "single arrowhead at end" };
    }
    if (arrowAtStart && arrowAtEnd) {
        return { flowFrom: null, flowTo: null, flowEvidence: "UNRESOLVED", reason: "arrowheads at both ends" };
    }
    return { flowFrom: null, flowTo: null, flowEvidence: "UNRESOLVED", reason: "no arrowhead declares direction" };
}

/** Resolves one connector against the id map. Missing ids stay unresolved with reasons. */
function resolveOneConnection(
    shapes: ParsedShapeAsset[],
    connectorIndex: number,
    indexById: Map<string, number>,
    scope: Set<number> | null
): ResolvedShapeConnection | null {
    const shape = shapes[connectorIndex];
    const connection = shape?.connection;
    if (!connection) return null;
    const { arrowAtStart, arrowAtEnd } = readArrowEnds(connection.arrow);
    const reasons: string[] = [];
    let inferred = false;
    const resolveEnd = (id: string | null, point: Point | null | undefined, label: string): number | null => {
        if (id !== null && indexById.has(id)) return indexById.get(id) as number;
        const touched = findShapeAtPoint(shapes, connectorIndex, point, scope);
        if (touched !== null) {
            inferred = true;
            return touched;
        }
        reasons.push(id === null ? `missing ${label} id` : `unknown ${label} target '${id}'`);
        return null;
    };
    const startIndex = resolveEnd(connection.startId, connection.startPoint, "stCxn");
    const endIndex = resolveEnd(connection.endId, connection.endPoint, "endCxn");
    const resolution: ConnectionResolution =
        startIndex === null || endIndex === null ? "UNRESOLVED" : inferred ? "INFERRED" : "CONFIRMED";
    const flow = deriveFlow(startIndex, endIndex, arrowAtStart, arrowAtEnd);
    return {
        connectorIndex,
        connectorId: shape.drawingId,
        startIndex,
        endIndex,
        resolution,
        reason: reasons.length > 0 ? reasons.join("; ") : null,
        arrowAtStart,
        arrowAtEnd,
        flowFrom: flow.flowFrom,
        flowTo: flow.flowTo,
        flowEvidence: flow.flowEvidence,
    };
}

/** Resolves every connector in the shape list against the full drawing scope. */
function resolveShapeConnections(shapes: ParsedShapeAsset[]): ResolvedShapeConnection[] {
    const indexById = buildIndexById(shapes, null);
    const resolved: ResolvedShapeConnection[] = [];
    shapes.forEach((_shape, index) => {
        const entry = resolveOneConnection(shapes, index, indexById, null);
        if (entry) resolved.push(entry);
    });
    return resolved;
}

export { resolveShapeConnections };
export type { ResolvedShapeConnection, ConnectionResolution, FlowEvidence };
