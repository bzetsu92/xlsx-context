/** Shared sheet-layout types. Narrative and section logic live beside this file. */
import type { NarrativeBlock, SectionBlock, TableCandidate } from "../../model/index.js";

/** A positioned piece of sheet content. Outputs extend it with their rendered text. */
export type LayoutEntry = {
    sortRow: number;
    sortCol: number;
    /** Assets are charts, images, and shape blocks. They join sections by position only. */
    kind: "narrative" | "table" | "asset";
    narrativeBlock?: NarrativeBlock;
    table?: TableCandidate;
};

export type GroupedEntries<T extends LayoutEntry> = {
    block: SectionBlock;
    entries: T[];
};

/** One entry of a sheet's reading order. */
export type ReadingOrderEntry = { kind: "table"; table: TableCandidate } | { kind: "narrative"; block: NarrativeBlock };
