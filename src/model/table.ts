/** Declared and detected tables with scores. */
/** Excel table object declared on a worksheet. */
export type ParsedTable = {
    sheetName: string;
    name: string;
    displayName: string;
    start: string;
    end: string;
    columns: string[];
    headerRowCount: number;
    totalsRowCount: number;
};

/** Detected table region with a score and scored reasons. */
export type TableCandidate = {
    startRow: number;
    startCol: number;
    endRow: number;
    endCol: number;
    score: number;
    reasonSummary: string[];
    /** False when the first row reads as data (for example "base1 | 10"). Renderers then emit no header. */
    hasHeaderRow?: boolean;
    /** Name of the Excel table when the file declares this range as a table. */
    name?: string;
};

/** Scored table range reported in a file summary. */
export type TableScoreDetail = {
    range: string;
    score: number;
    reasons: string[];
};
