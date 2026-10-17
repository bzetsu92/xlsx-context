import type {
    BorderFlags,
    CellStyleInfo,
    ParseWorkbookOptions,
    ParsedCell,
    ParsedRangeRef,
    ParsedSheet,
    ParsedWorkbook,
    ResolveCellValue,
    ResolveDefinedNameScalarValue,
    ResolveRangeEntries,
    TableCandidate,
} from "./model/index.js";
import {
    colToLetters,
    formatRange,
    lettersToCol,
    normalizeFormulaAddress,
    parseCellAddress,
    parseRangeAddress,
    parseRangeRef,
} from "./structure/addresses/address-utils.js";
import {
    decodeXmlText,
    getDirectChildByLocalName,
    getElementsByLocalName,
    getFirstDescendantByLocalName,
    getTextContent,
    xmlToDocument,
} from "./extraction/xml/xml-utils.js";
import { createRelsParser } from "./extraction/workbook/rels-parser.js";
import { unzipEntries } from "./extraction/zip/zip-io.js";
import { parseSharedStrings } from "./extraction/workbook/shared-strings.js";
import type { ParsedSharedString } from "./extraction/workbook/shared-strings.js";
import { parseCellStyles, readWorkbookLocale } from "./extraction/workbook/styles-parser.js";
import { parseWorksheet } from "./extraction/workbook/worksheet-parser.js";
import { parseWorkbook as parseWorkbookPackage } from "./extraction/workbook/workbook-loader.js";
import { normalizeStructuredTableKey, parseWorksheetTables } from "./structure/tables/worksheet-tables.js";
import { applyResolvedFormulaValue, formatCellDisplayValue } from "./structure/formatting/cell-format.js";
import { parseDrawingImages } from "./extraction/assets/sheet-images.js";
import { parseDrawingCharts } from "./extraction/assets/sheet-charts.js";
import {
    extractShapeBlocks,
    parseDrawingShapes,
    renderHierarchicalRawEntries,
} from "./extraction/assets/sheet-shapes.js";
import { renderShapeSvg } from "./extraction/assets/office-drawing.js";
import { normalizeMarkdownText } from "./outputs/markdown/markdown-normalize.js";
import { renderNarrativeBlock } from "./outputs/markdown/narrative-structure.js";
import { detectTableCandidates as detectTableCandidatesFromMap } from "./structure/tables/table-detector.js";
import { matrixFromCandidate } from "./structure/tables/table-layout.js";
import { normalizeTableDetectionMode } from "./outputs/markdown/markdown-options.js";
import { createOutputFileName, renderMarkdownTable } from "./outputs/markdown/markdown-export.js";
import { createSheetMarkdown } from "./outputs/markdown/sheet-markdown.js";
import { createFormulaEngine } from "./formula/evaluation/engine.js";
import { createFormulaAstBridge } from "./formula/evaluation/ast-bridge.js";
import type { FormulaEvaluationOptions } from "./formula/evaluation/evaluator.js";
import { createFormulaReferenceUtils } from "./formula/references/reference-utils.js";
import { resolveSimpleFormulaReferences } from "./formula/references/resolver.js";
import {
    findTopLevelOperatorIndex,
    parseWholeFunctionCall,
    splitFormulaArguments,
} from "./formula/references/formula-text.js";

// Workbook-model view for defined-name and structured-table resolvers.
type ResolveRangeRef = ((sheetName: string, text: string) => ParsedRangeRef | null) | null;

const EMPTY_BORDERS: BorderFlags = {
    top: false,
    bottom: false,
    left: false,
    right: false,
};

// Holds workbook resolvers during formula resolution. Installed per workbook and cleared after.
let resolveDefinedNameScalarValue: ResolveDefinedNameScalarValue = null;
let resolveDefinedNameRangeRef: ResolveRangeRef = null;
let resolveStructuredRangeRef: ResolveRangeRef = null;

// Excel measures shapes in EMUs. Fall back to one cell size when missing.
const DEFAULT_CELL_WIDTH_EMU = 609600;
const DEFAULT_CELL_HEIGHT_EMU = 190500;
const SHAPE_BLOCK_GAP_X_EMU = DEFAULT_CELL_WIDTH_EMU * 4;
const SHAPE_BLOCK_GAP_Y_EMU = DEFAULT_CELL_HEIGHT_EMU * 6;

const { parseRelationshipEntries, parseRelationships, buildRelsPath } = createRelsParser({
    xmlToDocument,
    decodeXmlText,
});
const {
    parseSimpleFormulaReference,
    parseSheetScopedDefinedNameReference,
    parseQualifiedRangeReference,
    normalizeFormulaSheetName,
    normalizeDefinedNameKey,
} = createFormulaReferenceUtils({
    normalizeFormulaAddress,
    parseRangeAddress,
    getDefinedNameRangeRef: () => resolveDefinedNameRangeRef,
    getStructuredRangeRef: () => resolveStructuredRangeRef,
});
const formulaAst = createFormulaAstBridge({
    normalizeFormulaAddress,
    parseSheetScopedDefinedNameReference,
    parseRangeAddress,
    parseCellAddress,
});
const { buildCellMap, convertSheetToMarkdown, convertWorkbookToMarkdownFiles, extractNarrativeBlocks } =
    createSheetMarkdown({
        renderNarrativeBlock,
        // Adapts the factory signature to the raw detector. Uses default score weights.
        detectTableCandidates: (
            sheet: ParsedSheet,
            buildCellMapForSheet: (sheet: ParsedSheet) => Map<string, ParsedCell>,
            tableDetectionMode: "balanced" | "border" | "planner-aware" = "balanced"
        ) => detectTableCandidatesFromMap(sheet, buildCellMapForSheet, undefined, tableDetectionMode),
        matrixFromCandidate,
        renderMarkdownTable,
        createOutputFileName,
        extractShapeBlocks,
        renderHierarchicalRawEntries,
        parseCellAddress,
        formatRange,
        colToLetters,
        normalizeMarkdownText,
        defaultCellWidthEmu: DEFAULT_CELL_WIDTH_EMU,
        defaultCellHeightEmu: DEFAULT_CELL_HEIGHT_EMU,
        shapeBlockGapXEmu: SHAPE_BLOCK_GAP_X_EMU,
        shapeBlockGapYEmu: SHAPE_BLOCK_GAP_Y_EMU,
    });
const { tryResolveFormulaExpressionDetailed, resolveScalarFormulaValue } = createFormulaEngine({
    getDefinedNameScalarValue: () => resolveDefinedNameScalarValue,
    // Forwards to the bridge below. Runs only during resolution, after engine creation.
    tryResolveFormulaExpressionWithAst: (
        expression: string,
        currentSheetName: string,
        resolveCellValue: (sheetName: string, address: string) => string,
        resolveRangeEntries?: (
            sheetName: string,
            rangeText: string
        ) => { rawValues: string[]; numericValues: number[] },
        currentAddress?: string,
        options?: FormulaEvaluationOptions
    ) =>
        tryResolveFormulaExpressionWithAst(
            expression,
            currentSheetName,
            resolveCellValue,
            resolveRangeEntries,
            currentAddress,
            options
        ),
    parseSimpleFormulaReference,
    parseSheetScopedDefinedNameReference,
});

// The spill operator has no workbook lookup on this path. Pass null.
// The resolver keeps the cached value when spill resolution is missing.
function tryResolveFormulaExpressionWithAst(
    expression: string,
    currentSheetName: string,
    resolveCellValue: ResolveCellValue,
    resolveRangeEntries?: ResolveRangeEntries,
    currentAddress?: string,
    options?: FormulaEvaluationOptions
): string | null {
    return formulaAst.tryResolveFormulaExpressionWithAst(
        expression,
        currentSheetName,
        resolveCellValue,
        resolveDefinedNameScalarValue,
        resolveDefinedNameRangeRef,
        resolveStructuredRangeRef,
        null,
        resolveRangeEntries,
        currentAddress,
        options
    );
}

/** Installs workbook resolvers for one resolution pass. Nulls clear them after the pass. */
function setDefinedNameResolvers(
    scalar: ResolveDefinedNameScalarValue,
    range: ResolveRangeRef,
    structured: ResolveRangeRef
): void {
    resolveDefinedNameScalarValue = scalar;
    resolveDefinedNameRangeRef = range;
    resolveStructuredRangeRef = structured;
}

function buildDrawingAssetDeps() {
    return {
        parseRelationships,
        buildRelsPath,
        xmlToDocument,
        decodeXmlText,
        getElementsByLocalName,
        getFirstDescendantByLocalName,
        getDirectChildByLocalName,
        getTextContent,
        colToLetters,
        drawingHelper: { renderShapeSvg },
        defaultCellWidthEmu: DEFAULT_CELL_WIDTH_EMU,
        defaultCellHeightEmu: DEFAULT_CELL_HEIGHT_EMU,
        shapeBlockGapXEmu: SHAPE_BLOCK_GAP_X_EMU,
        shapeBlockGapYEmu: SHAPE_BLOCK_GAP_Y_EMU,
    };
}

function createWorksheetParseDeps() {
    return {
        EMPTY_BORDERS,
        xmlToDocument,
        decodeXmlText,
        getTextContent,
        parseCellAddress,
        parseRangeRef,
        parseWorksheetTables,
        parseDrawingImages,
        parseDrawingCharts,
        parseDrawingShapes,
        parseRelationshipEntries,
        buildRelsPath,
        formatCellDisplayValue,
        buildAssetDeps: buildDrawingAssetDeps,
        lettersToCol,
        colToLetters,
    };
}

function createFormulaResolverDeps() {
    return {
        normalizeStructuredTableKey,
        normalizeFormulaSheetName,
        normalizeDefinedNameKey,
        normalizeFormulaAddress,
        parseSimpleFormulaReference,
        resolveScalarFormulaValue,
        parseQualifiedRangeReference,
        findTopLevelOperatorIndex,
        parseWholeFunctionCall,
        splitFormulaArguments,
        parseCellAddress,
        colToLetters,
        parseRangeAddress,
        tryResolveFormulaExpressionDetailed,
        applyResolvedFormulaValue,
        setDefinedNameResolvers,
    };
}

// Detect tables with default weights. Accept any mode string and normalize it.
export function detectTableCandidates(
    sheet: ParsedSheet,
    tableDetectionMode: "balanced" | "border" | "planner-aware" | string = "balanced"
): TableCandidate[] {
    return detectTableCandidatesFromMap(
        sheet,
        buildCellMap,
        undefined,
        normalizeTableDetectionMode(tableDetectionMode)
    );
}

/**
 * Parses .xlsx bytes into the shared workbook model.
 * Resolves formulas before returning, so outputs read final values.
 */
export async function parseWorkbook(
    arrayBuffer: ArrayBuffer,
    workbookName = "workbook.xlsx",
    options: ParseWorkbookOptions = {}
): Promise<ParsedWorkbook> {
    if (!arrayBuffer || arrayBuffer.byteLength === 0) {
        throw new Error("Empty workbook input. Expected .xlsx bytes.");
    }
    const worksheetParseDeps = createWorksheetParseDeps();
    const formulaResolverDeps = createFormulaResolverDeps();
    const parseShapes = options.includeShapeDetails !== false;
    const includeHidden = options.includeHidden !== false;
    return parseWorkbookPackage(arrayBuffer, workbookName, {
        unzipEntries,
        parseSharedStrings,
        parseCellStyles,
        readWorkbookLocale,
        includeHidden,
        parseRelationships,
        xmlToDocument,
        decodeXmlText,
        getTextContent,
        parseWorksheet: (
            files: Map<string, Uint8Array>,
            name: string,
            sheetPath: string,
            sheetIndex: number,
            sharedStrings: ParsedSharedString[],
            cellStyles: CellStyleInfo[],
            formatContext: { locale: string | null; date1904: boolean }
        ) =>
            parseWorksheet(
                files,
                name,
                sheetPath,
                sheetIndex,
                sharedStrings,
                cellStyles,
                { ...worksheetParseDeps, parseShapes, includeHidden },
                formatContext
            ),
        postProcessWorkbook: (workbook: ParsedWorkbook) => {
            const formatContext = { locale: workbook.locale ?? null, date1904: workbook.dateSystem === "1904" };
            resolveSimpleFormulaReferences(workbook, {
                ...formulaResolverDeps,
                applyResolvedFormulaValue: (cell, value, source) =>
                    applyResolvedFormulaValue(cell, value, source, formatContext),
            });
        },
    });
}

/** Markdown converters bound to the wired dependencies. */
export { convertSheetToMarkdown, convertWorkbookToMarkdownFiles, extractNarrativeBlocks };
