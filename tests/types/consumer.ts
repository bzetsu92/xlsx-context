import {
    buildWorkbookIR,
    convertWorkbookToMarkdown,
    createWhitespaceCounter,
    parseWorkbook,
    renderCompactMarkdown,
    type IrTable,
    type IrTableDetection,
    type MarkdownFile,
    type ParsedWorkbook,
    type TokenCounter,
    type WorkbookIR,
} from "xlsx-context";

export async function consume(bytes: Uint8Array): Promise<string> {
    const workbook: ParsedWorkbook = await parseWorkbook(bytes, { fileName: "book.xlsx", includeShapeDetails: false });
    const files: MarkdownFile[] = convertWorkbookToMarkdown(workbook, { outputMode: "display" });
    const ir: WorkbookIR = buildWorkbookIR(workbook, { tableDetectionMode: "border" });
    const firstTable: IrTable | undefined = ir.sheets[0]?.tables[0];
    const detection: IrTableDetection | undefined = firstTable?.detection;
    const counter: TokenCounter = createWhitespaceCounter();
    const text = renderCompactMarkdown(ir, { includeComments: false });
    return `${files.length}:${firstTable?.range ?? ""}:${detection?.score ?? 0}:${counter.count(text)}`;
}
