import { describe, expect, it } from "vitest";
import * as core from "../../../src/core.ts";
import * as aiCompactRenderer from "../../../src/outputs/ai-context/compact-renderer.ts";
import * as aiWorkbookIr from "../../../src/outputs/ai-context/workbook-ir.ts";
import { parseFixture } from "./helpers.js";

describe("renderCompactMarkdown()", () => {
    function bootAiApis() {
        return {
            core,
            ir: aiWorkbookIr,
            compact: aiCompactRenderer,
        };
    }

    it("renders deterministically", async () => {
        const apis = bootAiApis();
        const workbook = await parseFixture(apis.core, "table/table-basic-sample01.xlsx");
        const ir = apis.ir.buildWorkbookIR(workbook);

        expect(apis.compact.renderCompactMarkdown(ir)).toBe(apis.compact.renderCompactMarkdown(ir));
    });
    it("names the source table for charts and the joined shapes for connectors", async () => {
        const apis = bootAiApis();
        const chartIr = apis.ir.buildWorkbookIR(await parseFixture(apis.core, "chart/chart-mixed-sample01.xlsx"));
        const flowIr = apis.ir.buildWorkbookIR(await parseFixture(apis.core, "shape/shape-flowchart-sample01.xlsx"));

        const chartCompact = apis.compact.renderCompactMarkdown(chartIr);
        const flowCompact = apis.compact.renderCompactMarkdown(flowIr);

        expect(chartCompact).toContain(
            "- 利益率 (secondary axis): categories 'chart-mixed'!$B$4:$B$8; values 'chart-mixed'!$E$4:$E$8; table t001"
        );
        // The table already holds the numbers, so the chart line repeats none.
        expect(chartCompact).not.toContain("; data ");
        expect(flowCompact).toContain(
            "[I4] Shape sh005 (Straight Arrow Connector, 直線矢印コネクタ 9): sh001 -> sh002"
        );
        expect(flowCompact).toContain("[H3] Shape sh001 (Shape (flowChartTerminator), フローチャート: 端子 2): 開始");
    });
});
