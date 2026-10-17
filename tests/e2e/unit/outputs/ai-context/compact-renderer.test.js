import { describe, expect, it } from "vitest";
import * as aiCompactRenderer from "../../../../../src/outputs/ai-context/compact-renderer.ts";

describe("record table format", () => {
    const cell = (address, display) => ({
        address,
        row: Number(address.replace(/^[A-Z]+/, "")),
        col: 0,
        raw: display,
        display,
        valueType: "s",
        merged: false,
        formula: null,
    });
    const table = {
        id: "t001",
        range: "B2-E4",
        headers: [],
        rows: [
            [cell("B2", "Name"), cell("C2", ""), cell("D2", "Value"), cell("E2", "Value")],
            [cell("B3", "a|b"), cell("C3", "x"), cell("D3", "1"), cell("E3", "")],
            [cell("B4", ""), cell("C4", ""), cell("D4", ""), cell("E4", "")],
        ],
        source: { sheet: "S", range: "B2-E4", anchor: "" },
    };

    it("writes self-describing records with unique keys and no empty values", () => {
        const { header, lines } = aiCompactRenderer.renderTableLines(table, "records");

        expect(header).toBeNull();
        expect(lines).toEqual(["[B3] Name: a\\|b | C: x | Value (D): 1", "[B4]"]);
    });

    it("renders records in full compact text", () => {
        const ir = {
            workbook: "w",
            sheets: [
                {
                    name: "S",
                    index: 1,
                    tables: [table],
                    narratives: [],
                    comments: [],
                    charts: [],
                    images: [],
                    shapes: [],
                },
            ],
        };
        const text = aiCompactRenderer.renderCompactMarkdown(ir, { tableFormat: "records" });

        expect(text).toBe("# S\n\n### B2-E4\n[B3] Name: a\\|b | C: x | Value (D): 1\n[B4]");
    });
});
