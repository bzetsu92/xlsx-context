import { describe, expect, it } from "vitest";
import { renderMermaidFlowchart } from "../../../../../src/outputs/markdown/mermaid-flowchart.ts";

function shape(kind, text, name = "") {
    return { kind, text, name };
}

describe("xlsx-context mermaid flowchart", () => {
    it("maps presets to node shapes and arrows to edge styles", () => {
        const shapes = [
            shape("Shape (flowChartTerminator)", "Start"),
            shape("Shape (flowChartDecision)", "OK?"),
            shape("Shape (flowChartInputOutput)", "Data"),
            shape("Shape (flowChartMagneticDisk)", "DB"),
            shape("Shape (flowChartPredefinedProcess)", "Sub"),
            shape("Shape (hexagon)", "Prep"),
            shape("Shape (ellipse)", "Dot"),
            shape("Shape (roundRect)", "Soft"),
            shape("Rectangle", "Box"),
            shape("Shape (cloudCallout)", "Other"),
        ];
        const diagram = {
            nodes: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9],
            edges: [
                { from: 0, to: 1, connector: 9, arrow: "end" },
                { from: 1, to: 2, connector: 9, arrow: "none" },
                { from: 2, to: 3, connector: 9, arrow: "both" },
            ],
            looseConnectors: [],
            direction: "TD",
        };

        const lines = renderMermaidFlowchart(shapes, diagram).split("\n");

        expect(lines.slice(0, 12)).toEqual([
            "```mermaid",
            "flowchart TD",
            '    sh001(["Start"])',
            '    sh002{"OK?"}',
            '    sh003[/"Data"/]',
            '    sh004[("DB")]',
            '    sh005[["Sub"]]',
            '    sh006{{"Prep"}}',
            '    sh007(("Dot"))',
            '    sh008("Soft")',
            '    sh009["Box"]',
            '    sh010["Other"]',
        ]);
        // The connector text becomes the edge label on every edge it draws.
        expect(lines.slice(12)).toEqual([
            '    sh001 -->|"Other"| sh002',
            '    sh002 ---|"Other"| sh003',
            '    sh003 <-->|"Other"| sh004',
            "```",
        ]);
    });

    it("escapes label text so it cannot break the node or the code fence", () => {
        const shapes = [
            shape("Rectangle", 'a "b" <c> #d;\n```\nend'),
            shape("Rectangle", "", "Named"),
            shape("Rectangle", ""),
        ];
        const diagram = {
            nodes: [0, 1, 2],
            edges: [{ from: 0, to: 1, connector: 2, arrow: "end" }],
            looseConnectors: [],
            direction: "LR",
        };

        const output = renderMermaidFlowchart(shapes, diagram);

        expect(output).toContain('    sh001["a #quot;b#quot; #lt;c#gt; #35;d;<br>```<br>end"]');
        expect(output).toContain('    sh002["Named"]');
        expect(output).toContain('    sh003["sh003"]');
        expect(output.split("\n").filter((line) => line.startsWith("```"))).toEqual(["```mermaid", "```"]);
    });
});

function styledShape(kind, text, look, connection = null) {
    return { kind, text, name: text, look, connection, bbox: { left: 0, top: 0, right: 1, bottom: 1 } };
}

const STYLED_DIAGRAM = {
    nodes: [0, 1],
    edges: [{ from: 0, to: 1, connector: 2, arrow: "end" }],
    looseConnectors: [],
    direction: "LR",
};

describe("xlsx-context mermaid flowchart looks", () => {
    it("draws dashed connectors as dotted edges and copies colored node looks", () => {
        const shapes = [
            styledShape("Rectangle", "A", { fill: "#FFFFFF", line: "#000000", dash: null, lineWidthEmu: null }),
            styledShape("Rectangle", "B", { fill: "#FFC000", line: "#FF0000", dash: "dash", lineWidthEmu: null }),
            styledShape(
                "Connector",
                "",
                { fill: null, line: "#000000", dash: "sysDash", lineWidthEmu: null },
                { arrow: "end" }
            ),
        ];
        const text = renderMermaidFlowchart(shapes, STYLED_DIAGRAM);
        expect(text).toContain("sh001 -.-> sh002");
        expect(text).toContain("style sh002 fill:#FFC000,stroke:#FF0000,stroke-dasharray:5 5");
        expect(text).not.toContain("style sh001");
    });

    it("keeps solid connectors as plain edges", () => {
        const shapes = [
            styledShape("Rectangle", "A"),
            styledShape("Rectangle", "B"),
            styledShape("Connector", "", undefined, { arrow: "end" }),
        ];
        expect(renderMermaidFlowchart(shapes, STYLED_DIAGRAM)).toContain("sh001 --> sh002");
    });
});
