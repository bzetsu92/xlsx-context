import { describe, expect, it } from "vitest";
import { resolveShapeTargets } from "../../../../../src/structure/layout/shape-targets.ts";
import { columnLeftEmu, createUniformGrid, rowTopEmu } from "../../../../../src/extraction/workbook/sheet-grid.ts";
import { colToLetters } from "../../../../../src/structure/addresses/address-utils.ts";

const CELL_W = 190500;
const CELL_H = 190500;
const GRID = createUniformGrid(CELL_W, CELL_H);

function cell(address, text, extra = {}) {
    const match = /^([A-Z]+)(\d+)$/.exec(address);
    const col = [...match[1]].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
    return { address, row: Number(match[2]), col, outputValue: text, rawValue: text, valueType: "s", ...extra };
}

/** Box of a cell range, shrunk on all sides so it sits inside the cells. */
function boxAt(startCol, startRow, endCol = startCol, endRow = startRow, inset = 10000) {
    return {
        left: columnLeftEmu(GRID, startCol) + inset,
        top: rowTopEmu(GRID, startRow) + inset,
        right: columnLeftEmu(GRID, endCol + 1) - inset,
        bottom: rowTopEmu(GRID, endRow + 1) - inset,
    };
}

function center(col, row) {
    return { x: columnLeftEmu(GRID, col) + CELL_W / 2, y: rowTopEmu(GRID, row) + CELL_H / 2 };
}

function shape(drawingId, bbox, extra = {}) {
    return {
        sheetName: "Sheet1",
        anchor: "A1",
        drawingId,
        stableId: `xl/drawings/drawing1.xml#${drawingId}`,
        name: drawingId,
        kind: "Shape",
        text: "",
        connection: null,
        widthEmu: null,
        heightEmu: null,
        elementName: "xdr:sp",
        anchorElementName: "xdr:twoCellAnchor",
        rawEntries: [],
        bbox,
        svgFilename: null,
        svgPath: null,
        svgData: null,
        ...extra,
    };
}

function sheetOf(cells, shapes, merges = []) {
    return { name: "Sheet1", cells, shapes, merges, grid: GRID, images: [], charts: [], comments: [], tables: [] };
}

function targetsOf(sheet) {
    return resolveShapeTargets(sheet, colToLetters);
}

describe("shape targets", () => {
    it("points a callout tip at the text cell under it", () => {
        const sheet = sheetOf(
            [cell("C3", "注文番号")],
            [shape("1", boxAt(8, 1, 12, 2), { text: "note", calloutTip: center(3, 3) })]
        );
        expect(targetsOf(sheet)).toEqual([
            expect.objectContaining({
                role: "callout",
                cell: "C3",
                textAddress: "C3",
                text: "注文番号",
                method: "inside",
            }),
        ]);
    });

    it("finds text that overflows left from a right-aligned cell", () => {
        const right = { horizontal: "right", contained: false };
        const sheet = sheetOf(
            [cell("H5", "上記には消費税は含まれておりません。", { alignment: right })],
            [shape("1", boxAt(10, 1, 12, 2), { text: "n", calloutTip: center(4, 5) })]
        );
        expect(targetsOf(sheet)[0]).toMatchObject({ cell: "D5", textAddress: "H5", method: "overflow" });
    });

    it("keeps wrapped and numeric text inside its cell", () => {
        const wrapped = { horizontal: "general", contained: true };
        const sheet = sheetOf(
            [
                cell("B2", "ＸＸＸＸＸＸＸＸＸＸＸＸ", { alignment: wrapped }),
                cell("B6", "123456789012345", { valueType: "n" }),
            ],
            [
                shape("1", boxAt(10, 1, 12, 2), { text: "a", calloutTip: center(5, 2) }),
                shape("2", boxAt(10, 8, 12, 9), { text: "b", calloutTip: center(5, 6) }),
            ]
        );
        for (const target of targetsOf(sheet)) expect(target.method).not.toBe("overflow");
    });

    it("reports no text when a tip points at an empty area", () => {
        const sheet = sheetOf(
            [cell("A1", "far")],
            [shape("1", boxAt(30, 30, 34, 32), { text: "n", calloutTip: center(40, 40) })]
        );
        expect(targetsOf(sheet)[0]).toMatchObject({ cell: "AN40", textAddress: null, method: "none", nearby: [] });
    });

    it("lists other nearby texts after the match", () => {
        const sheet = sheetOf(
            [cell("C3", "TEL"), cell("C4", "FAX"), cell("C9", "far away")],
            [shape("1", boxAt(10, 1, 12, 2), { text: "n", calloutTip: center(3, 3) })]
        );
        const [target] = targetsOf(sheet);
        expect(target.textAddress).toBe("C3");
        expect(target.nearby.map((entry) => entry.address)).toEqual(["C4"]);
    });

    it("labels the text that starts right after a marker", () => {
        const sheet = sheetOf(
            [cell("D6", "〒"), cell("D7", "納入先住所１")],
            [shape("1", boxAt(3, 6), { text: "5" }), shape("2", boxAt(3, 7), { text: "６" })]
        );
        const targets = targetsOf(sheet);
        expect(targets.map((target) => [target.role, target.textAddress, target.method])).toEqual([
            ["marker", "D6", "right"],
            ["marker", "D7", "right"],
        ]);
    });

    it("prefers the text cell under a marker, also inside merged ranges", () => {
        const merge = { startRow: 4, startCol: 2, endRow: 4, endCol: 6, ref: "B4:F4" };
        const sheet = sheetOf(
            [cell("B4", "出荷番号"), cell("G4", "注文番号")],
            [shape("1", boxAt(5, 4), { text: "29" })],
            [merge]
        );
        expect(targetsOf(sheet)[0]).toMatchObject({
            textAddress: "B4",
            method: "inside",
            nearby: [{ address: "G4", text: "注文番号" }],
        });
    });

    it("falls back to the nearest text, such as a header above an empty cell", () => {
        const sheet = sheetOf([cell("E4", "金額"), cell("K5", "CHUBAN")], [shape("1", boxAt(5, 5), { text: "26" })]);
        expect(targetsOf(sheet)[0]).toMatchObject({ textAddress: "E4", method: "nearest" });
    });

    it("ignores text hidden under a merged range", () => {
        const merge = { startRow: 2, startCol: 2, endRow: 2, endCol: 4, ref: "B2:D2" };
        const sheet = sheetOf(
            [cell("B2", ""), cell("C2", "hidden")],
            [shape("1", boxAt(6, 1, 8, 1), { text: "n", calloutTip: center(3, 2) })],
            [merge]
        );
        expect(targetsOf(sheet)[0]).toMatchObject({ textAddress: null, method: "none" });
    });

    it("does not treat large shapes or long text as markers", () => {
        const sheet = sheetOf(
            [cell("D2", "item")],
            [shape("1", boxAt(1, 1, 3, 10), { text: "1" }), shape("2", boxAt(3, 2), { text: "123456" })]
        );
        expect(targetsOf(sheet)).toEqual([]);
    });

    it("resolves a loose line end and keeps the attached shape", () => {
        const box = shape("1", boxAt(10, 10, 14, 12), { text: "PRONES" });
        const line = shape("2", boxAt(5, 4, 10, 11, 0), {
            kind: "Straight Arrow Connector",
            elementName: "xdr:cxnSp",
            connection: {
                startId: "1",
                endId: null,
                arrow: "end",
                startPoint: { x: columnLeftEmu(GRID, 11), y: rowTopEmu(GRID, 11) },
                endPoint: center(5, 4),
            },
        });
        const sheet = sheetOf([cell("E4", "注文番号")], [box, line]);
        expect(targetsOf(sheet)).toEqual([
            expect.objectContaining({ role: "lineEnd", shapeIndex: 1, connectedShapeIndex: 0, textAddress: "E4" }),
        ]);
    });
});

describe("shape targets on large sheets", () => {
    it("handles a row with many text cells quickly and keeps neighbor limits", () => {
        const cells = [];
        for (let col = 1; col <= 30000; col += 1) {
            cells.push({ address: `R${col}`, row: 1, col, outputValue: `t${col}`, rawValue: "", valueType: "s" });
        }
        const tip = center(5000, 1);
        const sheet = sheetOf(cells, [shape("1", boxAt(1, 5, 3, 6), { text: "n", calloutTip: tip })]);
        const started = performance.now();
        const [target] = targetsOf(sheet);
        expect(performance.now() - started).toBeLessThan(2000);
        expect(target).toMatchObject({ method: "inside", text: "t5000" });
    });

    it("hides text inside a sheet-wide merge without walking its cells", () => {
        const merge = { startRow: 1, startCol: 1, endRow: 1048576, endCol: 16384, ref: "A1:XFD1048576" };
        const sheet = sheetOf(
            [cell("A1", "visible"), cell("C900000", "hidden")],
            [shape("1", boxAt(1, 1, 2, 2), { text: "n", calloutTip: center(3, 900000) })],
            [merge]
        );
        const started = performance.now();
        const [target] = targetsOf(sheet);
        expect(performance.now() - started).toBeLessThan(2000);
        expect(target).toMatchObject({ textAddress: "A1", text: "visible", method: "inside" });
        expect(target.nearby).toEqual([]);
    });
});

describe("shape targets in screen mockups", () => {
    const badge = { fill: "#FFFFFF", line: "#000000", dash: null, lineWidthEmu: 9525 };

    it("points a marker on a screenshot at that picture", () => {
        const sheet = sheetOf([], [shape("1", boxAt(5, 5), { kind: "Shape (roundRect)", text: "3", look: badge })]);
        sheet.images = [{ anchor: "B2", bbox: boxAt(2, 2, 20, 12, 0) }];
        expect(targetsOf(sheet)[0]).toMatchObject({ role: "marker", method: "image", targetImageIndex: 0 });
    });

    it("labels text drawn in a shape and names that shape", () => {
        const value = shape("2", boxAt(4, 5, 7, 5), { kind: "Rectangle", text: "100,000" });
        const sheet = sheetOf(
            [],
            [shape("1", boxAt(3, 5), { kind: "Shape (roundRect)", text: "311", look: badge }), value]
        );
        expect(targetsOf(sheet)[0]).toMatchObject({ method: "right", text: "100,000", targetShapeIndex: 1 });
    });

    it("leaves out groups of boxes that repeat one label", () => {
        const values = [1, 2, 3, 4, 5].map((row) =>
            shape(`v${row}`, boxAt(5, row * 2), { kind: "Shape (roundRect)", text: "1", look: badge })
        );
        expect(targetsOf(sheetOf([], values))).toEqual([]);
    });

    it("leaves out short labels without fill and line", () => {
        const plain = { fill: "none", line: "none", dash: null, lineWidthEmu: null };
        expect(
            targetsOf(sheetOf([], [shape("1", boxAt(5, 5), { kind: "Rectangle", text: "1", look: plain })]))
        ).toEqual([]);
    });

    it("prefers a picture that is closer than any text", () => {
        const sheet = sheetOf(
            [cell("H5", "far text")],
            [shape("1", boxAt(3, 5), { kind: "Shape (roundRect)", text: "13", look: badge })]
        );
        sheet.images = [{ anchor: "D5", bbox: boxAt(4, 4, 5, 6, 0) }];
        expect(targetsOf(sheet)[0]).toMatchObject({ method: "nearest", targetImageIndex: 0, text: null });
    });

    it("points a callout tip on a highlight box edge at that box", () => {
        const highlight = shape("2", boxAt(3, 5, 6, 6, 0), { kind: "Rectangle", text: "" });
        const tip = { x: columnLeftEmu(GRID, 3) - 20000, y: rowTopEmu(GRID, 5) + 5000 };
        const sheet = sheetOf([], [shape("1", boxAt(10, 1, 14, 2), { text: "note", calloutTip: tip }), highlight]);
        expect(targetsOf(sheet)[0]).toMatchObject({ role: "callout", targetShapeIndex: 1 });
    });
});
