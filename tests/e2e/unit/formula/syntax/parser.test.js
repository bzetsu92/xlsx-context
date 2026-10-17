import { describe, expect, it } from "vitest";
import { parseFormula } from "../../../../../src/formula/syntax/parser.ts";

function bootFormulaParser() {
    delete globalThis.__xlsxContextFormula;
    return { parseFormula };
}

describe("formula parser", () => {
    it("parses operator precedence into AST", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("=A1+B2*3");

        expect(ast).toEqual({
            type: "binary_op",
            operator: "+",
            left: {
                type: "cell",
                ref: "A1",
                sheet: null,
            },
            right: {
                type: "binary_op",
                operator: "*",
                left: {
                    type: "cell",
                    ref: "B2",
                    sheet: null,
                },
                right: {
                    type: "number",
                    value: 3,
                    raw: "3",
                },
            },
        });
    });

    it("parses postfix percent into AST", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("=(A1+B1)%");

        expect(ast).toEqual({
            type: "postfix_op",
            operator: "%",
            operand: {
                type: "binary_op",
                operator: "+",
                left: {
                    type: "cell",
                    ref: "A1",
                    sheet: null,
                },
                right: {
                    type: "cell",
                    ref: "B1",
                    sheet: null,
                },
            },
        });
    });

    it("parses spill postfix into AST", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("=A1#");

        expect(ast).toEqual({
            type: "postfix_op",
            operator: "#",
            operand: {
                type: "cell",
                ref: "A1",
                sheet: null,
            },
        });
    });

    it("parses array constants into AST", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("={1,2;3,4}");

        expect(ast).toEqual({
            type: "array_constant",
            rows: [
                [
                    { type: "number", value: 1, raw: "1" },
                    { type: "number", value: 2, raw: "2" },
                ],
                [
                    { type: "number", value: 3, raw: "3" },
                    { type: "number", value: 4, raw: "4" },
                ],
            ],
        });
    });

    it("parses array constants with expressions into AST", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("={1+2,A1;DATE(2024,3,17),4}");

        expect(ast).toEqual({
            type: "array_constant",
            rows: [
                [
                    {
                        type: "binary_op",
                        operator: "+",
                        left: { type: "number", value: 1, raw: "1" },
                        right: { type: "number", value: 2, raw: "2" },
                    },
                    {
                        type: "cell",
                        ref: "A1",
                        sheet: null,
                    },
                ],
                [
                    {
                        type: "function_call",
                        name: "DATE",
                        args: [
                            { type: "number", value: 2024, raw: "2024" },
                            { type: "number", value: 3, raw: "3" },
                            { type: "number", value: 17, raw: "17" },
                        ],
                    },
                    { type: "number", value: 4, raw: "4" },
                ],
            ],
        });
    });

    it("parses space intersection into AST", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("=A1:C3 B2:D4");

        expect(ast).toEqual({
            type: "binary_op",
            operator: " ",
            left: {
                type: "range",
                start: { type: "cell", ref: "A1", sheet: null },
                end: { type: "cell", ref: "C3", sheet: null },
            },
            right: {
                type: "range",
                start: { type: "cell", ref: "B2", sheet: null },
                end: { type: "cell", ref: "D4", sheet: null },
            },
        });
    });

    it("parses sheet-qualified ranges", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("='日本語シート'!A1:B3");

        expect(ast).toEqual({
            type: "range",
            start: {
                type: "cell",
                ref: "A1",
                sheet: "日本語シート",
            },
            end: {
                type: "cell",
                ref: "B3",
                sheet: "日本語シート",
            },
        });
    });

    it("parses absolute references and quoted-sheet refs", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("='週単位のビュー'!$H$14+A$2");

        expect(ast).toEqual({
            type: "binary_op",
            operator: "+",
            left: {
                type: "cell",
                ref: "$H$14",
                sheet: "週単位のビュー",
            },
            right: {
                type: "cell",
                ref: "A$2",
                sheet: null,
            },
        });
    });

    it("parses sheet-scoped names into dedicated AST nodes", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("=Other!LocalCross");

        expect(ast).toEqual({
            type: "scoped_name",
            sheet: "Other",
            name: "LocalCross",
        });
    });

    it("parses error constants", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("=#N/A");

        expect(ast).toEqual({
            type: "error",
            value: "#N/A",
        });
    });

    it("parses function calls with nested comparison", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula('=IF(成功=FALSE,"NG","OK")');

        expect(ast).toEqual({
            type: "function_call",
            name: "IF",
            args: [
                {
                    type: "binary_op",
                    operator: "=",
                    left: {
                        type: "name",
                        name: "成功",
                    },
                    right: {
                        type: "boolean",
                        value: false,
                        raw: "FALSE",
                    },
                },
                {
                    type: "string",
                    value: "NG",
                },
                {
                    type: "string",
                    value: "OK",
                },
            ],
        });
    });

    it("parses structured references", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula('=COUNTIF(課題[期日],">0")');

        expect(ast).toEqual({
            type: "function_call",
            name: "COUNTIF",
            args: [
                {
                    type: "structured_ref",
                    table: "課題",
                    column: "期日",
                },
                {
                    type: "string",
                    value: ">0",
                },
            ],
        });
    });

    it("parses row-qualified structured references", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("=チェックリスト[[#This Row],[数量]]");

        expect(ast).toEqual({
            type: "structured_ref",
            table: "チェックリスト",
            qualifier: "#This Row",
            column: "数量",
        });
    });

    it("parses structured reference columns containing question marks", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula("=タスク[[#This Row],[完了?]]");

        expect(ast).toEqual({
            type: "structured_ref",
            table: "タスク",
            qualifier: "#This Row",
            column: "完了?",
        });
    });
});
