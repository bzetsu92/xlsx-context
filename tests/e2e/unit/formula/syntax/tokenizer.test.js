import { describe, expect, it } from "vitest";
import { tokenizeFormula } from "../../../../../src/formula/syntax/tokenizer.ts";

function bootFormulaParser() {
    delete globalThis.__xlsxContextFormula;
    return { tokenizeFormula };
}

describe("formula tokenizer", () => {
    it("tokenizes arithmetic and references", () => {
        const api = bootFormulaParser();
        const tokens = api.tokenizeFormula("=A1+B2*3");

        expect(tokens.map((token) => `${token.type}:${token.value}`)).toEqual([
            "cell:A1",
            "operator:+",
            "cell:B2",
            "operator:*",
            "number:3",
        ]);
    });

    it("tokenizes percent operator", () => {
        const api = bootFormulaParser();
        const tokens = api.tokenizeFormula("=10%+A1");

        expect(tokens.map((token) => `${token.type}:${token.value}`)).toEqual([
            "number:10",
            "operator:%",
            "operator:+",
            "cell:A1",
        ]);
    });

    it("tokenizes spill postfix operator", () => {
        const api = bootFormulaParser();
        const tokens = api.tokenizeFormula("=A1#");

        expect(tokens.map((token) => `${token.type}:${token.value}`)).toEqual(["cell:A1", "operator:#"]);
    });

    it("tokenizes array constants", () => {
        const api = bootFormulaParser();
        const tokens = api.tokenizeFormula("={1,2;3,4}");

        expect(tokens.map((token) => `${token.type}:${token.value}`)).toEqual([
            "lbrace:{",
            "number:1",
            "comma:,",
            "number:2",
            "semicolon:;",
            "number:3",
            "comma:,",
            "number:4",
            "rbrace:}",
        ]);
    });

    it("tokenizes space intersection", () => {
        const api = bootFormulaParser();
        const tokens = api.tokenizeFormula("=A1:C3 B2:D4");

        expect(tokens.map((token) => `${token.type}:${token.value}`)).toEqual([
            "cell:A1",
            "colon::",
            "cell:C3",
            "operator: ",
            "cell:B2",
            "colon::",
            "cell:D4",
        ]);
    });

    it("tokenizes absolute cell references", () => {
        const api = bootFormulaParser();
        const tokens = api.tokenizeFormula("=$A$1+A$2+$H3");

        expect(tokens.map((token) => `${token.type}:${token.value}`)).toEqual([
            "cell:$A$1",
            "operator:+",
            "cell:A$2",
            "operator:+",
            "cell:$H3",
        ]);
    });
});
