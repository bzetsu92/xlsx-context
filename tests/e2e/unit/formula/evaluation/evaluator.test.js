import { describe, expect, it } from "vitest";
import { evaluateFormulaAst } from "../../../../../src/formula/evaluation/evaluator.ts";
import { parseFormula } from "../../../../../src/formula/syntax/parser.ts";

function bootFormulaParser() {
    delete globalThis.__xlsxContextFormula;
    return { parseFormula, evaluateFormulaAst };
}

describe("formula evaluator", () => {
    it("evaluates arithmetic and IF via AST", () => {
        const api = bootFormulaParser();
        const ast = api.parseFormula('=IF(A1+B1>10,"BIG","SMALL")');
        const value = api.evaluateFormulaAst(ast, {
            resolveCell(ref) {
                if (ref === "A1") {
                    return 7;
                }
                if (ref === "B1") {
                    return 5;
                }
                return null;
            },
        });

        expect(value).toBe("BIG");
    });

    it("evaluates IFERROR and logical functions via AST", () => {
        const api = bootFormulaParser();
        const ifErrorValue = api.evaluateFormulaAst(api.parseFormula('=IFERROR(#N/A,"ALT")'));
        const andValue = api.evaluateFormulaAst(api.parseFormula("=AND(TRUE,1,A1>0)"), {
            resolveCell(ref) {
                return ref === "A1" ? 5 : null;
            },
        });
        const orValue = api.evaluateFormulaAst(api.parseFormula("=OR(FALSE,A1<0,A2=3)"), {
            resolveCell(ref) {
                if (ref === "A1") return 5;
                if (ref === "A2") return 3;
                return null;
            },
        });
        const notValue = api.evaluateFormulaAst(api.parseFormula("=NOT(FALSE)"));

        expect(ifErrorValue).toBe("ALT");
        expect(andValue).toBe(true);
        expect(orValue).toBe(true);
        expect(notValue).toBe(true);
    });

    it("evaluates postfix percent via AST", () => {
        const api = bootFormulaParser();
        const value = api.evaluateFormulaAst(api.parseFormula("=(A1+B1)%"), {
            resolveCell(ref) {
                if (ref === "A1") {
                    return 7;
                }
                if (ref === "B1") {
                    return 5;
                }
                return null;
            },
        });

        expect(value).toBe(0.12);
    });

    it("evaluates spill postfix via AST", () => {
        const api = bootFormulaParser();
        const value = api.evaluateFormulaAst(api.parseFormula("=A1#"), {
            resolveSpill(ref, sheet) {
                if (ref === "A1" && sheet === null) {
                    return [
                        [1, 2],
                        [3, 4],
                    ];
                }
                return null;
            },
        });

        expect(value).toEqual([
            [1, 2],
            [3, 4],
        ]);
    });

    it("evaluates DATE and VALUE via AST", () => {
        const api = bootFormulaParser();
        const dateValue = api.evaluateFormulaAst(api.parseFormula("=DATE(2024,3,17)"));
        const parsedValue = api.evaluateFormulaAst(api.parseFormula('=VALUE("1,234.5")'));
        const parsedDate = api.evaluateFormulaAst(api.parseFormula('=VALUE("2024/03/17")'));

        expect(dateValue).toBe(45368);
        expect(parsedValue).toBe(1234.5);
        expect(parsedDate).toBe(45368);
    });

    it("evaluates SUM and SUBSTITUTE via AST", () => {
        const api = bootFormulaParser();
        const sumValue = api.evaluateFormulaAst(api.parseFormula("=SUM(A1:A3,5)"), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[10], [20], [30]];
                }
                return [];
            },
        });
        const substituteAll = api.evaluateFormulaAst(api.parseFormula('=SUBSTITUTE("A-B-B","B","X")'));
        const substituteOne = api.evaluateFormulaAst(api.parseFormula('=SUBSTITUTE("A-B-B","B","X",2)'));

        expect(sumValue).toBe(65);
        expect(substituteAll).toBe("A-X-X");
        expect(substituteOne).toBe("A-B-X");
    });

    it("evaluates array constants via AST", () => {
        const api = bootFormulaParser();
        const arrayValue = api.evaluateFormulaAst(api.parseFormula("={1,2;3,4}"));
        const sumValue = api.evaluateFormulaAst(api.parseFormula("=SUM({1,2;3,4})"));

        expect(arrayValue).toEqual([
            [1, 2],
            [3, 4],
        ]);
        expect(sumValue).toBe(10);
    });

    it("evaluates array constants with expressions via AST", () => {
        const api = bootFormulaParser();
        const value = api.evaluateFormulaAst(api.parseFormula("={1+2,A1;DATE(2024,3,17),4}"), {
            resolveCell(ref) {
                if (ref === "A1") {
                    return 10;
                }
                return null;
            },
        });

        expect(value).toEqual([
            [3, 10],
            [45368, 4],
        ]);
    });

    it("evaluates space intersection via AST", () => {
        const api = bootFormulaParser();
        const value = api.evaluateFormulaAst(api.parseFormula("=A1:C3 B2:D4"), {
            resolveRange(startRef, endRef) {
                if (startRef === "B2" && endRef === "C3") {
                    return [
                        [22, 23],
                        [32, 33],
                    ];
                }
                return [];
            },
        });

        expect(value).toEqual([
            [22, 23],
            [32, 33],
        ]);
    });

    it("evaluates SUMPRODUCT via AST", () => {
        const api = bootFormulaParser();
        const rangeValue = api.evaluateFormulaAst(api.parseFormula("=SUMPRODUCT(A1:A3,B1:B3)"), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[1], [2], [3]];
                }
                if (startRef === "B1" && endRef === "B3") {
                    return [[10], [20], [30]];
                }
                return [];
            },
        });
        const structuredValue = api.evaluateFormulaAst(api.parseFormula("=SUMPRODUCT(課題[数量],課題[単価])"), {
            resolveStructuredRef(_table, column) {
                switch (column) {
                    case "数量":
                        return [[1], [2], [3]];
                    case "単価":
                        return [[100], [200], [300]];
                    default:
                        return [];
                }
            },
        });

        expect(rangeValue).toBe(140);
        expect(structuredValue).toBe(1400);
    });

    it("evaluates VLOOKUP, HLOOKUP, XLOOKUP via AST", () => {
        const api = bootFormulaParser();
        const vlookupValue = api.evaluateFormulaAst(api.parseFormula('=VLOOKUP("K2",A1:B3,2,FALSE)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "B3") {
                    return [
                        ["K1", 10],
                        ["K2", 20],
                        ["K3", 30],
                    ];
                }
                return [];
            },
        });
        const hlookupValue = api.evaluateFormulaAst(api.parseFormula('=HLOOKUP("K2",A1:C2,2,FALSE)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "C2") {
                    return [
                        ["K1", "K2", "K3"],
                        [10, 20, 30],
                    ];
                }
                return [];
            },
        });
        const xlookupValue = api.evaluateFormulaAst(api.parseFormula('=XLOOKUP("K2",A1:A3,B1:B3,"NF")'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [["K1"], ["K2"], ["K3"]];
                }
                if (startRef === "B1" && endRef === "B3") {
                    return [[10], [20], [30]];
                }
                return [];
            },
        });

        expect(vlookupValue).toBe(20);
        expect(hlookupValue).toBe(20);
        expect(xlookupValue).toBe(20);
    });

    it("evaluates XLOOKUP match_mode and search_mode via AST", () => {
        const api = bootFormulaParser();
        const nextSmaller = api.evaluateFormulaAst(api.parseFormula('=XLOOKUP(25,A1:A3,B1:B3,"NF",-1)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[10], [20], [30]];
                }
                if (startRef === "B1" && endRef === "B3") {
                    return [["A"], ["B"], ["C"]];
                }
                return [];
            },
        });
        const nextLarger = api.evaluateFormulaAst(api.parseFormula('=XLOOKUP(25,A1:A3,B1:B3,"NF",1)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[10], [20], [30]];
                }
                if (startRef === "B1" && endRef === "B3") {
                    return [["A"], ["B"], ["C"]];
                }
                return [];
            },
        });
        const reverseSearch = api.evaluateFormulaAst(api.parseFormula('=XLOOKUP("K2",A1:A4,B1:B4,"NF",0,-1)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A4") {
                    return [["K1"], ["K2"], ["K2"], ["K3"]];
                }
                if (startRef === "B1" && endRef === "B4") {
                    return [[10], [20], [200], [30]];
                }
                return [];
            },
        });

        expect(nextSmaller).toBe("B");
        expect(nextLarger).toBe("C");
        expect(reverseSearch).toBe(200);
    });

    it("evaluates XLOOKUP wildcard match_mode via AST", () => {
        const api = bootFormulaParser();
        const wildcardValue = api.evaluateFormulaAst(api.parseFormula('=XLOOKUP("K*",A1:A3,B1:B3,"NF",2)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [["K10"], ["AX"], ["K20"]];
                }
                if (startRef === "B1" && endRef === "B3") {
                    return [["A"], ["B"], ["C"]];
                }
                return [];
            },
        });
        const escapedWildcardValue = api.evaluateFormulaAst(api.parseFormula('=XLOOKUP("K~*",A1:A2,B1:B2,"NF",2)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A2") {
                    return [["K*"], ["K20"]];
                }
                if (startRef === "B1" && endRef === "B2") {
                    return [["literal"], ["other"]];
                }
                return [];
            },
        });

        expect(wildcardValue).toBe("A");
        expect(escapedWildcardValue).toBe("literal");
    });

    it("evaluates XLOOKUP binary search modes via AST", () => {
        const api = bootFormulaParser();
        const binaryExact = api.evaluateFormulaAst(api.parseFormula('=XLOOKUP(20,A1:A3,B1:B3,"NF",0,2)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[10], [20], [30]];
                }
                if (startRef === "B1" && endRef === "B3") {
                    return [["A"], ["B"], ["C"]];
                }
                return [];
            },
        });
        const binaryNextSmaller = api.evaluateFormulaAst(api.parseFormula('=XLOOKUP(25,A1:A3,B1:B3,"NF",-1,2)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[10], [20], [30]];
                }
                if (startRef === "B1" && endRef === "B3") {
                    return [["A"], ["B"], ["C"]];
                }
                return [];
            },
        });
        const binaryDescendingNextLarger = api.evaluateFormulaAst(
            api.parseFormula('=XLOOKUP(25,A1:A3,B1:B3,"NF",1,-2)'),
            {
                resolveRange(startRef, endRef) {
                    if (startRef === "A1" && endRef === "A3") {
                        return [[30], [20], [10]];
                    }
                    if (startRef === "B1" && endRef === "B3") {
                        return [["C"], ["B"], ["A"]];
                    }
                    return [];
                },
            }
        );

        expect(binaryExact).toBe("B");
        expect(binaryNextSmaller).toBe("B");
        expect(binaryDescendingNextLarger).toBe("C");
    });

    it("handles XLOOKUP binary search edge cases via AST", () => {
        const api = bootFormulaParser();
        const ascendingTooSmall = api.evaluateFormulaAst(api.parseFormula('=XLOOKUP(5,A1:A3,B1:B3,"NF",-1,2)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[10], [20], [30]];
                }
                if (startRef === "B1" && endRef === "B3") {
                    return [["A"], ["B"], ["C"]];
                }
                return [];
            },
        });
        const descendingTooLarge = api.evaluateFormulaAst(api.parseFormula('=XLOOKUP(35,A1:A3,B1:B3,"NF",1,-2)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[30], [20], [10]];
                }
                if (startRef === "B1" && endRef === "B3") {
                    return [["C"], ["B"], ["A"]];
                }
                return [];
            },
        });
        const descendingNextSmaller = api.evaluateFormulaAst(api.parseFormula('=XLOOKUP(25,A1:A3,B1:B3,"NF",-1,-2)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[30], [20], [10]];
                }
                if (startRef === "B1" && endRef === "B3") {
                    return [["C"], ["B"], ["A"]];
                }
                return [];
            },
        });

        expect(ascendingTooSmall).toBe("NF");
        expect(descendingTooLarge).toBe("NF");
        expect(descendingNextSmaller).toBe("B");
    });

    it("evaluates approximate VLOOKUP and HLOOKUP via AST", () => {
        const api = bootFormulaParser();
        const vlookupApproxValue = api.evaluateFormulaAst(api.parseFormula("=VLOOKUP(25,A1:B3,2,TRUE)"), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "B3") {
                    return [
                        [10, "A"],
                        [20, "B"],
                        [30, "C"],
                    ];
                }
                return [];
            },
        });
        const hlookupApproxValue = api.evaluateFormulaAst(api.parseFormula('=HLOOKUP("K25",A1:C2,2)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "C2") {
                    return [
                        ["K10", "K20", "K30"],
                        ["A", "B", "C"],
                    ];
                }
                return [];
            },
        });

        expect(vlookupApproxValue).toBe("B");
        expect(hlookupApproxValue).toBe("B");
    });

    it("handles approximate VLOOKUP and HLOOKUP edge cases via AST", () => {
        const api = bootFormulaParser();
        const vlookupTooSmall = api.evaluateFormulaAst(api.parseFormula("=VLOOKUP(5,A1:B3,2,TRUE)"), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "B3") {
                    return [
                        [10, "A"],
                        [20, "B"],
                        [30, "C"],
                    ];
                }
                return [];
            },
        });
        const hlookupExactDefault = api.evaluateFormulaAst(api.parseFormula('=HLOOKUP("K20",A1:C2,2)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "C2") {
                    return [
                        ["K10", "K20", "K30"],
                        ["A", "B", "C"],
                    ];
                }
                return [];
            },
        });

        expect(vlookupTooSmall).toBe("#N/A");
        expect(hlookupExactDefault).toBe("B");
    });

    it("evaluates EOMONTH via AST", () => {
        const api = bootFormulaParser();
        const eoMonthValue = api.evaluateFormulaAst(api.parseFormula("=EOMONTH(DATE(2024,3,17),1)"));

        expect(eoMonthValue).toBe(45412);
    });

    it("evaluates ROUND, ROUNDUP, ROUNDDOWN, INT via AST", () => {
        const api = bootFormulaParser();
        const roundValue = api.evaluateFormulaAst(api.parseFormula("=ROUND(2.675,2)"));
        const roundUpValue = api.evaluateFormulaAst(api.parseFormula("=ROUNDUP(2.611,1)"));
        const roundDownValue = api.evaluateFormulaAst(api.parseFormula("=ROUNDDOWN(2.699,1)"));
        const intValue = api.evaluateFormulaAst(api.parseFormula("=INT(2.9)"));

        expect(roundValue).toBe(2.68);
        expect(roundUpValue).toBe(2.7);
        expect(roundDownValue).toBe(2.6);
        expect(intValue).toBe(2);
    });

    it("evaluates ABS and ROW via AST", () => {
        const api = bootFormulaParser();
        const absValue = api.evaluateFormulaAst(api.parseFormula("=ABS(-12.5)"));
        const rowValue = api.evaluateFormulaAst(api.parseFormula("=ROW(A$12)"));

        expect(absValue).toBe(12.5);
        expect(rowValue).toBe(12);
    });

    it("evaluates ROW() and COLUMN() without explicit references via AST", () => {
        const api = bootFormulaParser();
        const rowValue = api.evaluateFormulaAst(api.parseFormula("=ROW()"), {
            currentCellRef: "C12",
        });
        const columnValue = api.evaluateFormulaAst(api.parseFormula("=COLUMN()"), {
            currentCellRef: "C12",
        });

        expect(rowValue).toBe(12);
        expect(columnValue).toBe(3);
    });

    it("evaluates ISBLANK, ISTEXT, ISERROR, ISNA via AST", () => {
        const api = bootFormulaParser();
        const isBlankValue = api.evaluateFormulaAst(api.parseFormula('=ISBLANK("")'));
        const isTextValue = api.evaluateFormulaAst(api.parseFormula('=ISTEXT("ABC")'));
        const isErrorValue = api.evaluateFormulaAst(api.parseFormula("=ISERROR(#VALUE!)"));
        const isNaValue = api.evaluateFormulaAst(api.parseFormula("=ISNA(#N/A)"));

        expect(isBlankValue).toBe(true);
        expect(isTextValue).toBe(true);
        expect(isErrorValue).toBe(true);
        expect(isNaValue).toBe(true);
    });

    it("evaluates MAX and AVERAGE via AST", () => {
        const api = bootFormulaParser();
        const maxValue = api.evaluateFormulaAst(api.parseFormula("=MAX(A1:A3,5)"), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[10], [2], [30]];
                }
                return [];
            },
        });
        const averageValue = api.evaluateFormulaAst(api.parseFormula("=AVERAGE(A1:A3,8)"), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[10], [20], [30]];
                }
                return [];
            },
        });

        expect(maxValue).toBe(30);
        expect(averageValue).toBe(17);
    });

    it("evaluates COUNT and COUNTA via AST", () => {
        const api = bootFormulaParser();
        const countValue = api.evaluateFormulaAst(api.parseFormula("=COUNT(A1:A5,10)"), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A5") {
                    return [[1], [""], ["20"], ["text"], [null]];
                }
                return [];
            },
        });
        const countAValue = api.evaluateFormulaAst(api.parseFormula("=COUNTA(A1:A5,10)"), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A5") {
                    return [[1], [""], ["20"], ["text"], [null]];
                }
                return [];
            },
        });

        expect(countValue).toBe(3);
        expect(countAValue).toBe(4);
    });

    it("evaluates LEFT, RIGHT, MID, TRIM, REPLACE via AST", () => {
        const api = bootFormulaParser();
        const leftValue = api.evaluateFormulaAst(api.parseFormula('=LEFT("ABCDE",2)'));
        const rightValue = api.evaluateFormulaAst(api.parseFormula('=RIGHT("ABCDE",3)'));
        const midValue = api.evaluateFormulaAst(api.parseFormula('=MID("ABCDE",2,2)'));
        const trimValue = api.evaluateFormulaAst(api.parseFormula('=TRIM("  A   B  C  ")'));
        const replaceValue = api.evaluateFormulaAst(api.parseFormula('=REPLACE("ABCDE",2,2,"ZZ")'));

        expect(leftValue).toBe("AB");
        expect(rightValue).toBe("CDE");
        expect(midValue).toBe("BC");
        expect(trimValue).toBe("A B C");
        expect(replaceValue).toBe("AZZDE");
    });

    it("evaluates YEAR, LOWER, FIND, SEARCH via AST", () => {
        const api = bootFormulaParser();
        const yearValue = api.evaluateFormulaAst(api.parseFormula("=YEAR(DATE(2024,3,17))"));
        const lowerValue = api.evaluateFormulaAst(api.parseFormula('=LOWER("AbC")'));
        const findValue = api.evaluateFormulaAst(api.parseFormula('=FIND("BC","ABCDE")'));
        const searchValue = api.evaluateFormulaAst(api.parseFormula('=SEARCH("bc","ABCDE")'));

        expect(yearValue).toBe(2024);
        expect(lowerValue).toBe("abc");
        expect(findValue).toBe(2);
        expect(searchValue).toBe(2);
    });

    it("evaluates TODAY, WEEKDAY, DATEVALUE, LEN, DAY, MONTH via AST", () => {
        const api = bootFormulaParser();
        const currentDate = new Date(Date.UTC(2024, 2, 17));
        const todayValue = api.evaluateFormulaAst(api.parseFormula("=TODAY()"), { currentDate });
        const weekdayValue = api.evaluateFormulaAst(api.parseFormula("=WEEKDAY(DATE(2024,3,17))"));
        const weekdayMondayBase = api.evaluateFormulaAst(api.parseFormula("=WEEKDAY(DATE(2024,3,17),2)"));
        const dateValue = api.evaluateFormulaAst(api.parseFormula('=DATEVALUE("2024/03/17")'));
        const lenValue = api.evaluateFormulaAst(api.parseFormula('=LEN("ABC123")'));
        const dayValue = api.evaluateFormulaAst(api.parseFormula("=DAY(DATE(2024,3,17))"));
        const monthValue = api.evaluateFormulaAst(api.parseFormula("=MONTH(DATE(2024,3,17))"));

        expect(todayValue).toBe(45368);
        expect(weekdayValue).toBe(1);
        expect(weekdayMondayBase).toBe(7);
        expect(dateValue).toBe(45368);
        expect(lenValue).toBe(6);
        expect(dayValue).toBe(17);
        expect(monthValue).toBe(3);
    });

    it("evaluates SUBTOTAL, UPPER, CONCATENATE, ISNUMBER, NA, MIN, COLUMN, EDATE via AST", () => {
        const api = bootFormulaParser();
        const subtotalValue = api.evaluateFormulaAst(api.parseFormula("=SUBTOTAL(9,A1:A3)"), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[10], [20], [30]];
                }
                return [];
            },
        });
        const upperValue = api.evaluateFormulaAst(api.parseFormula('=UPPER("abc")'));
        const concatValue = api.evaluateFormulaAst(api.parseFormula('=CONCATENATE("A","-","B")'));
        const isNumberTrue = api.evaluateFormulaAst(api.parseFormula("=ISNUMBER(1234)"));
        const isNumberFalse = api.evaluateFormulaAst(api.parseFormula('=ISNUMBER("")'));
        const naValue = api.evaluateFormulaAst(api.parseFormula("=NA()"));
        const minValue = api.evaluateFormulaAst(api.parseFormula("=MIN(A1:A3,5)"), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A3") {
                    return [[10], [2], [30]];
                }
                return [];
            },
        });
        const columnValue = api.evaluateFormulaAst(api.parseFormula("=COLUMN(A$2)"));
        const eDateValue = api.evaluateFormulaAst(api.parseFormula("=EDATE(DATE(2024,3,17),1)"));

        expect(subtotalValue).toBe(60);
        expect(upperValue).toBe("ABC");
        expect(concatValue).toBe("A-B");
        expect(isNumberTrue).toBe(true);
        expect(isNumberFalse).toBe(false);
        expect(naValue).toBe("#N/A");
        expect(minValue).toBe(2);
        expect(columnValue).toBe(1);
        expect(eDateValue).toBe(45399);
    });

    it("evaluates REPT and name/structured ref callbacks via AST", () => {
        const api = bootFormulaParser();
        const repeated = api.evaluateFormulaAst(api.parseFormula('=REPT("A",成功=TRUE)'), {
            resolveName(name) {
                return name === "成功" ? true : null;
            },
        });
        const structured = api.evaluateFormulaAst(api.parseFormula("=課題[期日]"), {
            resolveStructuredRef(table, column) {
                return `${table}:${column}`;
            },
        });

        expect(repeated).toBe("A");
        expect(structured).toBe("課題:期日");
    });

    it("evaluates sheet-scoped names via AST", () => {
        const api = bootFormulaParser();
        const value = api.evaluateFormulaAst(api.parseFormula("=Other!LocalCross"), {
            resolveScopedName(sheet, name) {
                return `${sheet}:${name}`;
            },
        });

        expect(value).toBe("Other:LocalCross");
    });

    it("evaluates MATCH and INDEX via AST", () => {
        const api = bootFormulaParser();
        const matchValue = api.evaluateFormulaAst(api.parseFormula('=MATCH("K2",A1:C1,0)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "C1") {
                    return ["K1", "K2", "K3"];
                }
                return [];
            },
        });
        const indexValue = api.evaluateFormulaAst(api.parseFormula("=INDEX(A2:C2,1,2)"), {
            resolveRange(startRef, endRef) {
                if (startRef === "A2" && endRef === "C2") {
                    return [["V1", "V2", "V3"]];
                }
                return [];
            },
        });

        expect(matchValue).toBe(2);
        expect(indexValue).toBe("V2");
    });

    it("evaluates TEXT via AST", () => {
        const api = bootFormulaParser();
        const padded = api.evaluateFormulaAst(api.parseFormula('=TEXT(10,"0000")'));
        const grouped = api.evaluateFormulaAst(api.parseFormula('=TEXT(1234.5,"#,##0.00")'));
        const dateText = api.evaluateFormulaAst(api.parseFormula('=TEXT(DATE(2024,3,17),"yyyy/mm/dd")'));

        expect(padded).toBe("0010");
        expect(grouped).toBe("1,234.50");
        expect(dateText).toBe("2024/03/17");
    });

    it("evaluates COUNTIF and SUMIF via AST", () => {
        const api = bootFormulaParser();
        const countIf = api.evaluateFormulaAst(api.parseFormula('=COUNTIF(A1:A4,">2")'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A4") {
                    return [[1], [2], [3], [4]];
                }
                return [];
            },
        });
        const sumIf = api.evaluateFormulaAst(api.parseFormula('=SUMIF(A1:A4,">2",B1:B4)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A4") {
                    return [[1], [2], [3], [4]];
                }
                if (startRef === "B1" && endRef === "B4") {
                    return [[10], [20], [30], [40]];
                }
                return [];
            },
        });

        expect(countIf).toBe(2);
        expect(sumIf).toBe(70);
    });

    it("evaluates SUMIFS with structured reference via AST", () => {
        const api = bootFormulaParser();
        const sumIfs = api.evaluateFormulaAst(
            api.parseFormula('=SUMIFS(課題[金額],課題[状態],"完了",課題[担当],"A")'),
            {
                resolveStructuredRef(_table, column) {
                    switch (column) {
                        case "金額":
                            return [[10], [20], [30]];
                        case "状態":
                            return [["完了"], ["未完了"], ["完了"]];
                        case "担当":
                            return [["A"], ["A"], ["B"]];
                        default:
                            return [];
                    }
                },
            }
        );

        expect(sumIfs).toBe(10);
    });

    it("evaluates COUNTIFS via AST", () => {
        const api = bootFormulaParser();
        const countIfs = api.evaluateFormulaAst(api.parseFormula('=COUNTIFS(A1:A4,">1",B1:B4,"A")'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A4") {
                    return [[1], [2], [3], [4]];
                }
                if (startRef === "B1" && endRef === "B4") {
                    return [["A"], ["A"], ["B"], ["A"]];
                }
                return [];
            },
        });

        expect(countIfs).toBe(2);
    });

    it("evaluates AVERAGEIF and AVERAGEIFS via AST", () => {
        const api = bootFormulaParser();
        const averageIf = api.evaluateFormulaAst(api.parseFormula('=AVERAGEIF(A1:A4,">2",B1:B4)'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A4") {
                    return [[1], [2], [3], [4]];
                }
                if (startRef === "B1" && endRef === "B4") {
                    return [[10], [20], [30], [50]];
                }
                return [];
            },
        });
        const averageIfs = api.evaluateFormulaAst(api.parseFormula('=AVERAGEIFS(B1:B4,A1:A4,">1",C1:C4,"X")'), {
            resolveRange(startRef, endRef) {
                if (startRef === "A1" && endRef === "A4") {
                    return [[1], [2], [3], [4]];
                }
                if (startRef === "B1" && endRef === "B4") {
                    return [[10], [20], [30], [50]];
                }
                if (startRef === "C1" && endRef === "C4") {
                    return [["X"], ["X"], ["Y"], ["X"]];
                }
                return [];
            },
        });

        expect(averageIf).toBe(40);
        expect(averageIfs).toBe((20 + 50) / 2);
    });

    it("evaluates CHOOSE lazily and returns #VALUE! for an out-of-range index", () => {
        const api = bootFormulaParser();
        const evaluate = (formula) =>
            api.evaluateFormulaAst(api.parseFormula(formula), {
                resolveCell(ref) {
                    if (ref === "A1") return 3;
                    throw new Error(`unexpected read of ${ref}`);
                },
            });

        expect(evaluate('=CHOOSE(2,"A","B",B9)')).toBe("B");
        expect(evaluate('=CHOOSE(A1,B9,"x","C")')).toBe("C");
        expect(evaluate('=CHOOSE(2.9,"A","B")')).toBe("B");
        expect(evaluate('=CHOOSE(0,"A")')).toBe("#VALUE!");
        expect(evaluate('=CHOOSE(3,"A","B")')).toBe("#VALUE!");
    });

    it("parses ISO, Japanese and date-time text in VALUE, DATEVALUE and date-part functions", () => {
        const api = bootFormulaParser();
        const evaluate = (formula) => api.evaluateFormulaAst(api.parseFormula(formula));

        expect(evaluate('=VALUE("2000-03-13")')).toBe(36598);
        expect(evaluate('=VALUE("3月14日")')).toBe(36599);
        expect(evaluate('=VALUE("2000年3月")')).toBe(36586);
        expect(evaluate('=VALUE("2000-03-13 12:00")')).toBe(36598.5);
        expect(evaluate('=VALUE("1,024")')).toBe(1024);
        expect(evaluate('=DATEVALUE("2000-03-13 12:00")')).toBe(36598);
        expect(evaluate('=DATEVALUE("2000.3.13")')).toBe(36598);
        expect(evaluate('=DAY("2024年3月5日")')).toBe(5);
        expect(evaluate('=MONTH("3月14日")')).toBe(3);
        expect(evaluate('=YEAR("2024-01")')).toBe(2024);
    });

    it("accepts future-function prefixes stored in files", () => {
        const api = bootFormulaParser();
        const evaluate = (formula) =>
            api.evaluateFormulaAst(api.parseFormula(formula), {
                resolveRange: () => [
                    ["K1", "K2"],
                    [10, 20],
                ],
            });

        expect(evaluate('=_xlfn.XLOOKUP("K2",{"K1","K2"},{10,20})')).toBe(20);
        expect(evaluate('=_xlfn._xlws.CHOOSE(1,"A","B")')).toBe("A");
        expect(evaluate('=_XLFN.CONCATENATE("a","b")')).toBe("ab");
        expect(() => evaluate("=_xlfn.UNKNOWNFN(1)")).toThrow(/Unsupported formula function: _xlfn.UNKNOWNFN/);
    });
});
