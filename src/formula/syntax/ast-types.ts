/** Shared types for the formula tokenizer, parser and evaluator. */

export type FormulaTokenType =
    | "number"
    | "string"
    | "boolean"
    | "error"
    | "identifier"
    | "cell"
    | "quoted_identifier"
    | "lparen"
    | "rparen"
    | "lbrace"
    | "rbrace"
    | "comma"
    | "semicolon"
    | "colon"
    | "bang"
    | "lbracket"
    | "rbracket"
    | "operator";

export interface FormulaToken {
    type: FormulaTokenType;
    value: string;
    start: number;
    end: number;
}

export type FormulaAstNode =
    | { type: "number"; value: number; raw: string }
    | { type: "string"; value: string }
    | { type: "boolean"; value: boolean; raw: string }
    | { type: "error"; value: string }
    | { type: "array_constant"; rows: FormulaAstNode[][] }
    | { type: "name"; name: string }
    | { type: "scoped_name"; sheet: string; name: string }
    | { type: "cell"; ref: string; sheet: string | null }
    | { type: "range"; start: FormulaAstNode; end: FormulaAstNode }
    | { type: "structured_ref"; table: string; column: string; qualifier?: string | null }
    | { type: "function_call"; name: string; args: FormulaAstNode[] }
    | { type: "unary_op"; operator: string; operand: FormulaAstNode }
    | { type: "postfix_op"; operator: string; operand: FormulaAstNode }
    | { type: "binary_op"; operator: string; left: FormulaAstNode; right: FormulaAstNode };
