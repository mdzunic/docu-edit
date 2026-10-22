// The formula AST (SPEC-119 "The AST") — what formula/parse.ts builds from
// a canonical formula's token stream and what SPEC-120 evaluates. THIS FILE
// is the normative source for `FormulaNode`.
//
// - Every node records its token span: `from` and `to` are inclusive indices
//   into the token array `tokenize(text)` returns for the parsed text, so a
//   rewrite can map a node back to the tokens (and their whitespace) it came
//   from. A `missing` argument covers no token: its span is empty
//   (`to === from - 1`, `from` being the separator or `)` that ends it).
// - Nodes are never mutated after parse — parseFormulaAt shares one AST
//   between every cell of a shared formula (formula/parse.ts).
// - References are ReferenceNodes (formula/references.ts): 0-based parts with
//   `$` flags, a relative part relative to the cell the text was parsed for
//   (`origin`).
// - Names, structured references and external references are parsed, never
//   interpreted here: SPEC-133 gives names and tables their meaning, and an
//   external reference shows as stored.
import type { CellError, ErrorCode } from "../model";
import type { ReferenceNode } from "./references";

/** Inclusive token indices (see the module header). */
export interface Span {
  from: number;
  to: number;
}

/** An array constant's element: a (signed) number, string, boolean or error. */
export type ArrayElement = number | string | boolean | CellError;

/** The binary operators, incl. range `:`, intersection `" "` and union `","`. */
export type BinaryOp = "+" | "-" | "*" | "/" | "^" | "&" | "=" | "<>" | "<" | "<=" | ">" | ">=" | ":" | " " | ",";

/** Prefix `+`/`-` and postfix `%`. */
export type UnaryOp = "+" | "-" | "%";

export type FormulaNode = Span &
  (
    | { kind: "number"; value: number }
    | { kind: "string"; value: string }
    | { kind: "boolean"; value: boolean }
    /** `text` is the error as written: "Sheet2!#REF!", "#REF!A1", "#n/a". */
    | { kind: "error"; code: ErrorCode; text: string }
    | { kind: "array"; rows: ArrayElement[][] }
    | { kind: "ref"; ref: ReferenceNode }
    | { kind: "name"; name: string; sheet?: string }
    /** `name` upper-case; `args` keeps empty slots as `missing` nodes. */
    | { kind: "function"; name: string; args: FormulaNode[] }
    | { kind: "missing" }
    | { kind: "unary"; op: UnaryOp; operand: FormulaNode }
    | { kind: "binary"; op: BinaryOp; left: FormulaNode; right: FormulaNode }
    | { kind: "paren"; inner: FormulaNode }
    /** `table` is the name before the first `[`; `text` the whole token. */
    | { kind: "structured"; table?: string; text: string }
    | { kind: "external"; text: string }
  );

/** The node of one kind (`NodeOf<"ref">`). */
export type NodeOf<K extends FormulaNode["kind"]> = Extract<FormulaNode, { kind: K }>;
