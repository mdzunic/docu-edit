// SPEC-119 "The AST", "Precedence", "Literals": parseFormula builds the
// FormulaNode union — Excel's precedence asserted node for node (as fully
// parenthesised prints), every literal form, arrays and their equal-length
// rows, missing arguments, names, structured and external nodes, token
// spans, FormulaSyntaxError offsets and the MAX_PARSE_DEPTH bound — and
// parseFormulaAt shares one frozen AST among every shifted copy of a
// formula through the LRU keyed by relativeFormulaKey.
import { afterEach, describe, expect, it } from "vitest";
import { CANONICAL_CORPUS } from "../fixtures/formulaCorpus";
import { ERROR_CODES } from "../model";
import type { FormulaNode } from "./ast";
import {
  clearFormulaAstCache,
  FORMULA_AST_CACHE_SIZE,
  formulaAstCacheSize,
  FormulaSyntaxError,
  MAX_PARSE_DEPTH,
  parseFormula,
  parseFormulaAt,
} from "./parse";
import { formatReference } from "./print";
import { tokenize } from "./tokens";

/** A node fully parenthesised: every operator application in (…), so the
 * tree's shape is the string. */
function shape(node: FormulaNode): string {
  switch (node.kind) {
    case "binary":
      return `(${shape(node.left)}${node.op}${shape(node.right)})`;
    case "unary":
      return node.op === "%" ? `(${shape(node.operand)}%)` : `(${node.op}${shape(node.operand)})`;
    case "paren":
      return `[${shape(node.inner)}]`;
    case "number":
      return String(node.value);
    case "ref":
      return formatReference(node.ref);
    case "name":
      return node.name;
    case "function":
      return `${node.name}(${node.args.map(shape).join(",")})`;
    case "missing":
      return "_";
    default:
      return node.kind;
  }
}

const parsedShape = (text: string): string => shape(parseFormula(text));

function syntaxError(text: string): FormulaSyntaxError {
  try {
    parseFormula(text);
  } catch (error) {
    if (error instanceof FormulaSyntaxError) {
      return error;
    }
    throw error;
  }
  throw new Error(`"${text}" parsed`);
}

/** `1` inside `depth` parentheses. */
function nested(depth: number): string {
  return `${"(".repeat(depth)}1${")".repeat(depth)}`;
}

afterEach(() => {
  clearFormulaAstCache();
});

describe("precedence (SPEC-119's table, asserted node for node)", () => {
  it.each([
    ["-2^2", "((-2)^2)"],
    ["2^-2", "(2^(-2))"],
    ["2^3^2", "((2^3)^2)"],
    ["2^3%", "(2^(3%))"],
    ["-A1:A3", "(-A1:A3)"],
    ["1-2-3", "((1-2)-3)"],
    ["1+2*3", "(1+(2*3))"],
    ["1*2+3", "((1*2)+3)"],
    ["8/4/2", "((8/4)/2)"],
    ["1+2&3", "((1+2)&3)"],
    ["1&2=3", "((1&2)=3)"],
    ["A1=B1=C1", "((A1=B1)=C1)"],
    ["1<2<>3>=4", "(((1<2)<>3)>=4)"],
    ["1+-+-1", "(1+(-(+(-1))))"],
    ["-50%", "((-50)%)"],
    ["50%%", "((50%)%)"],
    ["-2%^2", "(((-2)%)^2)"],
    ["A1:B2 B1:C3", "(A1:B2 B1:C3)"],
    ["(A1,B1 C1)", "[(A1,(B1 C1))]"],
    ["-(A1,B1)", "(-[(A1,B1)])"],
    ["A1:B2:C3", "(A1:B2:C3)"],
    ["INDEX(A:A,1):B2:C3", "(INDEX(A:A,1):B2:C3)"],
    ["INDEX(A:A,1):INDEX(B:B,2):C3", "((INDEX(A:A,1):INDEX(B:B,2)):C3)"],
    ["A1:INDEX(B:B,5)", "(A1:INDEX(B:B,5))"],
    ["Rate:Rate2", "(Rate:Rate2)"],
    ["A1 B1:INDEX(C:C,1)", "(A1 (B1:INDEX(C:C,1)))"],
    ["A1,B1,C1", "((A1,B1),C1)"],
    ["(1+2)*3", "([(1+2)]*3)"],
  ])("%s → %s", (text, expected) => {
    expect(parsedShape(text)).toBe(expected);
  });

  it("keeps a range operator's operands to references", () => {
    expect(syntaxError("A1:5").offset).toBe(3);
    expect(syntaxError(`"a":A1`).offset).toBe(0);
    expect(syntaxError("A1:-B1").offset).toBe(3);
    expect(parsedShape("Sheet1!#REF!:A1")).toBe("(error:A1)");
    expect(syntaxError("#N/A:A1").offset).toBe(0);
  });
});

describe("literals", () => {
  it("numbers: value = Number(text); 1E400 is a syntax error", () => {
    for (const [text, value] of [["1", 1], ["1.5", 1.5], [".5", 0.5], ["1.", 1], ["1E3", 1000], ["1e-3", 0.001], ["1.5E+10", 1.5e10]] as const) {
      expect(parseFormula(text), text).toEqual({ kind: "number", value, from: 0, to: 0 });
    }
    const error = syntaxError("1+1E400");
    expect(error).toBeInstanceOf(FormulaSyntaxError);
    expect(error.offset).toBe(2);
  });

  it(`strings: "" is one quote; an unterminated string is a syntax error`, () => {
    expect(parseFormula(`"He said ""hi"""`)).toMatchObject({ kind: "string", value: `He said "hi"` });
    expect(parseFormula(`""`)).toMatchObject({ kind: "string", value: "" });
    expect(syntaxError(`A1&"abc`).offset).toBe(3);
  });

  it("booleans in any case", () => {
    expect(parseFormula("true")).toMatchObject({ kind: "boolean", value: true });
    expect(parseFormula("False")).toMatchObject({ kind: "boolean", value: false });
  });

  it("the fifteen error literals, any case, their reference spellings kept in text", () => {
    for (const code of ERROR_CODES) {
      expect(parseFormula(code.toLowerCase()), code).toMatchObject({ kind: "error", code, text: code.toLowerCase() });
    }
    expect(parseFormula("Sheet2!#REF!")).toMatchObject({ kind: "error", code: "#REF!", text: "Sheet2!#REF!" });
    expect(parseFormula("'My Sheet'!#ref!")).toMatchObject({ kind: "error", code: "#REF!", text: "'My Sheet'!#ref!" });
    expect(parseFormula("#REF!A1")).toMatchObject({ kind: "error", code: "#REF!", text: "#REF!A1" });
  });

  it("array constants: rows of equal length; numbers (signed), strings, booleans and errors only", () => {
    expect(parseFormula("{1,2;3,4}")).toEqual({ kind: "array", rows: [[1, 2], [3, 4]], from: 0, to: 8 });
    expect(parseFormula(`{"a",TRUE;#N/A,-1.5}`)).toMatchObject({ kind: "array", rows: [["a", true], [{ error: "#N/A" }, -1.5]] });
    expect(parseFormula("{+1,- 2,.5}")).toMatchObject({ rows: [[1, -2, 0.5]] });
    expect(parseFormula(`{""""}`)).toMatchObject({ rows: [[`"`]] });
    const ragged = syntaxError("{1,2;3}");
    expect(ragged.offset).toBe(6);
    expect(ragged.message).toMatch(/same number/);
    expect(syntaxError("{1;2,3}").offset).toBe(6);
    expect(syntaxError("{A1}").offset).toBe(1);
    expect(syntaxError("{1+1}").offset).toBe(2);
    expect(syntaxError("{1,,2}").offset).toBe(3);
    expect(syntaxError("{Sheet1!#REF!}").offset).toBe(1);
    expect(syntaxError("{1E400}").offset).toBe(1);
    expect(syntaxError("{1,2").offset).toBe(4);
  });
});

describe("calls, names, structured and external nodes", () => {
  it("keeps missing arguments: IF(A1,,1) has three, F() none, IFERROR(1/0,) two", () => {
    const ifNode = parseFormula("IF(A1,,1)");
    expect(ifNode.kind === "function" && ifNode.args.map((arg) => arg.kind)).toEqual(["ref", "missing", "number"]);
    expect(parseFormula("F()")).toEqual({ kind: "function", name: "F", args: [], from: 0, to: 2 });
    const iferror = parseFormula("IFERROR(1/0,)");
    expect(iferror.kind === "function" && iferror.args.map((arg) => arg.kind)).toEqual(["binary", "missing"]);
    const both = parseFormula("F(,)");
    expect(both.kind === "function" && both.args).toEqual([
      { kind: "missing", from: 2, to: 1 },
      { kind: "missing", from: 3, to: 2 },
    ]);
  });

  it("upper-cases function names", () => {
    expect(parseFormula("sum(1)")).toMatchObject({ kind: "function", name: "SUM" });
    expect(parseFormula("_xlfn.concat(1)")).toMatchObject({ name: "_XLFN.CONCAT" });
  });

  it("parses names, with an optional sheet scope", () => {
    expect(parseFormula("Rate")).toEqual({ kind: "name", name: "Rate", from: 0, to: 0 });
    expect(parseFormula("Sheet1!Rate")).toEqual({ kind: "name", name: "Rate", sheet: "Sheet1", from: 0, to: 0 });
    expect(parseFormula("'O''Brien'!Rate")).toMatchObject({ name: "Rate", sheet: "O'Brien" });
    expect(syntaxError("Jan:Mar!Rate").offset).toBe(0);
  });

  it("keeps structured and external references as source text — parsed, never interpreted", () => {
    expect(parseFormula("Table1[[#This Row],[Amount]]")).toEqual({
      kind: "structured",
      table: "Table1",
      text: "Table1[[#This Row],[Amount]]",
      from: 0,
      to: 0,
    });
    expect(parseFormula("[@Qty]")).toEqual({ kind: "structured", text: "[@Qty]", from: 0, to: 0 });
    expect(parseFormula("[1]Sheet1!$A$1")).toEqual({ kind: "external", text: "[1]Sheet1!$A$1", from: 0, to: 0 });
    expect(syntaxError("1+T[[x]").offset).toBe(2);
  });

  it("parses every canonical corpus formula", () => {
    for (const text of CANONICAL_CORPUS) {
      expect(() => parseFormula(text), text).not.toThrow();
    }
  });
});

describe("spans", () => {
  it("records each node's inclusive token span", () => {
    const text = "SUM( A1 , -B1%)*2";
    const tokens = tokenize(text);
    const node = parseFormula(text);
    expect(node).toMatchObject({ kind: "binary", op: "*", from: 0, to: 9 });
    const call = node.kind === "binary" ? node.left : node;
    expect(call).toMatchObject({ kind: "function", from: 0, to: 7 });
    const args = call.kind === "function" ? call.args : [];
    expect(args[0]).toMatchObject({ kind: "ref", from: 2, to: 2 });
    expect(args[1]).toMatchObject({ kind: "unary", op: "%", from: 4, to: 6 });
    expect(tokens.slice(4, 7).map((token) => token.text)).toEqual(["-", "B1", "%"]);
  });
});

describe("syntax errors", () => {
  it("carries the UTF-16 offset of the offending token", () => {
    expect(syntaxError("").offset).toBe(0);
    expect(syntaxError("1+*2").offset).toBe(2);
    expect(syntaxError("SUM(1,2").offset).toBe(7);
    expect(syntaxError("SUM(1;2)").offset).toBe(5);
    expect(syntaxError("(1+2").offset).toBe(4);
    expect(syntaxError("1+2)").offset).toBe(3);
    expect(syntaxError(" @A1").offset).toBe(1);
    expect(syntaxError("A1 2").offset).toBe(3);
    expect(syntaxError("Ü+🙂").offset).toBe(2);
    expect(syntaxError("=A1").offset).toBe(0);
    expect(syntaxError("LAMBDA(x,x+1)(2)").offset).toBe(13);
  });

  it(`bounds nesting at MAX_PARSE_DEPTH = 256 — deeper is a FormulaSyntaxError, never a stack overflow`, () => {
    expect(MAX_PARSE_DEPTH).toBe(256);
    expect(() => parseFormula(nested(256))).not.toThrow();
    expect(syntaxError(nested(257)).offset).toBe(256);
    expect(syntaxError(nested(100_000))).toBeInstanceOf(FormulaSyntaxError);
    expect(syntaxError(`${"F(".repeat(5000)}1${")".repeat(5000)}`)).toBeInstanceOf(FormulaSyntaxError);
    expect(syntaxError(`${"-".repeat(100_000)}1`)).toBeInstanceOf(FormulaSyntaxError);
  });
});

describe("parseFormulaAt — the AST cache", () => {
  it("shares one AST among the 100,000 cells of one shared formula (cache size 1)", () => {
    expect(FORMULA_AST_CACHE_SIZE).toBe(50_000);
    const first = parseFormulaAt("B2*$F$1", { row: 1, col: 2 });
    if (first instanceof FormulaSyntaxError) {
      throw first;
    }
    for (let row = 1; row <= 100_000; row += 1) {
      const result = parseFormulaAt(`B${row + 1}*$F$1`, { row, col: 2 });
      if (result instanceof FormulaSyntaxError) {
        throw result;
      }
      expect(result.node).toBe(first.node);
    }
    expect(formulaAstCacheSize()).toBe(1);
  });

  it("returns the origin the AST was parsed for — a relative part resolves as part + (host − origin)", () => {
    const first = parseFormulaAt("B2*$F$1", { row: 1, col: 2 });
    const later = parseFormulaAt("B9*$F$1", { row: 8, col: 2 });
    if (first instanceof FormulaSyntaxError || later instanceof FormulaSyntaxError) {
      throw new Error("unexpected syntax error");
    }
    expect(later.node).toBe(first.node);
    expect(later.origin).toEqual({ row: 1, col: 2 });
    const ref = later.node.kind === "binary" && later.node.left.kind === "ref" ? later.node.left.ref : null;
    expect(ref!.start.row! + (8 - later.origin.row)).toBe(8); // B9
  });

  it("never lets a cached node be mutated (frozen, deeply)", () => {
    const result = parseFormulaAt("SUM(A1:B2,{1,2})", { row: 0, col: 0 });
    if (result instanceof FormulaSyntaxError) {
      throw result;
    }
    expect(Object.isFrozen(result.node)).toBe(true);
    const call = result.node;
    expect(call.kind === "function" && Object.isFrozen(call.args)).toBe(true);
    expect(call.kind === "function" && call.args[0]!.kind === "ref" && Object.isFrozen(call.args[0]!.ref.start)).toBe(true);
    expect(() => {
      (result.node as { kind: string }).kind = "number";
    }).toThrow(TypeError);
  });

  it("parses different formulas, absolute references and whitespace variants as their own keys", () => {
    const a = parseFormulaAt("$B$2*2", { row: 1, col: 2 });
    const b = parseFormulaAt("$B$2*2", { row: 5, col: 2 });
    const c = parseFormulaAt("$B$3*2", { row: 2, col: 2 });
    expect(a).not.toBeInstanceOf(FormulaSyntaxError);
    expect((a as { node: FormulaNode }).node).toBe((b as { node: FormulaNode }).node);
    expect((a as { node: FormulaNode }).node).not.toBe((c as { node: FormulaNode }).node);
    expect(formulaAstCacheSize()).toBe(2);
  });

  it("returns a syntax error rather than throwing, and caches only parses", () => {
    const error = parseFormulaAt("SUM(1,", { row: 0, col: 0 });
    expect(error).toBeInstanceOf(FormulaSyntaxError);
    expect((error as FormulaSyntaxError).offset).toBe(6);
    expect(formulaAstCacheSize()).toBe(0);
  });

  it("never shares an AST between texts that only coincide once whitespace is dropped (`12` and `1 2`)", () => {
    expect(parseFormulaAt("12", { row: 0, col: 0 })).not.toBeInstanceOf(FormulaSyntaxError);
    expect(parseFormulaAt("1 2", { row: 0, col: 0 })).toBeInstanceOf(FormulaSyntaxError);
    expect(parseFormulaAt(`"a" "b"`, { row: 0, col: 0 })).toBeInstanceOf(FormulaSyntaxError);
  });

  it("evicts the least recently used AST past FORMULA_AST_CACHE_SIZE", () => {
    const keep = parseFormulaAt("1+0", { row: 0, col: 0 }) as { node: FormulaNode };
    const evicted = parseFormulaAt("2+0", { row: 0, col: 0 }) as { node: FormulaNode };
    for (let n = 1; n < FORMULA_AST_CACHE_SIZE - 1; n += 1) {
      parseFormulaAt(`9+${n}`, { row: 0, col: 0 });
      if (n === 10) {
        parseFormulaAt("1+0", { row: 0, col: 0 }); // touched: most recently used
      }
    }
    expect(formulaAstCacheSize()).toBe(FORMULA_AST_CACHE_SIZE);
    parseFormulaAt("3+0", { row: 0, col: 0 });
    expect(formulaAstCacheSize()).toBe(FORMULA_AST_CACHE_SIZE);
    expect((parseFormulaAt("1+0", { row: 0, col: 0 }) as { node: FormulaNode }).node).toBe(keep.node);
    expect((parseFormulaAt("2+0", { row: 0, col: 0 }) as { node: FormulaNode }).node).not.toBe(evicted.node);
  });
});
