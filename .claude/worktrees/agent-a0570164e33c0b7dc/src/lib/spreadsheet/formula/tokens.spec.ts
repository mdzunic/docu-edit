// SPEC-119 "Tokens", "Token rules": the tokenizer is total and lossless —
// joinTokens(tokenize(t)) === t for every corpus formula (fixture formulas
// included) and for 10,000 seeded random strings over the formula alphabet,
// none of which makes it throw — and it classifies per the token rules:
// one `ref` token per reference (cell, range, whole column/row, sheet-
// qualified, 3-D) carrying its ReferenceNode, functions whatever they spell,
// booleans, the fifteen errors and their reference forms, external and
// structured tokens, the intersection space, and the three meanings of a
// comma.
import { describe, expect, it } from "vitest";
import { CANONICAL_CORPUS, EXCEL_CORPUS } from "../fixtures/formulaCorpus";
import { ERROR_CODES } from "../model";
import { joinTokens, tokenize, type Token } from "./tokens";

/** "kind:text" for every token but the end. */
function kinds(text: string): string[] {
  return tokenize(text)
    .filter((token) => token.kind !== "end")
    .map((token) => `${token.kind}:${token.text}`);
}

function only(text: string): Token {
  const tokens = tokenize(text);
  expect(tokens.map((token) => token.kind), text).toHaveLength(2);
  return tokens[0]!;
}

/** mulberry32 — a small seeded PRNG, so the fuzz is the same every run. */
function seeded(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** The formula alphabet: letters (one non-ASCII, one astral), digits, every
 * operator and bracket, quotes, `$ ! # @ _ \ .`, and whitespace incl. CR/LF. */
const ALPHABET = [..."ABCDXFRCaex019", "Ü", "\u{1D400}", ..."$:!'\",;(){}[]#@._\\+-*/^&=<>%~|?", " ", " ", "\t", "\r", "\n"];

describe("the round trip", () => {
  it("joinTokens(tokenize(t)) === t for every corpus formula", () => {
    for (const text of [...CANONICAL_CORPUS, ...EXCEL_CORPUS]) {
      expect(joinTokens(tokenize(text)), text).toBe(text);
    }
  });

  it("is total: 10,000 seeded random strings never throw and always round-trip, offsets and trivia exact", () => {
    const random = seeded(119);
    for (let n = 0; n < 10_000; n += 1) {
      let text = "";
      const length = Math.floor(random() * 40);
      for (let i = 0; i < length; i += 1) {
        text += ALPHABET[Math.floor(random() * ALPHABET.length)];
      }
      const tokens = tokenize(text);
      expect(joinTokens(tokens), JSON.stringify(text)).toBe(text);
      expect(tokens.at(-1)!.kind).toBe("end");
      for (const token of tokens) {
        expect(text.slice(token.start, token.start + token.text.length)).toBe(token.text);
        expect(token.lead).toMatch(/^[ \t\r\n]*$/);
      }
    }
  });

  it("carries each token's preceding whitespace as its lead, the trailing run on `end`", () => {
    const tokens = tokenize("  SUM( A1 ,\tB1 )\r\n ");
    expect(tokens.map((token) => [token.kind, token.lead, token.text, token.start])).toEqual([
      ["function", "  ", "SUM", 2],
      ["open", "", "(", 5],
      ["ref", " ", "A1", 7],
      ["argSep", " ", ",", 10],
      ["ref", "\t", "B1", 12],
      ["close", " ", ")", 15],
      ["end", "\r\n ", "", 19],
    ]);
  });

  it("never throws on unterminated strings, quotes and brackets — the parser decides", () => {
    expect(only(`"abc`)).toMatchObject({ kind: "string", text: `"abc`, unterminated: true });
    expect(only(`Table1[[#Data],[A`)).toMatchObject({ kind: "structured", unterminated: true });
    expect(kinds(`'abc`)).toEqual(["unknown:'", "name:abc"]);
    expect(kinds(`'abc'+1`)).toEqual(["unknown:'abc'", "operator:+", "number:1"]);
    expect(kinds(`@A1;!`)).toEqual(["unknown:@", "ref:A1", "unknown:;", "unknown:!"]);
  });
});

describe("references are one ref token carrying their ReferenceNode", () => {
  it("cells and ranges, with $ markers, any case", () => {
    expect(only("B3").ref).toEqual({ start: { row: 2, col: 1, rowAbs: false, colAbs: false } });
    expect(only("$b$3").ref).toEqual({ start: { row: 2, col: 1, rowAbs: true, colAbs: true } });
    expect(only("A$1:$C3").ref).toEqual({
      start: { row: 0, col: 0, rowAbs: true, colAbs: false },
      end: { row: 2, col: 2, rowAbs: false, colAbs: true },
    });
    expect(only("XFD1048576").ref).toEqual({ start: { row: 1_048_575, col: 16_383, rowAbs: false, colAbs: false } });
  });

  it("whole columns and rows — only as ranges", () => {
    expect(only("A:A").ref).toEqual({ start: { col: 0, rowAbs: false, colAbs: false }, end: { col: 0, rowAbs: false, colAbs: false } });
    expect(only("$A:$C").ref).toEqual({ start: { col: 0, rowAbs: false, colAbs: true }, end: { col: 2, rowAbs: false, colAbs: true } });
    expect(only("1:1").ref).toEqual({ start: { row: 0, rowAbs: false, colAbs: false }, end: { row: 0, rowAbs: false, colAbs: false } });
    expect(only("$2:$4").ref).toEqual({ start: { row: 1, rowAbs: true, colAbs: false }, end: { row: 3, rowAbs: true, colAbs: false } });
    expect(kinds("A")).toEqual(["name:A"]);
    expect(kinds("1")).toEqual(["number:1"]);
  });

  it("sheet-qualified: bare, quoted with '' escapes, and a name or #REF! after the prefix", () => {
    expect(only("Sheet1!A1").ref).toEqual({ sheet: "Sheet1", start: { row: 0, col: 0, rowAbs: false, colAbs: false } });
    expect(only("'My Sheet'!A1:B2").ref).toMatchObject({ sheet: "My Sheet", end: { row: 1, col: 1 } });
    expect(only("'O''Brien'!A1").ref).toMatchObject({ sheet: "O'Brien" });
    expect(only("'Übersicht'!$A$1").ref).toMatchObject({ sheet: "Übersicht" });
    expect(only("Sheet1!Rate")).toMatchObject({ kind: "name", text: "Sheet1!Rate" });
  });

  it("3-D: Sheet1:Sheet3!A1 and 'Q1 2026:Q4 2026'!C3", () => {
    expect(only("Sheet1:Sheet3!A1").ref).toMatchObject({ sheet: "Sheet1", sheetEnd: "Sheet3" });
    expect(only("'Q1 2026:Q4 2026'!C3").ref).toMatchObject({ sheet: "Q1 2026", sheetEnd: "Q4 2026", start: { row: 2, col: 2 } });
    expect(only("Jan:Mar!B2:C3").ref).toMatchObject({ sheet: "Jan", sheetEnd: "Mar", end: { row: 2, col: 2 } });
  });

  it("joins greedily with no whitespace around `:` — a sheet-qualified pair is two refs and an operator", () => {
    expect(kinds("Sheet1!A1:Sheet1!B2")).toEqual(["ref:Sheet1!A1", "operator::", "ref:Sheet1!B2"]);
    expect(kinds("A1 : B2")).toEqual(["ref:A1", "operator::", "ref:B2"]);
    expect(kinds("A1:INDEX(B:B,5)")).toEqual(["ref:A1", "operator::", "function:INDEX", "open:(", "ref:B:B", "argSep:,", "number:5", "close:)"]);
  });

  it("names everything outside the grid or not cell-shaped: XFE1, A0, R1C1, ABCD1, A1B", () => {
    for (const text of ["XFE1", "A0", "R1C1", "ABCD1", "A1B", "Rate", "A1048577", "_xlpm.x", "\\path", "A1.5"]) {
      expect(kinds(text), text).toEqual([`name:${text}`]);
    }
  });
});

describe("functions, booleans, errors, externals, structured references", () => {
  it("an identifier immediately followed by ( is a function, even when it spells a cell", () => {
    expect(kinds("LOG10(1)+ATAN2(1,1)+DAYS360(A1,B1)")).toEqual([
      "function:LOG10", "open:(", "number:1", "close:)", "operator:+",
      "function:ATAN2", "open:(", "number:1", "argSep:,", "number:1", "close:)", "operator:+",
      "function:DAYS360", "open:(", "ref:A1", "argSep:,", "ref:B1", "close:)",
    ]);
    expect(kinds("_xlfn._xlws.SORT(A1)")[0]).toBe("function:_xlfn._xlws.SORT");
    expect(kinds("SUM (A1)")).toEqual(["name:SUM", "operator: ", "open:(", "ref:A1", "close:)"]);
  });

  it("TRUE/FALSE not followed by ( are booleans in any case", () => {
    expect(kinds("true+TRUE()+False")).toEqual(["boolean:true", "operator:+", "function:TRUE", "open:(", "close:)", "operator:+", "boolean:False"]);
  });

  it("all fifteen error codes, any case, and the reference forms Sheet1!#REF!, 'My Sheet'!#REF! and #REF!A1", () => {
    for (const code of ERROR_CODES) {
      expect(kinds(code), code).toEqual([`error:${code}`]);
      expect(kinds(code.toLowerCase()), code).toEqual([`error:${code.toLowerCase()}`]);
    }
    expect(kinds("Sheet1!#REF!+'My Sheet'!#REF!+#REF!A1+#REF!$B$2:C3+Jan:Mar!#REF!")).toEqual([
      "error:Sheet1!#REF!", "operator:+", "error:'My Sheet'!#REF!", "operator:+", "error:#REF!A1", "operator:+",
      "error:#REF!$B$2:C3", "operator:+", "error:Jan:Mar!#REF!",
    ]);
    expect(kinds("#WHAT")).toEqual(["unknown:#", "name:WHAT"]);
  });

  it("[1]Sheet1!A1, [1]!Name, '[1]My Sheet'!A1 and an OpenFormula external reference are one opaque external token", () => {
    expect(kinds("[1]Sheet1!A1+[1]!Name+'[1]My Sheet'!$A$1:B2+[2]Data:More!C3")).toEqual([
      "external:[1]Sheet1!A1", "operator:+", "external:[1]!Name", "operator:+", "external:'[1]My Sheet'!$A$1:B2", "operator:+", "external:[2]Data:More!C3",
    ]);
    expect(kinds("['file:///home/u/b.ods'#$Sheet1.A1]*2")).toEqual(["external:['file:///home/u/b.ods'#$Sheet1.A1]", "operator:*", "number:2"]);
    expect(only("[1]Sheet1!A1").ref).toBeUndefined();
  });

  it("Table1[[#This Row],[Amount]] and [@Qty] are one structured token: balanced brackets, ' escaping the next character", () => {
    expect(kinds("Table1[[#This Row],[Amount]]*[@Qty]")).toEqual(["structured:Table1[[#This Row],[Amount]]", "operator:*", "structured:[@Qty]"]);
    expect(kinds("Table1[[#Headers],[Col'#1]]")).toEqual(["structured:Table1[[#Headers],[Col'#1]]"]);
    expect(kinds("T[Col'[x]")).toEqual(["structured:T[Col'[x]"]);
    expect(only("T[Col']]")).toEqual({ kind: "structured", text: "T[Col']]", lead: "", start: 0 });
  });
});

describe("operators, separators and numbers", () => {
  it("every operator, two-character comparisons included", () => {
    expect(kinds("1+2-3*4/5^6&7=8<>9<10<=11>12>=13%")).toEqual([
      "number:1", "operator:+", "number:2", "operator:-", "number:3", "operator:*", "number:4", "operator:/", "number:5",
      "operator:^", "number:6", "operator:&", "number:7", "operator:=", "number:8", "operator:<>", "number:9", "operator:<",
      "number:10", "operator:<=", "number:11", "operator:>", "number:12", "operator:>=", "number:13", "operator:%",
    ]);
  });

  it("a space between a reference-ending and a reference-starting token is the intersection operator — its last space", () => {
    const tokens = tokenize("SUM(A1:C3  B2:D4)");
    expect(tokens.map((token) => [token.kind, token.lead, token.text])).toEqual([
      ["function", "", "SUM"],
      ["open", "", "("],
      ["ref", "", "A1:C3"],
      ["operator", " ", " "],
      ["ref", "", "B2:D4"],
      ["close", "", ")"],
      ["end", "", ""],
    ]);
    expect(kinds("(A1:B2) Rate")).toEqual(["open:(", "ref:A1:B2", "close:)", "operator: ", "name:Rate"]);
    expect(kinds("Sales[@Qty] [1]!X")).toEqual(["structured:Sales[@Qty]", "operator: ", "external:[1]!X"]);
    // Whitespace with no space, or next to a non-reference token, is trivia.
    expect(kinds("A1\nB1")).toEqual(["ref:A1", "ref:B1"]);
    expect(kinds("A1 + B1")).toEqual(["ref:A1", "operator:+", "ref:B1"]);
    expect(kinds("1 A1")).toEqual(["number:1", "ref:A1"]);
    const split = tokenize("A1 \t B1");
    expect(split.map((token) => [token.kind, token.lead])).toEqual([["ref", ""], ["operator", " \t"], ["ref", ""], ["end", ""]]);
  });

  it("a comma is argSep directly inside a call, arrayColSep inside {…}, union elsewhere; ; is arrayRowSep inside {…}", () => {
    expect(kinds("SUM(A1,(B1,C1),{1,2;3,4}),D1")).toEqual([
      "function:SUM", "open:(", "ref:A1", "argSep:,", "open:(", "ref:B1", "union:,", "ref:C1", "close:)", "argSep:,",
      "arrayOpen:{", "number:1", "arrayColSep:,", "number:2", "arrayRowSep:;", "number:3", "arrayColSep:,", "number:4", "arrayClose:}",
      "close:)", "union:,", "ref:D1",
    ]);
    expect(kinds("A1;B1")).toEqual(["ref:A1", "unknown:;", "ref:B1"]);
  });

  it("numbers: 1, 1.5, .5, 1., 1E3, 1e-3, 1.5E+10; a sign belongs to the number only inside {…}", () => {
    for (const text of ["1", "1.5", ".5", "1.", "1E3", "1e-3", "1.5E+10", "1E400"]) {
      expect(kinds(text), text).toEqual([`number:${text}`]);
    }
    expect(kinds("-1")).toEqual(["operator:-", "number:1"]);
    expect(kinds("{-1,+2.5,-.5}")).toEqual(["arrayOpen:{", "number:-1", "arrayColSep:,", "number:+2.5", "arrayColSep:,", "number:-.5", "arrayClose:}"]);
    expect(kinds("{- 1}")).toEqual(["arrayOpen:{", "operator:-", "number:1", "arrayClose:}"]);
  });
});
