// SPEC-121 "ODS range addresses": odsRangeAddressToCanonical and
// canonicalToOdsRangeAddress in both directions — the documented examples,
// quoted names (spaces, '' escapes, names only one side must quote), `$`
// handling (a sheet's dropped, a part's kept), 3-D, unions, whole columns and
// rows, and every refusal. Canonical output always tokenizes back to ref
// tokens naming the same sheets, canonical → ods → canonical round-trips,
// and splitOdsRangeAddressList keeps every separator byte, so a rewrite can
// leave untouched entries byte-identical.
import { describe, expect, it } from "vitest";
import { tokenize } from "../formula/tokens";
import {
  canonicalToOdsRangeAddress,
  formatOdsRangeAddress,
  odsRangeAddressToCanonical,
  parseOdsRangeAddress,
  splitOdsRangeAddressList,
} from "./rangeAddress";

describe("odsRangeAddressToCanonical", () => {
  it.each([
    ["Sheet1.$A$1:.$B$3", "Sheet1!$A$1:$B$3"],
    ["$'Q1 Data'.A1:'Q1 Data'.B3", "'Q1 Data'!A1:B3"],
    ["Sheet1.A1:Sheet1.B3 Sheet1.D1:Sheet1.D3", "(Sheet1!A1:B3,Sheet1!D1:D3)"],
  ])("the documented examples: %s → %s", (ods, canonical) => {
    expect(odsRangeAddressToCanonical(ods)).toBe(canonical);
  });

  it.each([
    ["$Sheet1.A1", "Sheet1!A1"],
    ["Sheet1.A1", "Sheet1!A1"],
    ["$Sheet1.$A$1", "Sheet1!$A$1"],
    ["$Sheet1.$A1:.B$2", "Sheet1!$A1:B$2"],
    ["$Sheet1.A1:$Sheet1.B2", "Sheet1!A1:B2"],
    ["Sheet1.A1:SHEET1.B2", "Sheet1!A1:B2"],
    ["Sheet1.A1:.A1", "Sheet1!A1:A1"],
    ["Sheet1.a1:.b2", "Sheet1!A1:B2"],
    ["$Sheet1.XFD1048576", "Sheet1!XFD1048576"],
    [".A1", "A1"],
    [".$A$1:.B2", "$A$1:B2"],
  ])("cells and ranges — a sheet's $ dropped, a part's kept, the start's sheet repeated (any case) is one sheet: %s → %s", (ods, canonical) => {
    expect(odsRangeAddressToCanonical(ods)).toBe(canonical);
  });

  it.each([
    ["$'My Sheet'.A1", "'My Sheet'!A1"],
    ["$'O''Brien'.A1:.B2", "'O''Brien'!A1:B2"],
    ["'It''s ''here'''.A1", "'It''s ''here'''!A1"],
    ["$'Sheet1'.A1", "Sheet1!A1"],
    ["$'Q1.Data'.A1", "Q1.Data!A1"],
    ["$2026.A1", "'2026'!A1"],
    ["$A1.B2", "'A1'!B2"],
    ["$TRUE.A1", "'TRUE'!A1"],
    ["$R1C1.A1", "'R1C1'!A1"],
    ["$Sheet-1.A1", "'Sheet-1'!A1"],
    ["$Übersicht.A1", "Übersicht!A1"],
    ["$'#REF!'.A1", "'#REF!'!A1"],
  ])("quoted and bare names, re-quoted as canonical text quotes them: %s → %s", (ods, canonical) => {
    expect(odsRangeAddressToCanonical(ods)).toBe(canonical);
  });

  it.each([
    ["$Sheet1.A1:$Sheet3.B2", "Sheet1:Sheet3!A1:B2"],
    ["Sheet1.A1:Sheet3.B2", "Sheet1:Sheet3!A1:B2"],
    ["$Jan.B5:$Dec.B5", "Jan:Dec!B5"],
    ["$Jan.$B$5:$Dec.B5", "Jan:Dec!$B$5:B5"],
    ["$Jan.A1:$'Q4 2026'.B2", "'Jan:Q4 2026'!A1:B2"],
    ["$'O''Brien'.A1:$Zed.B2", "'O''Brien:Zed'!A1:B2"],
    ["$Jan.A:$Mar.C", "Jan:Mar!A:C"],
  ])("3-D: another end sheet spans the sheets, a single cell is written once, the prefix quoted whole: %s → %s", (ods, canonical) => {
    expect(odsRangeAddressToCanonical(ods)).toBe(canonical);
  });

  it.each([
    [".A:.C", "A:C"],
    ["Sheet1.$A:.$C", "Sheet1!$A:$C"],
    ["$Sheet1.A:.A", "Sheet1!A:A"],
    ["$Sheet1.1:.3", "Sheet1!1:3"],
    [".$2:.$4", "$2:$4"],
    ["$Sheet1.$1:$Sheet1.$1", "Sheet1!$1:$1"],
  ])("whole columns and rows, as ranges: %s → %s", (ods, canonical) => {
    expect(odsRangeAddressToCanonical(ods)).toBe(canonical);
  });

  it.each([
    ["$Sheet1.A1   $Sheet2.B2", "(Sheet1!A1,Sheet2!B2)"],
    ["$'My Sheet'.A1 $'My Sheet'.B2:.C3", "('My Sheet'!A1,'My Sheet'!B2:C3)"],
    ["$Sheet1.A1 .B2 $Sheet1.A:.A $Jan.C3:$Mar.C3", "(Sheet1!A1,B2,Sheet1!A:A,Jan:Mar!C3)"],
    [" $Sheet1.A1  ", "Sheet1!A1"],
  ])("a list is a union in parentheses — split only at spaces outside quotes, a leading or trailing run ignored: %s → %s", (ods, canonical) => {
    expect(odsRangeAddressToCanonical(ods)).toBe(canonical);
  });

  it.each([
    "",
    "   ",
    "A1",
    "Sheet1!A1",
    "Sheet1",
    "Sheet1.",
    ".",
    "Sheet1.A1:",
    ":Sheet1.A1",
    "Sheet1.A1:.B2:.C3",
    "Sheet1.A1:B2",
    ".A1:.B",
    ".A:.C1",
    ".A",
    ".1",
    "Sheet1.$A$",
    "Sheet1.A1B2",
    ".A1:$Sheet1.B2",
    "$.A1",
    "$$Sheet1.A1",
    "$#REF!.A1",
    "#REF!.A1",
    "$#ref!.$A$1:.$B$5",
    "$Sheet1.A1:$#REF!.B2",
    ".#REF!",
    "$Sheet1.#REF!",
    "Sheet1.XFE1",
    "Sheet1.ABCD1",
    "Sheet1.A0",
    "Sheet1.A1048577",
    "'Sheet1.A1",
    "$'My Sheet.A1",
    "''.A1",
    "'a:b'.A1",
    "'[1]Sheet1'.A1",
    "$Sheet1.A1:$'[1]x'.B2",
    "Sheet1.A1\tSheet2.B2",
    "$Sheet1.A1 $Sheet1.X1Y",
  ])("refuses %j", (ods) => {
    expect(odsRangeAddressToCanonical(ods)).toBeNull();
  });

  it("gives text tokens.ts reads back as ref tokens (and the union's separators and parentheses) naming the same sheets", () => {
    const lists = [
      "Sheet1.$A$1:.$B$3",
      "$'Q1 Data'.A1:'Q1 Data'.B3 $'O''Brien'.C3",
      "$Sheet1.A1:$Sheet3.B2 $Jan.B5:$Dec.B5 $'Jan'.A1:$'Q4 2026'.B2",
      ".A:.C .$2:.$4 $Sheet1.$A:.$C",
      "$2026.A1 $A1.B2 $TRUE.C3 $R1C1.D4 $Sheet-1.E5 $Übersicht.F6 $'Q1.Data'.G7 $'#REF!'.H8 $'Sheet1'.I9",
    ];
    for (const list of lists) {
      const canonical = odsRangeAddressToCanonical(list);
      expect(canonical, list).not.toBeNull();
      const tokens = tokenize(canonical!).filter((token) => token.kind !== "end");
      expect(tokens.every((token) => ["ref", "union", "open", "close"].includes(token.kind)), canonical!).toBe(true);
      const entries = splitOdsRangeAddressList(list).filter((piece, index) => index % 2 === 0 && piece !== "");
      expect(tokens.filter((token) => token.kind === "ref").map((token) => token.ref)).toEqual(entries.map(parseOdsRangeAddress));
    }
  });
});

describe("canonicalToOdsRangeAddress", () => {
  it.each([
    ["Sheet1!$A$1:$B$3", "$Sheet1.$A$1:.$B$3"],
    ["'Q1 Data'!A1:B3", "$'Q1 Data'.A1:.B3"],
    ["(Sheet1!A1:B3,Sheet1!D1:D3)", "$Sheet1.A1:.B3 $Sheet1.D1:.D3"],
    ["Sheet1!A1", "$Sheet1.A1"],
    ["A1:B2", ".A1:.B2"],
    ["$A$1", ".$A$1"],
    ["sheet1!a1:b2", "$sheet1.A1:.B2"],
  ])("the sheet on the start only, absolute; the parts' $ as canonical has them; a single cell has no end: %s → %s", (canonical, ods) => {
    expect(canonicalToOdsRangeAddress(canonical)).toBe(ods);
  });

  it.each([
    ["Sheet1:Sheet3!A1:B2", "$Sheet1.A1:$Sheet3.B2"],
    ["Jan:Dec!B5", "$Jan.B5:$Dec.B5"],
    ["'Q1 2026:Q4 2026'!C3", "$'Q1 2026'.C3:$'Q4 2026'.C3"],
    ["Jan:Mar!$A:$C", "$Jan.$A:$Mar.$C"],
  ])("3-D: both ends carry their sheet, a single cell as a range: %s → %s", (canonical, ods) => {
    expect(canonicalToOdsRangeAddress(canonical)).toBe(ods);
  });

  it.each([
    ["Sheet1!A:C", "$Sheet1.A:.C"],
    ["Sheet1!$A:$A", "$Sheet1.$A:.$A"],
    ["$1:$3", ".$1:.$3"],
    ["Sheet1!2:2", "$Sheet1.2:.2"],
  ])("whole columns and rows: %s → %s", (canonical, ods) => {
    expect(canonicalToOdsRangeAddress(canonical)).toBe(ods);
  });

  it.each([
    ["'O''Brien'!A1", "$'O''Brien'.A1"],
    ["'2026'!A1", "$'2026'.A1"],
    ["'A1'!B2", "$'A1'.B2"],
    ["'TRUE'!A1", "$'TRUE'.A1"],
    ["'Sheet-1'!A1", "$'Sheet-1'.A1"],
    ["'#REF!'!A1", "$'#REF!'.A1"],
    ["My.Sheet!A1", "$'My.Sheet'.A1"],
    ["Übersicht!A1", "$Übersicht.A1"],
  ])("quotes by the shared predicate — and a `.`, the ODS separator — with ' doubled: %s → %s", (canonical, ods) => {
    expect(canonicalToOdsRangeAddress(canonical)).toBe(ods);
  });

  it.each([
    ["Sheet1!A1,Sheet2!B2", "$Sheet1.A1 $Sheet2.B2"],
    ["(Sheet1!A1)", "$Sheet1.A1"],
    [" ( Sheet1!A1 ,\tB2:C3 ) ", "$Sheet1.A1 .B2:.C3"],
    ["(A1,'My Sheet'!B2,Jan:Mar!C3)", ".A1 $'My Sheet'.B2 $Jan.C3:$Mar.C3"],
  ])("a union, with or without one pair of parentheses, whitespace aside: %j → %j", (canonical, ods) => {
    expect(canonicalToOdsRangeAddress(canonical)).toBe(ods);
  });

  it.each([
    "",
    "  ",
    "1",
    "\"A1\"",
    "Rate",
    "Sheet1!Rate",
    "SUM(A1)",
    "A1+B1",
    "-A1",
    "A1%",
    "A1 B1",
    "A1:B2:C3",
    "Sheet1!A1:Sheet1!B2",
    "#REF!",
    "Sheet2!#REF!",
    "#REF!A1",
    "((A1,B1))",
    "(A1,B1",
    "A1,",
    ",A1",
    "A1,,B1",
    "()",
    "[1]Sheet1!A1",
    "Table1[Col]",
    "{1,2}",
    "''!A1",
  ])("refuses %j", (canonical) => {
    expect(canonicalToOdsRangeAddress(canonical)).toBeNull();
  });
});

describe("round trips", () => {
  it("canonical → ods → canonical is the identity on formatReference's text and unions of it", () => {
    const corpus = [
      "Sheet1!$A$1:$B$3",
      "'Q1 Data'!A1:B3",
      "(Sheet1!A1:B3,Sheet1!D1:D3)",
      "Sheet1!A1",
      "Sheet1!A1:A1",
      "A1",
      "A1:B2",
      "$A1:B$2",
      "XFD1048576",
      "Sheet1:Sheet3!A1:B2",
      "Jan:Dec!B5",
      "Jan:Dec!$B$5",
      "'Q1 2026:Q4 2026'!C3",
      "'O''Brien'!A1",
      "'It''s ''here'''!A1:C3",
      "'2026'!A1",
      "'A1'!B2",
      "'TRUE'!A1",
      "My.Sheet!A1",
      "Übersicht!A1",
      "'#REF!'!A1",
      "Sheet1!A:C",
      "Sheet1!$A:$A",
      "$1:$3",
      "Jan:Mar!A:C",
      "Jan:Mar!1:1",
      "(A1,B2:C3,'My Sheet'!D4,Jan:Mar!E5,Sheet1!F:F)",
    ];
    for (const canonical of corpus) {
      const ods = canonicalToOdsRangeAddress(canonical);
      expect(ods, canonical).not.toBeNull();
      expect(odsRangeAddressToCanonical(ods!), ods!).toBe(canonical);
    }
  });

  it("ods → canonical → ods is the identity on the form this module writes", () => {
    const corpus = [
      "$Sheet1.$A$1:.$B$3",
      "$'Q1 Data'.A1:.B3",
      "$Sheet1.A1:.B3 $Sheet1.D1:.D3",
      ".A1:.B2",
      ".A1",
      "$Jan.B5:$Dec.B5",
      "$'Q1 2026'.C3:$'Q4 2026'.D4",
      "$Sheet1.$A:.$C",
      "$Sheet1.$1:.$3",
      "$'O''Brien'.A1",
      "$'My.Sheet'.A1",
    ];
    for (const ods of corpus) {
      expect(canonicalToOdsRangeAddress(odsRangeAddressToCanonical(ods)!), ods).toBe(ods);
    }
  });

  it("changes exactly the known non-identities", () => {
    expect(canonicalToOdsRangeAddress(odsRangeAddressToCanonical("Sheet1.A1:Sheet1.B3")!)).toBe("$Sheet1.A1:.B3");
    expect(canonicalToOdsRangeAddress(odsRangeAddressToCanonical("Sheet1.a1   'Sheet2'.b2")!)).toBe("$Sheet1.A1 $Sheet2.B2");
    expect(odsRangeAddressToCanonical(canonicalToOdsRangeAddress("Jan:Dec!B5:B5")!)).toBe("Jan:Dec!B5");
    expect(odsRangeAddressToCanonical(canonicalToOdsRangeAddress("Sheet1!A1,Sheet2!B2")!)).toBe("(Sheet1!A1,Sheet2!B2)");
  });
});

describe("one entry and the list", () => {
  it("parseOdsRangeAddress reads one entry into the ReferenceNode tokens.ts would give its canonical text; formatOdsRangeAddress writes it back", () => {
    expect(parseOdsRangeAddress("$'Q1 Data'.$A1:'Q1 Data'.B$3")).toEqual({
      sheet: "Q1 Data",
      start: { row: 0, col: 0, rowAbs: false, colAbs: true },
      end: { row: 2, col: 1, rowAbs: true, colAbs: false },
    });
    expect(parseOdsRangeAddress(".$A:.C")).toEqual({ start: { col: 0, rowAbs: false, colAbs: true }, end: { col: 2, rowAbs: false, colAbs: false } });
    expect(parseOdsRangeAddress("$S.$2:.3")).toEqual({ sheet: "S", start: { row: 1, rowAbs: true, colAbs: false }, end: { row: 2, rowAbs: false, colAbs: false } });
    expect(parseOdsRangeAddress("$Jan.B5:$Dec.B5")).toEqual({ sheet: "Jan", sheetEnd: "Dec", start: { row: 4, col: 1, rowAbs: false, colAbs: false } });
    expect(parseOdsRangeAddress("$S.A1 $S.B2")).toBeNull();
    expect(formatOdsRangeAddress({ sheet: "Jan", sheetEnd: "Dec", start: { row: 4, col: 1, rowAbs: false, colAbs: false } })).toBe("$Jan.B5:$Dec.B5");
    expect(formatOdsRangeAddress({ start: { col: 0, rowAbs: false, colAbs: true } })).toBe(".$A:.$A");
  });

  it("splitOdsRangeAddressList: entries at even indices, the space runs as written at odd ones, quotes respected; joining gives the list back", () => {
    expect(splitOdsRangeAddressList("$S.A1   $'My Sheet'.B2 .C3")).toEqual(["$S.A1", "   ", "$'My Sheet'.B2", " ", ".C3"]);
    expect(splitOdsRangeAddressList(" $'It''s here'.A1 ")).toEqual(["", " ", "$'It''s here'.A1", " ", ""]);
    expect(splitOdsRangeAddressList("")).toEqual([""]);
    expect(splitOdsRangeAddressList("$'unterminated .A1 .B2")).toEqual(["$'unterminated .A1 .B2"]);
    for (const list of ["a b", "  a  b  ", "'x y'.A1 'z''w v'.B2", "\ta\tb"]) {
      expect(splitOdsRangeAddressList(list).join(""), list).toBe(list);
    }
  });
});

describe("SPEC-132: chart address forms", () => {
  it("reads LibreOffice's chart addresses — a series label cell, a values range, a quoted sheet, the plot area's union — and writes them back in canonicalToOdsRangeAddress's form", () => {
    expect(odsRangeAddressToCanonical("Data.B1")).toBe("Data!B1");
    expect(odsRangeAddressToCanonical("Data.B2:Data.B5")).toBe("Data!B2:B5");
    expect(odsRangeAddressToCanonical("$'Q1 Data'.$B$2:.$B$13")).toBe("'Q1 Data'!$B$2:$B$13");
    expect(odsRangeAddressToCanonical("Data.A1:Data.A5 Data.C1:Data.D5")).toBe("(Data!A1:A5,Data!C1:D5)");
    expect(canonicalToOdsRangeAddress("Data!B1")).toBe("$Data.B1");
    expect(canonicalToOdsRangeAddress("'Q1 Data'!$B$2:$B$13")).toBe("$'Q1 Data'.$B$2:.$B$13");
    expect(canonicalToOdsRangeAddress("(Data!A1:A5,Data!C1:D5)")).toBe("$Data.A1:.A5 $Data.C1:.D5");
    for (const address of ["Data.B2:Data.B5", "$'Q1 Data'.$B$2:.$B$13", "Data.A1:Data.A5 Data.C1:Data.D5"]) {
      const canonical = odsRangeAddressToCanonical(address)!;
      expect(odsRangeAddressToCanonical(canonicalToOdsRangeAddress(canonical)!), address).toBe(canonical);
    }
  });
});
